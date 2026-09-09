import os
import sys
import json
import numpy as np
import calendar
import threading
from pathlib import Path
from threading import Lock
from datetime import date, datetime

# Resolve packaged, deployed src layout, and repo src layout.
_MODULE_ROOT = Path(__file__).resolve().parent
_SOURCE_ROOT = _MODULE_ROOT.parent
_PRODUCT_ROOT = _SOURCE_ROOT.parent
_BUNDLE_ROOT = Path(getattr(sys, "_MEIPASS", _MODULE_ROOT)).resolve()
_EXE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else None
_DEPLOY_ROOT = Path(os.environ.get("ARCRHO_DEPLOY_ROOT", r"E:\ArcRho Server"))

if "ARCRHO_ROOT" not in os.environ and "ADAS_ROOT" not in os.environ:
    if _EXE_DIR and _EXE_DIR.name.lower() == "apps":
        os.environ["ARCRHO_ROOT"] = str(_EXE_DIR.parent)
    elif _EXE_DIR and _EXE_DIR.parent.name.lower() == "apps":
        os.environ["ARCRHO_ROOT"] = str(_EXE_DIR.parent.parent)
    elif not getattr(sys, "frozen", False):
        os.environ["ARCRHO_ROOT"] = str(_DEPLOY_ROOT)

for _path in (_PRODUCT_ROOT, _SOURCE_ROOT, _BUNDLE_ROOT):
    if str(_path) not in sys.path:
        sys.path.insert(0, str(_path))

import pandas as pd
from utils import (
    function_brand,
    get_project_root,
    is_vector_function,
    resolve_app_path,
)
from arcrho_engine.general_utils import (
    DLOOKUP,
    _generate_period_range,
    _parse_date_to_yyyymm,
    get_current_time,
    split_formula,
    write_lists_to_csv,
)
from arcrho_engine.runtime_log import (
    ENGINE_REQUEST_LOG_FILENAME,
    append_runtime_log,
)
from arcrho_engine.data_processing_rules import (
    DataProcessingConfigurationError,
    DataProcessingRulesError,
    build_configuration_file_signature,
    build_reserving_class_catalog,
    build_weighted_source_frame,
    compile_data_processing_rules,
    resolve_request_path,
)

debug_mode = 1
device_name = os.environ.get("COMPUTERNAME")
ts = datetime.now().strftime("%y%m%d-%H%M%S-%f")[:-3]
robot_id =  f'{device_name}@' + os.getlogin() + "@" + ts
id_folder = str(resolve_app_path("engine", "instances"))
id_path = str(Path(id_folder) / f"{robot_id}.json")


PROJECT_CONFIG = {}   # Project configuration (Source Table, Dataset Types, Reserving Class Types)
DATA_DICT = {}  # CSV Data Table Files
DATA_DICT_LOCK = Lock() # for atomic swap
PROJECT_CONFIG_LOCK = Lock()
DATA_DICT_LOAD_ORDER = []  # Track load order for oldest removal

# Cache for project-specific settings to avoid repeated file reads
PROJECT_SETTINGS_CACHE = {}


class ProjectSettingsError(RuntimeError):
    """Raised when required project settings are missing or invalid."""


def remove_old_instances():
    folder = Path(id_folder)
    if not folder.exists():
        return
    for f in folder.iterdir():
        if f.is_file():
            is_instance_file = f.suffix.lower() in {".json", ".txt"}
            modified_date = datetime.fromtimestamp(f.stat().st_mtime).date()
            if is_instance_file and modified_date < date.today():
                f.unlink()


# Mirror of arcrho_api.field_mapping_contract. The Engine's calculation path
# runs before the canonical bundle is on sys.path, so the granularity rule and
# the two names it is recorded under are duplicated here and
# server-components/tests/test_engine_source_granularity.py fails if the mirror
# ever drifts from the canonical owner.
SOURCE_PERIOD_MONTHS_FIELD = "source_period_months"
DATE_ROLE_ORIGIN = "Origin Date"
ANNUAL_PERIOD_MONTHS = 12
MONTHLY_PERIOD_MONTHS = 1


def _period_months_from_date_value(value):
    """Months per period the date-role value *value* is written at.

    A four-digit value is a year; every other readable value is a YYYYMM
    month. An unreadable value returns 0, meaning "this column says nothing".
    """
    try:
        period = int(value)
    except (TypeError, ValueError):
        return 0
    if len(str(abs(period))) == 4:
        return ANNUAL_PERIOD_MONTHS
    return MONTHLY_PERIOD_MONTHS


def _recorded_origin_period_months(project_name):
    """Months per origin period the project's field mapping records, or 0."""
    try:
        payload = _read_json(_project_json_paths(project_name)["source_table"])
        recorded = payload.get(SOURCE_PERIOD_MONTHS_FIELD)
        return int(recorded[DATE_ROLE_ORIGIN])
    except (OSError, ValueError, TypeError, KeyError, AttributeError):
        return 0


def _granularity_name(months):
    return 'annual' if months == ANNUAL_PERIOD_MONTHS else 'monthly'


def _resolve_date_granularity(project_name, df, date_cols):
    """How fine this project's origin dates are, preferring its own record.

    The field mapping records the answer when the project's mapping was last
    saved, and that record is authoritative: it is what a generated dataset's
    stored shape is written from. Reading the column here is still worth doing
    as a cross-check, because a source table swapped for one at a different
    granularity would otherwise go unnoticed.
    """
    detected = 0
    if df is not None and date_cols is not None:
        try:
            detected = _period_months_from_date_value(df[date_cols[0]].dropna().iloc[0])
        except (IndexError, KeyError, AttributeError):
            detected = 0
    recorded = _recorded_origin_period_months(project_name)
    if not recorded:
        return _granularity_name(detected or MONTHLY_PERIOD_MONTHS)
    if detected and detected != recorded:
        append_runtime_log(
            get_project_root(),
            ENGINE_REQUEST_LOG_FILENAME,
            f"Project [{project_name}] records {recorded}-month origin periods but its "
            f"source table reads as {detected}-month; using the recorded value. "
            "Save the project's field mapping to record the table's own granularity.",
        )
    return _granularity_name(recorded)


def _load_project_settings(project_name, df=None, date_cols=None):
    """
    Load required project-specific settings from general_settings.json.
    Uses cache to avoid repeated file reads.

    Args:
        project_name: Name of the project
        df: Optional DataFrame used only to detect date granularity
        date_cols: Optional list of [origin_date_col, dev_date_col] names

    Returns:
        Dictionary with keys: origin_start, origin_end, dev_end (all in YYYYMM format)

    Raises:
        ProjectSettingsError: If general_settings.json is missing, invalid, or missing
            origin_start_date, origin_end_date, or development_end_date.
    """
    # Build path to project settings file
    settings_path = get_project_root() / "projects" / project_name / "general_settings.json"

    # Check cache: reuse if file hasn't been modified since last load
    if project_name in PROJECT_SETTINGS_CACHE:
        cached = PROJECT_SETTINGS_CACHE[project_name]
        if settings_path.exists():
            current_mtime = os.path.getmtime(settings_path)
            if cached.get('_mtime') == current_mtime:
                return cached
        else:
            PROJECT_SETTINGS_CACHE.pop(project_name, None)

    if not settings_path.exists():
        raise ProjectSettingsError(f"Project settings not defined for [{project_name}]")

    required_fields = {
        'origin_start_date': 'origin_start',
        'origin_end_date': 'origin_end',
        'development_end_date': 'dev_end',
    }

    try:
        with open(settings_path, mode="r", encoding="utf-8") as f:
            json_data = json.load(f)

        missing_fields = [
            field for field in required_fields
            if field not in json_data or str(json_data[field]).strip() == ''
        ]
        if missing_fields:
            raise ProjectSettingsError(
                f"Project settings not defined for [{project_name}]: missing {', '.join(missing_fields)}"
            )

        settings = {
            setting_key: _parse_date_to_yyyymm(json_data[field])
            for field, setting_key in required_fields.items()
        }

        print(f"Loaded settings from JSON for [{project_name}]: origin {settings['origin_start']}-{settings['origin_end']}, dev_end {settings['dev_end']}")
    except ProjectSettingsError:
        raise
    except Exception as e:
        raise ProjectSettingsError(f"Project settings not defined for [{project_name}]: {e}") from e

    settings['date_granularity'] = _resolve_date_granularity(project_name, df, date_cols)

    # Cache the settings along with file mtime for staleness detection
    if settings_path.exists():
        settings['_mtime'] = os.path.getmtime(settings_path)
    PROJECT_SETTINGS_CACHE[project_name] = settings

    return settings


def _enforce_data_dict_limit(max_tables=10):
    """
    Enforce the max table limit in DATA_DICT.
    Remove the oldest table if count >= max_tables before adding a new one.
    Should be called BEFORE adding a new table while holding DATA_DICT_LOCK.
    """
    # Count actual dataframes (exclude " - Version" entries)
    table_count = sum(1 for key in DATA_DICT.keys() if not key.endswith(" - Version"))

    if table_count >= max_tables:
        # Find oldest table from load order
        if DATA_DICT_LOAD_ORDER:
            oldest_table = DATA_DICT_LOAD_ORDER.pop(0)
            if oldest_table in DATA_DICT:
                del DATA_DICT[oldest_table]
            if oldest_table + " - Version" in DATA_DICT:
                del DATA_DICT[oldest_table + " - Version"]
            print(f"Removed oldest table from cache: {oldest_table}")


# Mirror of arcrho_api.source_table_contract. The engine ships as its own
# frozen bundle and cannot import that package, so these two names are
# duplicated here and frontend/tests/test_source_table_contract.py fails if the
# mirror ever drifts from the canonical owner.
SOURCE_IMPORT_DIR = "source"
MASTER_TABLE_FILE = "master_table.csv"


def _project_json_paths(project_name):
    project_dir = get_project_root() / "projects" / project_name
    return {
        "source_table": project_dir / "field_mapping.json",
        "dataset_types": project_dir / "dataset_types.json",
        "reserving_class_types": project_dir / "reserving_class_types.json",
        "data_processing_rules": project_dir / "data_processing_rules.json",
    }


def _project_dir(project_name):
    return get_project_root() / "projects" / str(project_name or "").strip()


def project_exists(project_name):
    return _project_dir(project_name).is_dir()


def _read_json(json_path):
    with open(json_path, mode="r", encoding="utf-8-sig") as f:
        return json.load(f)


def get_project_table_path(project_name):
    """Path of the project's imported master table.

    The engine never reads the external CSV path or the SQL Server table shown
    in Project Settings. Both import routes write the same fixed copy, so this
    resolves purely from the project folder.
    """
    master_path = _project_dir(project_name) / SOURCE_IMPORT_DIR / MASTER_TABLE_FILE
    if not master_path.exists():
        raise FileNotFoundError(
            f"No source table has been imported for project {project_name}: {master_path}. "
            "Import the table in Project Settings > Source Data."
        )
    return str(master_path)


def _data_table_cache_key(csv_path):
    return os.path.normcase(os.path.abspath(str(csv_path)))


def _source_columns_for_project_config(project_name):
    try:
        table_path = get_project_table_path(project_name)
    except FileNotFoundError as exc:
        raise DataProcessingConfigurationError(str(exc)) from exc

    try:
        return list(pd.read_csv(table_path, nrows=0).columns)
    except Exception as exc:
        raise DataProcessingConfigurationError(
            f"Cannot read source-table columns for project [{project_name}]: {exc}"
        ) from exc


def _json_table_to_df(json_obj):
    if isinstance(json_obj, dict) and "columns" in json_obj and "rows" in json_obj:
        return pd.DataFrame(json_obj.get("rows", []), columns=json_obj.get("columns", [])).fillna('')
    return pd.DataFrame(json_obj).fillna('')


def _source_table_df_from_json(json_obj):
    rows = json_obj.get("rows", []) if isinstance(json_obj, dict) else json_obj
    normalized_rows = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        normalized_rows.append({
            "Column Name": row.get("field_name", ""),
            "Significances": row.get("significance", ""),
            "Level": row.get("level", ""),
        })
    return pd.DataFrame(normalized_rows, columns=["Column Name", "Significances", "Level"]).fillna('')


def _get_project_config_signature(project_name):
    json_paths = _project_json_paths(project_name)
    return build_configuration_file_signature(
        json_paths,
        required_keys={"source_table", "dataset_types", "reserving_class_types"},
    )


def _get_vps_last_modified_time(project_name):
    """Compatibility name for the composite project-configuration signature."""
    return _get_project_config_signature(project_name)


def _build_project_config(project_name, json_paths):
    try:
        source_table_json = _read_json(json_paths["source_table"])
        dataset_types_json = _read_json(json_paths["dataset_types"])
        reserving_class_types_json = _read_json(json_paths["reserving_class_types"])
    except (json.JSONDecodeError, OSError) as exc:
        raise DataProcessingConfigurationError(
            f"Cannot read valid project configuration JSON for "
            f"project [{project_name}]: {exc}"
        ) from exc
    rules_json = None
    rules_path = json_paths["data_processing_rules"]
    if rules_path.exists():
        try:
            rules_json = _read_json(rules_path)
        except (json.JSONDecodeError, OSError) as exc:
            raise DataProcessingRulesError(
                f"Cannot read valid data_processing_rules.json for "
                f"project [{project_name}]: {exc}"
            ) from exc

    reserving_class_catalog = build_reserving_class_catalog(
        source_table_json,
        reserving_class_types_json,
    )
    compiled_rules = compile_data_processing_rules(
        rules_json,
        catalog=reserving_class_catalog,
        field_mapping_payload=source_table_json,
        dataset_types_payload=dataset_types_json,
        source_columns=_source_columns_for_project_config(project_name),
    )
    project_config = {
        "Source Table": _source_table_df_from_json(source_table_json),
        "Dataset Types": _json_table_to_df(dataset_types_json),
        "Reserving Class Types": _json_table_to_df(reserving_class_types_json),
        "Reserving Class Catalog": reserving_class_catalog,
        "Data Processing Rules": compiled_rules,
    }
    return project_config


def load_to_PROJECT_CONFIG(project_name, settings_file=None):
    json_paths = _project_json_paths(project_name)
    print(f"Loading JSON settings for [{project_name}] @ {get_current_time()}")

    for _attempt in range(3):
        signature_before = _get_project_config_signature(project_name)
        project_config = _build_project_config(project_name, json_paths)
        signature_after = _get_project_config_signature(project_name)
        if signature_before == signature_after:
            PROJECT_CONFIG[project_name] = project_config
            PROJECT_CONFIG[project_name + " - Version"] = signature_after
            return

    raise DataProcessingConfigurationError(
        f"Project configuration for [{project_name}] changed repeatedly while "
        "the data engine was loading it. Retry the request."
    )


def load_to_DATA_DICT(csv_path):
    print(f"Loading Data Table {csv_path} @ {get_current_time()}")
    key = _data_table_cache_key(csv_path)
    _enforce_data_dict_limit(max_tables=10)
    DATA_DICT[key] = pd.read_csv(csv_path, float_precision="round_trip")
    DATA_DICT[key + " - Version"] = datetime.now()
    if key not in DATA_DICT_LOAD_ORDER:
        DATA_DICT_LOAD_ORDER.append(key)
    print(f"Data Table Loaded @ {get_current_time()}")


def load_dataframe(data_csv_path):
    '''
    Add a new table to DATA_DICT
    '''
    print(get_current_time())
    print(f'Loading Data Table -- [{os.path.basename(data_csv_path)}]')
    df = pd.read_csv(data_csv_path, float_precision="round_trip") # build off-thread
    with DATA_DICT_LOCK:
        key = _data_table_cache_key(data_csv_path)
        _enforce_data_dict_limit(max_tables=10)
        DATA_DICT[key] = df
        if key not in DATA_DICT_LOAD_ORDER:
            DATA_DICT_LOAD_ORDER.append(key)

    print(get_current_time())
    print(f'Data Table Loaded -- [{os.path.basename(data_csv_path)}]')


def load_dataframe_in_thread(data_csv_path):
    t = threading.Thread(target=load_dataframe, args=(data_csv_path,), daemon=True)
    t.start()


def _calc_age(acc_yrmo, sys_yrmo):
    # Detect format by digit count: 4 digits = YYYY (annual), 6 digits = YYYYMM (monthly)
    if len(str(int(acc_yrmo))) == 4:  # YYYY annual format
        return (int(sys_yrmo) - int(acc_yrmo)) * 12 + 1
    # YYYYMM monthly format (original logic)
    acc_yr = acc_yrmo//100
    sys_yr = sys_yrmo//100
    acc_mo = acc_yrmo % 100
    sys_mo = sys_yrmo % 100

    return 12*(sys_yr-acc_yr) + sys_mo-acc_mo + 1


def _get_org_label(date_val, org_len):
    # Detect format by digit count: 4 digits = YYYY (annual), 6 digits = YYYYMM (monthly)
    if len(str(int(date_val))) == 4:  # YYYY annual format
        return int(date_val)  # org_len==12 is always the case for annual data

    # YYYYMM monthly format (original logic)
    yyyymm = date_val
    year = int(yyyymm // 100)
    month = int(yyyymm % 100)

    if org_len == 1:
        return yyyymm
        # return "'" + datetime.strptime(str(yyyymm), "%Y%m").strftime("%b %Y")

    elif org_len == 3:
        return f"{year} Q{(month+2)//3}"
    elif org_len == 6:
        return f"{year} H{(month+5)//6}"
    elif org_len == 12:
        return year


def _request_bool(value, default=False):
    if isinstance(value, bool):
        return value
    text = str(value if value is not None else "").strip().lower()
    if text in {"true", "yes", "1"}:
        return True
    if text in {"false", "no", "0"}:
        return False
    return default


def _as_month_period(value):
    n = int(value)
    if len(str(n)) == 4:
        return n * 100 + 1
    return n


def _add_months(yyyymm, months):
    year = int(yyyymm) // 100
    month = int(yyyymm) % 100
    total = year * 12 + (month - 1) + int(months)
    return (total // 12) * 100 + (total % 12) + 1


def _month_period_label(start_yyyymm, end_yyyymm):
    start = int(start_yyyymm)
    end = int(end_yyyymm)
    return f"{start % 100:02d}/{start // 100 % 100:02d} - {end % 100:02d}/{end // 100 % 100:02d}"


def _development_age_labels(project_settings, dev_len):
    acc_yrmo_all = _generate_period_range(
        project_settings['origin_start'], project_settings['origin_end'],
        project_settings.get('date_granularity', 'monthly'))
    is_annual = project_settings.get('date_granularity') == 'annual'
    dev_cnt = len(acc_yrmo_all) if is_annual else round(len(acc_yrmo_all)/dev_len)
    first_mon = int(project_settings['dev_end'] % 100)

    dev_label = list(range(first_mon, dev_cnt*dev_len+1, dev_len))
    for i in range(1, 999):
        prior_mon = first_mon - dev_len*i
        if prior_mon > 0:
            dev_label = [prior_mon] + dev_label
        else:
            break
    return dev_label


def _calendar_periods_from_dev_labels(project_settings, dev_label):
    origin_start = _as_month_period(project_settings['origin_start'])
    periods = []
    prior_end = None
    for age in dev_label:
        end_period = _add_months(origin_start, int(age) - 1)
        start_period = origin_start if prior_end is None else _add_months(prior_end, 1)
        periods.append({
            "start": start_period,
            "end": end_period,
            "label": _month_period_label(start_period, end_period),
        })
        prior_end = end_period
    return periods


def _reshape_triangle_to_calendar(df, org_index_grp, dev_label, project_settings):
    calendar_periods = _calendar_periods_from_dev_labels(project_settings, dev_label)
    calendar_labels = [period["label"] for period in calendar_periods]
    end_to_col = {period["end"]: idx for idx, period in enumerate(calendar_periods)}
    out = pd.DataFrame(np.nan, index=df.index, columns=calendar_labels)

    for row_idx, group in enumerate(org_index_grp):
        if row_idx >= len(df.index) or not group:
            continue
        origin_start = _as_month_period(group[0])
        for dev_age in dev_label:
            if dev_age not in df.columns:
                continue
            value = df.iat[row_idx, df.columns.get_loc(dev_age)]
            if pd.isna(value):
                continue
            dev_end = _add_months(origin_start, int(dev_age) - 1)
            col_idx = end_to_col.get(dev_end)
            if col_idx is None:
                continue
            out.iat[row_idx, col_idx] = value
    return out


def vector_to_triangle(df: pd.Series | pd.DataFrame, colnames=None) -> pd.DataFrame:
    """
    Convert a vector (Series or 1-col DataFrame) to a triangle DataFrame.
    If the input is already square (n×n), return it unchanged.
    
    colnames: optional list/Index to use as column names.
              If None, defaults to using the row index.
    """
    # Case 1: Already a square DataFrame → do nothing
    if isinstance(df, pd.DataFrame) and df.shape[0] == df.shape[1]:
        return df

    # Convert to Series for uniform processing
    if isinstance(df, pd.DataFrame):
        if df.shape[1] != 1:
            raise ValueError("DataFrame must have exactly one column or be square.")
        s = df.iloc[:, 0]
    elif isinstance(df, pd.Series):
        s = df
    else:
        raise TypeError("Input must be a pandas Series or 1-column DataFrame.")

    # Default column names = index
    if colnames is None:
        colnames = s.index
    else:
        # if len(colnames) != len(s): raise ValueError("Length of colnames must match length of vector.")
        pass

    # Expand vector → row-constant matrix
    idx = s.index
    arr = np.repeat(s.values.reshape(-1, 1), len(colnames), axis=1)

    return pd.DataFrame(arr, index=idx, columns=colnames, dtype=float)


def eval_triangle_formula(triangles: dict[str, pd.DataFrame],
                          formula: str,
                          div0_to_zero: bool = True) -> pd.DataFrame:
    """
    triangles: dict like {'A': tri_A, 'B': tri_B, ...} where each value is a pivoted DF
    formula:   e.g. 'D = A/B*1000' or 'A/B*1000' or 'A + B*C'
    div0_to_zero: if True, convert inf/NaN from division-by-zero to 0
    """
    # allow 'D = A/B*1000' or just 'A/B*1000'
    rhs = formula.split('=', 1)[-1].strip()

    # safety: no builtins; variables come from triangles dict
    env = {"__builtins__": {}}

    # element-wise eval; pandas aligns on index & columns automatically
    result = eval(rhs, env, triangles)

    if div0_to_zero:
        result = result.replace([np.inf, -np.inf], np.nan).fillna(0)

    # ensure numeric dtype (optional)
    return result.astype(float)


def _get_df(project_name):
    table_path = get_project_table_path(project_name)
    table_key = _data_table_cache_key(table_path)

    # DATA table cache (guarded)
    with DATA_DICT_LOCK:
        need_load = (table_key not in DATA_DICT) or (DATA_DICT.get(table_key + " - Version") is None) \
                    or (DATA_DICT[table_key + " - Version"] < datetime.fromtimestamp(os.path.getmtime(table_path)))

    if need_load:
        # build outside lock if you want, but simplest is just load here
        with DATA_DICT_LOCK:
            load_to_DATA_DICT(table_path)

    # VPS cache (guarded)
    with PROJECT_CONFIG_LOCK:
        if project_name not in PROJECT_CONFIG:
            load_to_PROJECT_CONFIG(project_name)

    return DATA_DICT[table_key]


def _get_dataset_info(arg):
    # This apply to both vector and triangle
    project_name = arg['ProjectName']
    path = arg['Path']
    dataset_name = arg['DatasetName']

    df = _get_df(project_name)

    # Set user defined name (ResQ) to actual SQL table col names
    df_info = PROJECT_CONFIG[project_name]['Dataset Types']
    
    if dataset_name in df_info['Name'].values:
        source = df_info.loc[df_info['Name'] == dataset_name, 'Source'].iloc[0]
    else:
        raise DataProcessingConfigurationError(
            f"Dataset type [{dataset_name}] is not defined for project [{project_name}]."
        )
    
    output_data_format = df_info.loc[df_info['Name'] == dataset_name, 'Data Format'].iloc[0]

    # find all required table and column names
    df_info = PROJECT_CONFIG[project_name]['Source Table']
    required_datasets = split_formula(source)
    catalog = PROJECT_CONFIG[project_name]['Reserving Class Catalog']

    date_cols = []
    date_cols.append(DLOOKUP(df_info, 'Origin Date', 'Significances', 'Column Name'))
    date_cols.append(DLOOKUP(df_info, 'Development Date', 'Significances', 'Column Name'))

    # Load project-specific date settings (with fallback to data-derived values)
    project_settings = _load_project_settings(project_name, df, date_cols)
    max_sys_yrmo = project_settings['dev_end']

    required_datasets = list(
        dict.fromkeys(c for c in required_datasets if c in df.columns)
    )
    if not required_datasets:
        raise DataProcessingConfigurationError(
            f"Dataset type [{dataset_name}] does not resolve to any numeric source "
            "columns in the project source table."
        )

    request_context, selected_coefficients = resolve_request_path(catalog, path)
    compiled_rules = PROJECT_CONFIG[project_name]['Data Processing Rules']

    return df, date_cols, required_datasets, \
           selected_coefficients, request_context, compiled_rules.rules, \
           source, output_data_format, max_sys_yrmo


def UDF_ADASProjectSettings(arg):
    project_name = arg['ProjectName']
    df = _get_df(project_name)

    df_info = PROJECT_CONFIG[project_name]['Source Table']
    date_cols = []
    date_cols.append(DLOOKUP(df_info, 'Origin Date', 'Significances', 'Column Name'))
    date_cols.append(DLOOKUP(df_info, 'Development Date', 'Significances', 'Column Name'))

    # Load project-specific date settings (with fallback to data-derived values)
    project_settings = _load_project_settings(project_name, df, date_cols)
    origin_start = project_settings['origin_start']
    origin_end = project_settings['origin_end']
    dev_end = project_settings['dev_end']

    data_list = [
        ['Name', project_name], 
        ['Origin Type', 'Accident'], 
        ['Origin Start Date', date(origin_start // 100, origin_start % 100, 1)], 
        ['Origin End Date', date(origin_end // 100, origin_end % 100, calendar.monthrange(origin_end // 100, origin_end % 100)[1])], 
        ['Development End Date', date(dev_end // 100, dev_end % 100, calendar.monthrange(dev_end // 100, dev_end % 100)[1])], 
        ['Origin Length', 12], 
        ['Development Length', 12], 
        ['Folder', f'{function_brand(arg.get("Function"))} Virtual Project']
    ]
    write_lists_to_csv(arg['DataPath'], data_list)


def UDF_ADASHeaders(arg):
    # Calculate Age & Origin Labels
    project_name = arg['ProjectName']
    org_len = int(arg['PeriodLength'])
    dev_len = int(arg['PeriodLength'])
    period_type = int(arg['periodType'])
    calendar_mode = _request_bool(arg.get('Calendar'), False)

    df = _get_df(project_name)
    df_info = PROJECT_CONFIG[project_name]['Source Table']
    date_cols = []
    date_cols.append(DLOOKUP(df_info, 'Origin Date', 'Significances', 'Column Name'))
    date_cols.append(DLOOKUP(df_info, 'Development Date', 'Significances', 'Column Name'))

    # Load project-specific date settings (with fallback to data-derived values)
    project_settings = _load_project_settings(project_name, df, date_cols)

    if period_type == 0: # Origin Period

        # Generate period range from project configuration (handles both annual and monthly)
        acc_yrmo_all = _generate_period_range(
            project_settings['origin_start'], project_settings['origin_end'],
            project_settings.get('date_granularity', 'monthly'))
        # Compute slicing step: for annual, group by 1 year; for monthly, group by org_len months
        org_step = 1 if project_settings.get('date_granularity') == 'annual' else org_len
        org_index_grp = [tuple(acc_yrmo_all[i: i+org_step]) for i in range(0, len(acc_yrmo_all), org_step)]

        org_label = [_get_org_label(i[0], org_len) for i in org_index_grp]

        return write_lists_to_csv(arg['DataPath'], [org_label])
    
    elif period_type == 1: # Development Period

        if (dev_len == 'Default') or (org_len % dev_len != 0):
            dev_len = org_len

        dev_label = _development_age_labels(project_settings, dev_len)

        if calendar_mode:
            dev_label = [period["label"] for period in _calendar_periods_from_dev_labels(project_settings, dev_label)]
        else:
            dev_label = list(map(lambda x:f"{x}m", dev_label))
        return write_lists_to_csv(arg['DataPath'], [dev_label])
    
    else:
        return write_lists_to_csv(arg['DataPath'], [['(invalid input: periodType)']])


def UDF_ADASTri(arg):
    org_len = arg['OriginLength']
    dev_len = arg['DevelopmentLength']
    cumulative = _request_bool(arg.get('Cumulative'), True)
    calendar_mode = _request_bool(arg.get('Calendar'), False)
    project_name = arg['ProjectName']

    # initialize
    if org_len == 'Default': org_len = 12

    # Get a subset dataframe based on a user's request
    df, date_cols, required_datasets, selected_coefficients, \
    request_context, processing_rules, \
    source, output_data_format, max_sys_yrmo = _get_dataset_info(arg)

    # Load project-specific date settings (with fallback to data-derived values)
    # Note: _get_dataset_info already loads settings, but we reload here for local use
    project_settings = _load_project_settings(project_name, df, date_cols) 

    max_sys_month = max_sys_yrmo % 100

    df1 = build_weighted_source_frame(
        df,
        passthrough_columns=[column for column in date_cols if column],
        source_measures=required_datasets,
        selected_coefficients=selected_coefficients,
        request_context=request_context,
        rules=processing_rules,
    )

    # Check if Development Date column is missing (optional when not in field_mapping)
    has_dev_date = date_cols[1] != '' and date_cols[1] in df1.columns
    if not has_dev_date:
        # Single-column triangle: set dev_len = 1 (unless explicitly set)
        if dev_len == 'Default' or dev_len == org_len:
            dev_len = 1

    # Prepare for grouping by origin period and development age
    if (dev_len == 'Default') or (org_len % dev_len != 0):
        dev_len = org_len

    # Use project configuration to calculate development counts
    acc_yrmo_all = _generate_period_range(
        project_settings['origin_start'], project_settings['origin_end'],
        project_settings.get('date_granularity', 'monthly'))
    # Fix dev_cnt: for annual, number of periods = number of years; for monthly, divide by dev_len
    is_annual = project_settings.get('date_granularity') == 'annual'
    dev_cnt = len(acc_yrmo_all) if is_annual else round(len(acc_yrmo_all)/dev_len)
    first_mon = int(project_settings['dev_end'] % 100)

    # When Development Date is missing, use MMM YYYY format from dev_end config
    if not has_dev_date:
        dev_end = project_settings['dev_end']
        dev_year = dev_end // 100
        dev_month = dev_end % 100
        # Convert YYYYMM to MMM YYYY format (e.g., 202603 -> Mar 2026)
        import calendar
        month_abbr = calendar.month_abbr[dev_month]
        dev_label = [f"{month_abbr} {dev_year}"]
    else:
        dev_label = _development_age_labels(project_settings, dev_len)
    # Compute slicing step: for annual, group by 1 year; for monthly, group by org_len months
    org_step = 1 if is_annual else org_len
    org_index_grp = [tuple(acc_yrmo_all[i: i+org_step]) for i in range(0, len(acc_yrmo_all), org_step)]
    org_index_map = {val: group[0] for group in org_index_grp for val in group}
    org_label = [_get_org_label(i[0], org_len) for i in org_index_grp]
    
    df1['Org*Grp'] = df1[date_cols[0]].apply(lambda x: _get_org_label(x, org_len))

    df1['Org*Start'] = df1[date_cols[0]].map(org_index_map)
    # When Development Date is missing, use dev_end from config for Age* (single column triangle)
    if has_dev_date:
        df1['Age*'] = df1[['Org*Start', date_cols[1]]].apply(lambda row: _calc_age(row.iloc[0], row.iloc[1]), axis=1)
    else:
        df1['Age*'] = df1['Org*Start'].apply(lambda x: _calc_age(x, project_settings['dev_end']))

    # When Development Date is missing (single column), all rows map to the single dev_label value
    if not has_dev_date:
        df1['Age*Grp'] = dev_label[0]  # Single column: all rows get the same label
    else:
        df1['Age*Grp'] = df1['Age*'].apply(lambda x: min([i for i in dev_label if i >= x]))

    df1 = df1.groupby(['Org*Grp', 'Age*Grp'])[required_datasets].sum().reset_index()

    # Create individual non-calculated triangles
    triangles  = {}
    
    for name in required_datasets:
        df2 = df1.pivot_table(
            index = df1['Org*Grp'], 
            columns = df1['Age*Grp'], 
            values = name,
            aggfunc = 'sum', 
            fill_value = 0
        )
        df2 = df2.reindex(index=org_label, columns=dev_label).fillna(0)

        if cumulative == True: 
            df2 = df2.cumsum(axis=1)

        data_format = DLOOKUP(PROJECT_CONFIG[arg['ProjectName']]['Dataset Types'], name, 'Source', 'Data Format')
        if data_format == 'Vector':
            df2 = vector_to_triangle(df2.iloc[:, [0]], dev_label)

        triangles[name] = df2
    
    # Calculated Triangle
    df2 = eval_triangle_formula(triangles, source)  

    # Clean Format
    if output_data_format == 'Vector' or is_vector_function(arg['Function']):
        # A vector has no development axis, so every origin period keeps its
        # value: the source table may carry full-year inputs (planned premium,
        # exposure) for origin periods after the Development End Date.
        df2 = df2.iloc[:, [0]].fillna(0)
        _export_dataframe(df2, arg)
        return

    n_rows = df2.shape[0]

    for i, acc in enumerate(df2.index):
        max_dev_age = (n_rows - i) * int((org_len/dev_len))
        
        # if org_len == 3 and dev_len == 1:
        if dev_len == 1:
            max_dev_age = max_dev_age - (12 - max_sys_month)

        if dev_len == 3:
            if max_sys_month in [1, 2, 3]:
                max_dev_age = max_dev_age - 3
            elif max_sys_month in [4, 5, 6]:
                max_dev_age = max_dev_age - 2
            elif max_sys_month in [7, 8, 9]:
                max_dev_age = max_dev_age - 1

        if dev_len == 6 and max_sys_month <= 6:
            max_dev_age = max_dev_age - 1

        if max_dev_age < 0:
            max_dev_age = 0

        df2.loc[acc, dev_label[max_dev_age:]] = np.nan

    if calendar_mode:
        df2 = _reshape_triangle_to_calendar(df2, org_index_grp, dev_label, project_settings)

    # Output
    _export_dataframe(df2, arg)


def _export_dataframe(df, arg):
    data_path = arg['DataPath']
    file_name = os.path.basename(data_path)
    folder = os.path.dirname(data_path)
    tmp_folder = folder + '\\tmp'
    tmp_data_path = tmp_folder + '\\' + file_name
    
    try:
        if not os.path.exists(folder):
            os.makedirs(folder)
        if not os.path.exists(tmp_folder):
            os.makedirs(tmp_folder)
    except:
        pass

    df.to_csv(tmp_data_path, index=False, header=False)

    if os.path.exists(data_path):
        os.remove(data_path)

    os.rename(tmp_data_path, data_path)



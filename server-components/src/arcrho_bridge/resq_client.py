import re
from pathlib import Path
import math
import threading
from datetime import timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import pythoncom
import win32com.client
import win32timezone  # noqa: F401 - required by pywin32 COM date conversion in frozen builds.

from arcrho_bridge.bridge_utils import read_json, write_json, write_json_with_compact_rows


CONNECTION_NAME = "JGO_CO1SQLWPV22"
RESQ_CONFIG_SECTION = "resq"
DFM_OWNED_PATCH_FORMAT = "arcrho-dfm-owned-patch-v4"
RESULT_SELECTION_JSON_FORMAT = "arcrho-result-selection-v4"


def resq_connection_settings():
    """Read the shared ResQ service account from the deployed config.json.

    Every ArcRho user connects to ResQ as one service account rather than with
    their own Windows authentication, so the credentials live in the shared
    ``<ArcRho Server>\\config\\config.json`` under ``resq`` instead of being
    compiled into the Bridge exe. Rotating the account is a config edit that the
    next connect picks up, with no rebuild or redeploy.
    """

    # Imported lazily: utils resolves the deploy root at import time, and the
    # Bridge sets that up in main.py before any COM work starts.
    from utils import get_config_value

    def setting(key, default=""):
        value = get_config_value(f"{RESQ_CONFIG_SECTION}.{key}", default)
        return str(value).strip() if value is not None else ""

    connection_name = setting("connection_name") or CONNECTION_NAME
    user_name = setting("user_name")
    password = setting("password")
    if not user_name or not password:
        raise RuntimeError(
            "The shared ResQ service account is not configured: set "
            f"'{RESQ_CONFIG_SECTION}.user_name' and '{RESQ_CONFIG_SECTION}.password' "
            "in the ArcRho Server config.json."
        )
    return connection_name, user_name, password


def shared_resq_credentials():
    """The same service account, shaped for the queued import and sync sessions.

    Those sessions open their own COM connection inside whichever user's Bridge
    worker claimed the request. Left to the migration's defaults they would
    connect with that worker's Windows identity, and ResQ would show each user
    only the projects they hold, so the outcome of a queued request depended on
    who won the claim.
    """

    connection_name, user_name, password = resq_connection_settings()
    return {"connection_name": connection_name, "user_name": user_name, "password": password}


class ResQClient:
    def __init__(self):
        self.app = None
        self._disconnect_lock = threading.RLock()
        self._com_thread_id = None
        self._com_initialized = False

    def _ensure_com_initialized(self):
        thread_id = threading.get_ident()
        if self._com_initialized and self._com_thread_id == thread_id:
            return
        pythoncom.CoInitialize()
        self._com_initialized = True
        self._com_thread_id = thread_id

    def _uninitialize_com(self):
        if not self._com_initialized:
            return
        if self._com_thread_id != threading.get_ident():
            return
        pythoncom.CoUninitialize()
        self._com_initialized = False
        self._com_thread_id = None

    def _connect(self):
        with self._disconnect_lock:
            if self.app is not None and self._com_thread_id != threading.get_ident():
                raise RuntimeError("ResQ COM connection is owned by another bridge worker thread.")
            self._ensure_com_initialized()
            if self.app is None:
                try:
                    connection_name, user_name, password = resq_connection_settings()
                    self.app = win32com.client.Dispatch("ResQ3Automation.ResQApplication")
                    self.app.ConnectByName(connection_name, user_name, password)
                except Exception:
                    self.app = None
                    self._uninitialize_com()
                    raise
            return self.app

    def disconnect_if_idle(self):
        return

    def _disconnect(self):
        with self._disconnect_lock:
            if self.app is not None and self._com_thread_id != threading.get_ident():
                return
            app = self.app
            self.app = None
        if app is None:
            self._uninitialize_com()
            return
        try:
            app.Disconnect()
        except Exception:
            pass
        finally:
            self._uninitialize_com()

    def close(self):
        self._disconnect()

    def write_resq_reserving_class_import(self, request, *, progress_callback=None):
        """Run the canonical staged ResQ import from this Bridge worker.

        The importer owns its full-reserving-class COM session and canonical
        JSON writer. Keeping this as a small delegation prevents the Bridge
        client from becoming a second persisted-data producer while preserving
        the existing RPC methods and their shorter-lived COM sessions.
        """

        # The canonical migration creates its own COM objects, so initialize
        # COM on this worker-owned thread before delegating to it.
        self._ensure_com_initialized()

        from arcrho_bridge.resq_import_runner import run_reserving_class_import

        return run_reserving_class_import(
            request,
            progress_callback=progress_callback,
            resq_credentials=shared_resq_credentials(),
        )

    def write_resq_reserving_class_sync(self, request, *, progress_callback=None):
        """Run one canonical ArcRho/ResQ synchronization phase from this worker.

        The session owns its own full-reserving-class COM connection and the
        canonical JSON writers, exactly as the importer does, so this stays a
        delegation rather than a second producer of persisted data.
        """

        # The canonical session creates its own COM objects, so initialize COM
        # on this worker-owned thread before delegating to it.
        self._ensure_com_initialized()

        from arcrho_bridge.resq_sync_runner import run_reserving_class_sync

        return run_reserving_class_sync(
            request,
            progress_callback=progress_callback,
            resq_credentials=shared_resq_credentials(),
        )

    def write_dfm_payload(self, request):
        self._connect()
        try:
            dfm = self._dfm_method(request)
            output_vector = self._optional_value(dfm, "OutputVector", None)
            output_dataset = self._clean_label(request.get("OutputVector") or self._nested_name(dfm, "OutputVector"))
            output_type = self._clean_label(self._nested_name(output_vector, "DatasetType") if output_vector is not None else "")
            output_category = self._clean_label(
                self._nested_name(self._optional_value(output_vector, "DatasetType", None), "Category")
                if output_vector is not None else ""
            )
            average_data = self._average_data(dfm)
            origin_labels, data_development_labels = self._labels(dfm)
            ratio_development_labels = self._ratio_development_labels(data_development_labels)
            cell_notes = self._cell_notes_data(
                dfm,
                origin_labels,
                ratio_development_labels,
                average_data.get("label", []),
            )
            payload = {
                "payload_format": DFM_OWNED_PATCH_FORMAT,
                "details_tab": {
                    "name": self._clean_label(request.get("MethodName") or self._optional_value(dfm, "Name", "")),
                    "output_type": output_type or output_dataset,
                    "output_dataset": output_dataset,
                    "output dataset_category": output_category,
                    "output_category": output_category,
                    "input_triangle": self._nested_name(dfm, "InputTriangle"),
                    "origin_length": self._optional_value(dfm, "OriginLength", ""),
                    "development_length": self._optional_value(dfm, "DevelopmentLength", ""),
                    "decimal_places": self._optional_value(dfm, "RatioDecimalPlaces", request.get("DecimalPlaces", 4)),
                },
                "ratios_tab": {
                    "ratio_triangle": {
                        "origin_labels": origin_labels,
                        "development_labels": ratio_development_labels,
                        "excluded": self._excluded_ratio_pattern(dfm),
                    },
                    "average_formulas": average_data,
                    "cell_notes": cell_notes,
                },
                "results_tab": {
                    "ratio_basis_dataset": self._nested_name(dfm, "SummaryRatioBasis"),
                    "ultimate_ratio_decimal_places": self._optional_value(dfm, "SummaryRatioDecimalPlaces", 2),
                },
                "method_metadata": {
                    "last_modified": self._output_vector_modified(dfm),
                    "method_notes": self._method_notes(dfm),
                },
            }
            write_json_with_compact_rows(request["DataPath"], payload)
            return payload
        finally:
            self._disconnect()

    def write_sync_dfm_payload(self, request):
        self._connect()
        try:
            dfm = self._dfm_method(request)
            payload = read_json(request["MethodJsonPath"])
            excluded_count = self._sync_excluded_ratios(dfm, payload)
            user_entry_count = self._sync_user_entry_values(dfm, payload)
            selected_count = self._sync_selected_ratios(dfm, payload)
            cell_notes_changed = self._sync_cell_notes(dfm, payload)
            method_notes_changed = self._sync_method_notes(dfm, request)
            dfm.Save()
            payload = {
                "ok": True,
                "status": "passed",
                "message": "Remote database updated",
                "updated": {
                    "excluded ratios": excluded_count,
                    "selected ratios": selected_count,
                    "user entry values": user_entry_count,
                    "cell_notes": cell_notes_changed,
                    "method_notes": method_notes_changed,
                },
                # The save above gave this method a new ResQ ``Modified``.
                # Report it so ArcRho can record the same instant against its
                # own copy: the two now hold identical settings, and without
                # this the next sync review calls the remote newer and offers
                # to pull back what was just pushed. Read after Save() and in
                # the same spelling the DFM export uses.
                "last_modified": self._output_vector_modified(dfm),
            }
            write_json(request["DataPath"], payload)
            return payload
        finally:
            self._disconnect()

    def write_error(self, request, message):
        data_path = request.get("DataPath")
        if not data_path:
            return
        write_json(
            Path(data_path),
            {
                "ok": False,
                "status": "error",
                "message": str(message),
            },
        )

    def _dfm_method(self, request):
        project = self.app.Projects().Item(request["ProjectName"])
        reserving_class = project.ReservingClasses().Item(request["Path"])
        return reserving_class.DFMMethods().Item(request["MethodName"])

    def _excluded_ratio_pattern(self, dfm):
        rows = int(dfm.OriginCount)
        row_widths = [
            max(int(dfm.DevelopmentCount(origin_index)) - 1, 0)
            for origin_index in range(1, rows + 1)
        ]
        columns = max(row_widths, default=0)
        pattern = []
        for origin_index, ratio_count in enumerate(row_widths, start=1):
            row = []
            for development_index in range(1, columns + 1):
                if development_index <= ratio_count:
                    row.append(int(dfm.ExcludedRatios(origin_index, development_index)))
                else:
                    row.append(2)
            pattern.append(self._trim_trailing_mask_cells(row))
        return pattern

    def _sync_excluded_ratios(self, dfm, payload):
        ratio_triangle = self._dict_path(payload, ("ratios_tab", "ratio_triangle"))
        pattern = ratio_triangle.get("excluded") if isinstance(ratio_triangle, dict) else None
        if not isinstance(pattern, list):
            return 0

        origin_count = int(self._optional_value(dfm, "OriginCount", 0) or 0)
        updates = 0
        for origin_index, row in enumerate(pattern, start=1):
            if origin_index > origin_count or not isinstance(row, list):
                continue
            ratio_count = max(int(dfm.DevelopmentCount(origin_index)) - 1, 0)
            for development_index, raw_value in enumerate(row, start=1):
                if development_index > ratio_count:
                    break
                value = self._excluded_value(raw_value)
                if value is None:
                    continue
                dfm.SetExcludedRatios(OriginIndex=origin_index, DevIndex=development_index, arg2=value)
                updates += 1
        return updates

    def _excluded_value(self, value):
        if value in (0, False, "0", "false", "False"):
            return 0
        if value in (1, True, "1", "true", "True"):
            return 1
        return None

    def _sync_selected_ratios(self, dfm, payload):
        average_formulas = self._dict_path(payload, ("ratios_tab", "average_formulas"))
        labels = average_formulas.get("label") if isinstance(average_formulas, dict) else None
        selected = average_formulas.get("selected") if isinstance(average_formulas, dict) else None
        if not isinstance(labels, list) or not isinstance(selected, list):
            return 0

        label_to_display_index = self._average_formula_display_indexes(dfm)
        column_count = self._development_column_count(dfm)
        updates = 0
        for development_index in range(1, column_count + 1):
            selected_label = self._selected_label_for_column(labels, selected, development_index - 1)
            if not selected_label:
                continue
            display_index = label_to_display_index.get(selected_label)
            if display_index is None:
                continue
            dfm.SetSelectedRatios(DevIndex=development_index, arg1=display_index)
            updates += 1
        return updates

    def _sync_user_entry_values(self, dfm, payload):
        average_formulas = self._dict_path(payload, ("ratios_tab", "average_formulas"))
        labels = average_formulas.get("label") if isinstance(average_formulas, dict) else None
        values = average_formulas.get("values") if isinstance(average_formulas, dict) else None
        if not isinstance(labels, list) or not isinstance(values, list):
            return 0

        row_index = self._user_entry_payload_row_index(average_formulas, labels)
        if row_index is None or row_index >= len(values) or not isinstance(values[row_index], list):
            return 0

        avg_index = self._user_entry_resq_index(dfm)
        if avg_index is None:
            return 0

        column_count = self._development_column_count(dfm)
        updates = 0
        for development_index, raw_value in enumerate(values[row_index], start=1):
            if development_index > column_count:
                break
            value = self._positive_number(raw_value)
            if value is None:
                continue
            self._set_user_entry_average_ratio_value(dfm, development_index, avg_index, value)
            updates += 1
        return updates

    def _user_entry_payload_row_index(self, average_formulas, labels):
        settings = average_formulas.get("custom_average_formula_settings")
        average_types = settings.get("average_type") if isinstance(settings, dict) else None
        if isinstance(average_types, list):
            for index, average_type in enumerate(average_types):
                if str(average_type or "").strip().lower() == "user_entry":
                    return index

        for index, label in enumerate(labels):
            normalized = self._clean_label(label).lower()
            if normalized == "user entry" or normalized.startswith("user entry "):
                return index
        return None

    def _user_entry_resq_index(self, dfm):
        for api_index in range(1, 50):
            try:
                raw_name = str(dfm.AverageFormula(api_index))
            except Exception:
                break
            display_index, name = self._parse_average_formula_name(raw_name, api_index)
            normalized = self._clean_label(name).lower()
            if normalized == "user entry" or normalized.startswith("user entry "):
                return display_index
        return None

    def _positive_number(self, value):
        if value is None or isinstance(value, bool):
            return None
        try:
            number = float(value)
        except (TypeError, ValueError):
            return None
        if number <= 0:
            return None
        return number

    def _set_user_entry_average_ratio_value(self, dfm, development_index, avg_index, value):
        try:
            dfm.SetUserRatios(DevIndex=development_index, AvgIndex=avg_index, arg2=value)
        except Exception as exc:
            raise RuntimeError(f"Unable to update DFM User Entry value in ResQ with SetUserRatios: {exc}") from exc

    def _average_formula_display_indexes(self, dfm):
        out = {}
        for api_index in range(1, 50):
            try:
                raw_name = str(dfm.AverageFormula(api_index))
            except Exception:
                break
            display_index, name = self._parse_average_formula_name(raw_name, api_index)
            out.setdefault(name, display_index)
            if name == "User Entry":
                break
        return out

    def _selected_label_for_column(self, labels, selected, column_index):
        for row_index, row in enumerate(selected):
            if row_index >= len(labels) or not isinstance(row, list) or column_index >= len(row):
                continue
            if row[column_index] in (1, True, "1", "true", "True"):
                return str(labels[row_index])
        return ""

    def _sync_cell_notes(self, dfm, payload):
        cell_notes = self._dict_path(payload, ("ratios_tab", "cell_notes"))
        if not cell_notes:
            return False
        # ResQ exposes DFM CellNotes as a read-side formatted string. The current
        # bridge examples do not expose a safe per-cell note setter, so remote
        # write-back intentionally leaves cell notes unchanged.
        _ = dfm
        return "read-only"

    def _dict_path(self, payload, path):
        current = payload
        for key in path:
            if not isinstance(current, dict):
                return {}
            current = current.get(key)
        return current if isinstance(current, dict) else {}

    def _trim_trailing_mask_cells(self, row):
        out = list(row)
        while out and out[-1] == 2:
            out.pop()
        return out

    def _average_data(self, dfm):
        formula_rows = self._average_formula_rows(dfm)
        column_count = self._development_column_count(dfm)
        display_indexes = [row["display_index"] for row in formula_rows]
        selected_indexes = [
            self._selected_average_display_index(dfm, development_index, display_indexes)
            for development_index in range(1, column_count + 1)
        ]

        return {
            "label": [row["name"] for row in formula_rows],
            "custom_average_formula_settings": {
                "average_type": [row["average_type"] for row in formula_rows],
                "base": [row["base"] for row in formula_rows],
                "periods": [row["periods"] for row in formula_rows],
                "exclude": [row["exclude"] for row in formula_rows],
            },
            "selected": [
                [1 if selected_index == row["display_index"] else 0 for selected_index in selected_indexes]
                for row in formula_rows
            ],
            "values": self._user_entry_average_formula_values(dfm, formula_rows, column_count),
        }

    def _method_notes(self, dfm):
        # ResQ method-level Notes; the ArcRho output sidecar `notes` field is its
        # only persisted ArcRho owner, so this value stays in transient metadata.
        return str(self._optional_value(dfm, "Notes", "") or "")

    def _sync_method_notes(self, dfm, request):
        if "MethodNotes" not in request:
            # An omitted field means the local notes owner was unavailable;
            # leave ResQ method Notes unchanged.
            return 0
        notes = str(request.get("MethodNotes") or "")
        if not notes.strip():
            notes = ""
        # ResQ Notes require \r\n line breaks; a \n-only value renders as one line.
        normalized = re.sub(r"\r?\n", "\r\n", notes)
        if self._method_notes(dfm) == normalized:
            return 0
        dfm.Notes = normalized
        return 1

    def _cell_notes_data(self, dfm, origin_labels, ratio_development_labels, average_labels):
        lines = str(self._optional_value(dfm, "CellNotes", "") or "").splitlines()
        development_label_map = self._development_note_label_map(ratio_development_labels)
        origin_label_set = {self._label_key(label) for label in origin_labels}
        average_label_set = {self._label_key(label) for label in average_labels}
        out = {
            "ratio_main_table": {},
            "ratio_summary_table": {},
        }
        for line in lines:
            parsed = self._parse_cell_note_line(line)
            if not parsed:
                continue
            col_label = development_label_map.get(self._label_key(parsed["x_label"]), parsed["x_label"])
            row_label = parsed["y_label"]
            note = parsed["note"]
            if not col_label or not row_label or not note:
                continue
            row_key = self._label_key(row_label)
            table_key = "ratio_summary_table" if row_key in average_label_set and row_key not in origin_label_set else "ratio_main_table"
            out.setdefault(table_key, {}).setdefault(row_label, {})[col_label] = note
        return out

    def _parse_cell_note_line(self, line):
        text = str(line or "").strip()
        if not text:
            return None
        match = re.match(
            r'^\s*"(?P<tab>(?:[^"]|"")*)"\s*,\s*Cell\[(?P<x_label>.*?),\s*(?P<y_label>.*?)\]\s*,\s*"(?P<note>(?:[^"]|"")*)"',
            text,
        )
        if not match:
            return None
        return {
            "tab": self._unescape_resq_note_value(match.group("tab")),
            "x_label": self._clean_label(match.group("x_label")),
            "y_label": self._clean_label(match.group("y_label")),
            "note": self._unescape_resq_note_value(match.group("note")).strip(),
        }

    def _development_note_label_map(self, ratio_development_labels):
        out = {}
        for label in ratio_development_labels:
            display_label = self._clean_label(label)
            if not display_label:
                continue
            out[self._label_key(display_label)] = display_label
            without_index = re.sub(r"^\(\s*\d+\s*\)\s*", "", display_label).strip()
            if without_index:
                out.setdefault(self._label_key(without_index), display_label)
        return out

    def _unescape_resq_note_value(self, value):
        return str(value or "").replace('""', '"')

    def _clean_label(self, value):
        return re.sub(r"\s+", " ", str(value or "")).strip()

    def _label_key(self, value):
        return self._clean_label(value).lower()

    def _average_formula_rows(self, dfm):
        rows = []
        for api_index in range(1, 20):
            try:
                raw_name = str(dfm.AverageFormula(api_index))
            except Exception:
                break

            display_index, name = self._parse_average_formula_name(raw_name, api_index)
            row = {
                "api_index": api_index,
                "display_index": display_index,
                "name": name,
                "is_user_entry": name == "User Entry",
            }
            row.update(self._formula_metadata(name, row["is_user_entry"]))
            rows.append(row)
            if row["is_user_entry"]:
                break
        return rows

    def _parse_average_formula_name(self, raw_name, api_index):
        match = re.match(r"^\s*(\d+)\s*:\s*(.*?)\s*$", raw_name)
        if not match:
            return api_index - 1, raw_name.strip()
        return int(match.group(1)), match.group(2)

    def _selected_average_display_index(self, dfm, development_index, display_indexes):
        try:
            selected_index = int(dfm.SelectedRatios(development_index))
        except Exception:
            return None

        display_index_set = set(display_indexes)
        if selected_index in display_index_set:
            return selected_index
        if selected_index - 1 in display_index_set:
            return selected_index - 1
        return selected_index

    def _user_entry_average_formula_values(self, dfm, formula_rows, column_count):
        values = [[] for _ in formula_rows]
        for row_index, row in enumerate(formula_rows):
            if not row["is_user_entry"]:
                continue
            values[row_index] = [
                self._snapshot_value(self._average_ratio_value(dfm, development_index, row["api_index"]))
                for development_index in range(1, column_count + 1)
            ]
            break
        return values

    def _formula_metadata(self, name, is_user_entry):
        if is_user_entry:
            return {
                "average_type": "user_entry",
                "base": "simple",
                "periods": "all",
                "exclude": 0,
            }

        match = re.match(r"^(Simple|Volume) - (all|\d+)(?: Ex hi/lo)?$", name, re.IGNORECASE)
        if match:
            periods = match.group(2).lower()
            return {
                "average_type": "custom",
                "base": match.group(1).lower(),
                "periods": "all" if periods == "all" else int(periods),
                "exclude": 1 if "ex hi/lo" in name.lower() else 0,
            }

        return {
            "average_type": "custom",
            "base": self._formula_metadata_base(name),
            "periods": "all",
            "exclude": 0,
        }

    def _formula_metadata_base(self, name):
        base = re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")
        return base or "custom"

    def _average_ratio_value(self, dfm, development_index, api_index):
        try:
            return self._json_value(dfm.AverageRatioValues(development_index, api_index))
        except Exception:
            return None

    def _labels(self, dfm):
        origin_count = int(self._optional_value(dfm, "OriginCount", 0) or 0)
        development_count = self._development_column_count(dfm)
        origin_labels = self._indexed_values(dfm, ("OriginLabel", "OriginLabels"), origin_count)
        development_labels = self._indexed_values(
            dfm,
            ("DevelopmentLabel", "DevelopmentLabels", "DevLabel", "DevLabels"),
            development_count,
        )
        return origin_labels, development_labels

    def _ratio_development_labels(self, data_development_labels):
        if len(data_development_labels) < 2:
            return data_development_labels

        parsed = [self._development_label_number(label) for label in data_development_labels]
        if any(value is None for value in parsed):
            return data_development_labels

        labels = [
            f"({index}) {parsed[index - 1]}-{parsed[index]}"
            for index in range(1, len(parsed))
        ]
        labels.append(f"{parsed[-1]} - Ult")
        return labels

    def _development_label_number(self, label):
        if isinstance(label, (int, float)) and not isinstance(label, bool):
            return int(label)
        match = re.match(r"^\s*(\d+)", str(label))
        if not match:
            return None
        return int(match.group(1))

    def _indexed_values(self, obj, attr_names, count):
        for attr_name in attr_names:
            values = []
            for index in range(1, count + 1):
                try:
                    attr = getattr(obj, attr_name)
                    value = attr(index) if callable(attr) else attr[index - 1]
                    values.append(self._clean_label(value))
                except Exception:
                    values = []
                    break
            if values:
                return values
        return []

    def _development_column_count(self, dfm):
        rows = int(dfm.OriginCount)
        if rows <= 0:
            return 0
        return max(int(dfm.DevelopmentCount(origin_index)) for origin_index in range(1, rows + 1))

    def _optional_value(self, obj, attr_name, default):
        try:
            value = getattr(obj, attr_name)
            if callable(value):
                value = value()
            return value
        except Exception:
            return default

    def _nested_name(self, obj, attr_name):
        try:
            value = getattr(obj, attr_name)
            return self._clean_label(value.Name)
        except Exception:
            return ""

    def _nested_value(self, obj, attr_name, nested_attr_name, default):
        try:
            value = getattr(obj, attr_name)
            nested_value = getattr(value, nested_attr_name)
            if callable(nested_value):
                nested_value = nested_value()
            return nested_value
        except Exception:
            return default

    def _output_vector_modified(self, method):
        """The ``Modified`` ResQ stamped on a method's output vector, as an instant.

        COM hands the DATE back as a datetime labelled UTC although ResQ keeps
        a local wall-clock reading, so the label is dropped and the value is
        converted from this machine's zone. The app server persists what comes
        back through ``arcrho_api.timestamps``, and the next sync review then
        compares two absolute times rather than a clock reading with an instant.
        """
        try:
            modified = method.OutputVector.Modified
        except Exception:
            return ""
        if hasattr(modified, "replace") and hasattr(modified, "isoformat"):
            try:
                return modified.replace(tzinfo=None).astimezone(timezone.utc).isoformat()
            except Exception:
                pass
        normalized = self._json_value(modified)
        if self._has_json_value(normalized):
            return normalized
        return ""

    def _has_json_value(self, value):
        if value is None or isinstance(value, bool):
            return False
        if isinstance(value, str):
            return bool(value.strip())
        if isinstance(value, (int, float)):
            return value > 0
        return True

    def _json_value(self, value):
        if hasattr(value, "isoformat"):
            return value.isoformat()
        if value is None or isinstance(value, (str, int, float, bool)):
            return value
        try:
            return float(value)
        except Exception:
            pass
        return str(value)

    def _snapshot_value(self, value):
        value = self._json_value(value)
        if isinstance(value, bool) or value is None:
            return value
        if isinstance(value, (int, float)):
            return round(value, 4)
        return value

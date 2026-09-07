"""Upgrade a pre-v4 persisted ArcRho JSON payload to the v4 contract.

This is the one place the old spellings are still known. Every current reader
expects v4 and refuses anything else, so a workspace written before v4 must be
converted in place (``tools/migrate_persisted_json_v4.py``) before a v4 build
opens it. The functions here produce the *parsed* v4 payload for each file
kind; the caller normalizes it through the canonical contract module for that
kind and writes it with ``arcrho_api.io.persisted_json_text``, so a converted
file is by construction identical to what a fresh v4 save writes.

Rules applied (``docs/plans/completed/persisted_json_contract_v4.md``):

1. ``snake_case`` keys at every depth.
2. Every file carries a ``json_format`` stamped ``-v4``.
3. Timestamps are ISO-8601 UTC at millisecond precision with a ``Z``. A
   value without a zone in an old file is a wall-clock reading -- ResQ's
   ``Modified`` copied by the migration or the bridge -- in the zone of the
   machine that wrote it, so run the conversion on the Server PC, where
   ``arcrho_api.timestamps`` reads it as local time.
4. ``audit_log`` is last and follows the one audit policy.
5. Always-empty placeholder sections are deleted; a ``notes_tab`` with text is
   reported so its text can move to the dataset sidecar, never dropped.
6. Forced copies, paired code fields, ``origin_count``, ``user``, ``formula``
   and ``processing_by_csv`` are dropped (the DFM ratio labels stay).
7. Dependency entries carry a name and nothing machine-local.
"""

from __future__ import annotations

import re
from copy import deepcopy
from typing import Any, Mapping

from .bootstrap_contract import BST_JSON_FORMAT
from .bornhuetter_ferguson_contract import BF_JSON_FORMAT
from .cape_cod_contract import CC_JSON_FORMAT
from .dataset_index_contract import BS_CRA_JSON_FORMAT, BS_SR_JSON_FORMAT, RS_JSON_FORMAT
from .dfm_contract import DFM_JSON_FORMAT
from .sidecar_audit_contract import (
    AUDIT_ACTION_UPDATE,
    PROJECT_AUDIT_LOG_MAX_ENTRIES,
    normalize_audit_action,
    normalize_audit_log,
)
from .revision_contract import FINGERPRINT_HEX_LENGTH, FINGERPRINT_PREFIX
from .sidecar_core_contract import (
    dependency_entries,
    finalize_sidecar,
    stored_length_fields_from_display,
)
from .source_table_contract import SOURCE_IMPORT_JSON_FORMAT
from .timestamps import normalize_persisted_timestamp


PROJECT_AUDIT_LOG_JSON_FORMAT = "arcrho-project-audit-log-v4"
RUNTIME_CACHE_PROVENANCE_JSON_FORMAT = "arcrho-runtime-cache-provenance-v4"
DATASET_NUMBER_FORMATS_JSON_FORMAT = "arcrho-dataset-number-formats-v4"

# Old stamp -> v4 stamp, for every method kind.
METHOD_FORMAT_UPGRADES: dict[str, str] = {
    "arcrho-dfm-method-by-tab-v2": DFM_JSON_FORMAT,
    "arcrho-bornhuetter-ferguson-method-by-tab-v3": BF_JSON_FORMAT,
    "arcrho-cape-cod-method-by-tab-v1": CC_JSON_FORMAT,
    "arcrho-bootstrap-method-by-tab-v1": BST_JSON_FORMAT,
    "arcrho-result-selection-method-by-tab-v2": RS_JSON_FORMAT,
    "arcrho-berquist-sherman-sr-method-by-tab-v1": BS_SR_JSON_FORMAT,
    "arcrho-berquist-sherman-cra-method-by-tab-v1": BS_CRA_JSON_FORMAT,
}
CURRENT_METHOD_FORMATS = frozenset(METHOD_FORMAT_UPGRADES.values())

# Method stamps the app refused to open long before v4, because a later
# version of that kind replaced them and the one-time upgrade path for them
# was retired. They are not converted -- there is no contract left that can
# read them -- but they are recognized, so a converter can report them by name
# and rescue the notes inside them instead of failing on an unknown file.
UNCONVERTIBLE_METHOD_FORMATS: frozenset[str] = frozenset({
    "arcrho-dfm-method-by-tab-v1",
    "arcrho-bornhuetter-ferguson-method-by-tab-v2",
})

TIMESTAMP_KEYS = frozenset({
    "last_modified", "data_refreshed", "created", "updated_at", "event_date",
    "source_modified", "timestamp",
})
PLACEHOLDER_SECTIONS = frozenset({
    "chart_tab", "audit_log_tab", "ultimates_tab", "ratios_tab", "validation_tab", "results_tab",
})
# Core sidecar fields an older file may simply not carry. Each value here is
# the one the canonical readers already apply when a file leaves the field out
# -- ``normalize_status(None)`` is Current, a dataset with no method type has
# none, a table with no subtotal flag shows none -- so writing it changes
# nothing a reader would have seen, and without it the shared validator
# refuses the converted sidecar. Measured on ``NJ_Annual_Prod_202605_Fake``
# 2026-08-23: 314 of 2,079 sidecars are missing at least one of these.
SIDECAR_CORE_DEFAULTS: dict[str, Any] = {
    "method_type": "",
    "status": 0,
    "show_subtotal": False,
}

# The first of these a sidecar carries is what ``notes`` goes in front of when
# a rescued note has to be added to a sidecar that had no notes field, so the
# field lands where every canonical builder writes it.
_NOTES_FOLLOWED_BY = frozenset({"origin_labels", "development_labels", "precedents", "dependents"})

# Every canonical sidecar builder writes these five after the graph and before
# ``audit_log`` (``dfm_contract.build_dfm_output_sidecar`` and its siblings).
# The converter has to place them explicitly rather than let insertion order
# decide: a field the old file left out is appended wherever it happens to be
# filled, which is both unstable across two conversions and -- worse -- a
# different order from the one the app writes. A sidecar converted in one
# order and re-saved by the app in another is not one shape, which is the
# whole point of v4.
_SIDECAR_FIELDS_AFTER_GRAPH = (
    "created",
    "updated_at",
    "modified_by",
    "status",
    "publication_revision",
)

# Persisted DFM fields that v4 no longer stores (re-derived on read).
DFM_DROPPED_PATHS = (
    ("results_tab", "ratio_basis_origin_labels"),
    ("data_tab", "input_data_triangle_mask"),
)


class PersistedJsonUpgradeError(ValueError):
    """Raised when a payload cannot be recognized as an ArcRho file kind."""


class UnsupportedMethodFormatError(PersistedJsonUpgradeError):
    """Raised for a method the app already refused to open before v4.

    Distinct from its parent so a converter can tell a file that was dead
    before the conversion from one it does not recognize at all. Rescue the
    notes with :func:`stranded_method_notes` and leave the file alone.
    """


def snake_key(key: Any) -> str:
    """``'ratio basis origin labels'`` -> ``'ratio_basis_origin_labels'``; ``averageType`` -> ``average_type``."""

    text = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", str(key))
    return re.sub(r"[\s\-]+", "_", text.strip()).lower()


def _snake_keys(value: Any, *, keep_under: frozenset[str] = frozenset()) -> Any:
    """Rename field names recursively; the data-keyed maps in *keep_under* keep their keys."""

    if isinstance(value, Mapping):
        out: dict[str, Any] = {}
        for key, child in value.items():
            new_key = snake_key(key)
            out[new_key] = deepcopy(child) if new_key in keep_under else _snake_keys(child, keep_under=keep_under)
        return out
    if isinstance(value, list):
        return [_snake_keys(item, keep_under=keep_under) for item in value]
    return value


_FULL_FINGERPRINT = re.compile(rf"^{FINGERPRINT_PREFIX}([0-9a-f]{{{FINGERPRINT_HEX_LENGTH}}})[0-9a-f]+$")


def _shorten_fingerprints(value: Any) -> Any:
    """Cut every stored ``sha256:`` value down to the length rule 2a keeps.

    An old file holds the full 64-character digest of the same canonical text,
    so its first sixteen characters are exactly what the one producer emits
    today. Both sides of a comparison have to shorten together: a stored
    full-length value never compares equal to a recomputed short one, so a
    ``config_hash`` left long would mark every cached table stale.
    """

    if isinstance(value, str):
        match = _FULL_FINGERPRINT.match(value)
        return FINGERPRINT_PREFIX + match.group(1) if match else value
    if isinstance(value, Mapping):
        return {key: _shorten_fingerprints(child) for key, child in value.items()}
    if isinstance(value, list):
        return [_shorten_fingerprints(item) for item in value]
    return value


def _upgrade_timestamps(value: Any) -> Any:
    if isinstance(value, Mapping):
        return {
            key: (normalize_persisted_timestamp(child, default=str(child or "")) if key in TIMESTAMP_KEYS and isinstance(child, str) else _upgrade_timestamps(child))
            for key, child in value.items()
        }
    if isinstance(value, list):
        return [_upgrade_timestamps(item) for item in value]
    return value


def stranded_method_notes(payload: Mapping[str, Any]) -> str:
    """Return the ``notes_tab`` text of any method file, whatever its stamp.

    v4 keeps no notes section in a method file, so this text has to reach the
    output dataset's sidecar before the file is converted or set aside. It is
    read without regard to the stamp on purpose: the only files in the
    workspace that carry real commentary here are ones the app already refuses
    to open, and their text is the only copy that exists.
    """

    if not isinstance(payload, Mapping):
        return ""
    notes_tab = payload.get("notes_tab")
    if not isinstance(notes_tab, Mapping):
        notes_tab = payload.get("notes tab")
    if not isinstance(notes_tab, Mapping):
        return ""
    return str(notes_tab.get("notes") or "").strip()


def upgrade_method(payload: Mapping[str, Any]) -> tuple[dict[str, Any], str]:
    """Return the v4 method payload and any ``notes_tab`` text that must move to the sidecar."""

    if not isinstance(payload, Mapping):
        raise PersistedJsonUpgradeError("A method file must hold a JSON object.")
    stamp = str(payload.get("json_format") or payload.get("json format") or "").strip()
    if stamp in UNCONVERTIBLE_METHOD_FORMATS:
        raise UnsupportedMethodFormatError(
            f"{stamp!r} was already refused before v4; rescue its notes and leave the file."
        )
    json_format = METHOD_FORMAT_UPGRADES.get(stamp, stamp if stamp in CURRENT_METHOD_FORMATS else "")
    if not json_format:
        raise PersistedJsonUpgradeError(f"Unknown method json_format: {stamp!r}.")
    notes_text = stranded_method_notes(payload)
    renamed = _snake_keys(payload, keep_under=frozenset({"ratio_main_table", "ratio_summary_table"}))
    renamed.pop("json_format", None)
    renamed.pop("notes_tab", None)
    for section in PLACEHOLDER_SECTIONS:
        if isinstance(renamed.get(section), Mapping) and not renamed[section]:
            renamed.pop(section)
    renamed.pop("audit_log", None)
    if json_format == DFM_JSON_FORMAT:
        for path in DFM_DROPPED_PATHS:
            parent = renamed.get(path[0])
            if isinstance(parent, dict):
                parent.pop(path[1], None)
    upgraded = {"json_format": json_format, **_shorten_fingerprints(_upgrade_timestamps(renamed))}
    return upgraded, notes_text


def upgrade_dataset_sidecar(
    payload: Mapping[str, Any],
    *,
    publication_revision: Any = None,
) -> dict[str, Any]:
    """Return the v4 dataset sidecar: lower-case graph keys, slim entries, retired fields gone.

    Pass *publication_revision* for a method-output sidecar, taking it from the
    converted method file. The stored copy cannot simply be shortened: the hash
    vocabulary stopped depending on the persisted key spelling, so the value a
    v4 method computes is a different number, and a sidecar left holding the
    old one would report every method as saved but never republished.
    """

    if not isinstance(payload, Mapping):
        raise PersistedJsonUpgradeError("A sidecar must hold a JSON object.")
    source = dict(payload)
    precedents = source.pop("Precedents", source.pop("precedents", None))
    dependents = source.pop("Dependents", source.pop("dependents", None))
    renamed = _snake_keys(source)
    # Fill the core *before* the graph fields go back on. A field the old file
    # left out is appended at the end, and so are the two graph fields, so
    # filling afterwards puts ``status`` behind ``precedents`` the first time a
    # file is converted and in front of it the second time. Converting is then
    # not a fixed point, and the file does not match what a fresh save writes.
    if not str(renamed.get("modified_by") or "").strip():
        renamed["modified_by"] = str(source.get("user") or "").strip()
    for field, default in SIDECAR_CORE_DEFAULTS.items():
        renamed.setdefault(field, default)
    renamed.update(stored_length_fields_from_display(renamed))
    renamed["precedents"] = dependency_entries([
        {"dataset_name": item.get("dataset_name") or item.get("dataset_type_name") or item.get("name"), "method_type": item.get("method_type")}
        if isinstance(item, Mapping) else item
        for item in (precedents or [])
    ])
    renamed["dependents"] = dependency_entries([
        {"dataset_name": item.get("dataset_name") or item.get("dataset_type_name") or item.get("name"), "method_type": item.get("method_type")}
        if isinstance(item, Mapping) else item
        for item in (dependents or [])
    ])
    for field in _SIDECAR_FIELDS_AFTER_GRAPH:
        if field in renamed:
            renamed[field] = renamed.pop(field)
    if str(renamed.get("method_name") or "").strip():
        # A sidecar that names the method which wrote it holds derived values
        # by definition, and every canonical builder says so. Two DFM outputs
        # in ``NJ_Annual_Prod_202605_Fake`` say ``false``, which no builder
        # can produce; the conversion is where that gets straightened out.
        renamed["calculated"] = True
    fresh_revision = str(publication_revision or "").strip()
    if fresh_revision:
        renamed["publication_revision"] = fresh_revision
    return finalize_sidecar(_shorten_fingerprints(_upgrade_timestamps(renamed)))


def sidecar_with_method_notes(payload: Mapping[str, Any], notes: Any) -> dict[str, Any]:
    """Return the sidecar carrying notes stranded in a pre-v4 method file (Trap 1).

    v4 keeps no ``notes_tab`` in a method file, and three of the four that
    carry one hold commentary written nowhere else. :func:`upgrade_method`
    reads the text out; the converter passes it here together with the sidecar
    of the method's own output dataset, which is where the app already shows a
    dataset's notes. Text already in the sidecar is kept and the incoming text
    follows it after a blank line, and text already present is not appended
    again, so converting the same workspace twice cannot duplicate a note.
    """

    merged = dict(payload)
    incoming = str(notes or "").strip()
    existing = str(merged.get("notes") or "").strip()
    if not incoming or incoming in existing:
        return finalize_sidecar(merged)
    text = "\n\n".join(part for part in (existing, incoming) if part)
    if "notes" in merged:
        merged["notes"] = text
        return finalize_sidecar(merged)
    # A sidecar with no notes yet gets the field where a canonical builder
    # writes it, ahead of the labels and the graph, rather than on the end.
    # Appending would put ``notes`` behind ``precedents`` the first time the
    # file is converted and in front of it the second, and converting would
    # not be a fixed point.
    rebuilt: dict[str, Any] = {}
    for key, value in merged.items():
        if key in _NOTES_FOLLOWED_BY and "notes" not in rebuilt:
            rebuilt["notes"] = text
        rebuilt[key] = value
    rebuilt.setdefault("notes", text)
    return finalize_sidecar(rebuilt)


def upgrade_project_audit_log(payload: Mapping[str, Any]) -> dict[str, Any]:
    """Return ``audit_log.json`` in the sidecar record shape (Decision 4)."""

    source = payload if isinstance(payload, Mapping) else {}
    entries = source.get("audit_log")
    if not isinstance(entries, list):
        entries = []
        for raw in source.get("entries") or []:
            if not isinstance(raw, Mapping):
                continue
            text = str(raw.get("action") or "").strip()
            known = normalize_audit_action(text)
            is_known = known in {"Insert", "Update", "Auto Refresh"}
            entries.append({
                "event_date": str(raw.get("timestamp") or raw.get("event_date") or ""),
                "action": known if is_known else AUDIT_ACTION_UPDATE,
                "change_info": "" if is_known else text,
                "user": str(raw.get("user") or ""),
            })
    return {
        "json_format": PROJECT_AUDIT_LOG_JSON_FORMAT,
        "project_name": str(source.get("project_name") or ""),
        "updated_at": normalize_persisted_timestamp(source.get("updated_at"), default=str(source.get("updated_at") or "")),
        "audit_log": normalize_audit_log(_upgrade_timestamps(entries), max_entries=PROJECT_AUDIT_LOG_MAX_ENTRIES),
    }


def upgrade_runtime_cache_provenance(payload: Mapping[str, Any]) -> dict[str, Any]:
    """``format`` -> ``json_format``, restamped, with the processing hash shortened.

    ``csv_fingerprint.sha256`` is deliberately left at full length: it is a
    digest of the cached file beside it, produced and compared by the runtime
    cache alone, not one of the persisted fingerprints rule 2a shortens.
    """

    source = dict(payload) if isinstance(payload, Mapping) else {}
    source.pop("format", None)
    source.pop("json_format", None)
    return {
        "json_format": RUNTIME_CACHE_PROVENANCE_JSON_FORMAT,
        **_shorten_fingerprints(_snake_keys(source)),
    }


def upgrade_dataset_number_formats(payload: Mapping[str, Any]) -> dict[str, Any]:
    source = dict(payload) if isinstance(payload, Mapping) else {}
    source.pop("json_format", None)
    return {"json_format": DATASET_NUMBER_FORMATS_JSON_FORMAT, **_upgrade_timestamps(_snake_keys(source))}


def upgrade_source_import(payload: Mapping[str, Any]) -> dict[str, Any]:
    source = dict(payload) if isinstance(payload, Mapping) else {}
    source.pop("version", None)
    source.pop("json_format", None)
    return {"json_format": SOURCE_IMPORT_JSON_FORMAT, **_snake_keys(source)}


__all__ = [
    "CURRENT_METHOD_FORMATS",
    "DATASET_NUMBER_FORMATS_JSON_FORMAT",
    "METHOD_FORMAT_UPGRADES",
    "PROJECT_AUDIT_LOG_JSON_FORMAT",
    "PersistedJsonUpgradeError",
    "RUNTIME_CACHE_PROVENANCE_JSON_FORMAT",
    "UNCONVERTIBLE_METHOD_FORMATS",
    "UnsupportedMethodFormatError",
    "sidecar_with_method_notes",
    "snake_key",
    "stranded_method_notes",
    "upgrade_dataset_number_formats",
    "upgrade_dataset_sidecar",
    "upgrade_method",
    "upgrade_project_audit_log",
    "upgrade_runtime_cache_provenance",
    "upgrade_source_import",
]

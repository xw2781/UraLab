"""The one schema every dataset sidecar shares, and the check that enforces it.

A dataset sidecar and a method-output sidecar are one schema, not two. A
method opened in a Dataset Viewer window shows only its output triangle or
vector, read from the sidecar and CSV like any other dataset, so both must
carry the same core: the ``json_format`` stamp first, identity, formatting,
status, the dependency graph, and the ``audit_log`` as the last field. A
method-output sidecar adds only ``method_name`` on top -- with ``calculated``
true and its own ``method_type`` / ``source_kind`` values -- and, where the
method computes one, ``publication_revision``; nothing in the core may differ
between the two kinds.

Every sidecar builder -- the engine contract and the four method-output
contracts -- runs its payload through :func:`validate_sidecar_core` before
returning it, and every app-server and public-API write funnel runs the
payload it is about to serialize through :func:`finalize_sidecar`, so the
invariants are enforced where the bytes are produced rather than remembered
at each call site.

The dependency graph is persisted location-independently: an entry is
``{"dataset_name": ...}`` plus ``method_type`` when the linked dataset is a
method output, and -- reserved for cross-class and cross-project links --
``reserving_class`` / ``project`` only when the linked dataset lives outside
the file's own class or project. No path, no modification time.

Axis labels and notes are not part of the required core on purpose: an engine
sidecar derives its labels from the project header when the CSV is loaded and
has no notes until someone writes them, so requiring either would make that
builder invent values. When present they are checked for shape.

Two period lengths, following ResQ's naming
-------------------------------------------

``origin_length`` / ``development_length`` -- ``period_length`` on a vector --
are the **display** shape: the months per period the dataset is shown and
saved at, the counterpart of ResQ's ``OriginLength`` family.
``stored_origin_length`` / ``stored_development_length`` --
``stored_period_length`` on a vector -- are the **stored** shape: the months
per period of the CSV named by ``csv_file``, the counterpart of ResQ's
``StoredOriginLength``. Every sidecar carries the stored pair for its format.

The display shape must be a whole multiple of the stored shape on each axis,
because any coarser view is a roll-up of the stored data computed when the
dataset is read. Such a view is never written back: the CSV stays at the
stored shape, and only that file is ever the dataset's data.
"""

from __future__ import annotations

import re
from typing import Any, Iterable, Mapping

from .sidecar_audit_contract import normalize_audit_log


class SidecarContractError(ValueError):
    """Raised when a sidecar payload does not satisfy the shared core."""


DATASET_SIDECAR_JSON_FORMAT = "arcrho-dataset-sidecar-v4"

SIDECAR_JSON_FORMAT_FIELD = "json_format"
SIDECAR_AUDIT_LOG_FIELD = "audit_log"
SIDECAR_PRECEDENTS_FIELD = "precedents"
SIDECAR_DEPENDENTS_FIELD = "dependents"

# Fields every sidecar carries, whatever produced it.
SIDECAR_CORE_FIELDS: tuple[str, ...] = (
    SIDECAR_JSON_FORMAT_FIELD,
    "dataset_name",
    "dataset_type",
    "reserving_class",
    "project_name",
    "source_kind",
    "calculated",
    "data_format",
    "method_type",
    "status",
    "number_format",
    "decimal_places",
    "show_subtotal",
    "csv_file",
    "created",
    "updated_at",
    "modified_by",
    SIDECAR_PRECEDENTS_FIELD,
    SIDECAR_DEPENDENTS_FIELD,
    SIDECAR_AUDIT_LOG_FIELD,
)

# Fields only a method-output sidecar adds on top of the core. ``method_name``
# is what marks the sidecar as one; ``method_type`` is core, carried by every
# sidecar, and reads ``None`` on a dataset that no method wrote.
METHOD_OUTPUT_SIDECAR_FIELDS: tuple[str, ...] = (
    "method_name",
)

# A method-output sidecar may also carry the publication fingerprint of the
# method that wrote it, which is how the app tells "saved but never
# republished". It is optional because Berquist Sherman publishes no revision:
# it has no contract module and computes none, so its 14 output sidecars in
# ``NJ_Annual_Prod_202605_Fake`` name a method and stop there. A revision
# without a method name is always wrong, and that is what is checked.
METHOD_OUTPUT_PUBLICATION_FIELD = "publication_revision"

# The display shape, by data format, and the stored shape beside it.
SIDECAR_DISPLAY_PERIOD_FIELD = "period_length"
SIDECAR_DISPLAY_ORIGIN_FIELD = "origin_length"
SIDECAR_DISPLAY_DEVELOPMENT_FIELD = "development_length"
SIDECAR_STORED_PERIOD_FIELD = "stored_period_length"
SIDECAR_STORED_ORIGIN_FIELD = "stored_origin_length"
SIDECAR_STORED_DEVELOPMENT_FIELD = "stored_development_length"
# The display a dataset's cell links were written against. A link names a cell
# of the grid that was on screen when it was entered, so the display can move
# on and the links keep pointing into this one.
SIDECAR_LINKED_PERIOD_FIELD = "linked_period_length"
SIDECAR_LINKED_ORIGIN_FIELD = "linked_origin_length"
SIDECAR_LINKED_DEVELOPMENT_FIELD = "linked_development_length"

# What a length reads as when a file states none, which is how every consumer
# already reads a missing one.
DEFAULT_PERIOD_MONTHS = 12

# stored field -> the display field it must divide, per data format.
_VECTOR_PERIOD_FIELDS = ((SIDECAR_STORED_PERIOD_FIELD, SIDECAR_DISPLAY_PERIOD_FIELD),)
_TRIANGLE_PERIOD_FIELDS = (
    (SIDECAR_STORED_ORIGIN_FIELD, SIDECAR_DISPLAY_ORIGIN_FIELD),
    (SIDECAR_STORED_DEVELOPMENT_FIELD, SIDECAR_DISPLAY_DEVELOPMENT_FIELD),
)

# Fields v4 removed because they restated another field, nothing read them,
# or they bound a shared file to one machine. A writer may not bring them back.
RETIRED_SIDECAR_FIELDS: frozenset[str] = frozenset({
    "method_type_code",
    "data_format_code",
    "origin_count",
    "user",
    "formula",
    "processing_by_csv",
    "Precedents",  # Title Case predecessors of the graph keys
    "Dependents",
    "path",
    "dependencies",
})

# Optional core fields that, when present, must have this shape.
_LIST_FIELDS = ("origin_labels", "development_labels")
_ENTRY_FIELDS = ("dataset_name", "method_type", "reserving_class", "project")
_SNAKE_CASE = re.compile(r"^[a-z0-9]+(?:_[a-z0-9]+)*$")


def _clean(value: Any) -> str:
    return str(value if value is not None else "").strip()


def is_vector_format(data_format: Any) -> bool:
    """Whether *data_format* names the vector shape rather than a triangle."""

    return _clean(data_format).casefold() == "vector"


def stored_length_fields(
    data_format: Any,
    origin_length: Any,
    development_length: Any = None,
) -> dict[str, int]:
    """The stored-shape fields a sidecar of *data_format* carries.

    A vector carries ``stored_period_length``; a triangle carries
    ``stored_origin_length`` and ``stored_development_length``. Every producer
    builds them here, so which pair a format takes is written once.

    *origin_length* and *development_length* are the months per period of the
    CSV the sidecar names -- not the shape it is displayed at, which is the
    same only while the two have not been allowed to come apart.
    """

    if is_vector_format(data_format):
        return {SIDECAR_STORED_PERIOD_FIELD: int(origin_length)}
    return {
        SIDECAR_STORED_ORIGIN_FIELD: int(origin_length),
        SIDECAR_STORED_DEVELOPMENT_FIELD: int(development_length),
    }


def stored_length_fields_from_display(payload: Mapping[str, Any]) -> dict[str, int]:
    """The stored-shape fields of a sidecar that records only a display shape.

    A file written before there were two shapes recorded one length per axis,
    and it was both what the dataset was displayed at and what its CSV held,
    so the single shape becomes the stored one. A file that recorded none at
    all is read the way every consumer already reads a missing length: as
    annual. Both the pre-v4 conversion and the one-time backfill of existing
    server projects fill the stored fields this way.
    """

    def months(field: str) -> int:
        value = payload.get(field)
        try:
            months_value = int(value)
        except (TypeError, ValueError):
            return DEFAULT_PERIOD_MONTHS
        return months_value if months_value > 0 else DEFAULT_PERIOD_MONTHS

    if is_vector_format(payload.get("data_format")):
        return stored_length_fields("Vector", months(SIDECAR_DISPLAY_PERIOD_FIELD))
    return stored_length_fields(
        "Triangle",
        months(SIDECAR_DISPLAY_ORIGIN_FIELD),
        months(SIDECAR_DISPLAY_DEVELOPMENT_FIELD),
    )


def stored_lengths(payload: Mapping[str, Any]) -> tuple[int, int]:
    """The ``(origin, development)`` months per period of *payload*'s own CSV.

    The read-side counterpart of :func:`stored_length_fields`: every reader
    that opens the file ``csv_file`` names asks here what shape it is at, so
    which field a format carries stays written in one place. A vector keeps
    one stored length and reports it on both axes; a value that is missing or
    unusable reads ``0``, which every caller treats as "not stated".
    """

    if is_vector_format(payload.get("data_format")):
        period = _stored_months(payload.get(SIDECAR_STORED_PERIOD_FIELD))
        return period, period
    return (
        _stored_months(payload.get(SIDECAR_STORED_ORIGIN_FIELD)),
        _stored_months(payload.get(SIDECAR_STORED_DEVELOPMENT_FIELD)),
    )


def display_lengths(payload: Mapping[str, Any]) -> tuple[int, int]:
    """The ``(origin, development)`` months per period *payload* is shown at.

    The display-side twin of :func:`stored_lengths`: a reader that opens a
    dataset at the shape it was saved at asks here, so which field a format
    carries stays written in one place. A vector reports its period on both
    axes; a value that is missing or unusable reads ``0``.
    """

    if is_vector_format(payload.get("data_format")):
        period = _stored_months(payload.get(SIDECAR_DISPLAY_PERIOD_FIELD))
        return period, period
    return (
        _stored_months(payload.get(SIDECAR_DISPLAY_ORIGIN_FIELD)),
        _stored_months(payload.get(SIDECAR_DISPLAY_DEVELOPMENT_FIELD)),
    )


def linked_length_fields(
    data_format: Any,
    origin_length: Any,
    development_length: Any = None,
) -> dict[str, int]:
    """The linked-shape fields a sidecar of *data_format* carries.

    The same pair as :func:`stored_length_fields`, naming the display the
    dataset's cell links were written against. Only a sidecar that holds a
    link carries them.
    """

    if is_vector_format(data_format):
        return {SIDECAR_LINKED_PERIOD_FIELD: int(origin_length)}
    return {
        SIDECAR_LINKED_ORIGIN_FIELD: int(origin_length),
        SIDECAR_LINKED_DEVELOPMENT_FIELD: int(development_length),
    }


def linked_lengths(payload: Mapping[str, Any]) -> tuple[int, int]:
    """The ``(origin, development)`` months per period *payload*'s links name.

    A sidecar that states no linked shape was saved with its links at the
    display it records, so that pair answers for it.
    """

    if is_vector_format(payload.get("data_format")):
        period = _stored_months(payload.get(SIDECAR_LINKED_PERIOD_FIELD))
        return (period, period) if period else display_lengths(payload)
    origin = _stored_months(payload.get(SIDECAR_LINKED_ORIGIN_FIELD))
    development = _stored_months(payload.get(SIDECAR_LINKED_DEVELOPMENT_FIELD))
    return (origin, development) if origin and development else display_lengths(payload)


def _stored_months(value: Any) -> int:
    try:
        months = int(value)
    except (TypeError, ValueError):
        return 0
    return months if months > 0 else 0


def dependency_names(entries: Any) -> list[str]:
    """The unique dataset names of persisted dependency entries, in order."""

    names: list[str] = []
    seen: set[str] = set()
    items = entries if isinstance(entries, Iterable) and not isinstance(entries, (str, bytes, Mapping)) else ()
    for item in items:
        name = _clean(item.get("dataset_name")) if isinstance(item, Mapping) else _clean(item)
        key = name.casefold()
        if name and key not in seen:
            seen.add(key)
            names.append(name)
    return names


def dependency_entries(
    entries: Any,
    *,
    method_types: Mapping[str, Any] | None = None,
) -> list[dict[str, str]]:
    """Normalize names or entries to the one persisted dependency-entry shape.

    *entries* may hold plain names or entry mappings. ``method_type`` is kept
    from an entry that carries one and may be supplied through *method_types*
    (name -> method type); ``reserving_class`` / ``project`` are kept only
    when present and non-empty, which is what makes a same-class link small.
    """

    lookup = {
        _clean(name).casefold(): _clean(value)
        for name, value in (method_types or {}).items()
        if _clean(name)
    }
    out: list[dict[str, str]] = []
    seen: set[str] = set()
    items = entries if isinstance(entries, Iterable) and not isinstance(entries, (str, bytes, Mapping)) else ()
    for item in items:
        source = item if isinstance(item, Mapping) else {"dataset_name": item}
        name = _clean(source.get("dataset_name"))
        key = name.casefold()
        if not name or key in seen:
            continue
        seen.add(key)
        entry: dict[str, str] = {"dataset_name": name}
        method_type = _clean(source.get("method_type"))
        if method_type.casefold() == "none":
            method_type = ""
        method_type = method_type or lookup.get(key, "")
        if method_type and method_type.casefold() != "none":
            entry["method_type"] = method_type
        for scope in ("reserving_class", "project"):
            value = _clean(source.get(scope))
            if value:
                entry[scope] = value
        out.append(entry)
    return out


def finalize_sidecar(payload: Mapping[str, Any]) -> dict[str, Any]:
    """Return *payload* stamped ``json_format`` first and ``audit_log`` last.

    This is the projection every write funnel applies. Key order otherwise
    follows the payload, so a writer that already builds the canonical order
    is untouched and one that merged an older file gets the stamp and the log
    moved to where the contract keeps them.

    The two graph fields are normalized to the one dependency-entry shape on
    the way through, so a writer that assembles its payload by hand -- and so
    never runs a builder's :func:`validate_sidecar_core` -- cannot land bare
    names in ``precedents`` while the far side of the same link carries
    entries in ``dependents``.
    """

    ordered: dict[str, Any] = {SIDECAR_JSON_FORMAT_FIELD: DATASET_SIDECAR_JSON_FORMAT}
    for key, value in payload.items():
        if key in (SIDECAR_JSON_FORMAT_FIELD, SIDECAR_AUDIT_LOG_FIELD) or key in RETIRED_SIDECAR_FIELDS:
            continue
        if key in (SIDECAR_PRECEDENTS_FIELD, SIDECAR_DEPENDENTS_FIELD):
            value = dependency_entries(value)
        ordered[key] = value
    ordered[SIDECAR_AUDIT_LOG_FIELD] = normalize_audit_log(payload.get(SIDECAR_AUDIT_LOG_FIELD))
    return ordered


# Kept for callers that only need the ordering half of finalize_sidecar.
with_audit_log_last = finalize_sidecar


def _period_months(payload: Mapping[str, Any], field: str) -> int:
    value = payload.get(field)
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise SidecarContractError(
            f"Sidecar {field} must be a positive whole number of months; found {value!r}."
        )
    return value


def validate_period_lengths(payload: Mapping[str, Any]) -> None:
    """Assert the stored shape is present and the display shape a multiple of it.

    Part of :func:`validate_sidecar_core`, and separately callable so the
    one-time backfill of existing server projects can hold a file to this rule
    alone without also holding it to core fields another conversion fills in.
    """

    pairs = (
        _VECTOR_PERIOD_FIELDS
        if is_vector_format(payload.get("data_format"))
        else _TRIANGLE_PERIOD_FIELDS
    )
    missing = [stored for stored, _display in pairs if stored not in payload]
    if missing:
        raise SidecarContractError(
            "Sidecar is missing the stored period length: " + ", ".join(missing)
        )
    for stored, display in pairs:
        stored_months = _period_months(payload, stored)
        if display not in payload:
            continue
        display_months = _period_months(payload, display)
        if display_months % stored_months:
            raise SidecarContractError(
                f"Sidecar {display} ({display_months}) must be a whole multiple of "
                f"{stored} ({stored_months})."
            )


def validate_sidecar_core(payload: Mapping[str, Any]) -> dict[str, Any]:
    """Assert the shared core of a complete sidecar payload and return it.

    Raises :class:`SidecarContractError` naming the first violation. The
    payload is returned unchanged so a builder can ``return
    validate_sidecar_core(payload)``.
    """

    if not isinstance(payload, Mapping):
        raise SidecarContractError("A sidecar payload must be a JSON object.")
    missing = [field for field in SIDECAR_CORE_FIELDS if field not in payload]
    if missing:
        raise SidecarContractError("Sidecar is missing core fields: " + ", ".join(missing))
    keys = list(payload.keys())
    if keys[0] != SIDECAR_JSON_FORMAT_FIELD:
        raise SidecarContractError(f"Sidecar {SIDECAR_JSON_FORMAT_FIELD} must be the first field; found {keys[0]!r}.")
    if payload[SIDECAR_JSON_FORMAT_FIELD] != DATASET_SIDECAR_JSON_FORMAT:
        raise SidecarContractError(f"Sidecar json_format must be {DATASET_SIDECAR_JSON_FORMAT!r}.")
    if keys[-1] != SIDECAR_AUDIT_LOG_FIELD:
        raise SidecarContractError(
            f"Sidecar {SIDECAR_AUDIT_LOG_FIELD} must be the last field; found {keys[-1]!r}."
        )
    retired = [key for key in keys if key in RETIRED_SIDECAR_FIELDS]
    if retired:
        raise SidecarContractError("Sidecar carries retired fields: " + ", ".join(retired))
    bad_keys = [key for key in keys if not _SNAKE_CASE.match(key)]
    if bad_keys:
        raise SidecarContractError("Sidecar keys must be snake_case: " + ", ".join(bad_keys))
    for field in (SIDECAR_PRECEDENTS_FIELD, SIDECAR_DEPENDENTS_FIELD, SIDECAR_AUDIT_LOG_FIELD, *_LIST_FIELDS):
        if field in payload and not isinstance(payload[field], list):
            raise SidecarContractError(f"Sidecar {field} must be a list.")
    for field in (SIDECAR_PRECEDENTS_FIELD, SIDECAR_DEPENDENTS_FIELD):
        for entry in payload[field]:
            if not isinstance(entry, Mapping) or not _clean(entry.get("dataset_name")):
                raise SidecarContractError(f"Sidecar {field} entries must name a dataset.")
            unknown = [key for key in entry if key not in _ENTRY_FIELDS]
            if unknown:
                raise SidecarContractError(
                    f"Sidecar {field} entries may only carry {', '.join(_ENTRY_FIELDS)}; found {', '.join(unknown)}."
                )
    audit_log = payload[SIDECAR_AUDIT_LOG_FIELD]
    if normalize_audit_log(audit_log) != audit_log:
        raise SidecarContractError("Sidecar audit_log is not in the canonical policy form.")
    if not isinstance(payload["calculated"], bool):
        raise SidecarContractError("Sidecar calculated must be a boolean.")
    validate_period_lengths(payload)
    is_method_output = bool(_clean(payload.get("method_name")))
    if is_method_output and payload["calculated"] is not True:
        raise SidecarContractError("A method-output sidecar is always calculated.")
    if _clean(payload.get(METHOD_OUTPUT_PUBLICATION_FIELD)) and not is_method_output:
        raise SidecarContractError(
            f"Only a method-output sidecar carries {METHOD_OUTPUT_PUBLICATION_FIELD}."
        )
    return dict(payload)


__all__ = [
    "DATASET_SIDECAR_JSON_FORMAT",
    "DEFAULT_PERIOD_MONTHS",
    "METHOD_OUTPUT_PUBLICATION_FIELD",
    "METHOD_OUTPUT_SIDECAR_FIELDS",
    "RETIRED_SIDECAR_FIELDS",
    "SIDECAR_AUDIT_LOG_FIELD",
    "SIDECAR_CORE_FIELDS",
    "SIDECAR_DEPENDENTS_FIELD",
    "SIDECAR_DISPLAY_DEVELOPMENT_FIELD",
    "SIDECAR_DISPLAY_ORIGIN_FIELD",
    "SIDECAR_DISPLAY_PERIOD_FIELD",
    "SIDECAR_JSON_FORMAT_FIELD",
    "SIDECAR_LINKED_DEVELOPMENT_FIELD",
    "SIDECAR_LINKED_ORIGIN_FIELD",
    "SIDECAR_LINKED_PERIOD_FIELD",
    "SIDECAR_PRECEDENTS_FIELD",
    "SIDECAR_STORED_DEVELOPMENT_FIELD",
    "SIDECAR_STORED_ORIGIN_FIELD",
    "SIDECAR_STORED_PERIOD_FIELD",
    "SidecarContractError",
    "dependency_entries",
    "dependency_names",
    "finalize_sidecar",
    "is_vector_format",
    "linked_length_fields",
    "linked_lengths",
    "stored_length_fields",
    "stored_length_fields_from_display",
    "stored_lengths",
    "validate_period_lengths",
    "validate_sidecar_core",
    "with_audit_log_last",
]

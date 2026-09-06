"""Timestamp planning and durable baselines for ArcRho/ResQ synchronization.

This module owns the location-independent comparison contract used by the
``Sync Reserving Class with ResQ`` macro.  It deliberately does not know how
either side is inventoried or written; callers provide normalized inventory
items with parsed timestamps and use the resulting action plan to delegate to
the existing canonical import/export writers.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Mapping

from arcrho_api.io import persisted_json_text


SYNC_STATE_VERSION = 1
SYNC_RUNTIME_API_VERSION = 1
SYNC_STATE_DIR = Path("sync") / "resq"
TIMESTAMP_TOLERANCE_SECONDS = 0.000001

ACTION_ARCRHO_TO_RESQ = "arcrho_to_resq"
ACTION_RESQ_TO_ARCRHO = "resq_to_arcrho"

# The two ways a whole reserving class is moved by the Import and Export
# macros, as opposed to the per-row action the Sync macro chooses. One
# vocabulary serves the review table, the queue requests, and the saved
# selection document.
DIRECTION_IMPORT = "import"
DIRECTION_EXPORT = "export"
TRANSFER_DIRECTIONS = (DIRECTION_IMPORT, DIRECTION_EXPORT)
_DIRECTION_ACTIONS = {
    DIRECTION_IMPORT: ACTION_RESQ_TO_ARCRHO,
    DIRECTION_EXPORT: ACTION_ARCRHO_TO_RESQ,
}

# Which sides carry an edit made since the recorded baseline. A blank answer is
# not "nothing changed": it means no usable baseline exists to measure from.
CHANGED_NEITHER = "none"
CHANGED_ARCRHO = "arcrho"
CHANGED_RESQ = "resq"
CHANGED_BOTH = "both"

_SPACE_RE = re.compile(r"\s+")


def clean_name(value: Any) -> str:
    """Return the whitespace-normalized display form used on both sides."""

    return _SPACE_RE.sub(" ", str(value or "").strip())


def logical_key(value: Any) -> str:
    """Return the case-insensitive logical identity used to pair artifacts."""

    return clean_name(value).casefold()


def parse_timestamp(value: Any) -> float | None:
    """Parse ArcRho absolute times and ResQ local wall-clock times canonically."""

    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return _timestamp(value)
    raw = str(value).strip()
    if not raw:
        return None
    try:
        return _timestamp(float(raw))
    except ValueError:
        pass
    normalized = raw[:-1] + "+00:00" if raw.endswith("Z") else raw
    try:
        return _timestamp(datetime.fromisoformat(normalized).timestamp())
    except ValueError:
        return None


def _timestamp(value: Any) -> float | None:
    try:
        parsed = float(value)
    except (TypeError, ValueError):
        return None
    return parsed if parsed > 0 else None


def _timestamps_equal(left: Any, right: Any) -> bool:
    left_value = _timestamp(left)
    right_value = _timestamp(right)
    if left_value is None or right_value is None:
        return False
    return abs(left_value - right_value) <= TIMESTAMP_TOLERANCE_SECONDS


def _changed_from_baseline(item: Mapping[str, Any] | None, baseline: Mapping[str, Any], prefix: str) -> bool | None:
    present = item is not None
    if present != bool(baseline.get(f"{prefix}_present")):
        return True
    if not present:
        return False
    current = _timestamp(item.get("modified_timestamp"))
    previous = _timestamp(baseline.get(f"{prefix}_timestamp"))
    if current is None or previous is None:
        return None
    return not _timestamps_equal(current, previous)


def _row_id(key: str) -> str:
    return hashlib.sha256(key.encode("utf-8")).hexdigest()[:20]


def _state_signature(value: Mapping[str, Any] | None) -> dict[str, Any]:
    source = value if isinstance(value, Mapping) else {}
    return {
        "present": bool(source),
        "arcrho_present": bool(source.get("arcrho_present")),
        "resq_present": bool(source.get("resq_present")),
        "arcrho_timestamp": _timestamp(source.get("arcrho_timestamp")),
        "resq_timestamp": _timestamp(source.get("resq_timestamp")),
        "synced_at": str(source.get("synced_at") or ""),
    }


def _group_inventory(items: Iterable[Mapping[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    for raw in items:
        item = dict(raw)
        key = logical_key(item.get("name"))
        if key:
            grouped.setdefault(key, []).append(item)
    return grouped


def _support_for_action(
    action: str,
    arcrho: Mapping[str, Any] | None,
    resq: Mapping[str, Any] | None,
) -> tuple[bool, str]:
    if action == ACTION_ARCRHO_TO_RESQ:
        source = arcrho or {}
        target = resq or {}
        if not bool(source.get("can_export_to_resq", False)):
            return False, str(source.get("export_block_reason") or "ArcRho cannot export this item to ResQ.")
        if target and not bool(target.get("can_receive_from_arcrho", True)):
            return False, str(target.get("receive_block_reason") or "The ResQ item cannot be overwritten.")
        return True, ""
    if action == ACTION_RESQ_TO_ARCRHO:
        source = resq or {}
        if not bool(source.get("can_import_to_arcrho", False)):
            return False, str(source.get("import_block_reason") or "ArcRho cannot import this ResQ item.")
        return True, ""
    return False, "No synchronization action is available."


def newer_side(arcrho: Mapping[str, Any], resq: Mapping[str, Any]) -> str:
    """Which side of one item was modified last: ``arcrho``, ``resq``, or ``""``.

    Blank when either timestamp is unknown or the two match within tolerance.
    The Export macro uses this to warn about a ResQ copy the export would
    overwrite, so it compares the same epoch seconds the plan compares.
    """

    local_time = _timestamp(arcrho.get("modified_timestamp"))
    remote_time = _timestamp(resq.get("modified_timestamp"))
    if local_time is None or remote_time is None or _timestamps_equal(local_time, remote_time):
        return ""
    return "arcrho" if local_time > remote_time else "resq"


def export_supported(arcrho: Mapping[str, Any] | None, resq: Mapping[str, Any] | None) -> bool:
    """Whether an ArcRho-to-ResQ push of this item would write anything."""

    return _support_for_action(ACTION_ARCRHO_TO_RESQ, arcrho, resq)[0]


def transfer_direction(value: Any) -> str:
    """Return one of the two whole-class transfer directions, or raise."""

    direction = str(value or "").strip().casefold()
    if direction not in TRANSFER_DIRECTIONS:
        raise ValueError("Direction must be one of: " + ", ".join(TRANSFER_DIRECTIONS) + ".")
    return direction


def transfer_support(
    direction: Any,
    arcrho: Mapping[str, Any] | None,
    resq: Mapping[str, Any] | None,
) -> tuple[bool, str]:
    """Whether one item can move in a transfer direction, and why not.

    The review table shows every item either side holds, so an item the other
    side does not have at all reaches this with one side missing. An import
    can create it; an export cannot, because ResQ objects are written, never
    created.
    """

    normalized = transfer_direction(direction)
    if normalized == DIRECTION_EXPORT and not arcrho:
        return False, "ArcRho has no copy of this item to export."
    if normalized == DIRECTION_IMPORT and not resq:
        return False, "ResQ has no copy of this item to import."
    if normalized == DIRECTION_EXPORT and not resq:
        return False, "ResQ has no matching dataset or method to overwrite."
    return _support_for_action(_DIRECTION_ACTIONS[normalized], arcrho, resq)


def changed_since_baseline(
    arcrho: Mapping[str, Any] | None,
    resq: Mapping[str, Any] | None,
    baseline: Mapping[str, Any] | None,
) -> str:
    """Which sides were edited since the timestamp pair the last run recorded.

    ``CHANGED_ARCRHO``, ``CHANGED_RESQ``, ``CHANGED_BOTH`` or
    ``CHANGED_NEITHER`` against a usable baseline, and ``""`` when no usable
    baseline exists -- the pair was never recorded, or it is incomplete. A
    blank answer means the question cannot be answered from a baseline at all,
    which is why callers must fall back to comparing the two timestamps rather
    than reading it as "nothing changed".

    This is what separates a real ResQ edit from a ResQ timestamp that is only
    newer because the last export stamped it.

    A pair whose two timestamps match is ``CHANGED_NEITHER`` whatever the
    baseline says: only a copy of one side over the other leaves both with
    the same stamp, and an import records no baseline, so the pair the last
    export saved would otherwise call that copy an edit on both sides.
    """

    if not isinstance(baseline, Mapping) or not baseline:
        return ""
    if "present" in baseline and not baseline.get("present"):
        return ""
    local_changed = _changed_from_baseline(arcrho or None, baseline, "arcrho")
    remote_changed = _changed_from_baseline(resq or None, baseline, "resq")
    if local_changed is None or remote_changed is None:
        return ""
    if arcrho and resq and _timestamps_equal(
        arcrho.get("modified_timestamp"), resq.get("modified_timestamp")
    ):
        return CHANGED_NEITHER
    if local_changed and remote_changed:
        return CHANGED_BOTH
    if local_changed:
        return CHANGED_ARCRHO
    if remote_changed:
        return CHANGED_RESQ
    return CHANGED_NEITHER


_EXPORT_REVIEW_TEXT = {
    CHANGED_BOTH: (
        "Both changed",
        "Both sides changed since the last export; the ArcRho copy overwrites the ResQ change.",
    ),
    CHANGED_RESQ: (
        "ResQ changed",
        "Only ResQ changed since the last export; the ArcRho copy overwrites that change.",
    ),
    CHANGED_ARCRHO: ("ArcRho changed", "Only ArcRho changed since the last export."),
    CHANGED_NEITHER: ("Synchronized", "Neither side has changed since the two were last synchronized."),
}


def export_review(
    arcrho: Mapping[str, Any] | None,
    resq: Mapping[str, Any] | None,
    baseline: Mapping[str, Any] | None,
) -> dict[str, Any]:
    """How one item reads in an ArcRho-to-ResQ export review.

    The export only ever pushes ArcRho over ResQ, so an item is described by
    what changed since the recorded baseline rather than by the direction the
    Sync macro would choose for the whole reserving class. ``overwrites_edit``
    is the one verdict the export warns on: a ResQ change this push would
    destroy. Until a baseline exists there is nothing to measure against, so
    the raw timestamp comparison stands in and says so.
    """

    supported, block_reason = _support_for_action(ACTION_ARCRHO_TO_RESQ, arcrho, resq)
    changed = changed_since_baseline(arcrho, resq, baseline)
    if changed:
        status, detail = _EXPORT_REVIEW_TEXT[changed]
        overwrites_edit = changed in (CHANGED_RESQ, CHANGED_BOTH)
    else:
        side = newer_side(arcrho or {}, resq or {})
        overwrites_edit = side == "resq"
        if side:
            label = "ResQ" if side == "resq" else "ArcRho"
            status = f"{label} newer"
            detail = f"No baseline is recorded yet; {label} has the newer timestamp."
        elif _timestamp((arcrho or {}).get("modified_timestamp")) is None or _timestamp(
            (resq or {}).get("modified_timestamp")
        ) is None:
            status = "Unknown timestamp"
            detail = "No baseline is recorded yet and one side has no usable timestamp."
        else:
            status = "Same timestamp"
            detail = "No baseline is recorded yet; the two timestamps match."
    if not supported:
        detail = f"{detail} {block_reason}".strip()
    else:
        scope_note = clean_name((arcrho or {}).get("export_scope_note"))
        if scope_note:
            detail = f"{detail} {scope_note}".strip()
    return {
        "changed": changed,
        "status": status,
        "detail": detail,
        "supported": supported,
        "overwrites_edit": overwrites_edit,
    }


def plan_direction(rows: Iterable[Mapping[str, Any]]) -> dict[str, Any]:
    """Decide the one direction a whole reserving class is pushed in.

    Each side's timestamp is the latest modified timestamp among the items the
    review shows, and the newer side is the source for every row. Matching or
    unknown timestamps give no direction, so nothing is pushed.
    """

    latest: dict[str, float | None] = {"arcrho": None, "resq": None}
    for row in rows:
        for side in ("arcrho", "resq"):
            item = row.get(side)
            moment = _timestamp(item.get("modified_timestamp")) if isinstance(item, Mapping) else None
            if moment is not None and (latest[side] is None or moment > latest[side]):
                latest[side] = moment
    local_time = latest["arcrho"]
    remote_time = latest["resq"]
    direction = ""
    if local_time is not None and remote_time is not None and not _timestamps_equal(local_time, remote_time):
        direction = ACTION_ARCRHO_TO_RESQ if local_time > remote_time else ACTION_RESQ_TO_ARCRHO
    return {"direction": direction, "arcrho_timestamp": local_time, "resq_timestamp": remote_time}


def _comparison_action(
    arcrho: Mapping[str, Any],
    resq: Mapping[str, Any],
    baseline: Mapping[str, Any] | None,
    direction: str,
) -> tuple[str, str, str, bool]:
    """Return action, status, detail, and whether the row is marked for review.

    ``direction`` is the reserving class's direction. A row is pushed that way
    whenever it changed on either side; it is marked for review when the side
    being written over is the one that changed, because that push overwrites
    an edit rather than delivering one. Review never blocks the push.
    """

    local_time = _timestamp(arcrho.get("modified_timestamp"))
    remote_time = _timestamp(resq.get("modified_timestamp"))
    if local_time is None or remote_time is None:
        missing = []
        if local_time is None:
            missing.append("ArcRho")
        if remote_time is None:
            missing.append("ResQ")
        return "", "Unknown timestamp", f"{', '.join(missing)} timestamp is unavailable; the row is left alone.", False

    if baseline:
        local_changed = _changed_from_baseline(arcrho, baseline, "arcrho")
        remote_changed = _changed_from_baseline(resq, baseline, "resq")
        if local_changed is None or remote_changed is None:
            # An incomplete legacy/invalid baseline must never silently decide.
            return "", "Unknown baseline", "The saved synchronization baseline is incomplete; the row is left alone.", False
        if not local_changed and not remote_changed:
            return "", "Synchronized", "Neither side changed since the last accepted synchronization.", False
        status = "Both changed" if local_changed and remote_changed else ("ArcRho changed" if local_changed else "ResQ changed")
        detail = f"{status} since the last synchronization."
    else:
        if _timestamps_equal(local_time, remote_time):
            return "", "Same timestamp", "The timestamps match; content equality was not assumed.", False
        local_changed = local_time > remote_time
        remote_changed = not local_changed
        status = "ArcRho newer" if local_changed else "ResQ newer"
        detail = f"{status.split()[0]} has the newer timestamp."

    if not direction:
        return "", status, f"{detail} The reserving class has no newer side, so nothing is pushed.", False
    if direction == ACTION_ARCRHO_TO_RESQ:
        source, target, target_changed = "ArcRho", "ResQ", remote_changed
    else:
        source, target, target_changed = "ResQ", "ArcRho", local_changed
    if target_changed:
        return direction, status, f"{detail} The {source} copy overwrites this {target} change; review before applying.", True
    return direction, status, detail, False


def build_sync_plan(
    arcrho_items: Iterable[Mapping[str, Any]],
    resq_items: Iterable[Mapping[str, Any]],
    state: Mapping[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Build a deterministic, reviewable plan over the items both inventories hold.

    An item on one side only is not a synchronization candidate: a new dataset
    or method reaches the other side through an import, not through this
    review, so such items never become rows. Every row is pushed in the one
    direction ``plan_direction`` decides for the whole reserving class.
    """

    local_groups = _group_inventory(arcrho_items)
    remote_groups = _group_inventory(resq_items)
    state_items = state.get("items") if isinstance(state, Mapping) and isinstance(state.get("items"), Mapping) else {}
    rows: list[dict[str, Any]] = []
    comparable: list[tuple[dict[str, Any], Mapping[str, Any] | None]] = []
    for key in sorted(set(local_groups) & set(remote_groups)):
        local_candidates = local_groups[key]
        remote_candidates = remote_groups[key]
        display_name = clean_name(local_candidates[0].get("name"))
        baseline = state_items.get(key) if isinstance(state_items, Mapping) and isinstance(state_items.get(key), Mapping) else None
        row: dict[str, Any] = {
            "id": _row_id(key),
            "key": key,
            "name": display_name,
            "arcrho": local_candidates[0] if len(local_candidates) == 1 else None,
            "resq": remote_candidates[0] if len(remote_candidates) == 1 else None,
            "action": "",
            "status": "",
            "detail": "",
            "selected": False,
            "disabled": True,
            "review": False,
            # True once both sides are paired and agree on identity, so their
            # timestamps mean the same thing and may be compared.
            "comparable": False,
            "state_signature": _state_signature(baseline),
        }
        if len(local_candidates) > 1 or len(remote_candidates) > 1:
            row.update(
                status="Ambiguous name",
                detail=(
                    f"Found {len(local_candidates)} ArcRho and {len(remote_candidates)} ResQ items "
                    "with the same normalized name."
                ),
            )
            rows.append(row)
            continue

        arcrho = row["arcrho"]
        resq = row["resq"]
        local_kind = clean_name((arcrho or {}).get("kind"))
        remote_kind = clean_name((resq or {}).get("kind"))
        row["kind"] = local_kind or remote_kind or "Dataset"
        if logical_key(local_kind) != logical_key(remote_kind):
            row.update(
                status="Type mismatch",
                detail=f"ArcRho identifies this as {local_kind}; ResQ identifies it as {remote_kind}.",
            )
            rows.append(row)
            continue

        local_format = clean_name(arcrho.get("data_format"))
        remote_format = clean_name(resq.get("data_format"))
        if local_format and remote_format and logical_key(local_format) != logical_key(remote_format):
            row.update(
                status="Format mismatch",
                detail=f"ArcRho is {local_format}; ResQ is {remote_format}.",
            )
            rows.append(row)
            continue
        local_type = clean_name(arcrho.get("dataset_type"))
        remote_type = clean_name(resq.get("dataset_type"))
        if local_type and remote_type and logical_key(local_type) != logical_key(remote_type):
            row.update(
                status="Dataset Type mismatch",
                detail=f"ArcRho uses {local_type}; ResQ uses {remote_type}.",
            )
            rows.append(row)
            continue
        local_method_name = clean_name(arcrho.get("method_name"))
        remote_method_name = clean_name(resq.get("method_name"))
        if (
            logical_key(local_kind) != logical_key("Dataset")
            and local_method_name
            and remote_method_name
            and logical_key(local_method_name) != logical_key(remote_method_name)
        ):
            row.update(
                status="Method mismatch",
                detail=(
                    f"ArcRho method {local_method_name} and ResQ method "
                    f"{remote_method_name} produce the same output name."
                ),
            )
            rows.append(row)
            continue

        row["comparable"] = True
        rows.append(row)
        comparable.append((row, baseline))

    direction = plan_direction(rows)["direction"]
    for row, baseline in comparable:
        arcrho = row["arcrho"]
        resq = row["resq"]
        action, status, detail, review = _comparison_action(arcrho, resq, baseline, direction)
        row.update(action=action, status=status, detail=detail, review=review)
        if action:
            supported, reason = _support_for_action(action, arcrho, resq)
            if supported:
                # Every supported row rides with the reserving class; a review
                # mark is a warning to read, not a reason to leave the row out.
                row["disabled"] = False
                row["selected"] = True
                if action == ACTION_ARCRHO_TO_RESQ:
                    scope_note = clean_name(arcrho.get("export_scope_note"))
                    if scope_note:
                        row["status"] = f"{status}; supported fields only"
                        row["detail"] = f"{detail} {scope_note}".strip()
            else:
                row["status"] = f"{status}; unsupported"
                row["detail"] = f"{detail} {reason}".strip()
    return rows


def plan_signature(row: Mapping[str, Any]) -> dict[str, Any]:
    """Return the immutable observations rechecked after the review window."""

    def side(item: Any) -> dict[str, Any]:
        source = item if isinstance(item, Mapping) else {}
        return {
            "present": bool(item),
            "kind": clean_name(source.get("kind")),
            "data_format": clean_name(source.get("data_format")),
            "method_name": clean_name(source.get("method_name")),
            "dataset_type": clean_name(source.get("dataset_type")),
            "modified_timestamp": _timestamp(source.get("modified_timestamp")),
        }

    return {
        "key": str(row.get("key") or ""),
        "action": str(row.get("action") or ""),
        "disabled": bool(row.get("disabled")),
        "review": bool(row.get("review")),
        "state_signature": dict(row.get("state_signature") or {}),
        "arcrho": side(row.get("arcrho")),
        "resq": side(row.get("resq")),
    }


def _signature_sides(
    left: Mapping[str, Any], right: Mapping[str, Any], side_name: str
) -> tuple[Mapping[str, Any], Mapping[str, Any]]:
    a = left.get(side_name) if isinstance(left.get(side_name), Mapping) else {}
    b = right.get(side_name) if isinstance(right.get(side_name), Mapping) else {}
    return a, b


def _side_identity_equal(a: Mapping[str, Any], b: Mapping[str, Any]) -> bool:
    if bool(a.get("present")) != bool(b.get("present")):
        return False
    for field in ("kind", "data_format", "method_name", "dataset_type"):
        if clean_name(a.get(field)) != clean_name(b.get(field)):
            return False
    return True


def _side_timestamp_equal(a: Mapping[str, Any], b: Mapping[str, Any]) -> bool:
    a_time = _timestamp(a.get("modified_timestamp"))
    b_time = _timestamp(b.get("modified_timestamp"))
    if a_time is None or b_time is None:
        return a_time == b_time
    return _timestamps_equal(a_time, b_time)


def signatures_equal(left: Mapping[str, Any], right: Mapping[str, Any]) -> bool:
    if str(left.get("key") or "") != str(right.get("key") or ""):
        return False
    if str(left.get("action") or "") != str(right.get("action") or ""):
        return False
    if bool(left.get("disabled")) != bool(right.get("disabled")):
        return False
    if bool(left.get("review")) != bool(right.get("review")):
        return False
    if dict(left.get("state_signature") or {}) != dict(right.get("state_signature") or {}):
        return False
    for side_name in ("arcrho", "resq"):
        a, b = _signature_sides(left, right, side_name)
        if not _side_identity_equal(a, b) or not _side_timestamp_equal(a, b):
            return False
    return True


def write_signatures_equal(
    left: Mapping[str, Any], right: Mapping[str, Any], *, source_side: str
) -> bool:
    """Tell whether a row may still be written from ``source_side``.

    ``signatures_equal`` holds a whole row still, which is right while the
    review is open. Inside one write batch it is too strict: saving a DFM into
    ResQ makes ResQ recalculate every Result Selection downstream of it, and an
    import into ArcRho refreshes its dependents, so the batch itself re-stamps
    the target side of a later row and shifts its proposed action against the
    unchanged baseline. Here only the identity of both sides and the timestamp
    of the side being written from decide.
    """

    if source_side not in ("arcrho", "resq"):
        raise ValueError(f"Unknown source side: {source_side!r}")
    if str(left.get("key") or "") != str(right.get("key") or ""):
        return False
    for side_name in ("arcrho", "resq"):
        a, b = _signature_sides(left, right, side_name)
        if not _side_identity_equal(a, b):
            return False
        if side_name == source_side and not _side_timestamp_equal(a, b):
            return False
    return True


def sync_state_path(server_root: str | os.PathLike[str], project_name: Any, rc_path: Any, connection_name: Any) -> Path:
    """Return the project-owned state path without embedding machine-local paths."""

    project = clean_name(project_name)
    if not project or project in {".", ".."} or any(separator in project for separator in ("/", "\\")):
        raise ValueError("project_name must be one project folder name.")
    identity = "\0".join((project.casefold(), clean_name(rc_path).casefold(), clean_name(connection_name).casefold()))
    digest = hashlib.sha256(identity.encode("utf-8")).hexdigest()
    return Path(server_root) / "projects" / project / SYNC_STATE_DIR / f"{digest}.json"


def empty_sync_state(project_name: Any, rc_path: Any, connection_name: Any) -> dict[str, Any]:
    return {
        "version": SYNC_STATE_VERSION,
        "project_name": clean_name(project_name),
        "reserving_class": clean_name(rc_path),
        "connection_name": clean_name(connection_name),
        "updated_at": "",
        "items": {},
    }


def read_sync_state(path: str | os.PathLike[str], project_name: Any, rc_path: Any, connection_name: Any) -> dict[str, Any]:
    expected = empty_sync_state(project_name, rc_path, connection_name)
    source = Path(path)
    try:
        payload = json.loads(source.read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        return expected
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Could not read synchronization state {source}: {exc}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError(f"Synchronization state is invalid or belongs to another scope: {source}")
    if (
        payload.get("version") != SYNC_STATE_VERSION
        or clean_name(payload.get("project_name")) != expected["project_name"]
        or clean_name(payload.get("reserving_class")) != expected["reserving_class"]
        or clean_name(payload.get("connection_name")) != expected["connection_name"]
        or not isinstance(payload.get("items"), dict)
    ):
        raise RuntimeError(f"Synchronization state is invalid or belongs to another scope: {source}")
    return payload


def record_synced_items(
    state: Mapping[str, Any],
    keys: Iterable[str],
    arcrho_items: Iterable[Mapping[str, Any]],
    resq_items: Iterable[Mapping[str, Any]],
    *,
    synced_at: str | None = None,
) -> dict[str, Any]:
    """Return state updated only for keys that exist on both sides post-sync."""

    local = _group_inventory(arcrho_items)
    remote = _group_inventory(resq_items)
    updated = dict(state)
    entries = dict(state.get("items") or {}) if isinstance(state.get("items"), Mapping) else {}
    timestamp = str(synced_at or datetime.now(timezone.utc).isoformat()).strip()
    recorded: list[str] = []
    for raw_key in keys:
        key = logical_key(raw_key)
        local_items = local.get(key, [])
        remote_items = remote.get(key, [])
        if len(local_items) != 1 or len(remote_items) != 1:
            continue
        local_timestamp = _timestamp(local_items[0].get("modified_timestamp"))
        remote_timestamp = _timestamp(remote_items[0].get("modified_timestamp"))
        if local_timestamp is None or remote_timestamp is None:
            continue
        entries[key] = {
            "name": clean_name(local_items[0].get("name") or remote_items[0].get("name")),
            "kind": clean_name(local_items[0].get("kind") or remote_items[0].get("kind")),
            "arcrho_present": True,
            "resq_present": True,
            "arcrho_timestamp": local_timestamp,
            "resq_timestamp": remote_timestamp,
            "synced_at": timestamp,
        }
        recorded.append(key)
    updated["items"] = entries
    updated["updated_at"] = timestamp
    updated["_recorded_keys"] = recorded
    return updated


def absorb_propagated_changes(
    state: Mapping[str, Any],
    before_rows: Iterable[Mapping[str, Any]],
    after_rows: Iterable[Mapping[str, Any]],
    *,
    keys: Iterable[str],
    synced_at: str | None = None,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Baseline the moves a write batch itself caused downstream of its rows.

    ``keys`` names the review rows downstream of what the batch wrote and not
    written themselves. Both systems recalculate those from the inputs the
    batch just synchronized and re-stamp them, so a move between the batch's
    opening observation (``before_rows``) and its closing one (``after_rows``)
    is the batch's own doing, not an edit. Each side that moved takes its
    closing timestamp as the new baseline while a side that held still keeps
    its old one, so a change that was already pending before the batch stays
    pending. A row with no baseline yet is baselined only when it showed no
    difference before the batch. Returns the updated state and one record per
    absorbed row naming the sides that moved.
    """

    before = {str(row.get("key") or ""): row for row in before_rows}
    after = {str(row.get("key") or ""): row for row in after_rows}
    updated = dict(state)
    entries = dict(state.get("items") or {}) if isinstance(state.get("items"), Mapping) else {}
    timestamp = str(synced_at or datetime.now(timezone.utc).isoformat()).strip()
    absorbed: list[dict[str, Any]] = []

    def side(row: Mapping[str, Any], name: str) -> Mapping[str, Any]:
        value = row.get(name)
        return value if isinstance(value, Mapping) else {}

    for raw_key in keys:
        key = logical_key(raw_key)
        before_row = before.get(key)
        after_row = after.get(key)
        if before_row is None or after_row is None:
            continue
        moved: dict[str, float] = {}
        closing: dict[str, float | None] = {}
        for side_name in ("arcrho", "resq"):
            opening_time = _timestamp(side(before_row, side_name).get("modified_timestamp"))
            closing_time = _timestamp(side(after_row, side_name).get("modified_timestamp"))
            closing[side_name] = closing_time
            if closing_time is None:
                continue
            if opening_time is not None and _timestamps_equal(opening_time, closing_time):
                continue
            moved[side_name] = closing_time
        if not moved:
            continue
        entry = entries.get(key)
        if isinstance(entry, Mapping):
            new_entry = dict(entry)
            for side_name, closing_time in moved.items():
                new_entry[f"{side_name}_timestamp"] = closing_time
                new_entry[f"{side_name}_present"] = True
        else:
            opening_local = _timestamp(side(before_row, "arcrho").get("modified_timestamp"))
            opening_remote = _timestamp(side(before_row, "resq").get("modified_timestamp"))
            showed_no_difference = (
                not before_row.get("action")
                and opening_local is not None
                and opening_remote is not None
                and _timestamps_equal(opening_local, opening_remote)
            )
            if not showed_no_difference or closing["arcrho"] is None or closing["resq"] is None:
                continue
            new_entry = {
                "name": clean_name(before_row.get("name")),
                "kind": clean_name(before_row.get("kind")),
                "arcrho_present": True,
                "resq_present": True,
                "arcrho_timestamp": closing["arcrho"],
                "resq_timestamp": closing["resq"],
                "synced_at": timestamp,
            }
        new_entry["propagated_at"] = timestamp
        entries[key] = new_entry
        absorbed.append({
            "key": key,
            "name": clean_name(before_row.get("name")),
            "kind": clean_name(before_row.get("kind")),
            "sides": sorted(moved),
        })
    updated["items"] = entries
    if absorbed:
        updated["updated_at"] = timestamp
    return updated, absorbed


def write_sync_state(path: str | os.PathLike[str], state: Mapping[str, Any]) -> Path:
    """Atomically persist the one canonical synchronization-baseline document."""

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.{uuid.uuid4().hex}.tmp")
    try:
        payload = {key: value for key, value in state.items() if not str(key).startswith("_")}
        temporary.write_text(persisted_json_text(payload), encoding="utf-8", newline="\n")
        os.replace(temporary, target)
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
    return target

#!/usr/bin/env python3
"""Rewrite dataset sidecar cell links into their block form, in place.

A link used to be stored one entry per cell, so a manual input triangle linked
to a 120x116 Excel range cost 6,786 entries and roughly 650 KB -- a file no
reviewer can read and a read every dataset open pays for over the network.
``arcrho_api.dataset_link_contract`` now stores the same cells as blocks, one
rectangle per line, and the same file costs 116 lines and under 9 KB.

Readers do not accept the old shape, so a workspace written before the change
has to be converted before a current build opens it. This walks each project's
reserving classes, rewrites every sidecar whose links are still per-cell, and
proves each rewrite by expanding the text it just produced back to cells: a
converted file has to give back exactly the links that went in, or it is not
written.

Nothing is written without ``--apply``.

Usage
-----
    py -3.10 tools/migrate_dataset_link_blocks.py
    py -3.10 tools/migrate_dataset_link_blocks.py --project "NJ_Annual_Prod_2026 Q3-Aug"
    py -3.10 tools/migrate_dataset_link_blocks.py --apply
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Iterator

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "python-api" / "src"))

from arcrho_api.dataset_link_contract import (  # noqa: E402
    LINK_TARGET_FIELDS,
    compact_sidecar_links,
    expand_sidecar_links,
)
from arcrho_api.io import persisted_json_text  # noqa: E402

DEFAULT_WORKSPACE = r"E:\ArcRho Server"


def sidecar_files(workspace: Path, projects: list[str]) -> Iterator[Path]:
    """Every sidecar of every named project, or of the whole workspace."""

    root = workspace / "projects"
    names = projects or sorted(entry.name for entry in root.iterdir() if entry.is_dir())
    for name in names:
        for sidecar_dir in sorted((root / name / "data").glob("*/sidecars")):
            yield from sorted(sidecar_dir.glob("*.json"))


def per_cell_links(payload: Any) -> int:
    """How many link cells this payload still stores one entry per cell."""

    if not isinstance(payload, dict):
        return 0
    cells = 0
    for field in LINK_TARGET_FIELDS:
        for link in payload.get(field) or []:
            targets = link.get("target_cells") if isinstance(link, dict) else None
            if isinstance(targets, list) and targets and isinstance(targets[0], dict):
                cells += len(targets)
    return cells


def convert(path: Path, apply: bool) -> dict[str, Any] | None:
    """Convert one sidecar; returns what changed, or None when it need not."""

    before = path.read_text(encoding="utf-8")
    original = json.loads(before)
    cells = per_cell_links(original)
    if not cells:
        return None
    text = persisted_json_text(compact_sidecar_links(json.loads(before)))
    if expand_sidecar_links(json.loads(text)) != original:
        raise ValueError("the block form does not give back the links that went in")
    if apply:
        # Through a temporary file, as every other sidecar writer does: a run
        # stopped part way leaves the old file, never half of the new one.
        temp = path.with_name(f"{path.name}.{uuid.uuid4()}.tmp")
        try:
            temp.write_text(text, encoding="utf-8", newline="\n")
            os.replace(temp, path)
        finally:
            if temp.exists():
                temp.unlink()
    return {"file": str(path), "cells": cells, "bytes_before": len(before), "bytes_after": len(text)}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--workspace", default=DEFAULT_WORKSPACE, help="Workspace root (default: %(default)s)")
    parser.add_argument("--project", action="append", default=[], help="Project folder name; repeatable")
    parser.add_argument("--apply", action="store_true", help="Write the conversion (default is a dry run)")
    parser.add_argument("--report", default="", help="Write the full JSON report to this path")
    parser.add_argument("--workers", type=int, default=32, help="Files read at once (default: %(default)s)")
    args = parser.parse_args(argv)

    workspace = Path(args.workspace)
    if not (workspace / "projects").is_dir():
        print(f"Workspace not found: {workspace}", file=sys.stderr)
        return 2

    converted: list[dict[str, Any]] = []
    failures: list[dict[str, str]] = []
    paths = list(sidecar_files(workspace, args.project))
    scanned = len(paths)

    def run(path: Path) -> tuple[Path, dict[str, Any] | None, str]:
        try:
            return path, convert(path, args.apply), ""
        except Exception as err:  # a sidecar that cannot be read is reported, never skipped silently
            return path, None, f"{type(err).__name__}: {err}"

    # Every file on the share is a network round trip, so the walk pays them
    # in parallel rather than one awaited read at a time.
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        for path, result, error in pool.map(run, paths):
            if error:
                failures.append({"file": str(path), "error": error})
            elif result:
                converted.append(result)

    before = sum(entry["bytes_before"] for entry in converted)
    after = sum(entry["bytes_after"] for entry in converted)
    print(f"Workspace: {workspace}")
    print(f"Sidecars scanned:   {scanned:,}")
    print(f"Sidecars converted: {len(converted):,}{'' if args.apply else ' (dry run, nothing written)'}")
    print(f"Link cells stored:  {sum(entry['cells'] for entry in converted):,}")
    print(f"Bytes:              {before:,} -> {after:,}")
    for entry in sorted(converted, key=lambda item: item["bytes_before"] - item["bytes_after"], reverse=True)[:10]:
        print(f"  {entry['bytes_before']:>9,} -> {entry['bytes_after']:>7,}  {entry['file']}")
    for failure in failures:
        print(f"  FAILED  {failure['file']}: {failure['error']}", file=sys.stderr)

    if args.report:
        Path(args.report).write_text(
            json.dumps(
                {"workspace": str(workspace), "applied": args.apply, "converted": converted, "failures": failures},
                indent=2,
            ),
            encoding="utf-8",
        )
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())

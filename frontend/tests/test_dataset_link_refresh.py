"""Link-driven datasets in the dependency graph and the propagation walk.

A plain-input dataset whose cells are driven by ArcRho links is an
instance-level node of the dependency graph: its save records precedent edges
for the datasets its links read, the linked sources record it back as a
dependent, and the Engine's dependent walk re-evaluates the links when a
source is refreshed. Excel operands are the one soft spot — a workbook the
server host cannot read keeps the linked cells' last values and reports a
warning instead of failing the chain.
"""

from __future__ import annotations

import copy
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

FRONTEND_ROOT = Path(__file__).resolve().parents[1]
if str(FRONTEND_ROOT) not in sys.path:
    sys.path.insert(0, str(FRONTEND_ROOT))

from app_server.services import (  # noqa: E402
    calculated_dataset_service,
    dataset_link_refresh_service,
    dataset_service,
    dataset_sidecar_status_service,
)
from dependent_propagation_workspace_stub import IsolatedPropagationWorkspace  # noqa: E402


def _vector(name, values, *, source_kind="input", method_type="None", links=None):
    sidecar = {
        "dataset_name": name,
        "dataset_type": name,
        "data_format": "Vector",
        "source_kind": source_kind,
        "method_type": method_type,
        "period_length": len(values),
        "origin_labels": [str(2024 + i) for i in range(len(values))],
        "precedents": [],
        "dependents": [],
    }
    if links:
        sidecar.update(links)
    return {
        "sidecar": sidecar,
        "values": [[value] for value in values],
    }


class LinkEdgeOnSaveTests(unittest.TestCase):
    """A save with ArcRho links records the instance-level graph edges."""

    def setUp(self) -> None:
        self.propagation_workspace = IsolatedPropagationWorkspace().start()

    def tearDown(self) -> None:
        self.propagation_workspace.stop()

    def test_link_precedents_merge_and_the_far_side_is_updated(self) -> None:
        written = {}
        edge_updates = []

        def capture_csv_and_sidecar(_frame, _csv_path, _sidecar_path, payload):
            written["payload"] = copy.deepcopy(payload)

        rows = [
            {"name": "Vector C", "data_format": "Vector", "category": "", "calculated": False, "formula": "", "source": "", "generated": False},
            {"name": "Vector A", "data_format": "Vector", "category": "", "calculated": False, "formula": "", "source": "", "generated": False},
        ]
        with (
            patch.object(dataset_service, "_get_dataset_sidecar_path", return_value="sidecar.json"),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value={}),
            patch.object(dataset_service, "_is_app_calculated_dataset_type", return_value=(False, "")),
            patch.object(dataset_service, "_current_user_name", return_value="tester"),
            patch.object(dataset_service, "_write_dataset_csv_and_sidecar", side_effect=capture_csv_and_sidecar),
            patch.object(dataset_service.config, "get_project_dataset_cache_dir", return_value="cache"),
            patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=rows),
            patch.object(
                calculated_dataset_service,
                "sidecar_graph_fields",
                return_value={"precedents": [], "dependents": []},
            ),
            patch.object(
                dataset_service.dataset_sidecar_status_service,
                "update_precedent_dependents",
                side_effect=lambda *args, **kwargs: edge_updates.append(args) or [],
            ),
            patch.object(
                dataset_service.dataset_sidecar_status_service,
                "refresh_method_statuses_for_dependents",
                return_value=[],
            ),
            patch.object(dataset_service.dataset_instance_index_service, "rebuild_index"),
        ):
            dataset_service.save_dataset_sidecar(
                "Project",
                "Class",
                "Vector C",
                source_kind="input",
                method_type="None",
                data_format="Vector",
                origin_length=2,
                development_length=1,
                values=[[1.0], [2.0]],
                formula_links=[{
                    "formula": "=[Vector A][1:2] * 2",
                    "target_cells": [
                        {"row": 0, "column": 0, "result_row": 0, "result_column": 0},
                        {"row": 1, "column": 0, "result_row": 1, "result_column": 0},
                    ],
                }],
            )

        payload = written["payload"]
        self.assertIn({"dataset_name": "Vector A"}, payload["precedents"])
        self.assertEqual(
            [(args[2], list(args[3]), list(args[4])) for args in edge_updates],
            [("Vector C", [], ["Vector A"])],
        )


class LinkDrivenWalkTests(unittest.TestCase):
    """The walk's link wave follows dependents edges and guards cycles."""

    def _walk(self, sidecars, refresh_results, roots, visited=None):
        calls = []

        def fake_sidecar_path(_project, _rc, name):
            return name

        def fake_read(path):
            return copy.deepcopy(sidecars.get(path))

        def fake_refresh(_project, _rc, name):
            calls.append(name)
            return refresh_results[name]

        link_updates = {"refreshed": [], "failed": [], "warnings": [], "errors": []}
        visited_keys = set(visited or [])
        with (
            patch.object(dataset_sidecar_status_service, "sidecar_path", side_effect=fake_sidecar_path),
            patch.object(dataset_sidecar_status_service, "read_sidecar", side_effect=fake_read),
            patch.object(
                dataset_link_refresh_service,
                "refresh_dataset_links",
                side_effect=fake_refresh,
            ),
        ):
            fresh = calculated_dataset_service._refresh_link_driven_dependents(
                "Project", "Class", roots, visited_keys, link_updates
            )
        return fresh, link_updates, calls

    def test_chained_link_dependents_refresh_in_order(self) -> None:
        sidecars = {
            "Root": {"dependents": [{"dataset_name": "Linked A"}]},
            "Linked A": {
                "source_kind": "input",
                "method_type": "None",
                "formula_links": [{"formula": "=[Root][1:2]", "target_cells": []}],
                "dependents": [{"dataset_name": "Linked B"}],
            },
            "Linked B": {
                "source_kind": "input",
                "method_type": "None",
                "internal_links": [{"reference": "=[Linked A][1:2]", "target_cells": []}],
                "dependents": [],
            },
        }
        results = {
            "Linked A": {"ok": True, "refreshed": True, "changed": True, "warnings": []},
            "Linked B": {"ok": True, "refreshed": True, "changed": True, "warnings": []},
        }

        fresh, link_updates, calls = self._walk(sidecars, results, ["Root"])

        self.assertEqual(fresh, ["Linked A", "Linked B"])
        self.assertEqual(link_updates["refreshed"], ["Linked A", "Linked B"])
        self.assertEqual(calls, ["Linked A", "Linked B"])

    def test_the_cycle_guard_refreshes_each_dataset_at_most_once(self) -> None:
        # D 31 reads the Result Selection's published indicated and feeds the
        # same Result Selection: once the walk has visited it, the back edge
        # must not re-enter it.
        sidecars = {
            "Indicated": {"dependents": [{"dataset_name": "D 31"}]},
            "D 31": {
                "source_kind": "input",
                "method_type": "None",
                "formula_links": [{"formula": "=[Indicated][1:2]", "target_cells": []}],
                "dependents": [{"dataset_name": "Indicated"}, {"dataset_name": "D 31"}],
            },
        }
        results = {
            "D 31": {"ok": True, "refreshed": True, "changed": True, "warnings": []},
        }

        fresh, _link_updates, calls = self._walk(
            sidecars, results, ["Indicated"], visited={"indicated"}
        )

        self.assertEqual(fresh, ["D 31"])
        self.assertEqual(calls, ["D 31"])

    def test_failures_and_warnings_accumulate_without_stopping(self) -> None:
        sidecars = {
            "Root": {
                "dependents": [
                    {"dataset_name": "Broken"},
                    {"dataset_name": "Warned"},
                ],
            },
            "Broken": {
                "source_kind": "input",
                "method_type": "None",
                "formula_links": [{"formula": "=[Gone][1]", "target_cells": []}],
                "dependents": [],
            },
            "Warned": {
                "source_kind": "input",
                "method_type": "None",
                "formula_links": [{"formula": "='C:\\\\F\\\\[B.xlsx]S'!A1", "target_cells": []}],
                "dependents": [],
            },
        }
        results = {
            "Broken": {
                "ok": False,
                "refreshed": False,
                "reason": "link_error",
                "errors": ["Missing dependency: Gone"],
            },
            "Warned": {
                "ok": True,
                "refreshed": True,
                "changed": False,
                "warnings": [{"reference": "='C:...'!A1", "reason": "Excel value could not be read"}],
            },
        }

        fresh, link_updates, _calls = self._walk(sidecars, results, ["Root"])

        self.assertEqual(fresh, [])
        self.assertEqual(link_updates["failed"], ["Broken"])
        self.assertEqual(
            link_updates["errors"],
            [{
                "dataset_name": "Broken",
                "reason": "link_error",
                "errors": ["Missing dependency: Gone"],
            }],
        )
        self.assertEqual(len(link_updates["warnings"]), 1)
        self.assertEqual(link_updates["warnings"][0]["dataset_name"], "Warned")

    def test_method_outputs_and_non_input_dependents_are_left_alone(self) -> None:
        sidecars = {
            "Root": {
                "dependents": [
                    {"dataset_name": "A DFM Output"},
                    {"dataset_name": "Calculated"},
                    {"dataset_name": "Plain"},
                ],
            },
            "A DFM Output": {"source_kind": "dfm", "method_type": "DFM"},
            "Calculated": {"source_kind": "calculated", "method_type": "None"},
            "Plain": {"source_kind": "input", "method_type": "None"},
        }

        fresh, link_updates, calls = self._walk(sidecars, {}, ["Root"])

        self.assertEqual(fresh, [])
        self.assertEqual(calls, [])
        self.assertEqual(link_updates["refreshed"], [])


class RefreshDatasetLinksTests(unittest.TestCase):
    """Server-side evaluation matches the Links-tab semantics."""

    def _refresh(self, datasets, target_name, *, excel_results=None):
        written = {}

        def fake_sidecar_path(_project, _rc, name):
            return name

        def fake_read(path):
            return copy.deepcopy(datasets[path]["sidecar"]) if path in datasets else {}

        def fake_load(_project, _rc, name, **_kwargs):
            entry = datasets[name]
            sidecar = entry["sidecar"]
            return {
                "dataset_name": name,
                "data_format": sidecar.get("data_format", "Vector"),
                "values": copy.deepcopy(entry["values"]),
                "origin_labels": list(sidecar.get("origin_labels") or []),
                "dev_labels": ["Ultimate"],
                "path": f"{name}.csv",
            }

        def capture_write(frame, _csv_path, _sidecar_path, payload):
            written["values"] = frame.values.tolist()
            written["payload"] = copy.deepcopy(payload)

        def fake_excel_batch(items):
            return {"results": [excel_results[i] for i in range(len(items))]}

        with (
            patch.object(dataset_sidecar_status_service, "sidecar_path", side_effect=fake_sidecar_path),
            patch.object(dataset_sidecar_status_service, "read_sidecar", side_effect=fake_read),
            patch.object(dataset_service, "load_cached_dataset_values", side_effect=fake_load),
            patch.object(dataset_service, "_write_dataset_csv_and_sidecar", side_effect=capture_write),
            patch.object(
                dataset_link_refresh_service.excel_service,
                "excel_read_cells_batch",
                side_effect=fake_excel_batch,
            ),
        ):
            result = dataset_link_refresh_service.refresh_dataset_links(
                "Project", "Class", target_name
            )
        return result, written

    def test_a_dataset_formula_recomputes_only_the_owned_cells(self) -> None:
        datasets = {
            "Source": _vector("Source", [10.0, 20.0]),
            "Target": _vector(
                "Target",
                [999.0, 999.0, 77.0],
                links={
                    "formula_links": [{
                        "formula": "=[Source][1:2] * 2",
                        "target_cells": [
                            {"row": 0, "column": 0, "result_row": 0, "result_column": 0},
                            {"row": 1, "column": 0, "result_row": 1, "result_column": 0},
                        ],
                    }],
                },
            ),
        }

        result, written = self._refresh(datasets, "Target")

        self.assertTrue(result["ok"])
        self.assertTrue(result["changed"])
        self.assertEqual(result["warnings"], [])
        self.assertEqual(written["values"], [[20.0], [40.0], [77.0]])
        self.assertEqual(
            written["payload"]["audit_log"][-1]["action"],
            "Auto Refresh",
        )

    def test_an_unreadable_excel_operand_keeps_stale_values_and_warns(self) -> None:
        datasets = {
            "Source": _vector("Source", [10.0, 20.0]),
            "Target": _vector(
                "Target",
                [111.0, 222.0],
                links={
                    "formula_links": [{
                        "formula": "=[Source][1:2] * 'C:\\F\\[B.xlsx]S1'!A1",
                        "target_cells": [
                            {"row": 0, "column": 0, "result_row": 0, "result_column": 0},
                            {"row": 1, "column": 0, "result_row": 1, "result_column": 0},
                        ],
                    }],
                },
            ),
        }

        result, written = self._refresh(
            datasets,
            "Target",
            excel_results={0: {"ok": False, "error": "File not found: C:\\F\\B.xlsx"}},
        )

        self.assertTrue(result["ok"])
        self.assertTrue(result["refreshed"])
        self.assertFalse(result["changed"])
        self.assertEqual(len(result["warnings"]), 1)
        self.assertIn("keep their last values", result["warnings"][0]["reason"])
        self.assertNotIn("values", written)

    def test_a_missing_arcrho_source_fails_the_refresh(self) -> None:
        target = _vector(
            "Target",
            [1.0],
            links={
                "internal_links": [{
                    "reference": "=[Gone][1]",
                    "target_cells": [
                        {"row": 0, "column": 0, "source_row": 0, "source_column": 0},
                    ],
                }],
            },
        )

        def raising_load(_project, _rc, name, **_kwargs):
            from fastapi import HTTPException

            raise HTTPException(404, f"Cached dataset CSV not found for '{name}'.")

        with (
            patch.object(
                dataset_sidecar_status_service, "sidecar_path", side_effect=lambda *_a: "Target"
            ),
            patch.object(
                dataset_sidecar_status_service,
                "read_sidecar",
                return_value=copy.deepcopy(target["sidecar"]),
            ),
            patch.object(
                dataset_service,
                "load_cached_dataset_values",
                side_effect=lambda project, rc, name, **kwargs: (
                    {
                        "dataset_name": "Target",
                        "data_format": "Vector",
                        "values": copy.deepcopy(target["values"]),
                        "origin_labels": ["2024"],
                        "dev_labels": ["Ultimate"],
                        "path": "Target.csv",
                    }
                    if name == "Target"
                    else raising_load(project, rc, name)
                ),
            ),
        ):
            result = dataset_link_refresh_service.refresh_dataset_links(
                "Project", "Class", "Target"
            )

        self.assertFalse(result["ok"])
        self.assertEqual(result["reason"], "link_error")
        self.assertIn("Missing dependency: Gone", result["errors"][0])

    def test_the_target_is_read_at_the_shape_its_links_were_written_at(self) -> None:
        # A link names a cell of the display the dataset was shown at when the
        # link was written, so the refresh asks for the dataset at that shape
        # rather than at the file's own rows.
        datasets = {
            "Source": _vector("Source", [5.0]),
            "Target": _vector(
                "Target",
                [0.0],
                links={
                    "internal_links": [{
                        "reference": "=[Source][1:1]",
                        "target_cells": [{"row": 0, "column": 0, "source_row": 0, "source_column": 0}],
                    }],
                },
            ),
        }
        calls = []

        def fake_load(_project, _rc, name, **kwargs):
            calls.append((name, kwargs))
            return {
                "dataset_name": name,
                "data_format": "Vector",
                "values": copy.deepcopy(datasets[name]["values"]),
                "origin_labels": ["2024"],
                "dev_labels": ["Ultimate"],
                "path": f"{name}.csv",
            }

        with (
            patch.object(dataset_sidecar_status_service, "sidecar_path", side_effect=lambda _p, _rc, n: n),
            patch.object(
                dataset_sidecar_status_service,
                "read_sidecar",
                side_effect=lambda path: copy.deepcopy(datasets[path]["sidecar"]),
            ),
            patch.object(dataset_service, "load_cached_dataset_values", side_effect=fake_load),
            patch.object(dataset_service, "_write_dataset_csv_and_sidecar"),
        ):
            result = dataset_link_refresh_service.refresh_dataset_links("Project", "Class", "Target")

        self.assertTrue(result["ok"])
        self.assertEqual(calls[0][0], "Target")
        self.assertIs(calls[0][1].get("at_linked_shape"), True)

    def test_an_internal_link_copies_by_stored_source_coordinates(self) -> None:
        datasets = {
            "Source": _vector("Source", [5.0, 6.0, 7.0]),
            "Target": _vector(
                "Target",
                [0.0, 0.0],
                links={
                    "internal_links": [{
                        "reference": "=[Source][2:3]",
                        "target_cells": [
                            {"row": 0, "column": 0, "source_row": 1, "source_column": 0},
                            {"row": 1, "column": 0, "source_row": 2, "source_column": 0},
                        ],
                    }],
                },
            ),
        }

        result, written = self._refresh(datasets, "Target")

        self.assertTrue(result["ok"])
        self.assertEqual(written["values"], [[6.0], [7.0]])


if __name__ == "__main__":
    unittest.main()

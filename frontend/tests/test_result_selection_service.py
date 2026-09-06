from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from fastapi import HTTPException

FRONTEND_ROOT = Path(__file__).resolve().parents[1]
PYTHON_API_SRC = FRONTEND_ROOT.parent / "python-api" / "src"
for path in (FRONTEND_ROOT, PYTHON_API_SRC):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from arcrho_api.sidecar_core_contract import finalize_sidecar
from app_server import config
from app_server.services import (
    calculated_dataset_service,
    dataset_service,
    precedent_cache_service,
    result_selection_service,
)
from dependent_propagation_workspace_stub import IsolatedPropagationWorkspace


class ResultSelectionServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory(dir=str(FRONTEND_ROOT))
        self.root = Path(self.temp_dir.name)
        self.methods = self.root / "methods"
        self.datasets = self.root / "datasets"
        self.sidecars = self.root / "dataset_sidercars"
        for path in (self.methods, self.datasets, self.sidecars):
            path.mkdir(parents=True)
        settings = self.root / "general_settings.json"
        settings.write_text(
            '{"origin_start_date":"202301","origin_end_date":"202412","development_end_date":"202412"}',
            encoding="utf-8",
        )
        self.patchers = [
            IsolatedPropagationWorkspace(),
            mock.patch.object(config, "get_general_settings_path", return_value=str(settings)),
            mock.patch.object(config, "get_project_method_data_dir", return_value=str(self.methods)),
            mock.patch.object(config, "get_project_dataset_cache_dir", return_value=str(self.datasets)),
            mock.patch.object(config, "get_project_dataset_sidecar_dir", return_value=str(self.sidecars)),
        ]
        for patcher in self.patchers:
            patcher.start()

    def tearDown(self) -> None:
        for patcher in reversed(self.patchers):
            patcher.stop()
        self.temp_dir.cleanup()

    def write_json(self, path: Path, payload: dict) -> None:
        path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")

    def method_payload(self, *, basis: bool = False) -> dict:
        return {
            "json_format": result_selection_service.RESULT_SELECTION_JSON_FORMAT,
            "details_tab": {
                "name": "Selection",
                "output_type": "Selected Ultimate",
                "origin_length": 12,
                "ratio_basis_datasets": ["Premium"] if basis else [],
                "active_ratio_basis_dataset": "Premium" if basis else "",
                "show_ratios_as_percentages": True,
                "statistic_decimal_places": 1,
            },
            "method_tab": {
                "origin_labels": ["2025", "2026"],
                "show_weights": True,
                "loaded_datasets": [{
                    "name": "Paid",
                    "dataset_type": "Paid",
                    "data_format": "Vector",
                    "method_type": "None",
                    "category": "Loss",
                    "source_kind": "input",
                    "origin_length": 12,
                    "values": [10, 20],
                    "weights": [1, 1],
                }],
                "ratio_basis_values": [{"name": "Premium", "values": [100, 200]}] if basis else [],
                "calculated_ultimate": [10, 20],
                "selected_ultimate": [10, 99],
                "ultimate_overrides": [None, 99],
            },
            "results_tab": {},
            "validation_tab": {},
            "method_metadata": {"last_modified": "2026-01-01T00:00:00Z"},
        }

    def write_selection(self, *, basis: bool = False) -> None:
        self.write_json(self.methods / "RS@Selection.json", self.method_payload(basis=basis))
        self.write_json(self.sidecars / "Selection.json", {
            "dataset_name": "Selection",
            "dataset_type": "Selected Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Selection@12.csv",
            "status": 2,
            "precedents": ["Paid", *(["Premium"] if basis else [])],
            "dependents": [],
            "audit_log": [],
        })
        (self.datasets / "Selection@12.csv").write_text("10\n99\n", encoding="utf-8")

    def write_source(self, name: str, values: list[float], *, data_format: str = "Vector") -> None:
        filename = f"{name}@12.csv" if data_format == "Vector" else f"{name}@12@12@cum@dev.csv"
        rows = "\n".join(str(value) for value in values) + "\n"
        (self.datasets / filename).write_text(rows, encoding="utf-8")
        self.write_json(self.sidecars / f"{name}.json", {
            "dataset_name": name,
            "dataset_type": name,
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "input",
            "method_type": "None",
            "data_format": data_format,
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": filename,
            "dependents": ["Selection"],
        })

    def test_current_open_reads_only_own_method_and_sidecar_json(self) -> None:
        self.write_selection()
        original = result_selection_service._read_json
        paths = []

        def recording_read(path: str):
            paths.append(Path(path))
            return original(path)

        with mock.patch.object(result_selection_service, "_read_json", side_effect=recording_read):
            result = result_selection_service.load_result_selection(
                "Project",
                "Class",
                "Selection",
            )

        self.assertTrue(result["method_exists"])
        self.assertFalse(result["upgraded"])
        self.assertCountEqual(
            paths,
            [self.methods / "RS@Selection.json", self.sidecars / "Selection.json"],
        )

    def test_current_open_rejects_method_sidecar_precedent_mismatch(self) -> None:
        self.write_selection()
        sidecar_path = self.sidecars / "Selection.json"
        sidecar = json.loads(sidecar_path.read_text(encoding="utf-8"))
        sidecar["precedents"] = []
        self.write_json(sidecar_path, sidecar)

        with self.assertRaises(HTTPException) as raised:
            result_selection_service.load_result_selection("Project", "Class", "Selection")

        self.assertEqual(raised.exception.status_code, 409)
        self.assertIn("precedents", str(raised.exception.detail))

    def test_save_registers_sources_and_ratio_bases_before_propagation(self) -> None:
        payload = self.method_payload(basis=True)
        with mock.patch(
            "app_server.services.dataset_service.save_dataset_sidecar",
            return_value={"ok": True, "audit_log": []},
        ) as save_sidecar:
            result = result_selection_service.save_result_selection(
                "Project",
                "Class",
                payload,
                "note",
            )

        self.assertTrue(result["ok"])
        self.assertTrue((self.methods / "RS@Selection.json").exists())
        self.assertTrue((self.datasets / "Selection@12.csv").exists())
        self.assertEqual(save_sidecar.call_args.kwargs["precedents"], ["Paid", "Premium"])
        saved = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["json_format"], result_selection_service.RESULT_SELECTION_JSON_FORMAT)
        self.assertNotIn("ratio_basis", saved["details_tab"])
        self.assertNotIn("ratio_basis_dataset", saved["details_tab"])

    def test_save_rejects_a_stale_open_method_revision(self) -> None:
        self.write_selection()
        sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        sidecar["status"] = 0
        self.write_json(self.sidecars / "Selection.json", sidecar)
        loaded = result_selection_service.load_result_selection("Project", "Class", "Selection")
        changed = self.method_payload()
        # Somebody else edited the method itself, which is what the token guards.
        changed["method_tab"]["loaded_datasets"][0]["values"] = [11, 21]
        self.write_json(self.methods / "RS@Selection.json", changed)

        with self.assertRaises(HTTPException) as raised:
            result_selection_service.save_result_selection(
                "Project",
                "Class",
                loaded["method"],
                expected_revision=loaded["method_revision"],
            )

        self.assertEqual(raised.exception.status_code, 409)
        self.assertIn("changed on disk", str(raised.exception.detail))

    def test_a_timestamp_only_rewrite_does_not_make_an_open_method_stale(self) -> None:
        # An RPC upload records the time ResQ stamped on the copy it just saved.
        # The content is untouched, so the editor that has this method open must
        # keep the token it saves with.
        self.write_selection()
        self.write_source("Paid", [10, 20])
        sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        sidecar["status"] = 0
        self.write_json(self.sidecars / "Selection.json", sidecar)
        loaded = result_selection_service.load_result_selection("Project", "Class", "Selection")
        restamped = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        restamped["method_metadata"]["last_modified"] = "2026-02-01T00:00:00Z"
        restamped["method_metadata"]["data_refreshed"] = "2026-02-01T00:00:00Z"
        self.write_json(self.methods / "RS@Selection.json", restamped)

        with (
            mock.patch.object(calculated_dataset_service, "apply_sidecar_graph_fields"),
            mock.patch.object(calculated_dataset_service, "recalculate_dependents", return_value={"ok": True}),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.save_result_selection(
                "Project",
                "Class",
                loaded["method"],
                expected_revision=loaded["method_revision"],
            )

        self.assertTrue(result["ok"])

    def test_review_needed_save_revalidates_only_the_revised_precedent_list(self) -> None:
        self.write_selection()
        current = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        removed = {**current["method_tab"]["loaded_datasets"][0], "name": "Removed"}
        current["method_tab"]["loaded_datasets"].append(removed)
        self.write_json(self.methods / "RS@Selection.json", current)
        sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        sidecar["precedents"] = ["Paid", "Removed"]
        self.write_json(self.sidecars / "Selection.json", sidecar)
        self.write_source("Paid", [30, 40])
        incoming = self.method_payload()
        expected_revision = result_selection_service._method_revision(
            result_selection_service.normalize_method_payload(current, require_complete_basis=True)
        )

        with (
            mock.patch.object(calculated_dataset_service, "apply_sidecar_graph_fields"),
            mock.patch.object(calculated_dataset_service, "recalculate_dependents", return_value={"ok": True}),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.save_result_selection(
                "Project",
                "Class",
                incoming,
                expected_revision=expected_revision,
            )

        self.assertTrue(result["ok"])
        self.assertEqual(result["method"]["method_tab"]["loaded_datasets"][0]["values"], [30.0, 40.0])
        self.assertEqual(result["method"]["method_tab"]["selected_ultimate"], [30.0, 99])
        self.assertEqual((self.datasets / "Selection@12.csv").read_text(encoding="utf-8"), "30.0\n99\n")
        saved_sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(saved_sidecar["status"], 0)
        self.assertEqual(saved_sidecar["precedents"], [{"dataset_name": "Paid"}])

    def test_review_needed_save_warns_and_still_saves_with_unreviewed_precedent(self) -> None:
        self.write_selection()
        self.write_source("Paid", [30, 40])
        paid_path = self.sidecars / "Paid.json"
        paid = json.loads(paid_path.read_text(encoding="utf-8"))
        paid["source_kind"] = "dfm"
        paid["method_type"] = "DFM"
        paid["status"] = 2
        self.write_json(paid_path, paid)

        with (
            mock.patch.object(calculated_dataset_service, "apply_sidecar_graph_fields"),
            mock.patch.object(calculated_dataset_service, "recalculate_dependents", return_value={"ok": True}),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.save_result_selection(
                "Project",
                "Class",
                self.method_payload(),
            )

        self.assertTrue(result["ok"])
        self.assertEqual(result["unreviewed_precedents"], ["Paid"])
        self.assertEqual(result["unreviewed_precedent_count"], 1)
        saved_sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(saved_sidecar["status"], 0)

    def test_review_needed_save_names_missing_revised_precedents(self) -> None:
        self.write_selection()

        with self.assertRaises(HTTPException) as raised:
            result_selection_service.save_result_selection(
                "Project",
                "Class",
                self.method_payload(),
            )

        self.assertEqual(raised.exception.status_code, 409)
        self.assertIn("are missing: Paid", str(raised.exception.detail))

    def test_missing_new_precedent_rolls_back_method_csv_and_sidecar(self) -> None:
        with (
            mock.patch.object(calculated_dataset_service, "apply_sidecar_graph_fields"),
            mock.patch.object(calculated_dataset_service, "recalculate_dependents", return_value={"ok": True}),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            with self.assertRaises(FileNotFoundError):
                result_selection_service.save_result_selection(
                    "Project",
                    "Class",
                    self.method_payload(),
                )

        self.assertFalse((self.methods / "RS@Selection.json").exists())
        self.assertFalse((self.datasets / "Selection@12.csv").exists())
        self.assertFalse((self.sidecars / "Selection.json").exists())

    def test_save_failure_after_graph_registration_restores_source_and_output_files(self) -> None:
        self.write_source("Paid", [10, 20])
        paid_path = self.sidecars / "Paid.json"
        paid = json.loads(paid_path.read_text(encoding="utf-8"))
        paid["dependents"] = []
        self.write_json(paid_path, paid)

        with (
            mock.patch.object(calculated_dataset_service, "apply_sidecar_graph_fields"),
            mock.patch.object(
                result_selection_service.dataset_sidecar_status_service,
                "refresh_method_statuses_for_dependents",
                side_effect=OSError("status write failed"),
            ),
        ):
            with self.assertRaises(OSError):
                result_selection_service.save_result_selection(
                    "Project",
                    "Class",
                    self.method_payload(),
                )

        restored_paid = json.loads(paid_path.read_text(encoding="utf-8"))
        self.assertEqual(restored_paid["dependents"], [])
        self.assertFalse((self.methods / "RS@Selection.json").exists())
        self.assertFalse((self.datasets / "Selection@12.csv").exists())
        self.assertFalse((self.sidecars / "Selection.json").exists())

    def test_save_rejects_self_dependency_before_writing_files(self) -> None:
        payload = self.method_payload()
        payload["method_tab"]["loaded_datasets"][0]["name"] = "Selection"

        with self.assertRaises(HTTPException) as raised:
            result_selection_service.save_result_selection("Project", "Class", payload)

        self.assertEqual(raised.exception.status_code, 422)
        self.assertIn("own output", str(raised.exception.detail))
        self.assertFalse((self.methods / "RS@Selection.json").exists())
        self.assertFalse((self.datasets / "Selection@12.csv").exists())
        self.assertFalse((self.sidecars / "Selection.json").exists())

    def test_source_update_refreshes_method_output_and_preserves_override(self) -> None:
        self.write_selection()
        self.write_source("Paid", [30, 40])
        with (
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={"updated": []},
            ),
            mock.patch(
                "app_server.services.dataset_instance_index_service.rebuild_index",
            ),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"])
        self.assertEqual(result["updated"], [{"dataset_name": "Selection"}])
        saved = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["loaded_datasets"][0]["values"], [30.0, 40.0])
        self.assertEqual(saved["method_tab"]["calculated_ultimate"], [30.0, 40.0])
        self.assertEqual(saved["method_tab"]["selected_ultimate"], [30.0, 99])
        self.assertEqual((self.datasets / "Selection@12.csv").read_text(encoding="utf-8"), "30.0\n99\n")
        sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(sidecar["status"], 2)
        self.assertEqual(
            result["review_status_updates"],
            [{"dataset_name": "Selection", "status": 2}],
        )

    def test_a_refresh_records_itself_without_moving_the_user_save_stamp(self) -> None:
        # The ResQ sync reads ``last_modified`` as the last edit a person made.
        # A propagation refresh recomputes from inputs nobody edited here, so it
        # stamps ``data_refreshed`` instead, as DFM, BF and Cape Cod already do.
        from app_server.helpers import parse_method_last_modified_timestamp

        self.write_selection()
        before = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        self.write_source("Paid", [30, 40])
        with (
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={"updated": []},
            ),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"])
        saved = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(
            parse_method_last_modified_timestamp(saved["method_metadata"]["last_modified"]),
            parse_method_last_modified_timestamp(before["method_metadata"]["last_modified"]),
        )
        refreshed_at = parse_method_last_modified_timestamp(saved["method_metadata"]["data_refreshed"])
        self.assertIsNotNone(refreshed_at)
        self.assertGreater(refreshed_at, parse_method_last_modified_timestamp(before["method_metadata"]["last_modified"]))

    def test_dataset_save_refreshes_transitive_result_selection_chain_but_keeps_review_alerts(self) -> None:
        self.write_selection()
        first_sidecar_path = self.sidecars / "Selection.json"
        first_sidecar = json.loads(first_sidecar_path.read_text(encoding="utf-8"))
        first_sidecar["dependents"] = ["Selection Two"]
        self.write_json(first_sidecar_path, first_sidecar)

        second = self.method_payload()
        second["details_tab"]["name"] = "Selection Two"
        second["method_tab"]["loaded_datasets"][0].update({
            "name": "Selection",
            "dataset_type": "Selected Ultimate",
            "method_type": "Result Selection",
            "source_kind": "result_selection",
            "values": [10, 99],
        })
        second["method_tab"]["calculated_ultimate"] = [10, 99]
        second["method_tab"]["selected_ultimate"] = [10, 99]
        second["method_tab"]["ultimate_overrides"] = [None, None]
        self.write_json(self.methods / "RS@Selection Two.json", second)
        self.write_json(self.sidecars / "Selection Two.json", {
            "dataset_name": "Selection Two",
            "dataset_type": "Selected Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Selection Two@12.csv",
            "status": 2,
            "precedents": ["Selection"],
            "dependents": [],
            "audit_log": [],
        })
        (self.datasets / "Selection Two@12.csv").write_text("10\n99\n", encoding="utf-8")
        self.write_source("Paid", [10, 20])
        with (
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=[]),
            mock.patch.object(calculated_dataset_service, "_existing_downstream_keys", return_value=[]),
            mock.patch.object(
                calculated_dataset_service,
                "sidecar_graph_fields",
                return_value={"precedents": [], "dependents": []},
            ),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
            mock.patch.object(
                dataset_service.dependent_propagation_service,
                "require_engine_available",
            ),
            mock.patch.object(
                dataset_service.dependent_propagation_service,
                "enqueue_save_propagation",
                return_value={"ok": True, "job_id": "job-1", "status": "queued"},
            ) as enqueue,
        ):
            result = dataset_service.save_dataset_sidecar(
                "Project",
                "Class",
                "Paid",
                dataset_type="Paid",
                instance_name="Paid",
                source_kind="input",
                data_format="Vector",
                origin_length=12,
                development_length=12,
                origin_labels=["2025", "2026"],
                values=[[30], [40]],
            )
            # The save only enqueues the Engine job; run the same canonical
            # walk the Engine executes for the enqueued root.
            walk = calculated_dataset_service.recalculate_dependents(
                "Project", "Class", "Paid", "Paid"
            )

        self.assertTrue(result["propagation_ok"], result)
        enqueue.assert_called_once_with(
            "Project",
            "Class",
            [{"dataset_name": "Paid", "dataset_type": "Paid"}],
        )
        self.assertEqual(
            result["calculated_updates"],
            {"ok": True, "job_id": "job-1", "status": "queued"},
        )
        self.assertEqual(
            walk["result_selection_updates"]["updated"],
            [{"dataset_name": "Selection"}, {"dataset_name": "Selection Two"}],
        )
        saved = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["loaded_datasets"][0]["values"], [30.0, 40.0])
        self.assertEqual(saved["method_tab"]["selected_ultimate"], [30.0, 99])
        source_sidecar = json.loads((self.sidecars / "Paid.json").read_text(encoding="utf-8"))
        self.assertEqual(source_sidecar["status"], 0)
        sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(sidecar["status"], 2)
        downstream_method = json.loads(
            (self.methods / "RS@Selection Two.json").read_text(encoding="utf-8")
        )
        self.assertEqual(
            downstream_method["method_tab"]["loaded_datasets"][0]["values"],
            [30.0, 99.0],
        )
        downstream_sidecar = json.loads(
            (self.sidecars / "Selection Two.json").read_text(encoding="utf-8")
        )
        self.assertEqual(downstream_sidecar["status"], 2)

    def test_source_update_uses_matching_aggregate_period_instead_of_native_rows(self) -> None:
        self.write_selection()
        method = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        method["method_tab"]["loaded_datasets"][0]["origin_length"] = 3
        self.write_json(self.methods / "RS@Selection.json", method)
        (self.datasets / "Paid@3.csv").write_text("1\n2\n3\n4\n", encoding="utf-8")
        (self.datasets / "Paid@12.csv").write_text("300\n400\n", encoding="utf-8")
        self.write_json(self.sidecars / "Paid.json", {
            "dataset_name": "Paid",
            "dataset_type": "Paid",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "dfm",
            "method_type": "DFM",
            "data_format": "Vector",
            "period_length": 3,
            "stored_period_length": 3,
            "csv_file": "Paid@3.csv",
            "dependents": ["Selection"],
        })
        with (
            mock.patch("app_server.services.calculated_dataset_service.recalculate_dependents", return_value={"updated": []}),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"], result)
        saved = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["loaded_datasets"][0]["values"], [300.0, 400.0])

    def test_monthly_manual_source_is_rolled_up_and_a_stale_yearly_copy_ignored(self) -> None:
        self.write_selection()
        (self.datasets / "Paid@1.csv").write_text(
            "".join(f"{value}\n" for value in range(1, 25)), encoding="utf-8"
        )
        # Left behind by an earlier release; the stored monthly file is the
        # only copy a hand-entered dataset is ever read from.
        (self.datasets / "Paid@12.csv").write_text("1\n2\n", encoding="utf-8")
        self.write_json(self.sidecars / "Paid.json", {
            "dataset_name": "Paid",
            "dataset_type": "Paid",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "input",
            "method_type": "None",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 1,
            "csv_file": "Paid@1.csv",
            "dependents": ["Selection"],
        })
        with (
            mock.patch("app_server.services.calculated_dataset_service.recalculate_dependents", return_value={"updated": []}),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"], result)
        saved = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["loaded_datasets"][0]["values"], [78.0, 222.0])

    def test_review_needed_method_reloads_every_precedent_but_waits_for_save_acknowledgement(self) -> None:
        self.write_selection()
        method_path = self.methods / "RS@Selection.json"
        method = json.loads(method_path.read_text(encoding="utf-8"))
        first = method["method_tab"]["loaded_datasets"][0]
        first["name"] = "Paid A"
        first["values"] = [10, 20]
        second = {**first, "name": "Paid B", "values": [20, 30]}
        method["method_tab"]["loaded_datasets"] = [first, second]
        method["method_tab"]["ultimate_overrides"] = [None, None]
        self.write_json(method_path, method)
        sidecar_path = self.sidecars / "Selection.json"
        sidecar = json.loads(sidecar_path.read_text(encoding="utf-8"))
        sidecar["precedents"] = ["Paid A", "Paid B"]
        sidecar["status"] = 2
        self.write_json(sidecar_path, sidecar)
        self.write_source("Paid A", [30, 50])
        self.write_source("Paid B", [50, 70])

        with (
            mock.patch("app_server.services.calculated_dataset_service.recalculate_dependents", return_value={"ok": True, "updated": []}),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Paid B"])

        self.assertTrue(result["ok"], result)
        saved = json.loads(method_path.read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["loaded_datasets"][0]["values"], [30.0, 50.0])
        self.assertEqual(saved["method_tab"]["loaded_datasets"][1]["values"], [50.0, 70.0])
        self.assertEqual(saved["method_tab"]["selected_ultimate"], [40.0, 60.0])
        self.assertEqual(json.loads(sidecar_path.read_text(encoding="utf-8"))["status"], 2)

    def test_review_refresh_missing_precedent_sidecar_preserves_artifacts(self) -> None:
        self.write_selection()
        method_path = self.methods / "RS@Selection.json"
        method = json.loads(method_path.read_text(encoding="utf-8"))
        first = method["method_tab"]["loaded_datasets"][0]
        first["name"] = "Paid A"
        second = {**first, "name": "Paid B", "values": [20, 30]}
        method["method_tab"]["loaded_datasets"] = [first, second]
        method["method_tab"]["ultimate_overrides"] = [None, None]
        self.write_json(method_path, method)
        sidecar_path = self.sidecars / "Selection.json"
        sidecar = json.loads(sidecar_path.read_text(encoding="utf-8"))
        sidecar["precedents"] = ["Paid A", "Paid B"]
        sidecar["status"] = 2
        self.write_json(sidecar_path, sidecar)
        self.write_source("Paid B", [50, 70])
        artifact_paths = [
            method_path,
            self.datasets / "Selection@12.csv",
            sidecar_path,
        ]
        before = {path: path.read_bytes() for path in artifact_paths}

        result = result_selection_service.refresh_dependents("Project", "Class", ["Paid B"])

        self.assertFalse(result["ok"])
        self.assertEqual(result["updated"], [])
        self.assertIn("Required precedent sidecar is missing: Paid A", result["errors"][0]["reason"])
        self.assertEqual({path: path.read_bytes() for path in artifact_paths}, before)

    def test_failed_calculated_fan_in_does_not_publish_mixed_result_selection(self) -> None:
        self.write_selection()
        method_path = self.methods / "RS@Selection.json"
        method = json.loads(method_path.read_text(encoding="utf-8"))
        first = method["method_tab"]["loaded_datasets"][0]
        first["name"] = "Calculated A"
        second = {**first, "name": "Paid B", "values": [20, 30]}
        method["method_tab"]["loaded_datasets"] = [first, second]
        method["method_tab"]["ultimate_overrides"] = [None, None]
        self.write_json(method_path, method)
        sidecar_path = self.sidecars / "Selection.json"
        sidecar_before = json.loads(sidecar_path.read_text(encoding="utf-8"))
        sidecar_before["precedents"] = ["Calculated A", "Paid B"]
        sidecar_before["status"] = 0
        self.write_json(sidecar_path, sidecar_before)
        self.write_source("Calculated A", [10, 20])
        calculated_sidecar_path = self.sidecars / "Calculated A.json"
        calculated_sidecar = json.loads(calculated_sidecar_path.read_text(encoding="utf-8"))
        calculated_sidecar.update({
            "source_kind": "calculated",
            "calculated": True,
            "status": 0,
        })
        self.write_json(calculated_sidecar_path, calculated_sidecar)
        self.write_source("Paid B", [50, 70])
        method_before = method_path.read_bytes()
        output_before = (self.datasets / "Selection@12.csv").read_bytes()
        rows = [
            {"name": "Paid B", "calculated": False, "generated": False, "formula": ""},
            {"name": "Calculated A", "calculated": True, "generated": False, "formula": '"Paid B"'},
        ]

        with (
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=rows),
            mock.patch.object(
                calculated_dataset_service,
                "_existing_downstream_keys",
                return_value=["calculated a"],
            ),
            mock.patch.object(calculated_dataset_service, "recalculate_dataset", return_value={
                "ok": False,
                "dataset_type_name": "Calculated A",
                "reason": "formula_error",
                "errors": ["bad formula"],
            }),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = calculated_dataset_service.recalculate_dependents(
                "Project",
                "Class",
                "Paid B",
                "Paid B",
            )

        rs_result = result["result_selection_updates"]
        self.assertFalse(result["ok"])
        self.assertFalse(rs_result["ok"])
        self.assertEqual(rs_result["updated"], [])
        self.assertIn("Precedent refresh failed: Calculated A", rs_result["errors"][0]["reason"])
        self.assertEqual(method_path.read_bytes(), method_before)
        self.assertEqual((self.datasets / "Selection@12.csv").read_bytes(), output_before)
        # The review mark is the only change, written in the v4 sidecar shape:
        # the stamp first and the audit log last.
        expected_sidecar = finalize_sidecar({**sidecar_before, "status": 2})
        self.assertEqual(
            list(json.loads(sidecar_path.read_text(encoding="utf-8")).items()),
            list(expected_sidecar.items()),
        )

    def test_successful_calculated_output_is_not_blocked_as_a_non_rs_edge(self) -> None:
        self.write_selection()
        method_path = self.methods / "RS@Selection.json"
        method = json.loads(method_path.read_text(encoding="utf-8"))
        method["method_tab"]["loaded_datasets"][0]["name"] = "Calculated A"
        self.write_json(method_path, method)
        selection_sidecar_path = self.sidecars / "Selection.json"
        selection_sidecar = json.loads(selection_sidecar_path.read_text(encoding="utf-8"))
        selection_sidecar["precedents"] = ["Calculated A"]
        self.write_json(selection_sidecar_path, selection_sidecar)
        self.write_source("Calculated A", [30, 40])
        calculated_sidecar_path = self.sidecars / "Calculated A.json"
        calculated_sidecar = json.loads(calculated_sidecar_path.read_text(encoding="utf-8"))
        calculated_sidecar.update({"source_kind": "calculated", "calculated": True})
        self.write_json(calculated_sidecar_path, calculated_sidecar)
        self.write_source("Root", [1, 2])
        root_sidecar_path = self.sidecars / "Root.json"
        root_sidecar = json.loads(root_sidecar_path.read_text(encoding="utf-8"))
        root_sidecar["dependents"] = ["Calculated A"]
        self.write_json(root_sidecar_path, root_sidecar)

        with (
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={"ok": True, "updated": [], "skipped": []},
            ),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.refresh_dependents(
                "Project",
                "Class",
                ["Root", "Calculated A"],
            )

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["updated"], [{"dataset_name": "Selection"}])
        saved = json.loads(method_path.read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["loaded_datasets"][0]["values"], [30.0, 40.0])

    def test_stale_non_rs_method_output_is_not_traversed_as_fresh(self) -> None:
        method = self.method_payload()
        method["method_tab"]["loaded_datasets"][0].update({
            "name": "Stale DFM",
            "dataset_type": "DFM Ultimate",
            "method_type": "DFM",
            "source_kind": "dfm",
        })
        self.write_json(self.methods / "RS@Selection.json", method)
        self.write_json(self.sidecars / "Selection.json", {
            "dataset_name": "Selection",
            "dataset_type": "Selected Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Selection@12.csv",
            "status": 2,
            "precedents": ["Stale DFM"],
            "dependents": [],
        })
        self.write_json(self.sidecars / "Stale DFM.json", {
            "dataset_name": "Stale DFM",
            "dataset_type": "DFM Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "dfm",
            "method_type": "DFM",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Stale DFM@12.csv",
            "status": 2,
            "precedents": ["Root"],
            "dependents": ["Selection"],
        })
        self.write_source("Root", [30, 40])
        root = json.loads((self.sidecars / "Root.json").read_text(encoding="utf-8"))
        root["dependents"] = ["Stale DFM"]
        self.write_json(self.sidecars / "Root.json", root)
        before = (self.methods / "RS@Selection.json").read_bytes()

        result = result_selection_service.refresh_dependents("Project", "Class", ["Root"])

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["updated"], [])
        self.assertEqual((self.methods / "RS@Selection.json").read_bytes(), before)
        self.assertIn("non_result_selection_dependent", result["skipped"][0]["reason"])

    def test_obsolete_finer_period_cache_is_not_accepted_after_origin_change(self) -> None:
        obsolete = self.datasets / "Selection@3.csv"
        obsolete.write_text("1\n2\n", encoding="utf-8")
        sidecar = {
            "dataset_name": "Selection",
            "data_format": "Vector",
            "source_kind": "result_selection",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Selection@12.csv",
        }

        with self.assertRaisesRegex(RuntimeError, "cannot be derived"):
            precedent_cache_service.precedent_csv_path(
                "Project", "Class", "Selection", sidecar, 3, exact=True,
            )

    def test_diamond_dependency_reprocesses_a_child_after_fan_in_updates(self) -> None:
        def write_rs(name: str, sources: list[dict], dependents: list[str]) -> None:
            method = self.method_payload()
            method["details_tab"]["name"] = name
            method["method_tab"]["loaded_datasets"] = sources
            method["method_tab"]["ultimate_overrides"] = [None, None]
            method["method_tab"]["calculated_ultimate"] = [1, 1]
            method["method_tab"]["selected_ultimate"] = [1, 1]
            self.write_json(self.methods / f"RS@{name}.json", method)
            self.write_json(self.sidecars / f"{name}.json", {
                "dataset_name": name,
                "dataset_type": "Selected Ultimate",
                "project_name": "Project",
                "reserving_class": "Class",
                "source_kind": "result_selection",
                "method_type": "Result Selection",
                "data_format": "Vector",
                "period_length": 12,
                "stored_period_length": 12,
                "csv_file": f"{name}@12.csv",
                "status": 2,
                "precedents": [source["name"] for source in sources],
                "dependents": dependents,
                "audit_log": [],
            })
            (self.datasets / f"{name}@12.csv").write_text("1\n1\n", encoding="utf-8")

        def source(name: str, weight: float = 1.0) -> dict:
            return {
                "name": name,
                "dataset_type": name,
                "data_format": "Vector",
                "method_type": "None",
                "category": "Loss",
                "source_kind": "input",
                "origin_length": 12,
                "values": [1, 1],
                "weights": [weight, weight],
            }

        write_rs("X", [source("Root")], ["Z"])
        write_rs("Z", [source("Root", 0), source("X", 1)], ["Achild"])
        write_rs("Achild", [source("Z")], [])
        self.write_source("Root", [5, 5])
        root_sidecar = json.loads((self.sidecars / "Root.json").read_text(encoding="utf-8"))
        root_sidecar["dependents"] = ["X", "Z"]
        self.write_json(self.sidecars / "Root.json", root_sidecar)

        with (
            mock.patch("app_server.services.calculated_dataset_service.recalculate_dependents", return_value={"updated": []}),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Root"])

        self.assertTrue(result["ok"], result)
        child = json.loads((self.methods / "RS@Achild.json").read_text(encoding="utf-8"))
        self.assertEqual(child["method_tab"]["selected_ultimate"], [5.0, 5.0])

    def test_refresh_cascade_accumulates_failed_outputs_before_later_fan_in(self) -> None:
        def write_sidecar(name: str, method_type: str, dependents: list[str]) -> None:
            self.write_json(self.sidecars / f"{name}.json", {
                "dataset_name": name,
                "dataset_type": name,
                "project_name": "Project",
                "reserving_class": "Class",
                "source_kind": "result_selection" if method_type == "Result Selection" else "input",
                "method_type": method_type,
                "status": 0,
                "dependents": dependents,
            })

        write_sidecar("Root", "None", ["Broken", "Healthy"])
        write_sidecar("Broken", "Result Selection", ["Fan In"])
        write_sidecar("Healthy", "Result Selection", ["Fan In"])
        write_sidecar("Fan In", "Result Selection", [])
        blocked_at_fan_in: set[str] = set()

        def refresh_one(
            _project: str,
            _reserving: str,
            output_name: str,
            output_sidecar: dict,
            _changed_sidecars: dict,
            _cache: dict,
            *,
            allow_status_current: bool,
            blocked_precedent_keys: set[str],
            sidecar_snapshot: dict[str, dict],
        ) -> dict:
            self.assertTrue(allow_status_current)
            self.assertTrue(sidecar_snapshot)
            if output_name == "Broken":
                return {"ok": False, "dataset_name": output_name, "reason": "refresh failed"}
            if output_name == "Healthy":
                return {
                    "ok": True,
                    "dataset_name": output_name,
                    "updated": True,
                    "output_changed": True,
                    "sidecar": output_sidecar,
                }
            blocked_at_fan_in.update(blocked_precedent_keys)
            return {"ok": False, "dataset_name": output_name, "reason": "fan-in blocked"}

        with (
            mock.patch.object(result_selection_service, "_refresh_one_method", side_effect=refresh_one),
            mock.patch.object(
                result_selection_service.dataset_sidecar_status_service,
                "refresh_method_statuses_for_dependents",
                return_value=[],
            ),
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={
                    "ok": False,
                    "updated": [],
                    "skipped": [{
                        "dataset_type_name": "Calculated Broken",
                        "reason": "formula_error",
                    }],
                },
            ),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Root"])

        self.assertFalse(result["ok"])
        self.assertEqual(blocked_at_fan_in, {"broken", "calculated broken"})

    def test_refresh_reuses_graph_snapshot_and_reloads_only_status_mutations(self) -> None:
        self.write_selection()
        self.write_source("Paid", [30, 40])
        downstream = self.method_payload()
        downstream["details_tab"]["name"] = "Selection Two"
        downstream["method_tab"]["loaded_datasets"][0].update({
            "name": "Selection",
            "dataset_type": "Selected Ultimate",
            "method_type": "Result Selection",
            "source_kind": "result_selection",
            "values": [10, 99],
        })
        downstream["method_tab"]["calculated_ultimate"] = [10, 99]
        downstream["method_tab"]["selected_ultimate"] = [10, 99]
        self.write_json(self.methods / "RS@Selection Two.json", downstream)
        self.write_json(self.sidecars / "Selection Two.json", {
            "dataset_name": "Selection Two",
            "dataset_type": "Selected Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Selection Two@12.csv",
            "status": 0,
            "precedents": ["Selection"],
            "dependents": [],
            "audit_log": [],
        })
        (self.datasets / "Selection Two@12.csv").write_text("10\n99\n", encoding="utf-8")
        selection_path = self.sidecars / "Selection.json"
        selection = json.loads(selection_path.read_text(encoding="utf-8"))
        selection["dependents"] = ["Selection Two"]
        self.write_json(selection_path, selection)

        original_read_sidecars = result_selection_service._read_sidecars
        read_counts: dict[str, int] = {}

        def recording_read(project: str, reserving: str, names) -> dict:
            ordered = result_selection_service._unique_names(list(names))
            for name in ordered:
                key = result_selection_service._key(name)
                read_counts[key] = read_counts.get(key, 0) + 1
            return original_read_sidecars(project, reserving, ordered)

        with (
            mock.patch.object(result_selection_service, "_read_sidecars", side_effect=recording_read),
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={"ok": True, "updated": [], "skipped": []},
            ),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"], result)
        self.assertEqual(read_counts["paid"], 1)
        self.assertEqual(read_counts["selection"], 1)
        self.assertEqual(read_counts["selection two"], 2)
        saved = json.loads((self.methods / "RS@Selection Two.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["selected_ultimate"], [30.0, 99])

    def test_ratio_basis_only_update_persists_statistics_values_without_changing_output(self) -> None:
        self.write_selection(basis=True)
        self.write_source("Paid", [10, 20])
        self.write_source("Premium", [300, 400])
        with (
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={"updated": []},
            ),
            mock.patch(
                "app_server.services.dataset_instance_index_service.rebuild_index",
            ),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Premium"])

        self.assertTrue(result["ok"])
        saved = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(
            saved["method_tab"]["ratio_basis_values"],
            [{"name": "Premium", "values": [300.0, 400.0]}],
        )
        self.assertEqual(saved["method_tab"]["selected_ultimate"], [10.0, 99])
        self.assertEqual((self.datasets / "Selection@12.csv").read_text(encoding="utf-8"), "10.0\n99\n")

    def test_unchanged_upstream_values_keep_review_alert_without_rewriting_method(self) -> None:
        self.write_selection()
        self.write_source("Paid", [10, 20])
        downstream = self.method_payload()
        downstream["details_tab"]["name"] = "Selection Two"
        downstream["method_tab"]["loaded_datasets"][0]["name"] = "Selection"
        downstream["method_tab"]["loaded_datasets"][0]["dataset_type"] = "Selected Ultimate"
        downstream["method_tab"]["loaded_datasets"][0]["method_type"] = "Result Selection"
        downstream["method_tab"]["loaded_datasets"][0]["source_kind"] = "result_selection"
        downstream["method_tab"]["loaded_datasets"][0]["values"] = [10, 99]
        downstream["method_tab"]["calculated_ultimate"] = [10, 99]
        self.write_json(self.methods / "RS@Selection Two.json", downstream)
        self.write_json(self.sidecars / "Selection Two.json", {
            "dataset_name": "Selection Two",
            "dataset_type": "Selected Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Selection Two@12.csv",
            "status": 2,
            "precedents": ["Selection"],
            "dependents": [],
            "audit_log": [],
        })
        (self.datasets / "Selection Two@12.csv").write_text("10\n99\n", encoding="utf-8")
        selection_sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        selection_sidecar["dependents"] = ["Selection Two"]
        self.write_json(self.sidecars / "Selection.json", selection_sidecar)
        method_path = self.methods / "RS@Selection.json"
        method_before = method_path.read_bytes()
        downstream_method_path = self.methods / "RS@Selection Two.json"
        downstream_method_before = downstream_method_path.read_bytes()

        with (
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={"updated": []},
            ),
            mock.patch(
                "app_server.services.dataset_instance_index_service.rebuild_index",
            ) as rebuild_index,
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["updated"], [])
        self.assertEqual(
            result["status_refreshed"],
            [{"dataset_name": "Selection"}, {"dataset_name": "Selection Two"}],
        )
        self.assertEqual(method_path.read_bytes(), method_before)
        self.assertEqual(downstream_method_path.read_bytes(), downstream_method_before)
        saved_sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(saved_sidecar["status"], 2)
        downstream_sidecar = json.loads((self.sidecars / "Selection Two.json").read_text(encoding="utf-8"))
        self.assertEqual(downstream_sidecar["status"], 2)
        rebuild_index.assert_called_once_with("Project", "Class")

    def test_unchanged_output_does_not_block_its_non_rs_dependents_for_a_later_fan_in(self) -> None:
        # Mirrors the live NJ walk: a BF output is only status-refreshed, so
        # "F 91" (Selection) re-reads unchanged inputs; its calculated
        # dependent "G 23" (Calc A) is not a Result Selection and gets
        # skipped; then "G 91" (Selection Two), which loads both, must not be
        # refused with "Precedent refresh failed: G 23" over that skip.
        self.write_selection()
        self.write_source("Paid", [10, 20])
        self.write_source("Calc A", [30, 40])
        calc_sidecar_path = self.sidecars / "Calc A.json"
        calc_sidecar = json.loads(calc_sidecar_path.read_text(encoding="utf-8"))
        calc_sidecar.update({
            "source_kind": "calculated",
            "calculated": True,
            "status": 2,
            "precedents": ["Selection"],
            "dependents": ["Selection Two"],
        })
        self.write_json(calc_sidecar_path, calc_sidecar)
        downstream = self.method_payload()
        downstream["details_tab"]["name"] = "Selection Two"
        first = downstream["method_tab"]["loaded_datasets"][0]
        first.update({
            "name": "Selection",
            "dataset_type": "Selected Ultimate",
            "method_type": "Result Selection",
            "source_kind": "result_selection",
            "values": [10, 99],
        })
        second = {
            **first,
            "name": "Calc A",
            "dataset_type": "Calc A",
            "method_type": "None",
            "source_kind": "calculated",
            "values": [30, 40],
        }
        downstream["method_tab"]["loaded_datasets"] = [first, second]
        downstream["method_tab"]["calculated_ultimate"] = [20, 69.5]
        downstream["method_tab"]["selected_ultimate"] = [20, 99]
        self.write_json(self.methods / "RS@Selection Two.json", downstream)
        self.write_json(self.sidecars / "Selection Two.json", {
            "dataset_name": "Selection Two",
            "dataset_type": "Selected Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Selection Two@12.csv",
            "status": 2,
            "precedents": ["Selection", "Calc A"],
            "dependents": [],
            "audit_log": [],
        })
        (self.datasets / "Selection Two@12.csv").write_text("20\n99\n", encoding="utf-8")
        selection_sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        selection_sidecar["dependents"] = ["Calc A", "Selection Two"]
        self.write_json(self.sidecars / "Selection.json", selection_sidecar)
        downstream_method_path = self.methods / "RS@Selection Two.json"
        downstream_method_before = downstream_method_path.read_bytes()

        with (
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={"updated": []},
            ) as recalculate_dependents,
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertEqual(result["errors"], [])
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["updated"], [])
        self.assertEqual(
            result["status_refreshed"],
            [{"dataset_name": "Selection"}, {"dataset_name": "Selection Two"}],
        )
        skipped_by_name = {item["dataset_name"]: item["reason"] for item in result["skipped"]}
        self.assertEqual(skipped_by_name["Calc A"], "non_result_selection_dependent_inputs_unchanged")
        # Nothing changed in value, so no calculated cascade ran for the skip.
        recalculate_dependents.assert_not_called()
        self.assertEqual(downstream_method_path.read_bytes(), downstream_method_before)

    def test_dfm_visited_to_the_same_publication_is_not_a_failed_precedent(self) -> None:
        # Mirrors the live HOL walk after a data-processing-rules save: the
        # regenerated "Total Earned Exposure" (Paid) is the ratio basis of
        # "C 12 - CWP DFM" (CWP DFM), so the DFM wave recomputed it to the
        # same ultimate; "C 91" (Selection) loads both. Reached again here as
        # the root's dependent, the DFM must be taken as current, not blocked.
        self.write_selection()
        self.write_source("Paid", [10, 20])
        self.write_source("CWP DFM", [5, 6])
        dfm_sidecar_path = self.sidecars / "CWP DFM.json"
        dfm_sidecar = json.loads(dfm_sidecar_path.read_text(encoding="utf-8"))
        dfm_sidecar.update({
            "source_kind": "dfm",
            "method_type": "DFM",
            "method_name": "CWP DFM",
            "status": 0,
            "precedents": ["Paid"],
            "dependents": ["Selection"],
        })
        self.write_json(dfm_sidecar_path, dfm_sidecar)
        paid_sidecar = json.loads((self.sidecars / "Paid.json").read_text(encoding="utf-8"))
        paid_sidecar["dependents"] = ["CWP DFM", "Selection"]
        self.write_json(self.sidecars / "Paid.json", paid_sidecar)
        method = self.method_payload()
        first = method["method_tab"]["loaded_datasets"][0]
        method["method_tab"]["loaded_datasets"] = [first, {
            **first,
            "name": "CWP DFM",
            "dataset_type": "CWP DFM",
            "method_type": "DFM",
            "source_kind": "dfm",
            "values": [5, 6],
        }]
        # The stored selection predates the root's new values, so a refresh
        # that is allowed to run has something to publish.
        method["method_tab"]["loaded_datasets"][0]["values"] = [8, 16]
        method["method_tab"]["calculated_ultimate"] = [6.5, 11]
        method["method_tab"]["selected_ultimate"] = [7, 99]
        self.write_json(self.methods / "RS@Selection.json", method)
        selection_sidecar = json.loads((self.sidecars / "Selection.json").read_text(encoding="utf-8"))
        selection_sidecar["precedents"] = ["Paid", "CWP DFM"]
        self.write_json(self.sidecars / "Selection.json", selection_sidecar)
        (self.datasets / "Selection@12.csv").write_text("7\n99\n", encoding="utf-8")

        with (
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={"updated": []},
            ),
            mock.patch("app_server.services.dataset_instance_index_service.rebuild_index"),
        ):
            refused = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])
            result = result_selection_service.refresh_dependents(
                "Project",
                "Class",
                ["Paid"],
                unchanged_precedent_names=["CWP DFM"],
            )

        # Without the DFM wave's word the walk still blocks the DFM it cannot
        # refresh itself, and the Result Selection loading it is refused.
        self.assertFalse(refused["ok"])
        self.assertEqual(
            refused["errors"],
            [{"dataset_name": "Selection", "reason": "Precedent refresh failed: CWP DFM"}],
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["errors"], [])
        self.assertEqual(result["updated"], [{"dataset_name": "Selection"}], result)
        self.assertEqual(
            result["skipped"],
            [{"dataset_name": "CWP DFM", "reason": "non_result_selection_dependent_inputs_unchanged"}],
        )
        saved = json.loads((self.methods / "RS@Selection.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["calculated_ultimate"], [7.5, 13.0])
        self.assertEqual((self.datasets / "Selection@12.csv").read_text(encoding="utf-8"), "7.5\n99\n")

    def test_result_selection_refresh_propagates_transitively_once(self) -> None:
        first = self.method_payload()
        first["details_tab"]["name"] = "Selection One"
        first["method_tab"]["ultimate_overrides"] = [None, None]
        first["method_tab"]["selected_ultimate"] = [10, 20]
        self.write_json(self.methods / "RS@Selection One.json", first)
        self.write_json(self.sidecars / "Selection One.json", {
            "dataset_name": "Selection One",
            "dataset_type": "Selected Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Selection One@12.csv",
            "status": 2,
            "precedents": ["Paid"],
            "dependents": ["Selection Two"],
            "audit_log": [],
        })
        second = self.method_payload()
        second["details_tab"]["name"] = "Selection Two"
        second["method_tab"]["loaded_datasets"][0]["name"] = "Selection One"
        second["method_tab"]["loaded_datasets"][0]["source_kind"] = "result_selection"
        second["method_tab"]["ultimate_overrides"] = [None, None]
        second["method_tab"]["selected_ultimate"] = [10, 20]
        self.write_json(self.methods / "RS@Selection Two.json", second)
        self.write_json(self.sidecars / "Selection Two.json", {
            "dataset_name": "Selection Two",
            "dataset_type": "Selected Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "data_format": "Vector",
            "period_length": 12,
            "stored_period_length": 12,
            "csv_file": "Selection Two@12.csv",
            "status": 2,
            "precedents": ["Selection One"],
            "dependents": [],
            "audit_log": [],
        })
        self.write_source("Paid", [50, 60])
        paid = json.loads((self.sidecars / "Paid.json").read_text(encoding="utf-8"))
        paid["dependents"] = ["Selection One"]
        self.write_json(self.sidecars / "Paid.json", paid)

        with (
            mock.patch(
                "app_server.services.calculated_dataset_service.recalculate_dependents",
                return_value={"updated": []},
            ),
            mock.patch(
                "app_server.services.dataset_instance_index_service.rebuild_index",
            ),
        ):
            result = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"])
        self.assertEqual(
            result["updated"],
            [{"dataset_name": "Selection One"}, {"dataset_name": "Selection Two"}],
        )
        second_saved = json.loads((self.methods / "RS@Selection Two.json").read_text(encoding="utf-8"))
        self.assertEqual(second_saved["method_tab"]["selected_ultimate"], [50.0, 60.0])

    def test_refresh_read_failure_preserves_last_valid_method_csv_and_sidecar(self) -> None:
        self.write_selection()
        self.write_source("Paid", [30, 40])
        (self.datasets / "Paid@12.csv").unlink()
        paths = [
            self.methods / "RS@Selection.json",
            self.datasets / "Selection@12.csv",
            self.sidecars / "Selection.json",
        ]
        before = {path: path.read_bytes() for path in paths}

        result = result_selection_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertFalse(result["ok"])
        self.assertEqual(result["updated"], [])
        self.assertIn("Cached dataset CSV is missing", result["errors"][0]["reason"])
        self.assertEqual({path: path.read_bytes() for path in paths}, before)

    def test_calculated_cascade_always_invokes_result_selection_refresh(self) -> None:
        with (
            mock.patch.object(calculated_dataset_service, "_existing_downstream_keys", return_value=[]),
            mock.patch.object(calculated_dataset_service, "_calculated_rows_by_key", return_value={}),
            mock.patch(
                "app_server.services.dataset_instance_index_service.rebuild_index",
            ),
            mock.patch.object(
                result_selection_service,
                "refresh_dependents",
                return_value={"ok": True, "updated": [], "errors": []},
            ) as refresh,
        ):
            result = calculated_dataset_service.recalculate_dependents(
                "Project",
                "Class",
                "Paid Instance",
                "Paid Type",
            )

        self.assertTrue(result["ok"])
        refresh.assert_called_once_with(
            "Project",
            "Class",
            ["Paid Instance", "Paid Type"],
            rebuild_index=False,
            allow_status_current=True,
            blocked_precedent_names=[],
            unchanged_precedent_names=[],
            finalize_method_review_status=False,
        )

    def test_calculated_graph_rebuild_preserves_registered_method_dependents(self) -> None:
        payload = {
            "project_name": "Project",
            "reserving_class": "Class",
            "dataset_name": "Calculated",
            "dataset_type": "Calculated",
            "dependents": [{"dataset_name": "Selection"}],
        }
        with (
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=[]),
            mock.patch.object(
                calculated_dataset_service,
                "sidecar_graph_fields",
                return_value={
                    "precedents": [{"dataset_name": "Paid"}],
                    "dependents": [{"dataset_name": "Formula Output"}],
                },
            ),
            mock.patch(
                "app_server.services.dataset_sidecar_status_service.read_sidecar",
                return_value={"method_type": "Result Selection", "source_kind": "result_selection"},
            ),
        ):
            calculated_dataset_service.apply_sidecar_graph_fields(payload)

        self.assertEqual(
            payload["dependents"],
            [
                {"dataset_name": "Formula Output"},
                {"dataset_name": "Selection"},
            ],
        )

    def test_bulk_graph_refresh_preserves_method_owned_precedents(self) -> None:
        payload = {
            "project_name": "Project",
            "reserving_class": "Class",
            "dataset_name": "Selection",
            "dataset_type": "Selected Ultimate",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "precedents": ["Paid", "Premium"],
            "dependents": [],
        }
        with (
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=[]),
            mock.patch.object(
                calculated_dataset_service,
                "sidecar_graph_fields",
                return_value={
                    "precedents": [{"dataset_name": "Formula Input"}],
                    "dependents": [],
                },
            ),
        ):
            calculated_dataset_service.apply_sidecar_graph_fields(payload)

        self.assertEqual(payload["precedents"], ["Paid", "Premium"])


if __name__ == "__main__":
    unittest.main()

"""App-server behaviour for Bootstrap methods.

The interesting part of this domain is that a Bootstrap's first precedent is a
*method*: it embeds a DFM's observed triangle and selected ratios.  The
reserving-class dependency graph is keyed by dataset name, so these tests pin
that the DFM method is resolved to the dataset it publishes everywhere the graph
is touched, while the numbers are read from the DFM method JSON.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from fastapi import HTTPException


REPO_ROOT = Path(__file__).resolve().parents[2]
# Every test temp directory lives under one gitignored folder at the
# repository root, so a suite that dies before teardown cannot scatter
# tmp folders beside the code.
TEST_TEMP_ROOT = REPO_ROOT / "test"
TEST_TEMP_ROOT.mkdir(parents=True, exist_ok=True)
FRONTEND_ROOT = REPO_ROOT / "frontend"
PYTHON_API_SRC = REPO_ROOT / "python-api" / "src"
for path in (FRONTEND_ROOT, PYTHON_API_SRC):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from arcrho_api.bootstrap_contract import BST_JSON_FORMAT, normalize_bootstrap_method
from arcrho_api.dfm_contract import normalize_dfm_method
from arcrho_api.io import persisted_json_text
from arcrho_api.sidecar_core_contract import DATASET_SIDECAR_JSON_FORMAT, RETIRED_SIDECAR_FIELDS
from app_server.services import (
    bootstrap_service,
    calculated_dataset_service,
    dataset_sidecar_status_service,
)
from dependent_propagation_workspace_stub import IsolatedPropagationWorkspace

RESQ_FIXTURE_PATH = REPO_ROOT / "python-api" / "tests" / "fixtures" / "resq_bootstrap_f72a.json"
DFM_METHOD_NAME = "F 25 - Incurred DFM Bootstrap"
# Deliberately different from the DFM method name: the graph must reach the DFM
# by the dataset it publishes, not by the name the Bootstrap stores.
DFM_OUTPUT_DATASET = "F 25 Ultimate"
TARGET_NAME = "F 92 - Current Qtr Selected"
BOOTSTRAP_NAME = "F 72 A - Bootstrap Net incurred with PV"
# Keep the simulation cheap: these tests exercise plumbing, and the numbers are
# already pinned against live ResQ in python-api/tests/test_bootstrap_contract.py.
SIMULATION_COUNT = 200


class BootstrapServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=TEST_TEMP_ROOT)
        root = Path(self.temp.name)
        self.methods = root / "methods"
        self.datasets = root / "datasets"
        self.sidecars = root / "sidecars"
        for folder in (self.methods, self.datasets, self.sidecars):
            folder.mkdir()
        settings = root / "general_settings.json"
        settings.write_text(
            '{"origin_start_date":"202301","origin_end_date":"202412","development_end_date":"202412"}',
            encoding="utf-8",
        )
        self.patchers = [
            IsolatedPropagationWorkspace(),
            mock.patch.object(
                bootstrap_service.config,
                "get_general_settings_path",
                return_value=str(settings),
            ),
            mock.patch.object(
                bootstrap_service.config,
                "get_project_method_data_dir",
                return_value=str(self.methods),
            ),
            mock.patch.object(
                bootstrap_service.config,
                "get_project_dataset_cache_dir",
                return_value=str(self.datasets),
            ),
            mock.patch.object(
                dataset_sidecar_status_service,
                "sidecar_path",
                side_effect=lambda _p, _r, name: str(self.sidecars / f"{name}.json"),
            ),
            mock.patch.object(
                dataset_sidecar_status_service,
                "update_precedent_dependents",
                side_effect=self._record_graph_update,
            ),
            # The dataset-type catalogue and the reserving-class index both need
            # a real project folder; neither is what these tests are about.
            mock.patch.object(
                calculated_dataset_service, "_dataset_type_rows", return_value=[]
            ),
            mock.patch.object(
                calculated_dataset_service, "_existing_dataset_keys", return_value=set()
            ),
        ]
        for patcher in self.patchers:
            patcher.start()
        self.graph_updates: list[tuple] = []
        self.fixture = json.loads(RESQ_FIXTURE_PATH.read_text(encoding="utf-8"))
        self.case = self.fixture["methods"]["odp_single_scale"]
        self.reference = self.fixture["simulation_reference"]
        self.origin_labels = list(self.case["origin_labels"])
        self._write_dfm()
        self._write_target()

    def tearDown(self) -> None:
        for patcher in reversed(self.patchers):
            patcher.stop()
        self.temp.cleanup()

    # -- fixtures ---------------------------------------------------------

    def _record_graph_update(self, _project, _reserving, output, old, new, **kwargs):
        self.graph_updates.append((output, list(old), list(new), kwargs))
        return {}

    def _write_json(self, path: Path, payload: dict) -> None:
        path.write_text(persisted_json_text(payload), encoding="utf-8", newline="\n")

    def _dfm_payload(self, *, output_dataset: str = DFM_OUTPUT_DATASET) -> dict:
        triangle = self.case["observed_triangle"]
        development = [str(12 * (index + 1)) for index in range(len(self.origin_labels))]
        ratios = list(self.case["selected_ratios"])
        return normalize_dfm_method(
            {
                "json_format": "arcrho-dfm-v4",
                "details_tab": {
                    "name": DFM_METHOD_NAME,
                    "output_type": "F 00 - Ultimate Net Loss",
                    "output_dataset": output_dataset,
                    "output_category": "F Net Loss",
                    "input_triangle": "F 10 - Incurred",
                    "origin_length": 12,
                    "development_length": 12,
                },
                "data_tab": {
                    "origin_labels": self.origin_labels,
                    "development_labels": development,
                    "input_data_triangle_values": triangle,
                },
                "ratios_tab": {
                    "ratio_triangle": {"development_labels": development},
                    "average_formulas": {
                        "values": [ratios],
                        "selected": [[1] * len(ratios)],
                        "custom_average_formula_settings": {
                            "average_type": ["user_entry"],
                            "base": ["volume"],
                            "periods": ["all"],
                            "exclude": [0],
                        },
                    },
                },
                "results_tab": {},
                "method_metadata": {},
            },
            require_complete=False,
        )

    def _write_dfm(self, *, output_dataset: str = DFM_OUTPUT_DATASET) -> None:
        payload = self._dfm_payload(output_dataset=output_dataset)
        self._write_json(self.methods / f"DFM@{DFM_METHOD_NAME}.json", payload)
        self._write_json(
            self.sidecars / f"{output_dataset}.json",
            {
                "dataset_name": output_dataset,
                "dataset_type": "F 00 - Ultimate Net Loss",
                "method_name": DFM_METHOD_NAME,
                "method_type": dataset_sidecar_status_service.METHOD_TYPE_DFM,
                "source_kind": "dfm",
                "data_format": "Vector",
                "period_length": 12,
                "stored_period_length": 12,
                "csv_file": f"{output_dataset}@12.csv",
                "status": dataset_sidecar_status_service.STATUS_CURRENT,
                "origin_labels": self.origin_labels,
                "precedents": [],
                "dependents": [],
            },
        )

    def _target_values(self) -> list[float]:
        return [
            target + latest
            for target, latest in zip(
                self.reference["target_reserve_values"],
                self.reference["dfm_latest_values"],
            )
        ]

    def _write_target(self, values: list[float] | None = None) -> None:
        numbers = self._target_values() if values is None else values
        (self.datasets / f"{TARGET_NAME}@12.csv").write_text(
            "\n".join(str(value) for value in numbers) + "\n",
            encoding="utf-8",
            newline="\n",
        )
        self._write_json(
            self.sidecars / f"{TARGET_NAME}.json",
            {
                "dataset_name": TARGET_NAME,
                "dataset_type": "F 00 - Ultimate Net Loss",
                "method_name": TARGET_NAME,
                "method_type": dataset_sidecar_status_service.METHOD_TYPE_RESULT_SELECTION,
                "source_kind": "result_selection",
                "data_format": "Vector",
                "period_length": 12,
                "stored_period_length": 12,
                "csv_file": f"{TARGET_NAME}@12.csv",
                "status": dataset_sidecar_status_service.STATUS_CURRENT,
                "origin_labels": self.origin_labels,
                "precedents": [],
                "dependents": [],
            },
        )

    def method_payload(self, **overrides) -> dict:
        details = {
            "name": BOOTSTRAP_NAME,
            "output_type": "F 00 - Ultimate Net Loss",
            "dataset_category": "F Net Loss",
            "origin_length": 12,
            "development_length": 12,
            "model_type": self.case["model_type"],
            "dfm_method": DFM_METHOD_NAME,
        }
        details.update(overrides.pop("details_tab", {}))
        results = {
            "target_ultimate": TARGET_NAME,
            "target_scaling_methods": list(self.reference["target_scaling_methods"]),
        }
        results.update(overrides.pop("results_tab", {}))
        payload = {
            "json_format": BST_JSON_FORMAT,
            "details_tab": details,
            "residuals_tab": {},
            "simulation_tab": {
                "estimation_variance": self.reference["estimation_variance"],
                "process_variance": self.reference["process_variance"],
                "simulation_count": SIMULATION_COUNT,
                "random_seed": self.reference["random_seed"],
                "prevent_negative_data": self.reference["prevent_negative_data"],
                "negative_mean_action": self.reference["negative_mean_action"],
            },
            "results_tab": results,
            "output_tab": {},
            "method_metadata": {},
        }
        payload.update(overrides)
        return payload

    def save(self, payload: dict | None = None, **kwargs) -> dict:
        with mock.patch.object(
            calculated_dataset_service,
            "recalculate_dependents",
            return_value={"ok": True, "updated": [], "index_ok": True, "index_error": ""},
        ):
            return bootstrap_service.save_bootstrap_method(
                "Project",
                "Class",
                payload if payload is not None else self.method_payload(),
                **kwargs,
            )

    # -- tests ------------------------------------------------------------

    def test_save_publishes_the_method_json_the_csv_and_the_sidecar(self) -> None:
        response = self.save()

        self.assertTrue(response["ok"])
        self.assertEqual(response["method_name"], BOOTSTRAP_NAME)
        self.assertEqual(response["output_dataset"], BOOTSTRAP_NAME)
        method_file = self.methods / f"BST@{BOOTSTRAP_NAME}.json"
        csv_file = self.datasets / f"{BOOTSTRAP_NAME}@12.csv"
        sidecar_file = self.sidecars / f"{BOOTSTRAP_NAME}.json"
        for path in (method_file, csv_file, sidecar_file):
            self.assertTrue(path.exists(), path)
        self.assertEqual(
            len(csv_file.read_text(encoding="utf-8").splitlines()),
            len(self.origin_labels),
        )
        sidecar = json.loads(sidecar_file.read_text(encoding="utf-8"))
        # v4 core: the format stamp opens the file, the audit log closes it, and
        # the fields v4 retired (method_type_code, path, mtime, ...) stay out.
        self.assertEqual(list(sidecar)[0], "json_format")
        self.assertEqual(sidecar["json_format"], DATASET_SIDECAR_JSON_FORMAT)
        self.assertEqual(list(sidecar)[-1], "audit_log")
        self.assertEqual(RETIRED_SIDECAR_FIELDS & set(sidecar), set())
        self.assertEqual(sidecar["method_type"], "Bootstrap")
        self.assertEqual(sidecar["source_kind"], "bootstrap")
        self.assertEqual(sidecar["data_format"], "Vector")
        self.assertEqual(sidecar["period_length"], 12)
        self.assertEqual(sidecar["origin_labels"], self.origin_labels)
        self.assertEqual(
            sidecar["publication_revision"], response["publication_revision"]
        )

    def test_graph_precedents_use_the_dfm_output_dataset_not_the_method_name(self) -> None:
        self.save()

        sidecar = json.loads(
            (self.sidecars / f"{BOOTSTRAP_NAME}.json").read_text(encoding="utf-8")
        )
        self.assertEqual(
            dataset_sidecar_status_service.entry_names(sidecar["precedents"]),
            [DFM_OUTPUT_DATASET, TARGET_NAME],
        )
        self.assertEqual(len(self.graph_updates), 1)
        _output, old, new, kwargs = self.graph_updates[0]
        self.assertEqual(old, [])
        self.assertEqual(new, [DFM_OUTPUT_DATASET, TARGET_NAME])
        self.assertTrue(kwargs["require_new_precedents"])
        # The method JSON still stores the DFM by its method name.
        method = json.loads(
            (self.methods / f"BST@{BOOTSTRAP_NAME}.json").read_text(encoding="utf-8")
        )
        self.assertEqual(method["details_tab"]["dfm_method"], DFM_METHOD_NAME)

    def test_saved_method_embeds_the_dfm_snapshot_and_a_simulation_summary(self) -> None:
        response = self.save()

        details = response["method"]["details_tab"]
        results = response["method"]["results_tab"]
        # Triangle values are persisted at ArcRho's canonical six decimals.
        self.assertEqual(
            details["dfm_snapshot"]["observed_triangle"][0],
            [round(value, 6) for value in self.case["observed_triangle"][0]],
        )
        self.assertTrue(details["dfm_source_revision"].startswith("sha256:"))
        self.assertEqual(results["origin_labels"], self.origin_labels)
        self.assertEqual(
            results["simulation_summary"]["simulation_count"], SIMULATION_COUNT
        )
        self.assertEqual(
            results["simulation_summary"]["random_seed"],
            self.reference["random_seed"],
        )
        self.assertEqual(len(results["bootstrap_ultimate"]), len(self.origin_labels))
        self.assertTrue(all(value is not None for value in results["bootstrap_ultimate"]))
        # The simulated reserve array itself is never persisted.
        text = (self.methods / f"BST@{BOOTSTRAP_NAME}.json").read_text(encoding="utf-8")
        self.assertLess(len(text), 200 * 1024)

    def test_load_round_trips_the_saved_pair(self) -> None:
        saved = self.save()

        loaded = bootstrap_service.load_bootstrap_method("Project", "Class", BOOTSTRAP_NAME)

        self.assertEqual(loaded["publication_revision"], saved["publication_revision"])
        self.assertEqual(loaded["owned_revision"], saved["owned_revision"])
        self.assertEqual(
            persisted_json_text(loaded["method"]), persisted_json_text(saved["method"])
        )
        self.assertTrue(loaded["sidecar"]["exists"])

    def test_load_rejects_a_missing_method_and_a_missing_sidecar(self) -> None:
        with self.assertRaises(HTTPException) as missing:
            bootstrap_service.load_bootstrap_method("Project", "Class", "Nope")
        self.assertEqual(missing.exception.status_code, 404)

        self.save()
        (self.sidecars / f"{BOOTSTRAP_NAME}.json").unlink()
        with self.assertRaises(HTTPException) as orphan:
            bootstrap_service.load_bootstrap_method("Project", "Class", BOOTSTRAP_NAME)
        self.assertEqual(orphan.exception.status_code, 409)

    def test_load_rejects_an_unknown_json_format(self) -> None:
        self.save()
        path = self.methods / f"BST@{BOOTSTRAP_NAME}.json"
        payload = json.loads(path.read_text(encoding="utf-8"))
        payload["json_format"] = "arcrho-bootstrap-method-by-tab-v0"
        self._write_json(path, payload)

        with self.assertRaises(HTTPException) as ctx:
            bootstrap_service.load_bootstrap_method("Project", "Class", BOOTSTRAP_NAME)
        self.assertEqual(ctx.exception.status_code, 422)

    def test_stale_owned_revision_is_a_conflict_and_stale_derived_only_rebases(self) -> None:
        saved = self.save()

        payload = self.method_payload()
        with self.assertRaises(HTTPException) as ctx:
            self.save(payload, expected_owned_revision="sha256:stale")
        self.assertEqual(ctx.exception.status_code, 409)

        rebased = self.save(
            payload,
            expected_owned_revision=saved["owned_revision"],
            expected_derived_revision="sha256:stale",
        )
        self.assertTrue(rebased["derived_rebased"])

    def test_a_sidecar_owned_by_another_method_blocks_the_save(self) -> None:
        self._write_json(
            self.sidecars / f"{BOOTSTRAP_NAME}.json",
            {
                "dataset_name": BOOTSTRAP_NAME,
                "method_name": "Someone Else",
                "method_type": dataset_sidecar_status_service.METHOD_TYPE_CAPE_COD,
                "source_kind": "cape_cod",
                "data_format": "Vector",
                "precedents": [],
                "dependents": [],
            },
        )

        with self.assertRaises(HTTPException) as ctx:
            self.save()
        self.assertEqual(ctx.exception.status_code, 409)
        self.assertIn("already owned by", str(ctx.exception.detail))

    def test_a_missing_dfm_method_json_is_reported_against_the_dfm(self) -> None:
        (self.methods / f"DFM@{DFM_METHOD_NAME}.json").unlink()

        with self.assertRaises(HTTPException) as ctx:
            self.save()
        self.assertEqual(ctx.exception.status_code, 404)
        self.assertIn(DFM_METHOD_NAME, str(ctx.exception.detail))

    def test_a_non_dfm_precedent_is_rejected(self) -> None:
        sidecar_path = self.sidecars / f"{DFM_OUTPUT_DATASET}.json"
        sidecar = json.loads(sidecar_path.read_text(encoding="utf-8"))
        sidecar["method_type"] = dataset_sidecar_status_service.METHOD_TYPE_CAPE_COD
        sidecar["source_kind"] = "cape_cod"
        self._write_json(sidecar_path, sidecar)

        with self.assertRaises(HTTPException) as ctx:
            self.save()
        self.assertEqual(ctx.exception.status_code, 422)
        self.assertIn("must be a DFM method", str(ctx.exception.detail))

    def test_refresh_after_a_target_change_republishes_the_scaled_ultimate(self) -> None:
        saved = self.save()
        original = list(saved["method"]["results_tab"]["bootstrap_ultimate"])

        self._write_target([value * 1.5 for value in self._target_values()])
        with mock.patch.object(
            bootstrap_service,
            "_refresh_downstream_domains",
            return_value={"ok": True, "updated": []},
        ), mock.patch.object(
            bootstrap_service, "_lock", return_value=mock.MagicMock()
        ):
            refreshed = bootstrap_service.refresh_bootstrap_method(
                "Project", "Class", BOOTSTRAP_NAME
            )

        self.assertTrue(refreshed["updated"])
        self.assertTrue(refreshed["output_changed"])
        self.assertNotEqual(
            refreshed["method"]["results_tab"]["bootstrap_ultimate"], original
        )

    def test_a_finer_target_precedent_is_brought_to_the_method_period(self) -> None:
        """A target vector stored finer than the Bootstrap is rolled up or rebuilt."""

        self._write_target()
        sidecar_path = self.sidecars / f"{TARGET_NAME}.json"
        sidecar = json.loads(sidecar_path.read_text(encoding="utf-8"))
        rows = len(self.origin_labels)

        # Hand-entered and stored quarterly: four quarters sum to each year.
        sidecar.update({
            "source_kind": "input",
            "method_type": dataset_sidecar_status_service.METHOD_TYPE_NONE,
            "stored_period_length": 3,
            "csv_file": f"{TARGET_NAME}@3.csv",
        })
        self._write_json(sidecar_path, sidecar)
        quarters = [float(10 * index + quarter) for index in range(rows) for quarter in range(1, 5)]
        (self.datasets / f"{TARGET_NAME}@3.csv").write_text(
            "\n".join(str(value) for value in quarters) + "\n", encoding="utf-8", newline="\n"
        )
        snapshot = bootstrap_service._read_target_snapshot(
            "Project",
            "Class",
            TARGET_NAME,
            sidecar,
            origin_length=12,
            origin_labels=self.origin_labels,
        )
        self.assertEqual(snapshot["values"], [40 * index + 10 for index in range(rows)])
        self.assertEqual(snapshot["origin_labels"], self.origin_labels)

        # Engine-generated and stored monthly: rebuilt at the method's period.
        sidecar.update({"source_kind": "engine", "stored_period_length": 1})
        rebuilt = self.datasets / f"{TARGET_NAME}@12.rebuilt.csv"
        rebuilt.write_text(
            "\n".join(str(float(index)) for index in range(rows)) + "\n", encoding="utf-8", newline="\n"
        )
        with mock.patch.object(
            bootstrap_service.precedent_cache_service,
            "materialize_engine_source",
            return_value=str(rebuilt),
        ) as materialize:
            snapshot = bootstrap_service._read_target_snapshot(
                "Project",
                "Class",
                TARGET_NAME,
                sidecar,
                origin_length=12,
                origin_labels=self.origin_labels,
            )
        self.assertEqual(snapshot["values"], [float(index) for index in range(rows)])
        self.assertEqual(materialize.call_args.args[2:], (TARGET_NAME, sidecar, 12))

        # Coarser than the method: still refused.
        sidecar.update({"source_kind": "input", "stored_period_length": 36})
        with self.assertRaisesRegex(HTTPException, "uses 36-month origins; expected 12"):
            bootstrap_service._read_target_snapshot(
                "Project",
                "Class",
                TARGET_NAME,
                sidecar,
                origin_length=12,
                origin_labels=self.origin_labels,
            )

    def test_an_unchanged_dfm_snapshot_skips_the_simulation(self) -> None:
        self.save()
        sidecar = json.loads(
            (self.sidecars / f"{BOOTSTRAP_NAME}.json").read_text(encoding="utf-8")
        )

        result = bootstrap_service._refresh_one(
            "Project",
            "Class",
            BOOTSTRAP_NAME,
            sidecar,
            [DFM_OUTPUT_DATASET],
            blocked_precedent_keys=set(),
            sidecar_cache={},
            snapshot_cache={},
        )

        self.assertTrue(result["skipped"])
        self.assertEqual(result["reason"], "dfm_snapshot_unchanged")

    def test_an_unrelated_changed_dataset_is_a_stale_reverse_edge(self) -> None:
        self.save()
        sidecar = json.loads(
            (self.sidecars / f"{BOOTSTRAP_NAME}.json").read_text(encoding="utf-8")
        )

        result = bootstrap_service._refresh_one(
            "Project",
            "Class",
            BOOTSTRAP_NAME,
            sidecar,
            ["Something Else"],
            blocked_precedent_keys=set(),
            sidecar_cache={},
            snapshot_cache={},
        )

        self.assertEqual(result["reason"], "stale_reverse_dependency_edge")

    def test_bootstrap_downstream_cascade_excludes_only_its_own_wave(self) -> None:
        with mock.patch.object(
            calculated_dataset_service,
            "recalculate_dependents",
            return_value={"ok": True, "updated": []},
        ) as cascade:
            bootstrap_service._refresh_downstream_domains(
                "Project", "Class", BOOTSTRAP_NAME, "F 00 - Ultimate Net Loss"
            )

        cascade.assert_called_once_with(
            "Project",
            "Class",
            BOOTSTRAP_NAME,
            "F 00 - Ultimate Net Loss",
            include_bootstrap=False,
            finalize_method_review_status=True,
            rebuild_index=False,
        )

    def test_cascade_names_cover_every_earlier_domain(self) -> None:
        report = {
            "updated": [{"dataset_type_name": "Calculated Direct"}],
            "skipped": [{"dataset_type_name": "Calculated Failed"}],
            "dfm_updates": {"updated": [{"dataset_name": "DFM Output"}], "errors": []},
            "result_selection_updates": {
                "updated": [{"dataset_name": "RS Output"}],
                "errors": [{"dataset_name": "RS Failed"}],
                "downstream_fresh_names": ["RS Downstream"],
                "downstream_blocked_names": ["RS Blocked"],
            },
            "bornhuetter_ferguson_updates": {
                "status_refreshed": [{"dataset_name": "BF Restored"}],
                "errors": [{"dataset_name": "BF Failed"}],
            },
            "cape_cod_updates": {
                "updated": [{"dataset_name": "CC Output"}],
                "errors": [{"dataset_name": "CC Failed"}],
            },
        }

        fresh, failed = bootstrap_service._cascade_names(report)

        self.assertCountEqual(
            fresh,
            [
                "Calculated Direct",
                "DFM Output",
                "RS Output",
                "RS Downstream",
                "BF Restored",
                "CC Output",
            ],
        )
        self.assertCountEqual(
            failed, ["Calculated Failed", "RS Failed", "RS Blocked", "BF Failed", "CC Failed"]
        )

    def test_outer_cascade_feeds_every_earlier_wave_into_the_bootstrap_wave(self) -> None:
        from app_server.services import (
            bornhuetter_ferguson_service,
            cape_cod_service,
            dfm_service,
            result_selection_service,
        )

        waves = {
            dfm_service: {
                "ok": True,
                # output_changed is False on purpose: a DFM edit that leaves the
                # published ultimate untouched still changes the triangle and
                # ratios a Bootstrap embeds, so it must reach the wave anyway.
                "updated": [{"dataset_name": "DFM Output", "output_changed": False}],
                "status_refreshed": [],
                "errors": [],
            },
            result_selection_service: {
                "ok": True,
                "updated": [{"dataset_name": "RS Output"}],
                "status_refreshed": [],
                "errors": [],
            },
            bornhuetter_ferguson_service: {
                "ok": True,
                "updated": [{"dataset_name": "BF Output"}],
                "status_refreshed": [],
                "errors": [],
            },
            cape_cod_service: {
                "ok": True,
                "updated": [{"dataset_name": "CC Output"}],
                "status_refreshed": [],
                "errors": [],
            },
        }
        with mock.patch.object(
            calculated_dataset_service.dataset_instance_index_service,
            "rebuild_index",
            return_value=None,
        ), mock.patch.object(
            calculated_dataset_service.dataset_sidecar_status_service,
            "refresh_method_statuses_for_dependents",
            return_value=[],
        ), mock.patch.object(
            bootstrap_service, "refresh_dependents", return_value={"ok": True, "updated": []}
        ) as wave:
            with mock.patch.multiple(
                dfm_service, refresh_dependents=mock.Mock(return_value=waves[dfm_service])
            ), mock.patch.multiple(
                result_selection_service,
                refresh_dependents=mock.Mock(return_value=waves[result_selection_service]),
            ), mock.patch.multiple(
                bornhuetter_ferguson_service,
                refresh_dependents=mock.Mock(return_value=waves[bornhuetter_ferguson_service]),
            ), mock.patch.multiple(
                cape_cod_service, refresh_dependents=mock.Mock(return_value=waves[cape_cod_service])
            ):
                report = calculated_dataset_service.recalculate_dependents(
                    "Project", "Class", "Root Dataset", "Root Type"
                )

        self.assertIsNotNone(report["bootstrap_updates"])
        roots = wave.call_args.args[2]
        for name in ("Root Dataset", "Root Type", "DFM Output", "RS Output", "BF Output", "CC Output"):
            self.assertIn(name, roots)

    def test_a_dfm_publishing_under_its_own_name_still_resolves(self) -> None:
        (self.sidecars / f"{DFM_OUTPUT_DATASET}.json").unlink()
        self._write_dfm(output_dataset=DFM_METHOD_NAME)

        self.save()

        sidecar = json.loads(
            (self.sidecars / f"{BOOTSTRAP_NAME}.json").read_text(encoding="utf-8")
        )
        self.assertEqual(
            dataset_sidecar_status_service.entry_names(sidecar["precedents"]),
            [DFM_METHOD_NAME, TARGET_NAME],
        )

    def test_saving_twice_with_the_same_inputs_is_byte_stable(self) -> None:
        first = self.save()
        method_path = self.methods / f"BST@{BOOTSTRAP_NAME}.json"
        before = method_path.read_text(encoding="utf-8")

        second = self.save(
            self.method_payload(), expected_owned_revision=first["owned_revision"]
        )

        self.assertEqual(second["publication_revision"], first["publication_revision"])
        self.assertEqual(
            normalize_bootstrap_method(json.loads(method_path.read_text(encoding="utf-8")))[
                "results_tab"
            ]["bootstrap_ultimate"],
            normalize_bootstrap_method(json.loads(before))["results_tab"]["bootstrap_ultimate"],
        )


if __name__ == "__main__":  # pragma: no cover - convenience entry point
    unittest.main()

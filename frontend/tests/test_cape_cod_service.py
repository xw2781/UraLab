from __future__ import annotations

import copy
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

from arcrho_api.cape_cod_contract import (
    CC_JSON_FORMAT,
    build_cape_cod_output_sidecar,
    method_revisions,
    recalculate_cape_cod_method,
)
from app_server.services import (
    bornhuetter_ferguson_service,
    calculated_dataset_service,
    cape_cod_service,
    dataset_sidecar_status_service,
)
from dependent_propagation_workspace_stub import IsolatedPropagationWorkspace

RESQ_FIXTURE_PATH = REPO_ROOT / "python-api" / "tests" / "fixtures" / "resq_cape_cod_d53.json"
# Same bound as python-api/tests/test_cape_cod_contract.py: parity against the
# raw ResQ COM capture is limited by the ArcRho-wide six-decimal
# canonical-number policy applied to embedded snapshot values.
CANONICAL_TOL = 2e-6


class CapeCodServiceTests(unittest.TestCase):
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
                cape_cod_service.config,
                "get_general_settings_path",
                return_value=str(settings),
            ),
            mock.patch.object(
                cape_cod_service.config,
                "get_project_method_data_dir",
                return_value=str(self.methods),
            ),
            mock.patch.object(
                cape_cod_service.config,
                "get_project_dataset_cache_dir",
                return_value=str(self.datasets),
            ),
            mock.patch.object(
                dataset_sidecar_status_service,
                "sidecar_path",
                side_effect=lambda _p, _r, name: str(self.sidecars / f"{name}.json"),
            ),
        ]
        for patcher in self.patchers:
            patcher.start()

    def tearDown(self) -> None:
        for patcher in reversed(self.patchers):
            patcher.stop()
        self.temp.cleanup()

    @staticmethod
    def method_payload(
        *,
        latest_values: list[list[float | None]] | None = None,
        decay_factor: float = 1,
        trend_factor_overrides: list[float | None] | None = None,
    ) -> dict:
        method = recalculate_cape_cod_method(
            {
                "json_format": CC_JSON_FORMAT,
                "details_tab": {
                    "name": "CC Method",
                    "method_type": "Cape Cod",
                    "output_type": "CC Ultimate",
                    "dataset_category": "Loss",
                    "origin_length": 12,
                    "statistic_decimal_places": 1,
                },
                "method_tab": {
                    "latest_dataset": "Paid",
                    "exposure_dataset": "Exposure",
                    "prior_ultimate_dataset": "Prior Ultimate",
                    "prior_ultimate_mode": "latest_ultimates",
                    "trend_rate": 0,
                    "auto_trend_fit": False,
                    "decay_factor": decay_factor,
                    "scaling_type": "percentage",
                    "alternative_ultimate_calculation": False,
                    "trend_factor_overrides": [],
                },
            },
            source_snapshots={
                "latest": {
                    "name": "Paid",
                    "origin_labels": ["2024", "2025"],
                    "values": latest_values or [[100, 150], [200, None]],
                    "mask": [[True, True], [True, False]],
                },
                "exposure": {
                    "name": "Exposure",
                    "origin_labels": ["2024", "2025"],
                    "values": [[500], [600]],
                },
                "prior_ultimate": {
                    "name": "Prior Ultimate",
                    "origin_labels": ["2024", "2025"],
                    "values": [[300], [400]],
                },
            },
            timestamp="2026-01-01T00:00:00Z",
        )
        if trend_factor_overrides is not None:
            method["method_tab"]["trend_factor_overrides"] = trend_factor_overrides
            method = recalculate_cape_cod_method(
                method,
                timestamp="2026-01-01T00:00:00Z",
                update_refresh_timestamp=False,
            )
        return method

    @staticmethod
    def write_json(path: Path, payload: dict) -> None:
        path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")

    def output_sidecar(self, method: dict, *, status: int = 0) -> dict:
        method_name = method["details_tab"]["name"]
        origin_length = method["details_tab"]["origin_length"]
        sidecar = build_cape_cod_output_sidecar(
            method,
            project_name="Project",
            reserving_class="Class",
            csv_file=f"{method_name}@{origin_length}.csv",
            existing={},
            dependents=[],
            notes="method note",
            timestamp="2026-01-01T00:00:00Z",
            user="tester",
            status=status,
        )
        sidecar["status"] = status
        return sidecar

    def write_method_pair(self, method: dict | None = None, *, status: int = 0) -> dict:
        payload = method or self.method_payload()
        method_name = payload["details_tab"]["name"]
        origin_length = payload["details_tab"]["origin_length"]
        self.write_json(self.methods / f"CC@{method_name}.json", payload)
        self.write_json(
            self.sidecars / f"{method_name}.json",
            self.output_sidecar(payload, status=status),
        )
        (self.datasets / f"{method_name}@{origin_length}.csv").write_text(
            "\n".join(str(value) for value in payload["method_tab"]["cape_cod_ultimate"]) + "\n",
            encoding="utf-8",
        )
        return payload

    def write_source(
        self,
        name: str,
        csv_text: str,
        *,
        data_format: str,
        method_type: str = "None",
        dependents: list[str] | None = None,
        status: int = 0,
        include_origin_labels: bool = True,
    ) -> None:
        csv_file = f"{name}@12.csv"
        (self.datasets / csv_file).write_text(csv_text, encoding="utf-8")
        sidecar = {
            "dataset_name": name,
            "dataset_type": name,
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "dfm" if method_type == "DFM" else "input",
            "method_type": method_type,
            "data_format": data_format,
            "period_length": 12,
            **(
                {"stored_period_length": 12}
                if data_format == "Vector"
                else {"stored_origin_length": 12, "stored_development_length": 12}
            ),
            "csv_file": csv_file,
            "status": status,
            "precedents": [],
            "dependents": [
                {"dataset_name": item} for item in (dependents or [])
            ],
        }
        if include_origin_labels:
            sidecar["origin_labels"] = ["2024", "2025"]
        self.write_json(self.sidecars / f"{name}.json", sidecar)

    def write_all_sources(self, *, paid_csv: str = "100,150\n200,\n") -> None:
        self.write_source("Paid", paid_csv, data_format="Triangle", dependents=["CC Method"])
        self.write_source("Exposure", "500\n600\n", data_format="Vector", dependents=["CC Method"])
        self.write_source(
            "Prior Ultimate",
            "300\n400\n",
            data_format="Vector",
            dependents=["CC Method"],
        )

    def test_a_finer_exposure_precedent_is_brought_to_the_method_period(self) -> None:
        """A monthly exposure vector feeds a yearly Cape Cod instead of being refused."""

        self.write_all_sources()
        path = self.sidecars / "Exposure.json"
        sidecar = json.loads(path.read_text(encoding="utf-8"))

        # Hand-entered and stored quarterly: summed up to the method's year.
        sidecar["stored_period_length"] = 3
        self.write_json(path, sidecar)
        (self.datasets / "Exposure@12.csv").write_text(
            "100\n200\n300\n400\n50\n60\n70\n80\n", encoding="utf-8"
        )
        snapshot = cape_cod_service._source_snapshots(
            "Project", "Class", self.method_payload(), {"exposure"},
        )
        self.assertEqual(snapshot["exposure"]["values"], [[1000], [260]])

        # Engine-generated and stored monthly: rebuilt at the method's period.
        sidecar["source_kind"] = "engine"
        sidecar["stored_period_length"] = 1
        self.write_json(path, sidecar)
        rebuilt = self.datasets / "Exposure@12.rebuilt.csv"
        rebuilt.write_text("700\n800\n", encoding="utf-8")
        with mock.patch.object(
            cape_cod_service.precedent_cache_service,
            "materialize_engine_source",
            return_value=str(rebuilt),
        ) as materialize:
            snapshot = cape_cod_service._source_snapshots(
                "Project", "Class", self.method_payload(), {"exposure"},
            )
        self.assertEqual(snapshot["exposure"]["values"], [[700], [800]])
        self.assertEqual(materialize.call_args.args[2:], ("Exposure", sidecar, 12))

        # Coarser than the method: still refused, yearly figures cannot be split.
        sidecar["source_kind"] = "input"
        sidecar["stored_period_length"] = 36
        self.write_json(path, sidecar)
        with self.assertRaisesRegex(HTTPException, "uses 36-month origins; expected 12"):
            cape_cod_service._source_snapshots(
                "Project", "Class", self.method_payload(), {"exposure"},
            )

    def test_load_reads_method_sidecar_and_latest_source_only(self) -> None:
        self.write_method_pair()
        self.write_all_sources()
        original = cape_cod_service._read_json
        reads: list[str] = []

        def recording(path: str) -> dict:
            reads.append(str(Path(path)))
            return original(path)

        with (
            mock.patch.object(
                cape_cod_service,
                "_read_json",
                side_effect=recording,
            ),
            mock.patch.object(
                cape_cod_service,
                "cape_cod_output_variants",
                side_effect=AssertionError("output recalculation"),
            ),
        ):
            result = cape_cod_service.load_cape_cod_method("Project", "Class", "CC Method")

        self.assertTrue(result["ok"])
        # The as-if Ultimates triangle needs the Latest triangle, so the Latest
        # precedent is read; Exposure and Prior Ultimate stay untouched.
        self.assertCountEqual(reads, [
            str(self.methods / "CC@CC Method.json"),
            str(self.sidecars / "CC Method.json"),
            str(self.sidecars / "Paid.json"),
        ])
        triangle = result["ultimates_triangle"]
        self.assertEqual([len(row) for row in triangle], [2, 1])
        for row in triangle:
            for value in row:
                self.assertIsNotNone(value)

    def test_load_returns_null_triangle_when_latest_source_unavailable(self) -> None:
        self.write_method_pair()

        result = cape_cod_service.load_cape_cod_method("Project", "Class", "CC Method")

        self.assertTrue(result["ok"])
        self.assertIsNone(result["ultimates_triangle"])

    def test_load_returns_null_triangle_when_latest_triangle_is_irregular(self) -> None:
        self.write_method_pair()
        # Two origins but a single stored development column: the oldest origin
        # cannot supply its n - i = 2 leading cells, so the canonical contract
        # rejects the irregular triangle and the load degrades the field.
        self.write_source(
            "Paid",
            "100\n200\n",
            data_format="Triangle",
            dependents=["CC Method"],
        )

        result = cape_cod_service.load_cape_cod_method("Project", "Class", "CC Method")

        self.assertTrue(result["ok"])
        self.assertIsNone(result["ultimates_triangle"])

    def test_load_triangle_blanks_cells_without_observed_values(self) -> None:
        self.write_method_pair()
        # Row 1 keeps its regular single leading cell, but that cell is masked
        # off in the source, so only that as-if estimate stays blank.
        self.write_source(
            "Paid",
            "100,150\n,175\n",
            data_format="Triangle",
            dependents=["CC Method"],
        )

        result = cape_cod_service.load_cape_cod_method("Project", "Class", "CC Method")

        self.assertTrue(result["ok"])
        triangle = result["ultimates_triangle"]
        self.assertEqual([len(row) for row in triangle], [2, 1])
        self.assertIsNotNone(triangle[0][0])
        self.assertIsNone(triangle[1][0])

    def test_load_missing_method_is_a_404(self) -> None:
        with self.assertRaises(HTTPException) as raised:
            cape_cod_service.load_cape_cod_method("Project", "Class", "CC Method")

        self.assertEqual(raised.exception.status_code, 404)

    def test_unsupported_format_load_is_rejected_without_source_reads_or_writes(self) -> None:
        method = self.write_method_pair()
        method["json_format"] = "arcrho-cape-cod-method-by-tab-v0"
        method_path = self.methods / "CC@CC Method.json"
        self.write_json(method_path, method)
        method_before = method_path.read_bytes()
        sidecar_path = self.sidecars / "CC Method.json"
        sidecar_before = sidecar_path.read_bytes()

        with (
            mock.patch.object(
                cape_cod_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=AssertionError("source read"),
            ),
            mock.patch.object(
                cape_cod_service,
                "_publish",
                side_effect=AssertionError("publication write"),
            ),
            self.assertRaises(HTTPException) as raised,
        ):
            cape_cod_service.load_cape_cod_method("Project", "Class", "CC Method")

        self.assertEqual(raised.exception.status_code, 422)
        self.assertIn("Unsupported Cape Cod JSON format", str(raised.exception.detail))
        self.assertEqual(method_path.read_bytes(), method_before)
        self.assertEqual(sidecar_path.read_bytes(), sidecar_before)

    def test_load_rejects_method_sidecar_geometry_mismatch(self) -> None:
        self.write_method_pair()
        sidecar_path = self.sidecars / "CC Method.json"
        sidecar = json.loads(sidecar_path.read_text(encoding="utf-8"))
        sidecar["origin_labels"] = ["2023", "2025"]
        self.write_json(sidecar_path, sidecar)

        with self.assertRaises(HTTPException) as raised:
            cape_cod_service.load_cape_cod_method("Project", "Class", "CC Method")

        self.assertEqual(raised.exception.status_code, 409)
        self.assertIn("origin labels do not match", str(raised.exception.detail))

    def test_save_rebases_owned_settings_over_newer_disk_derived_snapshot(self) -> None:
        stale = self.method_payload()
        current = self.method_payload(latest_values=[[100, 175], [200, None]])
        self.write_method_pair(current)
        incoming = copy.deepcopy(stale)
        incoming["method_tab"]["decay_factor"] = 0.5
        source_roles: list[str] = []
        original = cape_cod_service._read_source_snapshot_from_sidecar

        def recording(*args, **kwargs):
            source_roles.append(str(kwargs.get("role")))
            return original(*args, **kwargs)

        with (
            mock.patch.object(
                cape_cod_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=recording,
            ),
            mock.patch.object(
                calculated_dataset_service,
                "recalculate_dependents",
                return_value={"ok": True, "updated": [], "index_ok": True},
            ),
        ):
            result = cape_cod_service.save_cape_cod_method(
                "Project",
                "Class",
                incoming,
                expected_owned_revision=method_revisions(stale)["owned_revision"],
                expected_derived_revision=method_revisions(stale)["derived_revision"],
            )

        self.assertTrue(result["derived_rebased"])
        # Only the ultimates-triangle diagnostic may touch the Latest source;
        # the save itself must reuse the embedded derived snapshots.
        self.assertLessEqual(set(source_roles), {"latest"})
        saved = json.loads((self.methods / "CC@CC Method.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["latest_values"], current["method_tab"]["latest_values"])
        self.assertEqual(saved["method_tab"]["decay_factor"], 0.5)

    def test_save_conflict_when_owned_settings_changed_on_disk(self) -> None:
        stale = self.method_payload()
        self.write_method_pair(self.method_payload(decay_factor=0.5))
        incoming = copy.deepcopy(stale)

        with self.assertRaises(HTTPException) as raised:
            cape_cod_service.save_cape_cod_method(
                "Project",
                "Class",
                incoming,
                expected_owned_revision=method_revisions(stale)["owned_revision"],
            )

        self.assertEqual(raised.exception.status_code, 409)
        self.assertIn("owned settings changed on disk", str(raised.exception.detail))

    def test_save_rejects_output_dataset_owned_by_another_method(self) -> None:
        method = self.write_method_pair()
        sidecar_path = self.sidecars / "CC Method.json"
        sidecar = json.loads(sidecar_path.read_text(encoding="utf-8"))
        sidecar["method_name"] = "Another Method"
        self.write_json(sidecar_path, sidecar)
        (self.methods / "CC@CC Method.json").unlink()

        with self.assertRaises(HTTPException) as raised:
            cape_cod_service.save_cape_cod_method("Project", "Class", method)

        self.assertEqual(raised.exception.status_code, 409)
        self.assertIn("already owned by 'Another Method'", str(raised.exception.detail))

    def test_no_op_save_submits_no_engine_propagation_job(self) -> None:
        method = self.write_method_pair()

        with (
            mock.patch.object(
                cape_cod_service.dependent_propagation_service,
                "require_reserving_class_writable",
            ),
            mock.patch.object(
                cape_cod_service.dependent_propagation_service,
                "enqueue_marked_save_propagation",
            ) as enqueue,
        ):
            result = cape_cod_service.save_cape_cod_method(
                "Project",
                "Class",
                method,
                expected_owned_revision=method_revisions(method)["owned_revision"],
            )

        self.assertEqual(result["sidecar"]["status"], 0)
        self.assertTrue(result["propagation_ok"])
        self.assertEqual(result["propagation"], {"ok": True, "status": "unchanged"})
        enqueue.assert_not_called()

    def test_review_needed_save_uses_embedded_snapshots_and_reports_precedents(self) -> None:
        method = self.write_method_pair(status=2)
        self.write_all_sources()
        prior_path = self.sidecars / "Prior Ultimate.json"
        prior_source = json.loads(prior_path.read_text(encoding="utf-8"))
        prior_source["method_type"] = "DFM"
        prior_source["source_kind"] = "dfm"
        prior_source["status"] = 2
        self.write_json(prior_path, prior_source)

        with mock.patch.object(
            calculated_dataset_service,
            "recalculate_dependents",
            return_value={"ok": True, "updated": [], "index_ok": True},
        ):
            result = cape_cod_service.save_cape_cod_method(
                "Project",
                "Class",
                method,
                expected_owned_revision=method_revisions(method)["owned_revision"],
            )

        self.assertEqual(result["sidecar"]["status"], 0)
        self.assertEqual(result["unreviewed_precedents"], ["Prior Ultimate"])
        self.assertEqual(result["unreviewed_precedent_count"], 1)

    def test_exposure_refresh_uses_method_origins_and_persists_new_values(self) -> None:
        self.write_method_pair(status=2)
        self.write_source(
            "Paid",
            "100,150\n200,\n",
            data_format="Triangle",
            dependents=["CC Method"],
            include_origin_labels=False,
        )
        self.write_source(
            "Exposure",
            "700\n800\n",
            data_format="Vector",
            dependents=["CC Method"],
            include_origin_labels=False,
        )
        self.write_source(
            "Prior Ultimate",
            "300\n400\n",
            data_format="Vector",
            dependents=["CC Method"],
            include_origin_labels=False,
        )
        source_reads: list[str] = []
        original = cape_cod_service._read_source_snapshot_from_sidecar

        def recording(*args, **kwargs):
            source_reads.append(str(args[2]))
            return original(*args, **kwargs)

        with (
            mock.patch.object(
                cape_cod_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=recording,
            ),
            mock.patch.object(
                cape_cod_service,
                "_refresh_downstream_domains",
                return_value={"ok": True, "updated": []},
            ),
        ):
            result = cape_cod_service.refresh_dependents(
                "Project", "Class", ["Exposure"], rebuild_index=False
            )

        self.assertTrue(result["ok"], result)
        self.assertEqual(source_reads, ["Exposure"])
        self.assertEqual(result["errors"], [])
        saved = json.loads((self.methods / "CC@CC Method.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["exposure_values"], [700, 800])
        self.assertEqual(saved["method_tab"]["origin_labels"], ["2024", "2025"])
        self.assertEqual(
            (self.datasets / "CC Method@12.csv").read_text(encoding="utf-8").splitlines(),
            [str(value) for value in saved["method_tab"]["cape_cod_ultimate"]],
        )

    def test_explicit_refresh_keeps_review_alert_until_save(self) -> None:
        self.write_method_pair(status=2)
        self.write_all_sources()

        with mock.patch.object(
            cape_cod_service,
            "_refresh_downstream_domains",
            return_value={"ok": True, "updated": []},
        ) as cascade:
            result = cape_cod_service.refresh_cape_cod_method(
                "Project",
                "Class",
                "CC Method",
            )

        self.assertFalse(result["output_changed"])
        self.assertFalse(result["status_refreshed"])
        cascade.assert_not_called()
        sidecar = json.loads((self.sidecars / "CC Method.json").read_text(encoding="utf-8"))
        self.assertEqual(
            sidecar["status"],
            dataset_sidecar_status_service.STATUS_REVIEW_NEEDED,
        )

    def test_failed_refresh_preserves_method_and_output_and_marks_review(self) -> None:
        self.write_method_pair()
        self.write_source(
            "Paid",
            "100,175\n",
            data_format="Triangle",
            dependents=["CC Method"],
        )
        method_path = self.methods / "CC@CC Method.json"
        output_path = self.datasets / "CC Method@12.csv"
        before_method = method_path.read_bytes()
        before_output = output_path.read_bytes()

        result = cape_cod_service.refresh_dependents(
            "Project", "Class", ["Paid"], rebuild_index=False
        )

        self.assertFalse(result["ok"])
        self.assertEqual(method_path.read_bytes(), before_method)
        self.assertEqual(output_path.read_bytes(), before_output)
        sidecar = json.loads((self.sidecars / "CC Method.json").read_text(encoding="utf-8"))
        self.assertEqual(sidecar["status"], dataset_sidecar_status_service.STATUS_REVIEW_NEEDED)

    def test_latest_refresh_rewrites_embedded_values_but_preserves_overrides(self) -> None:
        method = self.write_method_pair(
            self.method_payload(trend_factor_overrides=[1.5, None])
        )
        original_overrides = copy.deepcopy(method["method_tab"]["trend_factor_overrides"])
        self.assertEqual(original_overrides, [1.5, None])
        self.write_all_sources(paid_csv="100,175\n200,\n")
        old_revision = self.output_sidecar(method)["publication_revision"]
        with mock.patch.object(
            cape_cod_service,
            "_refresh_downstream_domains",
            return_value={"ok": True, "updated": []},
        ):
            result = cape_cod_service.refresh_dependents(
                "Project", "Class", ["Paid"], rebuild_index=False
            )

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["updated"][0]["dataset_name"], "CC Method")
        self.assertTrue(result["updated"][0]["output_changed"])
        saved = json.loads((self.methods / "CC@CC Method.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["latest_values"], [175, 200])
        self.assertEqual(
            saved["method_tab"]["trend_factor_overrides"],
            original_overrides,
        )
        self.assertNotEqual(
            saved["method_tab"]["cape_cod_ultimate"],
            method["method_tab"]["cape_cod_ultimate"],
        )
        sidecar = json.loads((self.sidecars / "CC Method.json").read_text(encoding="utf-8"))
        self.assertEqual(
            sidecar["status"],
            dataset_sidecar_status_service.STATUS_REVIEW_NEEDED,
        )
        self.assertEqual(
            result["review_status_updates"],
            [{"dataset_name": "CC Method", "status": 2}],
        )
        self.assertEqual(
            (self.datasets / "CC Method@12.csv").read_text(encoding="utf-8").splitlines(),
            [str(value) for value in saved["method_tab"]["cape_cod_ultimate"]],
        )
        self.assertNotEqual(sidecar["publication_revision"], old_revision)
        self.assertEqual(sidecar["publication_revision"], method_revisions(saved)["publication_revision"])

    def test_source_executor_is_bounded_for_network_reads(self) -> None:
        self.assertEqual(cape_cod_service.READ_MAX_WORKERS, 4)
        self.assertEqual(
            cape_cod_service._READ_EXECUTOR._max_workers,
            cape_cod_service.READ_MAX_WORKERS,
        )

    def test_refresh_report_order_is_deterministic(self) -> None:
        self.write_json(self.sidecars / "Paid.json", {
            "dataset_name": "Paid",
            "dependents": [
                {"dataset_name": "CC Z"},
                {"dataset_name": "CC A"},
            ],
        })
        for name in ("CC Z", "CC A"):
            self.write_json(self.sidecars / f"{name}.json", {
                "dataset_name": name,
                "method_name": name,
                "method_type": "Cape Cod",
                "source_kind": "cape_cod",
                "dependents": [],
                "status": 0,
            })

        def refreshed(_project, _reserving, output, sidecar, *_args, **_kwargs):
            return {
                "ok": True,
                "dataset_name": output,
                "dataset_type": output,
                "updated": True,
                "output_changed": False,
                "status_refreshed": False,
                "sidecar": sidecar,
            }

        with mock.patch.object(
            cape_cod_service,
            "_refresh_one",
            side_effect=refreshed,
        ):
            result = cape_cod_service.refresh_dependents(
                "Project", "Class", ["Paid"], rebuild_index=False
            )

        self.assertEqual(
            [item["dataset_name"] for item in result["updated"]],
            ["CC A", "CC Z"],
        )

    def test_staged_publish_rolls_back_and_replaces_sidecar_last(self) -> None:
        method_path = self.methods / "method.json"
        csv_path = self.datasets / "output.csv"
        sidecar_path = self.sidecars / "output.json"
        for path, text in (
            (method_path, "old-method\n"),
            (csv_path, "old-csv\n"),
            (sidecar_path, "old-sidecar\n"),
        ):
            path.write_text(text, encoding="utf-8")
        original = {path: path.read_bytes() for path in (method_path, csv_path, sidecar_path)}
        real_replace = cape_cod_service.os.replace
        targets: list[str] = []

        def replace(source: str, target: str) -> None:
            targets.append(target)
            if target == str(sidecar_path) and source.endswith(".tmp"):
                raise OSError("sidecar publish failed")
            real_replace(source, target)

        with mock.patch.object(
            cape_cod_service.os,
            "replace",
            side_effect=replace,
        ):
            with self.assertRaises(OSError):
                cape_cod_service._commit_text_files(
                    {
                        str(method_path): "new-method\n",
                        str(csv_path): "new-csv\n",
                        str(sidecar_path): "new-sidecar\n",
                    },
                    last_paths=[str(sidecar_path)],
                )

        self.assertEqual({path: path.read_bytes() for path in original}, original)
        sidecar_index = targets.index(str(sidecar_path))
        self.assertGreater(sidecar_index, targets.index(str(method_path)))
        self.assertGreater(sidecar_index, targets.index(str(csv_path)))

    def test_cc_nested_cascade_names_include_bf_rs_and_calculated_outputs(self) -> None:
        fresh, failed = cape_cod_service._cascade_names({
            "ok": False,
            "updated": [{"dataset_type_name": "Calculated Direct"}],
            "result_selection_updates": {
                "updated": [{"dataset_name": "RS Output"}],
                "downstream_fresh_names": ["Calculated After RS"],
                "downstream_blocked_names": ["Failed After RS"],
            },
            "bornhuetter_ferguson_updates": {
                "updated": [{"dataset_name": "BF Output"}],
                "status_refreshed": [{"dataset_name": "BF Restored"}],
                "errors": [{"dataset_name": "BF Failed"}],
            },
        })

        self.assertCountEqual(
            fresh,
            ["Calculated Direct", "RS Output", "Calculated After RS", "BF Output", "BF Restored"],
        )
        self.assertCountEqual(failed, ["Failed After RS", "BF Failed"])

    def test_cape_cod_downstream_cascade_excludes_its_own_and_later_waves(self) -> None:
        # Each wave's nested cascade suppresses itself and every wave that runs
        # after it; the outer cascade feeds those later waves from this wave's
        # fresh names instead, so nothing is refreshed twice.
        with mock.patch.object(
            calculated_dataset_service,
            "recalculate_dependents",
            return_value={"ok": True, "updated": []},
        ) as cascade:
            cape_cod_service._refresh_downstream_domains(
                "Project", "Class", "CC Method", "CC Ultimate"
            )

        cascade.assert_called_once_with(
            "Project",
            "Class",
            "CC Method",
            "CC Ultimate",
            include_cape_cod=False,
            include_bootstrap=False,
            finalize_method_review_status=True,
            rebuild_index=False,
        )

    def test_bf_downstream_cascade_excludes_cape_cod_and_bootstrap_waves(self) -> None:
        with mock.patch.object(
            calculated_dataset_service,
            "recalculate_dependents",
            return_value={"ok": True, "updated": []},
        ) as cascade:
            bornhuetter_ferguson_service._refresh_downstream_domains(
                "Project", "Class", "BF Method", "BF Ultimate"
            )

        cascade.assert_called_once_with(
            "Project",
            "Class",
            "BF Method",
            "BF Ultimate",
            include_bornhuetter_ferguson=False,
            include_cape_cod=False,
            include_bootstrap=False,
            finalize_method_review_status=True,
            rebuild_index=False,
        )

    def test_outer_cascade_passes_bf_outputs_to_cape_cod_wave(self) -> None:
        with (
            mock.patch(
                "app_server.services.dfm_service.refresh_dependents",
                return_value={"ok": True, "updated": [], "errors": []},
            ),
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=[]),
            mock.patch.object(calculated_dataset_service, "_existing_downstream_keys", return_value=[]),
            mock.patch(
                "app_server.services.result_selection_service.refresh_dependents",
                return_value={"ok": True, "updated": [], "status_refreshed": [], "errors": []},
            ),
            mock.patch.object(
                bornhuetter_ferguson_service,
                "refresh_dependents",
                return_value={
                    "ok": True,
                    "updated": [{"dataset_name": "BF Output"}],
                    "status_refreshed": [{"dataset_name": "BF Restored"}],
                    "errors": [{"dataset_name": "BF Failed"}],
                },
            ),
            mock.patch.object(
                cape_cod_service,
                "refresh_dependents",
                return_value={"ok": True, "updated": [], "errors": []},
            ) as refresh_cc,
            mock.patch.object(
                calculated_dataset_service.dataset_instance_index_service,
                "rebuild_index",
            ),
        ):
            result = calculated_dataset_service.recalculate_dependents(
                "Project", "Class", "Paid", "Paid"
            )

        self.assertIn("cape_cod_updates", result)
        roots = refresh_cc.call_args.args[2]
        blocked = refresh_cc.call_args.kwargs["blocked_precedent_names"]
        self.assertIn("BF Output", roots)
        self.assertIn("BF Restored", roots)
        self.assertIn("BF Failed", roots)
        self.assertIn("BF Failed", blocked)

    def test_aggregated_output_variants_are_published_alongside_native_csv(self) -> None:
        method = recalculate_cape_cod_method(
            {
                "json_format": CC_JSON_FORMAT,
                "details_tab": {
                    "name": "CC Half",
                    "method_type": "Cape Cod",
                    "output_type": "CC Ultimate",
                    "dataset_category": "Loss",
                    "origin_length": 6,
                    "statistic_decimal_places": 1,
                },
                "method_tab": {
                    "latest_dataset": "Paid",
                    "exposure_dataset": "Exposure",
                    "prior_ultimate_dataset": "Prior Ultimate",
                    "prior_ultimate_mode": "latest_ultimates",
                    "trend_rate": 0,
                    "auto_trend_fit": False,
                    "decay_factor": 1,
                    "scaling_type": "percentage",
                    "alternative_ultimate_calculation": False,
                    "trend_factor_overrides": [],
                },
            },
            source_snapshots={
                "latest": {
                    "name": "Paid",
                    "origin_labels": ["2024H1", "2024H2"],
                    "values": [[100, 150], [200, None]],
                    "mask": [[True, True], [True, False]],
                },
                "exposure": {
                    "name": "Exposure",
                    "origin_labels": ["2024H1", "2024H2"],
                    "values": [[500], [600]],
                },
                "prior_ultimate": {
                    "name": "Prior Ultimate",
                    "origin_labels": ["2024H1", "2024H2"],
                    "values": [[300], [400]],
                },
            },
            timestamp="2026-01-01T00:00:00Z",
        )
        method_name = method["details_tab"]["name"]
        self.write_json(self.methods / f"CC@{method_name}.json", method)
        sidecar = build_cape_cod_output_sidecar(
            method,
            project_name="Project",
            reserving_class="Class",
            csv_file=f"{method_name}@6.csv",
            existing={},
            dependents=[],
            timestamp="2026-01-01T00:00:00Z",
            user="tester",
        )
        self.write_json(self.sidecars / f"{method_name}.json", sidecar)

        with mock.patch.object(
            calculated_dataset_service,
            "recalculate_dependents",
            return_value={"ok": True, "updated": [], "index_ok": True},
        ):
            result = cape_cod_service.save_cape_cod_method(
                "Project",
                "Class",
                method,
                expected_owned_revision=method_revisions(method)["owned_revision"],
            )

        ultimates = method["method_tab"]["cape_cod_ultimate"]
        native = (self.datasets / "CC Half@6.csv").read_text(encoding="utf-8")
        self.assertEqual(native.splitlines(), [str(value) for value in ultimates])
        aggregated = (self.datasets / "CC Half@12.csv").read_text(encoding="utf-8")
        self.assertEqual(len(aggregated.splitlines()), 1)
        self.assertAlmostEqual(
            float(aggregated.strip()),
            float(ultimates[0]) + float(ultimates[1]),
            places=6,
        )
        self.assertEqual(
            [Path(path).name for path in result["aggregated_csv_paths"]],
            ["CC Half@12.csv"],
        )

    def test_load_reproduces_resq_verified_ultimates_triangle(self) -> None:
        fixture = json.loads(RESQ_FIXTURE_PATH.read_text(encoding="utf-8"))
        labels = list(fixture["origin_labels"])
        row_count = len(labels)
        padded_rows = [
            list(row) + [None] * (row_count - len(row))
            for row in fixture["latest_triangle"]
        ]
        method = recalculate_cape_cod_method(
            {
                "json_format": CC_JSON_FORMAT,
                "details_tab": {
                    "name": "CC D53",
                    "method_type": "Cape Cod",
                    "output_type": "CC Ultimate",
                    "dataset_category": "Loss",
                    "origin_length": fixture["method"]["origin_length"],
                    "statistic_decimal_places": fixture["method"]["decimal_places"],
                },
                "method_tab": {
                    "latest_dataset": "Gross Incurred",
                    "exposure_dataset": "Earned Exposure",
                    "prior_ultimate_dataset": "Prior DFM",
                    "prior_ultimate_mode": "latest_ultimates",
                    "trend_rate": 0,
                    "auto_trend_fit": fixture["method"]["auto_trend_fit"],
                    "decay_factor": fixture["method"]["decay_factor"],
                    "scaling_type": "percentage",
                    "alternative_ultimate_calculation": fixture["method"][
                        "alternative_ultimate_calculation"
                    ],
                    "trend_factor_overrides": [],
                },
            },
            source_snapshots={
                "latest": {
                    "name": "Gross Incurred",
                    "origin_labels": labels,
                    "values": padded_rows,
                    "mask": [[value is not None for value in row] for row in padded_rows],
                },
                "exposure": {
                    "name": "Earned Exposure",
                    "origin_labels": labels,
                    "values": [[value] for value in fixture["exposure_values"]],
                },
                "prior_ultimate": {
                    "name": "Prior DFM",
                    "origin_labels": labels,
                    "values": [[value] for value in fixture["prior_ultimate_values"]],
                },
            },
            timestamp="2026-01-01T00:00:00Z",
        )
        self.write_json(self.methods / "CC@CC D53.json", method)
        sidecar = build_cape_cod_output_sidecar(
            method,
            project_name="Project",
            reserving_class="Class",
            csv_file="CC D53@12.csv",
            existing={},
            dependents=[],
            timestamp="2026-01-01T00:00:00Z",
            user="tester",
        )
        self.write_json(self.sidecars / "CC D53.json", sidecar)
        latest_sidecar = {
            "dataset_name": "Gross Incurred",
            "dataset_type": "Gross Incurred",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "input",
            "method_type": "None",
            "data_format": "Triangle",
            "period_length": 12,
            "stored_origin_length": 12,
            "stored_development_length": 12,
            "csv_file": "Gross Incurred@12.csv",
            "status": 0,
            "precedents": [],
            "dependents": [{"dataset_name": "CC D53"}],
        }
        self.write_json(self.sidecars / "Gross Incurred.json", latest_sidecar)
        (self.datasets / "Gross Incurred@12.csv").write_text(
            "\n".join(
                ",".join("" if value is None else repr(value) for value in row)
                for row in padded_rows
            )
            + "\n",
            encoding="utf-8",
        )

        expected_ultimates = fixture["expected"]["cape_cod_ultimate"]
        for index, value in enumerate(method["method_tab"]["cape_cod_ultimate"]):
            self.assertLessEqual(
                abs(float(value) - float(expected_ultimates[index])),
                CANONICAL_TOL * max(1.0, abs(float(expected_ultimates[index]))),
                f"cape_cod_ultimate[{index}]",
            )

        result = cape_cod_service.load_cape_cod_method("Project", "Class", "CC D53")

        triangle = result["ultimates_triangle"]
        expected = fixture["expected_ultimates_triangle"]
        self.assertEqual(
            [len(row) for row in triangle],
            [len(row) for row in expected],
        )
        for row_index, (row, expected_row) in enumerate(zip(triangle, expected)):
            for column, (value, expected_value) in enumerate(zip(row, expected_row)):
                if expected_value is None:
                    self.assertIsNone(value, f"[{row_index}][{column}]")
                    continue
                self.assertIsNotNone(value, f"[{row_index}][{column}]")
                self.assertLessEqual(
                    abs(float(value) - float(expected_value)),
                    CANONICAL_TOL * max(1.0, abs(float(expected_value))),
                    f"[{row_index}][{column}]: {value!r} != {expected_value!r}",
                )


if __name__ == "__main__":
    unittest.main()

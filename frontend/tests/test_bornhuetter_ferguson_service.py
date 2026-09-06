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

from arcrho_api.dfm_contract import (
    DFM_JSON_FORMAT,
    default_average_formulas,
    recalculate_dfm_method,
)
from arcrho_api.bornhuetter_ferguson_contract import (
    BF_JSON_FORMAT,
    build_bornhuetter_ferguson_output_sidecar,
    method_revisions,
    recalculate_bornhuetter_ferguson_method,
)
from app_server.services import (
    bornhuetter_ferguson_service,
    calculated_dataset_service,
    dataset_sidecar_status_service,
    result_selection_service,
)


class BornhuetterFergusonServiceTests(unittest.TestCase):
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
            mock.patch.object(
                bornhuetter_ferguson_service.config,
                "get_general_settings_path",
                return_value=str(settings),
            ),
            mock.patch.object(
                bornhuetter_ferguson_service.config,
                "get_project_method_data_dir",
                return_value=str(self.methods),
            ),
            mock.patch.object(
                bornhuetter_ferguson_service.config,
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
    def method_payload(*, latest_values: list[list[float | None]] | None = None) -> dict:
        return recalculate_bornhuetter_ferguson_method(
            {
                "json_format": BF_JSON_FORMAT,
                "details_tab": {
                    "name": "BF Method",
                    "method_type": "Bornhuetter Ferguson",
                    "output_type": "BF Ultimate",
                    "dataset_category": "Loss",
                    "origin_length": 12,
                    "statistic_decimal_places": 1,
                },
                "method_tab": {
                    "latest_dataset": "Paid",
                    "dfm_dataset": "Development Output",
                    "show_weights": True,
                    "show_effective_weights": False,
                    "prior_datasets": [{"name": "Prior", "weights": [1, 1]}],
                },
            },
            source_snapshots={
                "latest": {
                    "name": "Paid",
                    "origin_labels": ["2024", "2025"],
                    "values": latest_values or [[100, 150], [200, None]],
                    "mask": [[True, True], [True, False]],
                },
                "dfm": {
                    "name": "Development Output",
                    "origin_labels": ["2024", "2025"],
                    "values": [[300], [400]],
                    "percentage_developed": [0.5, 0.5],
                },
                "priors": {
                    "Prior": {
                        "name": "Prior",
                        "origin_labels": ["2024", "2025"],
                        "values": [[500], [600]],
                    },
                },
            },
            timestamp="2026-01-01T00:00:00Z",
        )

    @staticmethod
    def dfm_method_payload() -> dict:
        """A DFM whose selected factors develop both origins to 50%.

        Its factors chain to a cumulative 2.0 at each origin's own development
        age, so it publishes the same 300/400 ultimates the other fixtures use
        and a percentage-developed pattern of 0.5 for both origins.
        """

        formulas = default_average_formulas()
        formulas["selected"] = [[0, 0, 0], [0, 0, 0], [1, 1, 1]]
        formulas["inputs"] = [["", "", ""], ["", "", ""], ["1", "2", "1"]]
        return recalculate_dfm_method(
            {
                "json_format": DFM_JSON_FORMAT,
                "details_tab": {
                    "name": "Development Output",
                    "output_dataset": "Development Output",
                    "output_type": "Development Output",
                    "input_triangle": "Paid",
                    "origin_length": 12,
                },
                "data_tab": {
                    "origin_labels": ["2024", "2025"],
                    "development_labels": ["12", "24", "36"],
                    "input_data_triangle_values": [[100, 150], [200]],
                },
                "ratios_tab": {"average_formulas": formulas},
                "results_tab": {},
            },
            timestamp="2026-01-01T00:00:00Z",
            update_refresh_timestamp=False,
        )

    @staticmethod
    def write_json(path: Path, payload: dict) -> None:
        path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")

    def output_sidecar(self, method: dict, *, status: int = 0) -> dict:
        sidecar = build_bornhuetter_ferguson_output_sidecar(
            method,
            project_name="Project",
            reserving_class="Class",
            csv_file="BF Method@12.csv",
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
        self.write_json(self.methods / "BF@BF Method.json", payload)
        self.write_json(self.sidecars / "BF Method.json", self.output_sidecar(payload, status=status))
        (self.datasets / "BF Method@12.csv").write_text(
            "\n".join(str(value) for value in payload["method_tab"]["new_ultimate"]) + "\n",
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
        if method_type == "DFM":
            # A dependent reads its percentage developed from the DFM method,
            # not by dividing the published ultimates, so a DFM source is only
            # complete once its method file is on disk beside the sidecar.
            self.write_json(self.methods / f"DFM@{name}.json", self.dfm_method_payload())

    def write_all_sources(self, *, paid_csv: str = "100,150\n200,\n") -> None:
        self.write_source("Paid", paid_csv, data_format="Triangle", dependents=["BF Method"])
        self.write_source(
            "Development Output",
            "300\n400\n",
            data_format="Vector",
            method_type="DFM",
            dependents=["BF Method"],
        )
        self.write_source("Prior", "500\n600\n", data_format="Vector", dependents=["BF Method"])

    def test_v3_load_reads_only_method_and_own_sidecar(self) -> None:
        self.write_method_pair()
        original = bornhuetter_ferguson_service._read_json
        reads: list[str] = []

        def recording(path: str) -> dict:
            reads.append(str(Path(path)))
            return original(path)

        with (
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_read_json",
                side_effect=recording,
            ),
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=AssertionError("source read"),
            ),
            mock.patch.object(
                bornhuetter_ferguson_service,
                "bornhuetter_ferguson_output_variants",
                side_effect=AssertionError("output recalculation"),
            ),
        ):
            result = bornhuetter_ferguson_service.load_bornhuetter_ferguson_method(
                "Project", "Class", "BF Method"
            )

        self.assertTrue(result["ok"])
        self.assertCountEqual(reads, [
            str(self.methods / "BF@BF Method.json"),
            str(self.sidecars / "BF Method.json"),
        ])

    def test_v2_load_is_rejected_without_source_reads_or_writes(self) -> None:
        method = self.write_method_pair()
        method["json_format"] = "arcrho-bornhuetter-ferguson-method-by-tab-v2"
        method_path = self.methods / "BF@BF Method.json"
        self.write_json(method_path, method)
        method_before = method_path.read_bytes()
        sidecar_path = self.sidecars / "BF Method.json"
        sidecar_before = sidecar_path.read_bytes()

        with (
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=AssertionError("source read"),
            ),
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_publish",
                side_effect=AssertionError("publication write"),
            ),
            self.assertRaises(HTTPException) as raised,
        ):
            bornhuetter_ferguson_service.load_bornhuetter_ferguson_method(
                "Project", "Class", "BF Method"
            )

        self.assertEqual(raised.exception.status_code, 422)
        self.assertIn("Unsupported BF JSON format", str(raised.exception.detail))
        self.assertEqual(method_path.read_bytes(), method_before)
        self.assertEqual(sidecar_path.read_bytes(), sidecar_before)

    def test_v3_load_rejects_method_sidecar_geometry_mismatch(self) -> None:
        self.write_method_pair()
        sidecar_path = self.sidecars / "BF Method.json"
        sidecar = json.loads(sidecar_path.read_text(encoding="utf-8"))
        sidecar["origin_labels"] = ["2023", "2025"]
        self.write_json(sidecar_path, sidecar)

        with self.assertRaises(HTTPException) as raised:
            bornhuetter_ferguson_service.load_bornhuetter_ferguson_method(
                "Project", "Class", "BF Method"
            )

        self.assertEqual(raised.exception.status_code, 409)
        self.assertIn("origin labels do not match", str(raised.exception.detail))

    def test_save_rebases_owned_weights_over_newer_disk_derived_snapshot(self) -> None:
        stale = self.method_payload()
        current = self.method_payload(latest_values=[[100, 175], [200, None]])
        self.write_method_pair(current)
        incoming = copy.deepcopy(stale)
        incoming["method_tab"]["prior_datasets"][0]["weights"] = [0.5, 1]

        with (
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=AssertionError("source read"),
            ),
            mock.patch.object(
                calculated_dataset_service,
                "recalculate_dependents",
                return_value={"ok": True, "updated": [], "index_ok": True},
            ),
        ):
            result = bornhuetter_ferguson_service.save_bornhuetter_ferguson_method(
                "Project",
                "Class",
                incoming,
                expected_owned_revision=method_revisions(stale)["owned_revision"],
                expected_derived_revision=method_revisions(stale)["derived_revision"],
            )

        self.assertTrue(result["derived_rebased"])
        saved = json.loads((self.methods / "BF@BF Method.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["latest_values"], current["method_tab"]["latest_values"])
        self.assertEqual(saved["method_tab"]["prior_datasets"][0]["weights"], [0.5, 1])

    def test_no_op_save_submits_no_engine_propagation_job(self) -> None:
        method = self.write_method_pair()

        with (
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=AssertionError("source read"),
            ),
            mock.patch.object(
                bornhuetter_ferguson_service.dependent_propagation_service,
                "require_reserving_class_writable",
            ),
            mock.patch.object(
                bornhuetter_ferguson_service.dependent_propagation_service,
                "enqueue_marked_save_propagation",
            ) as enqueue,
        ):
            result = bornhuetter_ferguson_service.save_bornhuetter_ferguson_method(
                "Project",
                "Class",
                method,
                expected_owned_revision=method_revisions(method)["owned_revision"],
            )

        self.assertEqual(result["sidecar"]["status"], 0)
        self.assertTrue(result["propagation_ok"])
        self.assertEqual(result["propagation"], {"ok": True, "status": "unchanged"})
        enqueue.assert_not_called()

    def test_review_needed_save_uses_embedded_snapshots_before_restoring_current(self) -> None:
        method = self.write_method_pair(status=2)
        self.write_all_sources()
        dfm_path = self.sidecars / "Development Output.json"
        dfm_source = json.loads(dfm_path.read_text(encoding="utf-8"))
        dfm_source["status"] = 2
        self.write_json(dfm_path, dfm_source)
        with (
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=AssertionError("source read"),
            ),
            mock.patch.object(
                calculated_dataset_service,
                "recalculate_dependents",
                return_value={"ok": True, "updated": [], "index_ok": True},
            ),
        ):
            result = bornhuetter_ferguson_service.save_bornhuetter_ferguson_method(
                "Project",
                "Class",
                method,
                expected_owned_revision=method_revisions(method)["owned_revision"],
            )

        self.assertEqual(result["sidecar"]["status"], 0)
        self.assertEqual(result["unreviewed_precedents"], ["Development Output"])
        self.assertEqual(result["unreviewed_precedent_count"], 1)

    def test_review_needed_prior_refresh_uses_method_origins_and_persists_new_values(self) -> None:
        self.write_method_pair(status=2)
        self.write_source(
            "Paid",
            "100,150\n200,\n",
            data_format="Triangle",
            dependents=["BF Method"],
            include_origin_labels=False,
        )
        self.write_source(
            "Development Output",
            "300\n400\n",
            data_format="Vector",
            method_type="DFM",
            dependents=["BF Method"],
            status=2,
            include_origin_labels=False,
        )
        self.write_source(
            "Prior",
            "700\n800\n",
            data_format="Vector",
            dependents=["BF Method"],
            include_origin_labels=False,
        )
        source_reads: list[str] = []
        original = bornhuetter_ferguson_service._read_source_snapshot_from_sidecar

        def recording(*args, **kwargs):
            source_reads.append(str(args[2]))
            return original(*args, **kwargs)

        with (
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=recording,
            ),
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_refresh_downstream_domains",
                return_value={"ok": True, "updated": []},
            ),
        ):
            result = bornhuetter_ferguson_service.refresh_dependents(
                "Project", "Class", ["Prior"], rebuild_index=False
            )

        self.assertTrue(result["ok"], result)
        self.assertEqual(source_reads, ["Prior"])
        self.assertEqual(result["errors"], [])
        saved = json.loads((self.methods / "BF@BF Method.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["prior_datasets"][0]["values"], [700, 800])
        self.assertEqual(saved["method_tab"]["origin_labels"], ["2024", "2025"])
        self.assertEqual(saved["method_tab"]["new_ultimate"], [500, 600])
        self.assertEqual(
            (self.datasets / "BF Method@12.csv").read_text(encoding="utf-8"),
            "500\n600\n",
        )

    def test_status_two_dfm_refresh_consumes_valid_publication(self) -> None:
        self.write_method_pair(status=2)
        self.write_source(
            "Paid",
            "100,150\n200,\n",
            data_format="Triangle",
            dependents=["BF Method"],
            include_origin_labels=False,
        )
        self.write_source(
            "Development Output",
            "350\n450\n",
            data_format="Vector",
            method_type="DFM",
            dependents=["BF Method"],
            status=2,
            include_origin_labels=False,
        )
        self.write_source(
            "Prior",
            "500\n600\n",
            data_format="Vector",
            dependents=["BF Method"],
            include_origin_labels=False,
        )
        source_reads: list[str] = []
        original = bornhuetter_ferguson_service._read_source_snapshot_from_sidecar

        def recording(*args, **kwargs):
            source_reads.append(str(args[2]))
            return original(*args, **kwargs)

        with (
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_read_source_snapshot_from_sidecar",
                side_effect=recording,
            ),
            mock.patch.object(
                bornhuetter_ferguson_service,
                "_refresh_downstream_domains",
                return_value={"ok": True, "updated": []},
            ),
        ):
            result = bornhuetter_ferguson_service.refresh_dependents(
                "Project", "Class", ["Development Output"], rebuild_index=False
            )

        self.assertTrue(result["ok"], result)
        self.assertEqual(source_reads, ["Development Output"])
        saved = json.loads((self.methods / "BF@BF Method.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["dfm_ultimate_values"], [350, 450])

    def test_source_snapshot_cache_is_scoped_to_method_origin_axis(self) -> None:
        self.write_source(
            "Prior",
            "500\n600\n",
            data_format="Vector",
            include_origin_labels=False,
        )
        first = self.method_payload()
        second = copy.deepcopy(first)
        second["method_tab"]["origin_labels"] = ["2023", "2025"]
        cache = {}

        first_snapshot = bornhuetter_ferguson_service._source_snapshots(
            "Project",
            "Class",
            first,
            {"priors"},
            snapshot_cache=cache,
        )
        second_snapshot = bornhuetter_ferguson_service._source_snapshots(
            "Project",
            "Class",
            second,
            {"priors"},
            snapshot_cache=cache,
        )

        self.assertEqual(first_snapshot["priors"]["Prior"]["origin_labels"], ["2024", "2025"])
        self.assertEqual(second_snapshot["priors"]["Prior"]["origin_labels"], ["2023", "2025"])
        self.assertEqual(len(cache), 2)

    def test_a_finer_hand_entered_precedent_is_rolled_up_to_the_method(self) -> None:
        """A quarterly vector shown yearly feeds the yearly method summed by year."""

        self.write_source("Prior", "100\n200\n300\n400\n50\n60\n70\n80\n", data_format="Vector")
        path = self.sidecars / "Prior.json"
        sidecar = json.loads(path.read_text(encoding="utf-8"))
        sidecar["stored_period_length"] = 3
        self.write_json(path, sidecar)

        snapshot = bornhuetter_ferguson_service._source_snapshots(
            "Project", "Class", self.method_payload(), {"priors"},
        )

        self.assertEqual(snapshot["priors"]["Prior"]["values"], [[1000], [260]])

        # The window was left on a yearly view of yearly data: nothing to roll
        # up, and the method loads the file as it is.
        sidecar["stored_period_length"] = 12
        sidecar["period_length"] = 36
        self.write_json(path, sidecar)
        (self.datasets / "Prior@12.csv").write_text("500\n600\n", encoding="utf-8")

        snapshot = bornhuetter_ferguson_service._source_snapshots(
            "Project", "Class", self.method_payload(), {"priors"},
        )

        self.assertEqual(snapshot["priors"]["Prior"]["values"], [[500], [600]])

    def test_a_coarser_precedent_is_still_refused(self) -> None:
        """Yearly figures cannot be split, so a three-year vector stays refused."""

        self.write_source("Prior", "500\n600\n", data_format="Vector")
        path = self.sidecars / "Prior.json"
        sidecar = json.loads(path.read_text(encoding="utf-8"))
        sidecar["stored_period_length"] = 36
        self.write_json(path, sidecar)

        with self.assertRaisesRegex(HTTPException, "uses 36-month origins; expected 12"):
            bornhuetter_ferguson_service._source_snapshots(
                "Project", "Class", self.method_payload(), {"priors"},
            )

    def test_a_finer_generated_precedent_is_rebuilt_at_the_method_period(self) -> None:
        """A monthly Engine vector is regenerated yearly rather than refused."""

        self.write_source("Prior", "1\n2\n3\n", data_format="Vector")
        path = self.sidecars / "Prior.json"
        sidecar = json.loads(path.read_text(encoding="utf-8"))
        sidecar["source_kind"] = "engine"
        sidecar["stored_period_length"] = 1
        self.write_json(path, sidecar)
        rebuilt = self.datasets / "Prior@12.rebuilt.csv"
        rebuilt.write_text("700\n800\n", encoding="utf-8")

        with mock.patch.object(
            bornhuetter_ferguson_service.precedent_cache_service,
            "materialize_engine_source",
            return_value=str(rebuilt),
        ) as materialize:
            snapshot = bornhuetter_ferguson_service._source_snapshots(
                "Project", "Class", self.method_payload(), {"priors"},
            )

        self.assertEqual(snapshot["priors"]["Prior"]["values"], [[700], [800]])
        self.assertEqual(materialize.call_args.args[2:], ("Prior", sidecar, 12))

    def test_explicit_refresh_keeps_review_alert_until_save(self) -> None:
        self.write_method_pair(status=2)
        self.write_all_sources()

        with mock.patch.object(
            bornhuetter_ferguson_service,
            "_refresh_downstream_domains",
            return_value={"ok": True, "updated": []},
        ) as cascade:
            result = bornhuetter_ferguson_service.refresh_bornhuetter_ferguson_method(
                "Project",
                "Class",
                "BF Method",
            )

        self.assertFalse(result["output_changed"])
        self.assertFalse(result["status_refreshed"])
        cascade.assert_not_called()
        sidecar = json.loads((self.sidecars / "BF Method.json").read_text(encoding="utf-8"))
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
            dependents=["BF Method"],
        )
        method_path = self.methods / "BF@BF Method.json"
        output_path = self.datasets / "BF Method@12.csv"
        before_method = method_path.read_bytes()
        before_output = output_path.read_bytes()

        result = bornhuetter_ferguson_service.refresh_dependents(
            "Project", "Class", ["Paid"], rebuild_index=False
        )

        self.assertFalse(result["ok"])
        self.assertEqual(method_path.read_bytes(), before_method)
        self.assertEqual(output_path.read_bytes(), before_output)
        sidecar = json.loads((self.sidecars / "BF Method.json").read_text(encoding="utf-8"))
        self.assertEqual(sidecar["status"], dataset_sidecar_status_service.STATUS_REVIEW_NEEDED)

    def test_paid_refresh_rewrites_embedded_values_and_output_but_preserves_weights(self) -> None:
        method = self.write_method_pair()
        original_weights = copy.deepcopy(method["method_tab"]["prior_datasets"][0]["weights"])
        self.write_all_sources(paid_csv="100,175\n200,\n")
        old_revision = self.output_sidecar(method)["publication_revision"]
        with mock.patch.object(
            bornhuetter_ferguson_service,
            "_refresh_downstream_domains",
            return_value={"ok": True, "updated": []},
        ):
            result = bornhuetter_ferguson_service.refresh_dependents(
                "Project", "Class", ["Paid"], rebuild_index=False
            )

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["updated"][0]["dataset_name"], "BF Method")
        self.assertTrue(result["updated"][0]["output_changed"])
        saved = json.loads((self.methods / "BF@BF Method.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["method_tab"]["latest_values"], [175, 200])
        self.assertEqual(
            saved["method_tab"]["prior_datasets"][0]["weights"],
            original_weights,
        )
        self.assertNotEqual(
            saved["method_tab"]["new_ultimate"],
            method["method_tab"]["new_ultimate"],
        )
        sidecar = json.loads((self.sidecars / "BF Method.json").read_text(encoding="utf-8"))
        self.assertEqual(
            sidecar["status"],
            dataset_sidecar_status_service.STATUS_REVIEW_NEEDED,
        )
        self.assertEqual(
            result["review_status_updates"],
            [{"dataset_name": "BF Method", "status": 2}],
        )
        self.assertEqual(
            (self.datasets / "BF Method@12.csv").read_text(encoding="utf-8").splitlines(),
            [str(value) for value in saved["method_tab"]["new_ultimate"]],
        )
        sidecar = json.loads((self.sidecars / "BF Method.json").read_text(encoding="utf-8"))
        self.assertNotEqual(sidecar["publication_revision"], old_revision)
        self.assertEqual(sidecar["publication_revision"], method_revisions(saved)["publication_revision"])

    def test_source_executor_is_bounded_for_network_reads(self) -> None:
        self.assertEqual(bornhuetter_ferguson_service.READ_MAX_WORKERS, 4)
        self.assertEqual(
            bornhuetter_ferguson_service._READ_EXECUTOR._max_workers,
            bornhuetter_ferguson_service.READ_MAX_WORKERS,
        )

    def test_refresh_report_order_is_deterministic(self) -> None:
        self.write_json(self.sidecars / "Paid.json", {
            "dataset_name": "Paid",
            "dependents": [
                {"dataset_name": "BF Z"},
                {"dataset_name": "BF A"},
            ],
        })
        for name in ("BF Z", "BF A"):
            self.write_json(self.sidecars / f"{name}.json", {
                "dataset_name": name,
                "method_name": name,
                "method_type": "Bornhuetter Ferguson",
                "source_kind": "bornhuetter_ferguson",
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
            bornhuetter_ferguson_service,
            "_refresh_one",
            side_effect=refreshed,
        ):
            result = bornhuetter_ferguson_service.refresh_dependents(
                "Project", "Class", ["Paid"], rebuild_index=False
            )

        self.assertEqual(
            [item["dataset_name"] for item in result["updated"]],
            ["BF A", "BF Z"],
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
        real_replace = bornhuetter_ferguson_service.os.replace
        targets: list[str] = []

        def replace(source: str, target: str) -> None:
            targets.append(target)
            if target == str(sidecar_path) and source.endswith(".tmp"):
                raise OSError("sidecar publish failed")
            real_replace(source, target)

        with mock.patch.object(
            bornhuetter_ferguson_service.os,
            "replace",
            side_effect=replace,
        ):
            with self.assertRaises(OSError):
                bornhuetter_ferguson_service._commit_text_files(
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

    def test_bf_nested_cascade_names_include_rs_calculated_and_dfm_outputs(self) -> None:
        fresh, failed = bornhuetter_ferguson_service._cascade_names({
            "ok": False,
            "updated": [{"dataset_type_name": "Calculated Direct"}],
            "result_selection_updates": {
                "updated": [{"dataset_name": "RS Output"}],
                "downstream_fresh_names": ["Calculated After RS", "DFM After RS"],
                "downstream_blocked_names": ["Failed After RS"],
            },
        })

        self.assertCountEqual(
            fresh,
            ["Calculated Direct", "RS Output", "Calculated After RS", "DFM After RS"],
        )
        self.assertEqual(failed, ["Failed After RS"])

    def test_outer_cascade_passes_rs_nested_outputs_to_bf_wave(self) -> None:
        with (
            mock.patch(
                "app_server.services.dfm_service.refresh_dependents",
                return_value={"ok": True, "updated": [], "errors": []},
            ),
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=[]),
            mock.patch.object(calculated_dataset_service, "_existing_downstream_keys", return_value=[]),
            mock.patch(
                "app_server.services.result_selection_service.refresh_dependents",
                return_value={
                    "ok": True,
                    "updated": [{"dataset_name": "RS Output"}],
                    "status_refreshed": [],
                    "errors": [],
                    "downstream_fresh_names": ["Calculated After RS", "DFM After RS"],
                    "downstream_blocked_names": ["Failed After RS"],
                },
            ),
            mock.patch.object(
                bornhuetter_ferguson_service,
                "refresh_dependents",
                return_value={"ok": True, "updated": [], "errors": []},
            ) as refresh_bf,
            mock.patch.object(
                calculated_dataset_service.dataset_instance_index_service,
                "rebuild_index",
            ),
        ):
            result = calculated_dataset_service.recalculate_dependents(
                "Project", "Class", "Paid", "Paid"
            )

        self.assertTrue(result["ok"], result)
        roots = refresh_bf.call_args.args[2]
        blocked = refresh_bf.call_args.kwargs["blocked_precedent_names"]
        self.assertIn("RS Output", roots)
        self.assertIn("Calculated After RS", roots)
        self.assertIn("DFM After RS", roots)
        self.assertIn("Failed After RS", blocked)

    def test_rs_report_exposes_nested_calculated_and_dfm_outputs_for_bf(self) -> None:
        self.write_json(self.sidecars / "Paid.json", {
            "dataset_name": "Paid",
            "dependents": [{"dataset_name": "RS Output"}],
        })
        rs_sidecar = {
            "dataset_name": "RS Output",
            "dataset_type": "Selected Ultimate",
            "method_name": "RS Method",
            "method_type": "Result Selection",
            "source_kind": "result_selection",
            "status": 0,
            "dependents": [],
        }
        self.write_json(self.sidecars / "RS Output.json", rs_sidecar)
        nested = {
            "ok": True,
            "updated": [{"dataset_type_name": "Calculated After RS"}],
            "skipped": [],
            "dfm_updates": {
                "ok": True,
                "updated": [{"dataset_name": "DFM After RS"}],
                "status_refreshed": [],
                "errors": [],
            },
        }
        with (
            mock.patch.object(result_selection_service, "_assert_acyclic_dependency_subgraph"),
            mock.patch.object(
                result_selection_service,
                "_refresh_one_method",
                return_value={
                    "ok": True,
                    "dataset_name": "RS Output",
                    "updated": True,
                    "output_changed": True,
                    "sidecar": rs_sidecar,
                },
            ),
            mock.patch.object(
                dataset_sidecar_status_service,
                "refresh_method_statuses_for_dependents",
                return_value=[],
            ),
            mock.patch.object(
                calculated_dataset_service,
                "recalculate_dependents",
                return_value=nested,
            ),
        ):
            result = result_selection_service.refresh_dependents(
                "Project", "Class", ["Paid"], rebuild_index=False
            )

        self.assertCountEqual(
            result["downstream_fresh_names"],
            ["Calculated After RS", "DFM After RS"],
        )


if __name__ == "__main__":
    unittest.main()

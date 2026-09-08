"""The automatic Berquist Sherman refresh the dependent-propagation walk runs."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO_ROOT = Path(__file__).resolve().parents[2]
FRONTEND_ROOT = REPO_ROOT / "frontend"
PYTHON_API_SRC = REPO_ROOT / "python-api" / "src"
for path in (FRONTEND_ROOT, PYTHON_API_SRC):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from arcrho_api.berquist_sherman_contract import (
    BS_SR_JSON_FORMAT,
    berquist_sherman_output_csv_text,
    calculate_settlement_rate,
)
from arcrho_api.sidecar_audit_contract import AUDIT_ACTION_AUTO_REFRESH
from app_server.services import (
    berquist_sherman_service,
    calculated_dataset_service,
    dataset_sidecar_status_service,
)

SR_TYPE = "B&S Settlement Rate Adjustment"
OUTPUT = "Paid - B&S Settlement Rate Adjustment"
PAID = [[100, 400, 900], [200, 500], [300]]
CLOSED = [[10, 20, 20], [8, 16], [5]]


def _csv(rows: list[list[float | None]], width: int) -> str:
    return "\n".join(
        ",".join("" if index >= len(row) or row[index] is None else str(row[index]) for index in range(width))
        for row in rows
    ) + "\n"


class BerquistShermanRefreshTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=str(FRONTEND_ROOT / "tests"))
        root = Path(self.temp.name)
        self.methods = root / "methods"
        self.datasets = root / "datasets"
        self.sidecars = root / "sidecars"
        for folder in (self.methods, self.datasets, self.sidecars):
            folder.mkdir()
        self.patchers = [
            mock.patch.object(
                dataset_sidecar_status_service.config,
                "get_project_method_data_dir",
                return_value=str(self.methods),
            ),
            mock.patch.object(
                berquist_sherman_service.config,
                "get_project_dataset_cache_dir",
                return_value=str(self.datasets),
            ),
            mock.patch.object(
                dataset_sidecar_status_service,
                "sidecar_path",
                side_effect=lambda _p, _r, name: str(self.sidecars / f"{name}.json"),
            ),
            mock.patch.object(
                berquist_sherman_service.user_identity_service,
                "get_current_display_name",
                return_value="Engine Walker",
            ),
        ]
        for patcher in self.patchers:
            patcher.start()

    def tearDown(self) -> None:
        for patcher in reversed(self.patchers):
            patcher.stop()
        self.temp.cleanup()

    # -- fixtures --------------------------------------------------------

    def write_json(self, path: Path, payload: dict) -> None:
        path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")

    def read_json(self, path: Path) -> dict:
        return json.loads(path.read_text(encoding="utf-8"))

    def write_source(self, name: str, csv_text: str, *, data_format: str, status: int = 0) -> None:
        csv_file = f"{name}@12@12@cum@dev.csv" if data_format == "Triangle" else f"{name}@12.csv"
        (self.datasets / csv_file).write_text(csv_text, encoding="utf-8")
        sidecar = {
            "dataset_name": name,
            "dataset_type": name,
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "input",
            "method_type": "None",
            "data_format": data_format,
            "csv_file": csv_file,
            "status": status,
            "precedents": [],
            "dependents": [{"dataset_name": OUTPUT}],
        }
        if data_format == "Triangle":
            sidecar["origin_length"] = 12
            sidecar["development_length"] = 12
            sidecar["stored_origin_length"] = 12
            sidecar["stored_development_length"] = 12
        else:
            sidecar["period_length"] = 12
            sidecar["stored_period_length"] = 12
        self.write_json(self.sidecars / f"{name}.json", sidecar)

    def method_payload(self) -> dict:
        return {
            "json_format": BS_SR_JSON_FORMAT,
            "details_tab": {
                "name": OUTPUT,
                "method_type": SR_TYPE,
                "output_type": "Gross Loss - ad hoc",
                "origin_length": 12,
                "development_length": 12,
            },
            "method_tab": {
                "origin_labels": ["2024", "2025", "2026"],
                "development_labels": ["12", "24", "36"],
                "paid_claims": "Paid",
                "closed_claim_numbers": "Closed",
                "ultimate_claim_numbers": "Ultimate Counts",
                "selected_proportion_settled": [0.25, 0.8, 1.0],
                "selected_proportion_is_default": [True, True, True],
                "selected_adjustment": [["pairs", "pairs", "pairs"], ["pairs", "pairs"], ["pairs"]],
                "loess_span": 7,
            },
            "method_metadata": {
                "method_type": SR_TYPE,
                "source_kind": "berquist_sherman_sr",
                "last_modified": "2026-01-01T00:00:00.000Z",
            },
        }

    def expected_output(self, ultimate: list[float]) -> list[list[float | None]]:
        tab = self.method_payload()["method_tab"]
        return calculate_settlement_rate({
            "paid_claims": PAID,
            "closed_claim_numbers": CLOSED,
            "ultimate_claim_numbers": ultimate,
            "selected_proportion_settled": tab["selected_proportion_settled"],
            "selected_proportion_is_default": tab["selected_proportion_is_default"],
            "selected_adjustment": tab["selected_adjustment"],
            "loess_span": tab["loess_span"],
        })["output"]

    def write_method_pair(self, *, ultimate: list[float], status: int = 2) -> None:
        self.write_json(self.methods / f"BSSR@{OUTPUT}.json", self.method_payload())
        (self.datasets / f"{OUTPUT}@12@12@cum@dev.csv").write_text(
            berquist_sherman_output_csv_text(self.expected_output(ultimate), 3),
            encoding="utf-8",
        )
        self.write_json(self.sidecars / f"{OUTPUT}.json", {
            "dataset_name": OUTPUT,
            "dataset_type": "Gross Loss - ad hoc",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "berquist_sherman_sr",
            "method_type": SR_TYPE,
            "method_name": OUTPUT,
            "data_format": "Triangle",
            "origin_length": 12,
            "development_length": 12,
            "csv_file": f"{OUTPUT}@12@12@cum@dev.csv",
            "status": status,
            "modified_by": "Wei, Xiao",
            "updated_at": "2026-01-01T00:00:00.000Z",
            "precedents": [
                {"dataset_name": "Paid"},
                {"dataset_name": "Closed"},
                {"dataset_name": "Ultimate Counts"},
            ],
            "dependents": [],
            "audit_log": [
                {"event_date": "2026-01-01T00:00:00.000Z", "action": "Update", "change_info": "Values", "user": "Wei, Xiao"},
            ],
        })

    def write_workspace(self, *, saved_ultimate: list[float], current_ultimate: list[float]) -> None:
        self.write_source("Paid", _csv(PAID, 3), data_format="Triangle")
        self.write_source("Closed", _csv(CLOSED, 3), data_format="Triangle")
        self.write_source("Ultimate Counts", _csv([[v] for v in current_ultimate], 1), data_format="Vector")
        self.write_method_pair(ultimate=saved_ultimate)

    # -- refresh_dependents ---------------------------------------------

    def test_a_moved_ultimate_vector_rewrites_the_output_csv_sidecar_and_method(self) -> None:
        self.write_workspace(saved_ultimate=[20, 20, 20], current_ultimate=[20, 20, 40])
        with mock.patch.object(
            berquist_sherman_service,
            "_refresh_downstream_domains",
            return_value={"ok": True, "updated": [{"dataset_type_name": "Calc After BS"}], "skipped": []},
        ) as cascade:
            report = berquist_sherman_service.refresh_dependents(
                "Project", "Class", ["Ultimate Counts"], rebuild_index=False, finalize_method_review_status=False
            )

        self.assertTrue(report["ok"], report)
        self.assertEqual(
            report["updated"],
            [{"dataset_name": OUTPUT, "dataset_type": "Gross Loss - ad hoc", "output_changed": True}],
        )
        self.assertEqual(report["downstream_fresh_names"], ["Calc After BS"])
        cascade.assert_called_once()
        self.assertEqual(cascade.call_args.args[2], OUTPUT)

        csv_text = (self.datasets / f"{OUTPUT}@12@12@cum@dev.csv").read_text(encoding="utf-8")
        self.assertEqual(csv_text, berquist_sherman_output_csv_text(self.expected_output([20, 20, 40]), 3))
        sidecar = self.read_json(self.sidecars / f"{OUTPUT}.json")
        self.assertEqual(sidecar["status"], 0)
        self.assertEqual(sidecar["modified_by"], "Engine Walker")
        self.assertNotEqual(sidecar["updated_at"], "2026-01-01T00:00:00.000Z")
        self.assertEqual(sidecar["audit_log"][-1]["action"], AUDIT_ACTION_AUTO_REFRESH)
        self.assertEqual(sidecar["audit_log"][-1]["user"], "Engine Walker")
        self.assertEqual(sidecar["audit_log"][-1]["event_date"], sidecar["updated_at"])
        self.assertEqual(list(sidecar)[-1], "audit_log")
        method = self.read_json(self.methods / f"BSSR@{OUTPUT}.json")
        self.assertEqual(method["method_metadata"]["last_modified"], sidecar["updated_at"])
        # The stored selections are the page's; the refresh does not rewrite them.
        self.assertEqual(method["method_tab"], self.method_payload()["method_tab"])

    def test_an_unchanged_output_only_restores_the_review_status(self) -> None:
        self.write_workspace(saved_ultimate=[20, 20, 20], current_ultimate=[20, 20, 20])
        csv_path = self.datasets / f"{OUTPUT}@12@12@cum@dev.csv"
        before_csv = csv_path.read_text(encoding="utf-8")
        with mock.patch.object(berquist_sherman_service, "_refresh_downstream_domains") as cascade:
            report = berquist_sherman_service.refresh_dependents(
                "Project", "Class", ["Ultimate Counts"], rebuild_index=False, finalize_method_review_status=False
            )

        self.assertTrue(report["ok"], report)
        self.assertEqual(report["updated"], [])
        self.assertEqual(report["status_refreshed"], [{"dataset_name": OUTPUT}])
        self.assertEqual(csv_path.read_text(encoding="utf-8"), before_csv)
        sidecar = self.read_json(self.sidecars / f"{OUTPUT}.json")
        self.assertEqual(sidecar["status"], 0)
        self.assertEqual(sidecar["updated_at"], "2026-01-01T00:00:00.000Z")
        self.assertEqual(len(sidecar["audit_log"]), 1)
        method = self.read_json(self.methods / f"BSSR@{OUTPUT}.json")
        self.assertEqual(method["method_metadata"]["last_modified"], "2026-01-01T00:00:00.000Z")
        # A restored status still tells the downstream domains the source is fresh.
        cascade.assert_called_once()

    def test_a_source_the_refresh_cannot_read_marks_the_output_for_review(self) -> None:
        self.write_workspace(saved_ultimate=[20, 20, 20], current_ultimate=[20, 20, 40])
        (self.datasets / "Closed@12@12@cum@dev.csv").unlink()
        sidecar_path = self.sidecars / f"{OUTPUT}.json"
        sidecar = self.read_json(sidecar_path)
        sidecar["status"] = 0
        self.write_json(sidecar_path, sidecar)
        csv_path = self.datasets / f"{OUTPUT}@12@12@cum@dev.csv"
        before_csv = csv_path.read_text(encoding="utf-8")

        with mock.patch.object(berquist_sherman_service, "_refresh_downstream_domains") as cascade:
            report = berquist_sherman_service.refresh_dependents(
                "Project", "Class", ["Ultimate Counts"], rebuild_index=False, finalize_method_review_status=False
            )

        self.assertFalse(report["ok"])
        self.assertEqual(report["errors"][0]["dataset_name"], OUTPUT)
        self.assertIn("Closed", report["errors"][0]["reason"])
        self.assertEqual(csv_path.read_text(encoding="utf-8"), before_csv)
        self.assertEqual(self.read_json(sidecar_path)["status"], 2)
        cascade.assert_not_called()

    def test_a_source_of_the_wrong_period_is_refused_like_the_page_refuses_it(self) -> None:
        self.write_workspace(saved_ultimate=[20, 20, 20], current_ultimate=[20, 20, 40])
        ultimate_path = self.sidecars / "Ultimate Counts.json"
        ultimate = self.read_json(ultimate_path)
        # Annual is a question about the grid the dataset is shown at, which is
        # the shape the page tests before it will take a source: a quarterly
        # vector is refused here for the same reason it is refused there.
        ultimate["period_length"] = 3
        ultimate["stored_period_length"] = 3
        self.write_json(ultimate_path, ultimate)

        with mock.patch.object(berquist_sherman_service, "_refresh_downstream_domains"):
            report = berquist_sherman_service.refresh_dependents(
                "Project", "Class", ["Ultimate Counts"], rebuild_index=False, finalize_method_review_status=False
            )

        self.assertFalse(report["ok"])
        self.assertIn("not an annual dataset", report["errors"][0]["reason"])

    def test_a_source_stored_finer_than_its_annual_display_is_rolled_up(self) -> None:
        self.write_workspace(saved_ultimate=[20, 20, 20], current_ultimate=[20, 20, 40])
        paid_path = self.sidecars / "Paid.json"
        paid = self.read_json(paid_path)
        # The triangle is entered and shown annually but held one column per
        # month, the shape an Excel-linked adjusted triangle is saved at. Its
        # own file is rolled up to the annual grid rather than read as 36
        # development periods of mostly cumulative zero.
        paid["stored_development_length"] = 1
        paid["csv_file"] = "Paid@12@1@cum@dev.csv"
        self.write_json(paid_path, paid)
        monthly = [
            [None] * 36,
            [None] * 36,
            [None] * 36,
        ]
        for row_index, row in enumerate(PAID):
            for column_index, value in enumerate(row):
                # Development periods count back from the valuation date, so
                # every row's annual figures sit in its own 12th, 24th and 36th
                # month whatever calendar date that row started at.
                monthly[row_index][11 + 12 * column_index] = value
        (self.datasets / "Paid@12@1@cum@dev.csv").write_text(_csv(monthly, 36), encoding="utf-8")
        (self.datasets / "Paid@12@12@cum@dev.csv").unlink()

        with mock.patch.object(
            berquist_sherman_service.dataset_service, "valuation_months", return_value=36
        ), mock.patch.object(berquist_sherman_service, "_refresh_downstream_domains"):
            report = berquist_sherman_service.refresh_dependents(
                "Project", "Class", ["Ultimate Counts"], rebuild_index=False, finalize_method_review_status=False
            )

        self.assertTrue(report["ok"], report)
        csv_text = (self.datasets / f"{OUTPUT}@12@12@cum@dev.csv").read_text(encoding="utf-8")
        self.assertEqual(csv_text, berquist_sherman_output_csv_text(self.expected_output([20, 20, 40]), 3))

    def test_a_generated_sources_source_table_granularity_is_not_read_as_its_shape(self) -> None:
        self.write_workspace(saved_ultimate=[20, 20, 20], current_ultimate=[20, 20, 40])
        ultimate_path = self.sidecars / "Ultimate Counts.json"
        ultimate = self.read_json(ultimate_path)
        # A generated dataset's stored pair is how fine the project's source
        # table is, not the shape of the annual cache beside it, so the annual
        # check must not read it.
        ultimate["source_kind"] = "engine"
        ultimate["stored_period_length"] = 1
        self.write_json(ultimate_path, ultimate)

        with mock.patch.object(berquist_sherman_service, "_refresh_downstream_domains"):
            report = berquist_sherman_service.refresh_dependents(
                "Project", "Class", ["Ultimate Counts"], rebuild_index=False, finalize_method_review_status=False
            )

        self.assertTrue(report["ok"], report)
        csv_text = (self.datasets / f"{OUTPUT}@12@12@cum@dev.csv").read_text(encoding="utf-8")
        self.assertEqual(csv_text, berquist_sherman_output_csv_text(self.expected_output([20, 20, 40]), 3))

    def test_a_generated_source_without_an_annual_cache_is_rebuilt_at_the_annual_period(self) -> None:
        self.write_workspace(saved_ultimate=[20, 20, 20], current_ultimate=[20, 20, 40])
        ultimate_path = self.sidecars / "Ultimate Counts.json"
        ultimate = self.read_json(ultimate_path)
        ultimate["source_kind"] = "engine"
        ultimate["stored_period_length"] = 1
        ultimate["csv_file"] = "Ultimate Counts@1.csv"
        self.write_json(ultimate_path, ultimate)
        annual_cache = self.datasets / "Ultimate Counts@12.csv"
        rebuilt = self.datasets / "Ultimate Counts@1.csv"
        rebuilt.write_text(annual_cache.read_text(encoding="utf-8"), encoding="utf-8")
        annual_cache.unlink()

        with mock.patch.object(berquist_sherman_service, "_refresh_downstream_domains"), \
                mock.patch.object(
                    berquist_sherman_service.precedent_cache_service,
                    "materialize_engine_source",
                    return_value=str(rebuilt),
                ) as materialize:
            report = berquist_sherman_service.refresh_dependents(
                "Project", "Class", ["Ultimate Counts"], rebuild_index=False, finalize_method_review_status=False
            )

        self.assertTrue(report["ok"], report)
        materialize.assert_called_once()
        self.assertEqual(materialize.call_args.args[2], "Ultimate Counts")
        self.assertEqual(materialize.call_args.args[4], 12)
        csv_text = (self.datasets / f"{OUTPUT}@12@12@cum@dev.csv").read_text(encoding="utf-8")
        self.assertEqual(csv_text, berquist_sherman_output_csv_text(self.expected_output([20, 20, 40]), 3))

    def test_a_non_bs_dependent_is_left_to_the_central_cascade(self) -> None:
        self.write_workspace(saved_ultimate=[20, 20, 20], current_ultimate=[20, 20, 40])
        self.write_json(self.sidecars / "DFM Over Paid.json", {
            "dataset_name": "DFM Over Paid",
            "dataset_type": "DFM Over Paid",
            "source_kind": "dfm",
            "method_type": "DFM",
            "data_format": "Vector",
            "status": 0,
            "precedents": [{"dataset_name": "Paid"}],
            "dependents": [],
        })
        paid_path = self.sidecars / "Paid.json"
        paid = self.read_json(paid_path)
        paid["dependents"] = [{"dataset_name": "DFM Over Paid"}]
        self.write_json(paid_path, paid)

        with mock.patch.object(berquist_sherman_service, "_refresh_downstream_domains"):
            report = berquist_sherman_service.refresh_dependents(
                "Project", "Class", ["Paid"], rebuild_index=False, finalize_method_review_status=False
            )

        self.assertTrue(report["ok"], report)
        self.assertEqual(report["updated"], [])
        self.assertEqual(
            report["skipped"],
            [{"dataset_name": "DFM Over Paid", "reason": "non_bs_dependent_handled_by_central_cascade"}],
        )

    # -- the walk --------------------------------------------------------

    def test_the_walk_runs_the_bs_wave_after_result_selection_and_feeds_the_later_waves(self) -> None:
        order: list[str] = []

        def record(name: str, payload: dict):
            def _refresh(*_args, **_kwargs):
                order.append(name)
                return payload

            return _refresh

        with (
            mock.patch(
                "app_server.services.dfm_service.refresh_dependents",
                side_effect=record("dfm", {"ok": True, "updated": [], "errors": []}),
            ),
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=[]),
            mock.patch.object(calculated_dataset_service, "_existing_downstream_keys", return_value=[]),
            mock.patch(
                "app_server.services.result_selection_service.refresh_dependents",
                side_effect=record("result_selection", {
                    "ok": True,
                    "updated": [{"dataset_name": "C 92 - Current Qtr Selected"}],
                    "status_refreshed": [],
                    "errors": [],
                    "downstream_fresh_names": [],
                    "downstream_blocked_names": [],
                }),
            ),
            mock.patch.object(
                berquist_sherman_service,
                "refresh_dependents",
                side_effect=record("berquist_sherman", {
                    "ok": False,
                    "updated": [{"dataset_name": OUTPUT, "dataset_type": "Gross Loss - ad hoc", "output_changed": True}],
                    "status_refreshed": [],
                    "errors": [{"dataset_name": "BS CRA Incurred", "reason": "Reported Claim Counts must be an annual triangle dataset: Reported"}],
                    "downstream_fresh_names": ["D 18 - BS Paid DFM"],
                    "downstream_blocked_names": ["D 19 - Blocked DFM"],
                }),
            ) as refresh_bs,
            mock.patch(
                "app_server.services.bornhuetter_ferguson_service.refresh_dependents",
                side_effect=record("bornhuetter_ferguson", {"ok": True, "updated": [], "errors": []}),
            ) as refresh_bf,
            mock.patch(
                "app_server.services.cape_cod_service.refresh_dependents",
                side_effect=record("cape_cod", {"ok": True, "updated": [], "errors": []}),
            ) as refresh_cc,
            mock.patch(
                "app_server.services.bootstrap_service.refresh_dependents",
                side_effect=record("bootstrap", {"ok": True, "updated": [], "errors": []}),
            ) as refresh_bst,
            mock.patch.object(calculated_dataset_service.dataset_instance_index_service, "rebuild_index"),
        ):
            result = calculated_dataset_service.recalculate_dependents(
                "Project", "Class", "Claim Counts--CWP", "Claim Counts--CWP"
            )

        self.assertEqual(
            order,
            ["dfm", "result_selection", "berquist_sherman", "bornhuetter_ferguson", "cape_cod", "bootstrap"],
        )
        # The B&S wave failed for one method, so the walk is not clean.
        self.assertFalse(result["ok"])
        self.assertEqual(result["berquist_sherman_updates"]["updated"][0]["dataset_name"], OUTPUT)
        self.assertIn("C 92 - Current Qtr Selected", refresh_bs.call_args.args[2])
        for refresher in (refresh_bf, refresh_cc, refresh_bst):
            roots = refresher.call_args.args[2]
            blocked = refresher.call_args.kwargs["blocked_precedent_names"]
            self.assertIn(OUTPUT, roots)
            self.assertIn("D 18 - BS Paid DFM", roots)
            self.assertIn("BS CRA Incurred", blocked)
            self.assertIn("D 19 - Blocked DFM", blocked)

    def test_the_walk_can_leave_the_bs_wave_out(self) -> None:
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
            mock.patch.object(berquist_sherman_service, "refresh_dependents") as refresh_bs,
            mock.patch.object(calculated_dataset_service.dataset_instance_index_service, "rebuild_index"),
        ):
            result = calculated_dataset_service.recalculate_dependents(
                "Project",
                "Class",
                "Paid",
                "Paid",
                include_berquist_sherman=False,
                include_bornhuetter_ferguson=False,
                include_cape_cod=False,
                include_bootstrap=False,
            )

        self.assertTrue(result["ok"], result)
        self.assertIsNone(result["berquist_sherman_updates"])
        refresh_bs.assert_not_called()


if __name__ == "__main__":
    unittest.main()

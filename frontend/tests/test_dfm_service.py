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

from arcrho_api.dfm_contract import (
    dfm_output_variants,
    method_revisions,
    normalize_dfm_method,
    recalculate_dfm_method,
)
from app_server.services import calculated_dataset_service, dataset_sidecar_status_service, dfm_service


class DfmServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=TEST_TEMP_ROOT)
        root = Path(self.temp.name)
        self.methods = root / "methods"
        self.datasets = root / "datasets"
        self.sidecars = root / "sidecars"
        for folder in (self.methods, self.datasets, self.sidecars):
            folder.mkdir()
        # Two years of monthly origins valued at the end of the second year,
        # which is the shape the monthly sources below are written at.
        settings = root / "general_settings.json"
        settings.write_text(
            '{"origin_start_date":"202301","origin_end_date":"202412","development_end_date":"202412"}',
            encoding="utf-8",
        )
        self.patchers = [
            mock.patch.object(dfm_service.config, "get_project_method_data_dir", return_value=str(self.methods)),
            mock.patch.object(dfm_service.config, "get_project_dataset_cache_dir", return_value=str(self.datasets)),
            mock.patch.object(dfm_service.config, "get_general_settings_path", return_value=str(settings)),
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
    def method_payload() -> dict:
        return recalculate_dfm_method(
            {
                "details_tab": {
                    "name": "Development",
                    "output_type": "Selected Ultimate",
                    "output_dataset": "Development Output",
                    "input_triangle": "Paid",
                    "origin_length": 12,
                    "development_length": 12,
                },
                "ratios_tab": {
                    "average_formulas": {
                        "label": ["User Entry"],
                        "custom_average_formula_settings": {"average_type": ["user_entry"]},
                        "selected": [[1, 1]],
                        "values": [[1.5, 1]],
                        "inputs": [["1.5", "1"]],
                    },
                    "cell_notes": {"ratio_main_table": {"2024": {"(1) 12-24": "keep"}}},
                },
                "results_tab": {"ratio_basis_dataset": "Premium"},
            },
            input_snapshot={
                "name": "Paid",
                "data_format": "Triangle",
                "origin_labels": ["2024", "2025"],
                "development_labels": ["12", "24"],
                "values": [[100, 150], [200, None]],
                "mask": [[True, True], [True, False]],
                "number_format": "#,##0",
                "decimal_places": 0,
                "revision": "paid-r1",
            },
            ratio_basis_snapshot={
                "name": "Premium",
                "data_format": "Vector",
                "origin_labels": ["2024", "2025"],
                "values": [1000, 1100],
                "number_format": "#,##0",
                "decimal_places": 0,
                "revision": "premium-r1",
            },
            timestamp="2026-01-01T00:00:00Z",
        )

    def write_json(self, path: Path, payload: dict) -> None:
        path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")

    def output_sidecar(self, method: dict, *, status: int = 0) -> dict:
        revisions = method_revisions(method)
        return {
            "dataset_name": "Development Output",
            "dataset_type": "Selected Ultimate",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "dfm",
            "method_type": "DFM",
            "method_name": "Development",
            "data_format": "Vector",
            "period_length": 12,
            "csv_file": "Development Output@12.csv",
            "origin_labels": ["2024", "2025"],
            "precedents": [{"dataset_name": "Paid"}, {"dataset_name": "Premium"}],
            "dependents": [],
            "status": status,
            "notes": "method note",
            "audit_log": [],
            "publication_revision": revisions["publication_revision"],
            "updated_at": "2026-01-01T00:00:00Z",
        }

    def write_method_pair(self, method: dict | None = None, *, status: int = 0) -> dict:
        payload = method or self.method_payload()
        self.write_json(self.methods / "DFM@Development.json", payload)
        self.write_json(self.sidecars / "Development Output.json", self.output_sidecar(payload, status=status))
        (self.datasets / "Development Output@12.csv").write_text("150\n300\n", encoding="utf-8")
        return payload

    def write_source(
        self,
        name: str,
        csv_text: str,
        *,
        data_format: str,
        dependents: list[str] | None = None,
        method_type: str = "None",
        status: int = 0,
    ) -> None:
        csv_file = f"{name}@12.csv"
        (self.datasets / csv_file).write_text(csv_text, encoding="utf-8")
        self.write_json(self.sidecars / f"{name}.json", {
            "dataset_name": name,
            "dataset_type": name,
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "dfm" if method_type == "DFM" else "input",
            "method_type": method_type,
            "data_format": data_format,
            "origin_length": 12,
            "development_length": 12,
            "period_length": 12 if data_format == "Vector" else None,
            "stored_origin_length": 12,
            "stored_development_length": 12,
            "stored_period_length": 12 if data_format == "Vector" else None,
            "origin_labels": ["2024", "2025"],
            "csv_file": csv_file,
            "number_format": "#,##0",
            "decimal_places": 0,
            "status": status,
            "precedents": [],
            "dependents": [
                {"dataset_name": item} for item in (dependents or [])
            ],
        })

    def write_monthly_source(self, name: str, *, dependents: list[str] | None = None) -> None:
        """Write a 24-month cumulative triangle: every cell is 100 x its age in months.

        Rolled up along the calendar this is the synthetic the plan measured,
        so the two annual years read 7,800 / 22,200 and 7,800 / blank.
        """

        csv_file = f"{name}@1@1@cum@dev.csv"
        rows = [
            [100 * (column + 1) for column in range(24 - row)]
            for row in range(24)
        ]
        (self.datasets / csv_file).write_text(
            "".join(",".join(str(value) for value in row) + "\n" for row in rows),
            encoding="utf-8",
        )
        self.write_json(self.sidecars / f"{name}.json", {
            "dataset_name": name,
            "dataset_type": name,
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "input",
            "method_type": "None",
            "data_format": "Triangle",
            "origin_length": 12,
            "development_length": 12,
            "stored_origin_length": 1,
            "stored_development_length": 1,
            "csv_file": csv_file,
            "number_format": "#,##0",
            "decimal_places": 0,
            "status": 0,
            "precedents": [],
            "dependents": [{"dataset_name": item} for item in (dependents or [])],
        })

    def test_dataset_references_resolve_labels_indices_and_reuse_dataset_reads(self) -> None:
        datasets = {
            "paid": {
                "dataset_name": "Paid",
                "data_format": "Triangle",
                "origin_labels": ["2023", "2024"],
                "dev_labels": ["12", "24"],
                "values": [[100, 150], [200, 275]],
            },
            "premium": {
                "dataset_name": "Premium",
                "data_format": "Vector",
                "origin_labels": ["2023", "2024"],
                "dev_labels": ["Ultimate"],
                "values": [[1000], [1100]],
            },
        }
        reads: list[str] = []

        def load(_project: str, _reserving_class: str, dataset_name: str) -> dict:
            reads.append(dataset_name)
            return datasets[dataset_name.casefold()]

        with mock.patch(
            "app_server.services.dataset_service.load_cached_dataset_values",
            side_effect=load,
        ):
            result = dfm_service.resolve_dfm_dataset_references(
                "Project",
                "Class",
                [
                    {"dataset_name": "Paid", "row_idx": "1", "col_idx": "2"},
                    {"dataset_name": "Paid", "row_idx": "2024", "col_idx": "24"},
                    {"dataset_name": "Premium", "row_idx": "2"},
                    {"dataset_name": "Premium", "row_idx": "2023", "col_idx": "1"},
                ],
            )

        self.assertEqual([item["value"] for item in result["results"]], [150, 275, 1100, 1000])
        self.assertEqual(
            [(item["row_label"], item["col_label"]) for item in result["results"]],
            [("2023", "24"), ("2024", "24"), ("2024", "Ultimate"), ("2023", "Ultimate")],
        )
        self.assertCountEqual(reads, ["Paid", "Premium"])

    def test_dataset_reference_requires_triangle_column_and_vector_column_one(self) -> None:
        triangle = {
            "dataset_name": "Paid",
            "data_format": "Triangle",
            "origin_labels": ["2024"],
            "dev_labels": ["12"],
            "values": [[100]],
        }
        vector = {
            "dataset_name": "Premium",
            "data_format": "Vector",
            "origin_labels": ["2024"],
            "dev_labels": ["Ultimate"],
            "values": [[1000]],
        }

        with mock.patch(
            "app_server.services.dataset_service.load_cached_dataset_values",
            side_effect=lambda _p, _r, name: triangle if name == "Paid" else vector,
        ):
            with self.assertRaisesRegex(HTTPException, "Column index is required"):
                dfm_service.resolve_dfm_dataset_references(
                    "Project", "Class", [{"dataset_name": "Paid", "row_idx": "1"}],
                )
            with self.assertRaisesRegex(HTTPException, "Column index 2 is outside"):
                dfm_service.resolve_dfm_dataset_references(
                    "Project",
                    "Class",
                    [{"dataset_name": "Premium", "row_idx": "1", "col_idx": "2"}],
                )

    def test_dataset_reference_negative_vector_indices_trim_only_trailing_blanks(self) -> None:
        vector = {
            "dataset_name": "Quarterly Premium",
            "data_format": "Vector",
            "origin_labels": [str(index) for index in range(1, 10)],
            "dev_labels": ["Ultimate"],
            "values": [[22], [33], [55], [66], [None], [45], [None], [76], [None]],
        }

        self.assertEqual(
            dfm_service._resolved_dataset_reference(
                {"dataset_name": "Quarterly Premium", "row_idx": "-1"},
                vector,
            )["value"],
            76,
        )
        self.assertEqual(
            dfm_service._resolved_dataset_reference(
                {"dataset_name": "Quarterly Premium", "row_idx": "-3"},
                vector,
            )["value"],
            45,
        )
        with self.assertRaisesRegex(HTTPException, r"\[Quarterly Premium\]\[7, Ultimate\] is blank"):
            dfm_service._resolved_dataset_reference(
                {"dataset_name": "Quarterly Premium", "row_idx": "-2"},
                vector,
            )

    def test_dataset_reference_negative_vector_indices_stop_at_the_valuation_period(self) -> None:
        # Every row holds a value, including the four after the Development End
        # Date; [-1] must still be the valuation period, not the last row.
        full_year = {
            "dataset_name": "Monthly Premium",
            "data_format": "Vector",
            "origin_labels": [f"2026-{month:02d}" for month in range(1, 13)],
            "dev_labels": ["Ultimate"],
            "values": [[100 + month] for month in range(1, 13)],
            "valuation_row_count": 8,
        }
        latest = dfm_service._resolved_dataset_reference(
            {"dataset_name": "Monthly Premium", "row_idx": "-1"}, full_year,
        )
        self.assertEqual((latest["row_label"], latest["value"]), ("2026-08", 108))
        self.assertEqual(
            dfm_service._resolved_dataset_reference(
                {"dataset_name": "Monthly Premium", "row_idx": "-3"}, full_year,
            )["row_label"],
            "2026-06",
        )
        with self.assertRaisesRegex(HTTPException, "outside the valid range"):
            dfm_service._resolved_dataset_reference(
                {"dataset_name": "Monthly Premium", "row_idx": "-9"}, full_year,
            )
        # Data that stops before the valuation period keeps the earlier boundary.
        short = dict(full_year, values=[[100 + month] if month <= 5 else [None] for month in range(1, 13)])
        self.assertEqual(
            dfm_service._resolved_dataset_reference(
                {"dataset_name": "Monthly Premium", "row_idx": "-1"}, short,
            )["row_label"],
            "2026-05",
        )

    def test_resolve_dataset_references_stamps_vectors_with_the_valuation_row_count(self) -> None:
        from app_server.services import dataset_service

        vector = {
            "dataset_name": "Monthly Premium",
            "data_format": "Vector",
            "origin_length": 1,
            "origin_labels": [str(month) for month in range(1, 13)],
            "dev_labels": ["Ultimate"],
            "values": [[100 + month] for month in range(1, 13)],
        }
        with (
            mock.patch.object(dataset_service, "load_cached_dataset_values", return_value=vector),
            mock.patch.object(dataset_service, "valuation_origin_row_count", return_value=8) as row_count,
        ):
            result = dfm_service.resolve_dfm_dataset_references(
                "Project",
                "Class",
                [
                    {"dataset_name": "Monthly Premium", "row_idx": "-1"},
                    {"dataset_name": "Monthly Premium", "row_idx": "-2"},
                ],
            )
        row_count.assert_called_once_with("Project", 1)
        self.assertEqual([item["value"] for item in result["results"]], [108, 107])

    def test_dataset_reference_negative_triangle_indices_follow_latest_valid_diagonal(self) -> None:
        triangle = {
            "dataset_name": "Quarterly Paid",
            "data_format": "Triangle",
            "origin_labels": ["Q1", "Q2", "Q3", "Q4", "Q5"],
            "dev_labels": ["3", "6", "9", "12", "15"],
            # The final two calendar diagonals are outside the valuation range.
            # A blank inside the valid geometry remains an addressable position.
            "values": [
                [10, None, 12, None, None],
                [20, 21, None, None, None],
                [30, None, None, None, None],
                [None, None, None, None, None],
                [None, None, None, None, None],
            ],
        }

        latest = dfm_service._resolved_dataset_reference(
            {"dataset_name": "Quarterly Paid", "row_idx": "-1", "col_idx": "-1"},
            triangle,
        )
        self.assertEqual((latest["row_label"], latest["col_label"], latest["value"]), ("Q3", "3", 30))
        prior = dfm_service._resolved_dataset_reference(
            {"dataset_name": "Quarterly Paid", "row_idx": "-2", "col_idx": "-1"},
            triangle,
        )
        self.assertEqual((prior["row_label"], prior["col_label"], prior["value"]), ("Q2", "6", 21))
        with self.assertRaisesRegex(HTTPException, "outside the valid range"):
            dfm_service._resolved_dataset_reference(
                {"dataset_name": "Quarterly Paid", "row_idx": "-4", "col_idx": "-1"},
                triangle,
            )

    def test_v2_load_reads_only_method_and_own_sidecar(self) -> None:
        self.write_method_pair()
        original = dfm_service._read_json
        reads: list[str] = []

        def recording(path: str) -> dict:
            reads.append(str(Path(path)))
            return original(path)

        with mock.patch.object(dfm_service, "_read_json", side_effect=recording):
            result = dfm_service.load_dfm_method(
                "Project",
                "Class",
                "Development",
                output_dataset="Development Output",
            )

        self.assertTrue(result["ok"])
        self.assertCountEqual(reads, [
            str(self.methods / "DFM@Development.json"),
            str(self.sidecars / "Development Output.json"),
        ])

    def test_load_without_output_sidecar_is_rejected_without_mutation(self) -> None:
        # With no declared output the method's own name stands in for it, and
        # no sidecar of that name exists here.
        method = self.method_payload()
        method["details_tab"].pop("output_dataset", None)
        method_path = self.methods / "DFM@Development.json"
        self.write_json(method_path, method)
        before = method_path.read_bytes()

        with self.assertRaises(HTTPException) as raised:
            dfm_service.load_dfm_method("Project", "Class", "Development")

        self.assertEqual(raised.exception.status_code, 409)
        self.assertEqual(method_path.read_bytes(), before)

    def test_existing_save_rebases_owned_patch_without_precedent_reads(self) -> None:
        method = self.write_method_pair(status=2)
        method["ratios_tab"]["cell_notes"]["ratio_main_table"]["2024"]["(1) 12-24"] = "updated"
        owned_revision = method_revisions(self.method_payload())["owned_revision"]
        with (
            mock.patch.object(dfm_service, "_load_source_snapshot", side_effect=AssertionError("source read")),
            mock.patch.object(
                dfm_service.dependent_propagation_service,
                "enqueue_marked_save_propagation",
                return_value={"ok": True, "job_id": "job-1", "status": "queued"},
            ) as enqueue,
            mock.patch.object(
                dfm_service.dependent_propagation_service,
                "require_reserving_class_writable",
            ),
        ):
            result = dfm_service.save_dfm_method(
                "Project",
                "Class",
                method,
                expected_owned_revision=owned_revision,
            )

        self.assertTrue(result["ok"])
        # An owned-only rebase leaves the publication revision unchanged, so a
        # no-op save submits no Engine propagation job.
        enqueue.assert_not_called()
        self.assertEqual(result["propagation"], {"ok": True, "status": "unchanged"})
        self.assertTrue(result["propagation_ok"])
        saved = json.loads((self.methods / "DFM@Development.json").read_text(encoding="utf-8"))
        self.assertEqual(
            saved["ratios_tab"]["cell_notes"]["ratio_main_table"]["2024"]["(1) 12-24"],
            "updated",
        )

    def test_explicit_save_warns_and_still_saves_with_unreviewed_precedent(self) -> None:
        method = self.write_method_pair(status=2)
        self.write_source(
            "Paid",
            "100,150\n200,\n",
            data_format="Triangle",
            method_type="DFM",
            status=2,
        )
        self.write_source("Premium", "1000\n1100\n", data_format="Vector")

        with mock.patch(
            "app_server.services.calculated_dataset_service.recalculate_dependents",
            return_value={"ok": True, "updated": []},
        ):
            result = dfm_service.save_dfm_method(
                "Project",
                "Class",
                method,
                expected_owned_revision=method_revisions(method)["owned_revision"],
            )

        self.assertTrue(result["ok"])
        self.assertEqual(result["sidecar"]["status"], 0)
        self.assertEqual(result["unreviewed_precedents"], ["Paid"])
        self.assertEqual(result["unreviewed_precedent_count"], 1)

    def test_dataset_formula_save_registers_graph_and_source_change_marks_review(self) -> None:
        method = self.write_method_pair()
        expected_owned_revision = method_revisions(method)["owned_revision"]
        for name, data_format, csv_text in (
            ("Paid", "Triangle", "100,150\n200,\n"),
            ("Premium", "Vector", "1000\n1100\n"),
            ("Accounting Cutoff", "Vector", "1.01\n1.02\n"),
        ):
            self.write_source(name, csv_text, data_format=data_format)
        method["ratios_tab"]["average_formulas"]["inputs"][0][0] = \
            '=[Accounting Cutoff][-1]'

        with mock.patch.object(
            dfm_service.dependent_propagation_service,
            "require_reserving_class_writable",
        ):
            result = dfm_service.save_dfm_method(
                "Project",
                "Class",
                method,
                expected_owned_revision=expected_owned_revision,
            )

        self.assertEqual(
            result["sidecar"]["precedents"],
            [
                {"dataset_name": "Paid"},
                {"dataset_name": "Premium"},
                {"dataset_name": "Accounting Cutoff"},
            ],
        )
        formula_source = json.loads(
            (self.sidecars / "Accounting Cutoff.json").read_text(encoding="utf-8")
        )
        self.assertEqual(
            formula_source["dependents"],
            [{"dataset_name": "Development Output"}],
        )

        with mock.patch(
            "app_server.services.dataset_service.load_cached_dataset_values",
            return_value={
                "dataset_name": "Accounting Cutoff",
                "data_format": "Vector",
                "origin_labels": ["2024", "2025"],
                "dev_labels": ["Ultimate"],
                "values": [[1.01], [1.02]],
            },
        ):
            refresh = dfm_service.refresh_dependents(
                "Project",
                "Class",
                ["Accounting Cutoff"],
            )

        self.assertTrue(refresh["ok"], refresh)
        saved_sidecar = json.loads(
            (self.sidecars / "Development Output.json").read_text(encoding="utf-8")
        )
        self.assertEqual(
            saved_sidecar["status"],
            dataset_sidecar_status_service.STATUS_REVIEW_NEEDED,
        )
        # The propagation walk re-resolves the dataset reference and persists
        # the refreshed evaluation and outputs; the opened DFM then shows the
        # updated values without becoming dirty.
        saved_method = json.loads(
            (self.methods / "DFM@Development.json").read_text(encoding="utf-8")
        )
        self.assertEqual(
            saved_method["ratios_tab"]["average_formulas"]["values"][0][0],
            1.02,
        )
        self.assertEqual(
            saved_method["results_tab"]["ultimate_vector"],
            [150, 204],
        )
        published = (self.datasets / "Development Output@12.csv").read_text(encoding="utf-8")
        self.assertEqual(
            [float(row) for row in published.strip().splitlines()],
            [150.0, 204.0],
        )

    def test_blank_dataset_reference_aborts_refresh_and_preserves_publication(self) -> None:
        method = self.method_payload()
        method["ratios_tab"]["average_formulas"]["inputs"][0][0] = \
            '=[Accounting Cutoff][2]'
        method = recalculate_dfm_method(method, timestamp="2026-01-01T00:00:00Z")
        method = self.write_method_pair(method)
        self.write_source("Paid", "100,150\n200,\n", data_format="Triangle")
        self.write_source(
            "Accounting Cutoff",
            "1.01\n\n",
            data_format="Vector",
            dependents=["Development Output"],
        )
        method_path = self.methods / "DFM@Development.json"
        output_path = self.datasets / "Development Output@12.csv"
        before_method = method_path.read_bytes()
        before_output = output_path.read_bytes()

        with mock.patch(
            "app_server.services.dataset_service.load_cached_dataset_values",
            return_value={
                "dataset_name": "Accounting Cutoff",
                "data_format": "Vector",
                "origin_labels": ["2024", "2025"],
                "dev_labels": ["Ultimate"],
                "values": [[1.01], [None]],
            },
        ):
            result = dfm_service.refresh_dependents("Project", "Class", ["Accounting Cutoff"])

        self.assertFalse(result["ok"], result)
        self.assertEqual(len(result["errors"]), 1)
        self.assertIn("blank or non-numeric", result["errors"][0]["reason"])
        self.assertEqual(method_path.read_bytes(), before_method)
        self.assertEqual(output_path.read_bytes(), before_output)
        saved_sidecar = json.loads(
            (self.sidecars / "Development Output.json").read_text(encoding="utf-8")
        )
        self.assertEqual(
            saved_sidecar["status"],
            dataset_sidecar_status_service.STATUS_REVIEW_NEEDED,
        )

    def test_existing_save_rejects_owned_revision_conflict_without_mutation(self) -> None:
        method = self.write_method_pair()
        method_path = self.methods / "DFM@Development.json"
        sidecar_path = self.sidecars / "Development Output.json"
        output_path = self.datasets / "Development Output@12.csv"
        before = {
            path: path.read_bytes()
            for path in (method_path, sidecar_path, output_path)
        }
        method["ratios_tab"]["cell_notes"]["ratio_main_table"]["2024"]["(1) 12-24"] = "conflict"

        with self.assertRaises(HTTPException) as raised:
            dfm_service.save_dfm_method(
                "Project",
                "Class",
                method,
                expected_owned_revision="stale-owned-revision",
            )

        self.assertEqual(raised.exception.status_code, 409)
        self.assertIn("owned settings changed", str(raised.exception.detail))
        for path, contents in before.items():
            self.assertEqual(path.read_bytes(), contents)

    def test_triangle_ratio_basis_uses_latest_available_value_per_origin(self) -> None:
        self.write_source("Premium", "100,150\n200,\n", data_format="Triangle")

        snapshot = dfm_service._load_source_snapshot(
            "Project", "Class", "Premium", vector=True
        )

        self.assertEqual(snapshot["data_format"], "Triangle")
        self.assertEqual(snapshot["origin_labels"], ["2024", "2025"])
        self.assertEqual(snapshot["values"], [150, 200])

    def test_input_refresh_uses_saved_method_axes_when_sidecar_labels_are_absent(self) -> None:
        method = self.write_method_pair()
        self.write_source(
            "Paid",
            "100,175\n200,\n",
            data_format="Triangle",
            dependents=["Development Output"],
        )
        source_path = self.sidecars / "Paid.json"
        source = json.loads(source_path.read_text(encoding="utf-8"))
        source.pop("origin_labels")
        self.write_json(source_path, source)

        result = dfm_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"], result)
        saved = json.loads((self.methods / "DFM@Development.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["data_tab"]["origin_labels"], method["data_tab"]["origin_labels"])
        self.assertEqual(
            saved["data_tab"]["development_labels"],
            method["data_tab"]["development_labels"],
        )
        # Persisted rows drop the trailing nulls that sit outside the triangle.
        self.assertEqual(saved["data_tab"]["input_data_triangle_values"], [[100, 175], [200]])
        self.assertNotIn("input_data_triangle_mask", saved["data_tab"])
        self.assertEqual(
            normalize_dfm_method(saved)["data_tab"]["input_data_triangle_values"],
            [[100, 175], [200, None]],
        )

    def test_basis_refresh_ignores_numeric_sidecar_labels_and_review_status(self) -> None:
        method = self.write_method_pair()
        self.write_source(
            "Premium",
            "2000\n2200\n",
            data_format="Vector",
            dependents=["Development Output"],
            method_type="DFM",
            status=2,
        )
        source_path = self.sidecars / "Premium.json"
        source = json.loads(source_path.read_text(encoding="utf-8"))
        source["origin_labels"] = ["1", "2"]
        self.write_json(source_path, source)

        result = dfm_service.refresh_dependents("Project", "Class", ["Premium"])

        self.assertTrue(result["ok"], result)
        saved = json.loads((self.methods / "DFM@Development.json").read_text(encoding="utf-8"))
        # The basis axis is not persisted; reading the file re-derives it from
        # the method's own origin labels rather than the sidecar's numeric ones.
        self.assertNotIn("ratio_basis_origin_labels", saved["results_tab"])
        self.assertEqual(
            normalize_dfm_method(saved)["results_tab"]["ratio_basis_origin_labels"],
            method["data_tab"]["origin_labels"],
        )
        self.assertEqual(saved["results_tab"]["ratio_basis_values"], [2000, 2200])

    def test_snapshot_cache_is_scoped_by_saved_method_axes(self) -> None:
        method = self.method_payload()
        self.write_source("Paid", "100,150\n200,\n", data_format="Triangle")
        source_path = self.sidecars / "Paid.json"
        source = json.loads(source_path.read_text(encoding="utf-8"))
        source.pop("origin_labels")
        self.write_json(source_path, source)
        snapshot_cache = {}

        first, _ = dfm_service._source_snapshots(
            "Project",
            "Class",
            method,
            load_input=True,
            load_basis=False,
            snapshot_cache=snapshot_cache,
        )
        second_method = json.loads(json.dumps(method))
        second_method["data_tab"]["origin_labels"] = ["2022", "2023"]
        second, _ = dfm_service._source_snapshots(
            "Project",
            "Class",
            second_method,
            load_input=True,
            load_basis=False,
            snapshot_cache=snapshot_cache,
        )

        self.assertEqual(first["origin_labels"], ["2024", "2025"])
        self.assertEqual(second["origin_labels"], ["2022", "2023"])
        self.assertEqual(len(snapshot_cache), 2)

    def test_missing_sidecar_labels_do_not_hide_row_or_period_mismatches(self) -> None:
        self.write_method_pair()
        method_path = self.methods / "DFM@Development.json"
        output_path = self.datasets / "Development Output@12.csv"
        before_method = method_path.read_bytes()
        before_output = output_path.read_bytes()
        self.write_source(
            "Paid",
            "100,150\n200,\n300,\n",
            data_format="Triangle",
            dependents=["Development Output"],
        )
        source_path = self.sidecars / "Paid.json"
        source = json.loads(source_path.read_text(encoding="utf-8"))
        source.pop("origin_labels")
        self.write_json(source_path, source)

        row_result = dfm_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertFalse(row_result["ok"])
        self.assertIn("has 3 rows; expected 2", row_result["errors"][0]["reason"])
        self.assertEqual(method_path.read_bytes(), before_method)
        self.assertEqual(output_path.read_bytes(), before_output)

        (self.datasets / "Paid@12.csv").write_text("100,150\n200,\n", encoding="utf-8")
        source["stored_origin_length"] = 3
        self.write_json(source_path, source)

        period_result = dfm_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertFalse(period_result["ok"])
        self.assertIn("incompatible origin period length", period_result["errors"][0]["reason"])
        self.assertEqual(method_path.read_bytes(), before_method)
        self.assertEqual(output_path.read_bytes(), before_output)

    def test_precedent_period_is_read_from_the_stored_shape_not_the_displayed_one(self) -> None:
        """A quarterly triangle shown yearly is still quarterly data to a method."""

        method = self.method_payload()
        self.write_source("Paid", "100,150\n200,\n", data_format="Triangle")
        source_path = self.sidecars / "Paid.json"
        source = json.loads(source_path.read_text(encoding="utf-8"))
        source["stored_origin_length"] = 3
        source["stored_development_length"] = 3
        self.write_json(source_path, source)

        # Read as the displayed twelve months this two-row CSV would be taken
        # as it stands; read as the stored three it is two quarters, which
        # make one partly filled year, not the two the method needs.
        with self.assertRaisesRegex(HTTPException, r"has 1 rows; expected 2"):
            dfm_service._source_snapshots(
                "Project",
                "Class",
                method,
                load_input=True,
                load_basis=False,
            )

        source["stored_origin_length"] = 12
        source["stored_development_length"] = 12
        source["origin_length"] = 36
        source["development_length"] = 36
        self.write_json(source_path, source)

        snapshot, _ = dfm_service._source_snapshots(
            "Project",
            "Class",
            method,
            load_input=True,
            load_basis=False,
        )

        self.assertEqual(snapshot["values"], [[100, 150], [200, None]])

    def test_monthly_manual_input_is_rolled_up_for_an_annual_method(self) -> None:
        method = self.method_payload()
        self.write_monthly_source("Paid")

        snapshot, _ = dfm_service._source_snapshots(
            "Project",
            "Class",
            method,
            load_input=True,
            load_basis=False,
        )

        self.assertEqual(snapshot["values"], [[7800, 22200], [7800, None]])
        self.assertEqual(snapshot["mask"], [[True, True], [True, False]])
        self.assertEqual(snapshot["development_labels"], ["12", "24"])
        # The roll-up happens in memory, so no coarser copy is left behind for
        # a later load to trust.
        self.assertEqual(sorted(item.name for item in self.datasets.glob("Paid*")), ["Paid@1@1@cum@dev.csv"])

    def test_manual_input_roll_up_ignores_a_stale_coarser_copy_on_disk(self) -> None:
        self.write_method_pair()
        self.write_monthly_source("Paid", dependents=["Development Output"])
        self.write_source("Premium", "1000\n1100\n", data_format="Vector")
        (self.datasets / "Paid@12@12@cum@dev.csv").write_text("1,2\n3,\n", encoding="utf-8")

        result = dfm_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"], result)
        saved = json.loads((self.methods / "DFM@Development.json").read_text(encoding="utf-8"))
        self.assertEqual(
            normalize_dfm_method(saved)["data_tab"]["input_data_triangle_values"],
            [[7800, 22200], [7800, None]],
        )

    def test_monthly_manual_input_is_rolled_up_when_a_save_picks_it_up(self) -> None:
        previous = self.method_payload()
        previous["details_tab"]["input_triangle"] = "Old Paid"
        previous = recalculate_dfm_method(previous, timestamp="2026-01-01T00:00:00Z")
        self.write_method_pair(previous)
        method = self.method_payload()
        self.write_monthly_source("Paid")
        self.write_source("Premium", "1000\n1100\n", data_format="Vector")

        with mock.patch.object(
            dfm_service.dependent_propagation_service,
            "require_reserving_class_writable",
        ):
            result = dfm_service.save_dfm_method(
                "Project",
                "Class",
                method,
                expected_owned_revision=method_revisions(previous)["owned_revision"],
            )

        self.assertTrue(result["ok"], result)
        saved = json.loads((self.methods / "DFM@Development.json").read_text(encoding="utf-8"))
        self.assertEqual(
            normalize_dfm_method(saved)["data_tab"]["input_data_triangle_values"],
            [[7800, 22200], [7800, None]],
        )

    def test_engine_generated_precedent_is_regenerated_at_the_method_period(self) -> None:
        method = self.method_payload()
        method["details_tab"]["origin_length"] = 3
        method["details_tab"]["development_length"] = 3
        self.write_source("Paid", "100,150\n200,\n", data_format="Triangle")
        source_path = self.sidecars / "Paid.json"
        source = json.loads(source_path.read_text(encoding="utf-8"))
        source["source_kind"] = "engine"
        source["csv_file"] = "Paid@12@12@cum@dev.csv"
        self.write_json(source_path, source)
        requests: list[tuple[dict, str, dict]] = []

        def run_arcrho_tri(pairs, path, **options):
            requests.append((dict(pairs), path, options))
            Path(path).write_text("10,15\n20,\n", encoding="utf-8")
            return {"ok": True, "status": "cache_missing"}

        with mock.patch(
            "app_server.services.arcrho_runtime_service.run_arcrho_tri",
            side_effect=run_arcrho_tri,
        ):
            snapshot, _ = dfm_service._source_snapshots(
                "Project",
                "Class",
                method,
                load_input=True,
                load_basis=False,
            )

        self.assertEqual(snapshot["values"], [[10, 15], [20, None]])
        pairs, path, options = requests[0]
        self.assertEqual(
            (pairs["Function"], pairs["InstanceName"], pairs["OriginLength"], pairs["DevelopmentLength"]),
            ("ArcRhoTri", "Paid", "3", "3"),
        )
        self.assertEqual(Path(path).name, "Paid@3@3@cum@dev.csv")
        self.assertFalse(options["write_sidecar"])
        self.assertEqual(
            json.loads(source_path.read_text(encoding="utf-8"))["csv_file"],
            "Paid@12@12@cum@dev.csv",
        )

    def test_output_csv_variants_are_projected_by_the_canonical_contract(self) -> None:
        method = self.write_method_pair()
        method["details_tab"]["origin_length"] = 3
        method["data_tab"]["origin_labels"] = ["2024 Q1", "2024 Q2", "2024 Q3", "2024 Q4"]
        method["results_tab"]["ultimate_vector"] = [10, 20, 30, 40]

        files = dfm_service._output_files("Project", "Class", method)
        actual = {
            int(Path(path).stem.rsplit("@", 1)[1]): [
                float(value) for value in text.splitlines() if value
            ]
            for path, text in files.items()
        }
        expected = {
            period: [float(value) for value in values if value is not None]
            for period, values in dfm_output_variants(method).items()
        }

        self.assertEqual(actual, expected)

    def test_precedent_sidecar_label_change_does_not_replace_saved_method_axis(self) -> None:
        method = self.write_method_pair()
        self.write_source(
            "Paid",
            "100,150\n200,\n",
            data_format="Triangle",
            dependents=["Development Output"],
        )
        self.write_source("Premium", "1000\n1100\n", data_format="Vector")
        for name in ("Paid", "Premium"):
            path = self.sidecars / f"{name}.json"
            sidecar = json.loads(path.read_text(encoding="utf-8"))
            sidecar["origin_labels"] = ["2023", "2025"]
            self.write_json(path, sidecar)

        result = dfm_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"], result)
        self.assertEqual(
            result["updated"],
            [{
                "dataset_name": "Development Output",
                "dataset_type": "Selected Ultimate",
                "output_changed": False,
            }],
        )
        saved = json.loads((self.methods / "DFM@Development.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["results_tab"]["ultimate_vector"], method["results_tab"]["ultimate_vector"])
        self.assertEqual(saved["data_tab"]["origin_labels"], method["data_tab"]["origin_labels"])
        # The method file was rewritten, so the output dataset's Last Modified
        # and Audit Log move with it even though the published ultimate held.
        sidecar = json.loads((self.sidecars / "Development Output.json").read_text(encoding="utf-8"))
        self.assertNotEqual(sidecar["updated_at"], "2026-01-01T00:00:00Z")
        self.assertEqual([entry["action"] for entry in sidecar["audit_log"]], ["Auto Refresh"])

    def test_basis_only_refresh_updates_method_without_rewriting_ultimate_csv(self) -> None:
        method = self.write_method_pair(status=0)
        sidecar_path = self.sidecars / "Development Output.json"
        sidecar = json.loads(sidecar_path.read_text(encoding="utf-8"))
        sidecar["dependents"] = [{"dataset_name": "Unrelated Downstream DFM"}]
        self.write_json(sidecar_path, sidecar)
        self.write_source("Paid", "100,150\n200,\n", data_format="Triangle")
        self.write_source(
            "Premium",
            "2000\n2200\n",
            data_format="Vector",
            dependents=["Development Output"],
        )
        output_path = self.datasets / "Development Output@12.csv"
        before_output = output_path.read_bytes()

        result = dfm_service.refresh_dependents("Project", "Class", ["Premium"])

        self.assertTrue(result["ok"], result)
        self.assertEqual(output_path.read_bytes(), before_output)
        saved = json.loads((self.methods / "DFM@Development.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["results_tab"]["ratio_basis_values"], [2000, 2200])
        self.assertEqual(saved["results_tab"]["ultimate_vector"], method["results_tab"]["ultimate_vector"])
        self.assertEqual(
            saved["ratios_tab"]["cell_notes"]["ratio_main_table"]["2024"]["(1) 12-24"],
            "keep",
        )
        sidecar = json.loads((self.sidecars / "Development Output.json").read_text(encoding="utf-8"))
        self.assertEqual(sidecar["status"], 2)
        self.assertNotEqual(sidecar["updated_at"], "2026-01-01T00:00:00Z")
        self.assertEqual(sidecar["audit_log"][-1]["action"], "Auto Refresh")
        self.assertEqual(
            result["review_status_updates"],
            [{"dataset_name": "Development Output", "status": 2}],
        )
        self.assertEqual(result["errors"], [])
        self.assertEqual(
            result["updated"],
            [{
                "dataset_name": "Development Output",
                "dataset_type": "Selected Ultimate",
                "output_changed": False,
            }],
        )

    def test_refresh_that_changes_nothing_keeps_last_modified_and_audit(self) -> None:
        self.write_method_pair()
        self.write_source(
            "Paid",
            "100,175\n200,\n",
            data_format="Triangle",
            dependents=["Development Output"],
        )
        first = dfm_service.refresh_dependents("Project", "Class", ["Paid"])
        self.assertTrue(first["ok"], first)
        sidecar_path = self.sidecars / "Development Output.json"
        stamped = json.loads(sidecar_path.read_text(encoding="utf-8"))
        self.assertNotEqual(stamped["updated_at"], "2026-01-01T00:00:00Z")

        with mock.patch.object(dfm_service, "_now", return_value="2030-01-01T00:00:00Z"):
            second = dfm_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(second["ok"], second)
        self.assertEqual(second["updated"], [])
        unchanged = json.loads(sidecar_path.read_text(encoding="utf-8"))
        self.assertEqual(unchanged["updated_at"], stamped["updated_at"])
        self.assertEqual(unchanged["audit_log"], stamped["audit_log"])

    def test_input_refresh_with_unchanged_origins_does_not_read_ratio_basis(self) -> None:
        self.write_method_pair()
        self.write_source(
            "Paid",
            "100,175\n200,\n",
            data_format="Triangle",
            dependents=["Development Output"],
        )

        original_load = dfm_service._load_source_snapshot
        source_reads: list[tuple[str, bool]] = []

        def record_source(*args, **kwargs):
            source_reads.append((str(args[2]), bool(kwargs.get("vector"))))
            return original_load(*args, **kwargs)

        with mock.patch.object(dfm_service, "_load_source_snapshot", side_effect=record_source):
            result = dfm_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"], result)
        self.assertEqual(source_reads, [("Paid", False)])
        self.assertFalse((self.sidecars / "Premium.json").exists())

    def test_incompatible_input_refresh_preserves_last_valid_artifacts_and_marks_review(self) -> None:
        self.write_method_pair()
        self.write_source(
            "Paid",
            "100,150,175\n200,,\n",
            data_format="Triangle",
            dependents=["Development Output"],
        )
        method_path = self.methods / "DFM@Development.json"
        output_path = self.datasets / "Development Output@12.csv"
        before_method = method_path.read_bytes()
        before_output = output_path.read_bytes()

        result = dfm_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertFalse(result["ok"])
        self.assertEqual(method_path.read_bytes(), before_method)
        self.assertEqual(output_path.read_bytes(), before_output)
        sidecar = json.loads((self.sidecars / "Development Output.json").read_text(encoding="utf-8"))
        self.assertEqual(sidecar["status"], dataset_sidecar_status_service.STATUS_REVIEW_NEEDED)

    def test_dfm_output_refreshes_downstream_dfm_ratio_basis_in_same_wave(self) -> None:
        first = self.write_method_pair(status=2)
        first_sidecar_path = self.sidecars / "Development Output.json"
        first_sidecar = json.loads(first_sidecar_path.read_text(encoding="utf-8"))
        first_sidecar["dependents"] = [{"dataset_name": "Second Output"}]
        self.write_json(first_sidecar_path, first_sidecar)
        self.write_source(
            "Paid",
            "100,180\n200,\n",
            data_format="Triangle",
            dependents=["Development Output"],
        )
        self.write_source("Premium", "1000\n1100\n", data_format="Vector")
        self.write_source("Incurred", "80,120\n160,\n", data_format="Triangle")
        second = recalculate_dfm_method(
            {
                "details_tab": {
                    "name": "Second",
                    "output_type": "Second Ultimate",
                    "output_dataset": "Second Output",
                    "input_triangle": "Incurred",
                    "origin_length": 12,
                    "development_length": 12,
                },
                "ratios_tab": {
                    "average_formulas": {
                        "label": ["User Entry"],
                        "custom_average_formula_settings": {"average_type": ["user_entry"]},
                        "selected": [[1, 1]],
                        "values": [[1.5, 1]],
                        "inputs": [["1.5", "1"]],
                    },
                },
                "results_tab": {"ratio_basis_dataset": "Development Output"},
            },
            input_snapshot={
                "name": "Incurred",
                "data_format": "Triangle",
                "origin_labels": ["2024", "2025"],
                "development_labels": ["12", "24"],
                "values": [[80, 120], [160, None]],
                "mask": [[True, True], [True, False]],
                "revision": "incurred-r1",
            },
            ratio_basis_snapshot={
                "name": "Development Output",
                "data_format": "Vector",
                "origin_labels": ["2024", "2025"],
                "values": first["results_tab"]["ultimate_vector"],
                "revision": "first-r1",
            },
            timestamp="2026-01-01T00:00:00Z",
        )
        self.write_json(self.methods / "DFM@Second.json", second)
        second_sidecar = {
            **self.output_sidecar(second, status=2),
            "dataset_name": "Second Output",
            "dataset_type": "Second Ultimate",
            "method_name": "Second",
            "csv_file": "Second Output@12.csv",
            "precedents": [
                {"dataset_name": "Incurred"},
                {"dataset_name": "Development Output"},
            ],
            "publication_revision": method_revisions(second)["publication_revision"],
        }
        self.write_json(self.sidecars / "Second Output.json", second_sidecar)
        (self.datasets / "Second Output@12.csv").write_text("120\n240\n", encoding="utf-8")

        result = dfm_service.refresh_dependents("Project", "Class", ["Paid"])

        self.assertTrue(result["ok"], result)
        by_name = {item["dataset_name"]: item for item in result["updated"]}
        self.assertTrue(by_name["Development Output"]["output_changed"])
        self.assertFalse(by_name["Second Output"]["output_changed"])
        first_saved = json.loads((self.methods / "DFM@Development.json").read_text(encoding="utf-8"))
        second_saved = json.loads((self.methods / "DFM@Second.json").read_text(encoding="utf-8"))
        self.assertEqual(
            second_saved["results_tab"]["ratio_basis_values"],
            first_saved["results_tab"]["ultimate_vector"],
        )

    def test_calculated_cascade_refreshes_dfm_before_calculated_and_rs(self) -> None:
        events: list[str] = []
        dfm_result = {
            "ok": True,
            "updated": [{
                "dataset_name": "Development Output",
                "dataset_type": "Selected Ultimate",
                "output_changed": True,
            }],
            "errors": [],
        }
        rows = [{
            "name": "Calculated Ultimate",
            "calculated": True,
            "generated": False,
            "formula": "[Selected Ultimate]",
        }]
        with (
            mock.patch.object(
                dfm_service,
                "refresh_dependents",
                side_effect=lambda *_args, **_kwargs: events.append("dfm") or dfm_result,
            ),
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=rows),
            mock.patch.object(
                calculated_dataset_service,
                "_existing_downstream_keys",
                side_effect=lambda _p, _r, roots, _rows: (
                    self.assertIn("Development Output", roots),
                    self.assertIn("Selected Ultimate", roots),
                    ["calculated ultimate"],
                )[-1],
            ),
            mock.patch.object(
                calculated_dataset_service,
                "recalculate_dataset",
                side_effect=lambda *_args, **_kwargs: events.append("calculated") or {
                    "ok": True,
                    "dataset_type_name": "Calculated Ultimate",
                },
            ),
            mock.patch(
                "app_server.services.result_selection_service.refresh_dependents",
                side_effect=lambda *_args, **_kwargs: events.append("rs") or {
                    "ok": True,
                    "updated": [],
                },
            ) as refresh_rs,
            mock.patch.object(calculated_dataset_service.dataset_instance_index_service, "rebuild_index"),
        ):
            result = calculated_dataset_service.recalculate_dependents(
                "Project", "Class", "Paid", "Paid"
            )

        self.assertTrue(result["ok"], result)
        self.assertEqual(events, ["dfm", "calculated", "dfm", "rs"])
        rs_roots = refresh_rs.call_args.args[2]
        self.assertIn("Development Output", rs_roots)
        self.assertIn("Selected Ultimate", rs_roots)

    def test_failed_dfm_root_blocks_only_its_calculated_and_rs_descendants(self) -> None:
        dfm_result = {
            "ok": False,
            "updated": [],
            "errors": [{
                "dataset_name": "Development Output",
                "dataset_type": "Selected Ultimate",
                "reason": "geometry mismatch",
            }],
        }
        rows = [{
            "name": "Calculated Ultimate",
            "calculated": True,
            "generated": False,
            "formula": "[Selected Ultimate]",
        }]
        with (
            mock.patch.object(dfm_service, "refresh_dependents", return_value=dfm_result),
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=rows),
            mock.patch.object(
                calculated_dataset_service,
                "_existing_downstream_keys",
                return_value=["calculated ultimate"],
            ),
            mock.patch.object(
                calculated_dataset_service,
                "_formula_components",
                return_value=["Selected Ultimate"],
            ),
            mock.patch.object(calculated_dataset_service, "recalculate_dataset") as recalculate,
            mock.patch(
                "app_server.services.result_selection_service.refresh_dependents",
                return_value={"ok": False, "updated": [], "errors": []},
            ) as refresh_rs,
            mock.patch.object(
                dataset_sidecar_status_service,
                "refresh_method_statuses_for_dependents",
                return_value=[],
            ),
            mock.patch.object(calculated_dataset_service.dataset_instance_index_service, "rebuild_index"),
        ):
            result = calculated_dataset_service.recalculate_dependents(
                "Project", "Class", "Paid", "Paid"
            )

        recalculate.assert_not_called()
        self.assertFalse(result["ok"])
        self.assertEqual(result["skipped"][0]["reason"], "upstream_calculation_failed")
        blocked = refresh_rs.call_args.kwargs["blocked_precedent_names"]
        self.assertIn("Development Output", blocked)
        self.assertIn("Selected Ultimate", blocked)

    def test_calculated_output_refreshes_dfm_before_later_calculated_descendant(self) -> None:
        events: list[str] = []
        dfm_results = iter([
            {"ok": True, "updated": [], "errors": []},
            {
                "ok": True,
                "updated": [{
                    "dataset_name": "Method B Output",
                    "dataset_type": "Method B Ultimate",
                    "output_changed": True,
                }],
                "errors": [],
            },
            {"ok": True, "updated": [], "errors": []},
        ])
        rows = [
            {"name": "Calculated C", "calculated": True, "generated": False, "formula": "[Paid]"},
            {
                "name": "Calculated D",
                "calculated": True,
                "generated": False,
                "formula": "[Method B Ultimate]",
            },
        ]

        def refresh_dfm(*_args, **_kwargs):
            events.append("dfm")
            return next(dfm_results)

        def downstream(_project, _reserving, roots, _rows):
            normalized = {str(item).casefold() for item in roots}
            return ["calculated d"] if "method b ultimate" in normalized else ["calculated c"]

        def recalculate(_project, _reserving, name, **_kwargs):
            events.append(name)
            return {"ok": True, "dataset_type_name": name}

        with (
            mock.patch.object(dfm_service, "refresh_dependents", side_effect=refresh_dfm),
            mock.patch.object(calculated_dataset_service, "_dataset_type_rows", return_value=rows),
            mock.patch.object(calculated_dataset_service, "_existing_downstream_keys", side_effect=downstream),
            mock.patch.object(calculated_dataset_service, "recalculate_dataset", side_effect=recalculate),
            mock.patch(
                "app_server.services.result_selection_service.refresh_dependents",
                side_effect=lambda *_args, **_kwargs: events.append("rs") or {"ok": True, "updated": []},
            ) as refresh_rs,
            mock.patch.object(calculated_dataset_service.dataset_instance_index_service, "rebuild_index"),
        ):
            result = calculated_dataset_service.recalculate_dependents(
                "Project", "Class", "Paid", "Paid"
            )

        self.assertTrue(result["ok"], result)
        self.assertEqual(
            events,
            ["dfm", "Calculated C", "dfm", "Calculated D", "dfm", "rs"],
        )
        self.assertIn("Method B Output", refresh_rs.call_args.args[2])
        self.assertIn("Method B Ultimate", refresh_rs.call_args.args[2])

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
        real_replace = dfm_service.os.replace
        targets: list[str] = []

        def replace(source: str, target: str) -> None:
            targets.append(target)
            if target == str(sidecar_path) and source.endswith(".tmp"):
                raise OSError("sidecar publish failed")
            real_replace(source, target)

        with mock.patch.object(dfm_service.os, "replace", side_effect=replace):
            with self.assertRaises(OSError):
                dfm_service._commit_text_files(
                    {
                        str(method_path): "new-method\n",
                        str(csv_path): "new-csv\n",
                        str(sidecar_path): "new-sidecar\n",
                    },
                    last_paths=[str(sidecar_path)],
                )

        self.assertEqual(
            {path: path.read_bytes() for path in original},
            original,
        )
        first_sidecar_target = targets.index(str(sidecar_path))
        self.assertGreater(first_sidecar_target, targets.index(str(method_path)))
        self.assertGreater(first_sidecar_target, targets.index(str(csv_path)))

    def test_staged_publish_skips_unchanged_files(self) -> None:
        path = self.methods / "unchanged.json"
        path.write_bytes(b"same\n")
        with mock.patch.object(dfm_service.os, "replace") as replace:
            changed = dfm_service._commit_text_files({str(path): "same\n"})
        self.assertEqual(changed, [])
        replace.assert_not_called()


if __name__ == "__main__":
    unittest.main()

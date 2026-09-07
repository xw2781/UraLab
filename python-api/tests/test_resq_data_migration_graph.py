from __future__ import annotations

import importlib.util
import importlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock


_TMP_ROOT = Path(__file__).resolve().parent / "logs" / "tmp"
_MIGRATION_PATH = Path(__file__).resolve().parents[1] / "migration" / "resq_data_migration.py"


def load_migration_module():
    spec = importlib.util.spec_from_file_location("resq_data_migration_under_test", _MIGRATION_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load resq_data_migration.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ResqDataMigrationGraphTests(unittest.TestCase):
    def setUp(self) -> None:
        _TMP_ROOT.mkdir(parents=True, exist_ok=True)
        self.tmp = tempfile.TemporaryDirectory(dir=str(_TMP_ROOT))
        self.root = Path(self.tmp.name) / "ArcRho Server"
        self.project_dir = self.root / "projects" / "Demo"
        self.rc_dir = self.project_dir / "data" / "Auto_%5C_PP"
        self.datasets_dir = self.rc_dir / "datasets"
        self.methods_dir = self.rc_dir / "methods"
        self.sidecars_dir = self.rc_dir / "sidecars"
        self.datasets_dir.mkdir(parents=True)
        self.methods_dir.mkdir()
        self.sidecars_dir.mkdir()

        self.module = load_migration_module()
        self.module.SERVER_ROOT = self.root
        self.module.PROJECT_NAME = "Demo"
        self.module.PROJECT_DATA_DIR = self.project_dir / "data"
        self.catalog = importlib.import_module("resq_migration.catalog")
        self.catalog.configure_catalog(
            server_root=self.root,
            project_name="Demo",
            rs_json_format=self.module.RS_JSON_FORMAT,
            method_data_dir=self.module.METHOD_DATA_DIR,
        )
        self.extractors = importlib.import_module("resq_migration.extractors")
        self.extractors.configure_extractors(
            project_name="Demo",
            rs_json_format=self.module.RS_JSON_FORMAT,
            method_data_dir=self.module.METHOD_DATA_DIR,
        )

        (self.project_dir / "dataset_types.json").write_text(json.dumps({
            "columns": ["Formula", "Generated", "Name", "Calculated", "Data Format", "Category", "Source"],
            "rows": [
                ["", True, "Paid Loss", False, "Triangle", "Loss", ""],
                ["", True, "DFM Ultimate", False, "Vector", "Loss", ""],
                ["", True, "Generated Premium", False, "Vector", "Premium", "Generated_Premium"],
                ["\"Paid Loss\" + \"DFM Ultimate\"", False, "Net Ultimate", True, "Vector", "Loss", ""],
                ["\"Net Ultimate\" * 1.1", False, "Loaded Ultimate", True, "Vector", "Loss", ""],
                ["", False, "Prior Qtr Indicated", False, "Vector", "Loss", ""],
            ],
        }), encoding="utf-8")

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_formula_graph_fields_include_resolved_dependency_info(self) -> None:
        paid_csv = self.datasets_dir / "Paid Loss@12@12@cum@dev.csv"
        paid_csv.write_text("1,2\n", encoding="utf-8")
        (self.sidecars_dir / "Paid Loss.json").write_text(json.dumps({
            "dataset_name": "Paid Loss",
            "dataset_type": "Paid Loss",
            "source_kind": "engine",
            "formula": "",
        }), encoding="utf-8")
        dfm_path = self.methods_dir / "DFM@Selected DFM.json"
        dfm_input = str(self.datasets_dir / "Paid Loss@12@12@cum@dev.csv")
        dfm_path.write_text(json.dumps({
            "json_format": self.module.DFM_JSON_FORMAT,
            "details_tab": {"name": "Selected DFM", "output_type": "DFM Ultimate"},
            "data_tab": {"input data triangle csv path": dfm_input},
        }), encoding="utf-8")
        (self.sidecars_dir / "Loaded Ultimate.json").write_text(json.dumps({
            "dataset_name": "Loaded Ultimate",
            "dataset_type": "Loaded Ultimate",
            "source_kind": "calculated",
            "formula": "\"Net Ultimate\" * 1.1",
        }), encoding="utf-8")

        graph = self.catalog._dataset_type_graph_fields("Net Ultimate", self.rc_dir)

        # v4 keeps the persisted graph location-independent: an entry names the
        # dataset and nothing else, even though the inventory could resolve its
        # file, source kind, method input, or formula.
        self.assertEqual(
            graph["precedents"],
            [{"dataset_name": "Paid Loss"}, {"dataset_name": "DFM Ultimate"}],
        )
        self.assertEqual(graph["dependents"], [{"dataset_name": "Loaded Ultimate"}])

    def test_formula_graph_omits_absent_rc_dependents(self) -> None:
        (self.sidecars_dir / "Net Ultimate.json").write_text(json.dumps({
            "dataset_name": "Net Ultimate",
            "dataset_type": "Net Ultimate",
            "source_kind": "calculated",
            "formula": "\"Paid Loss\" + \"DFM Ultimate\"",
        }), encoding="utf-8")

        graph = self.catalog._dataset_type_graph_fields("Net Ultimate", self.rc_dir)

        self.assertEqual(graph["dependents"], [])

    def test_bulk_graph_refresh_scans_physical_inventory_once(self) -> None:
        (self.datasets_dir / "Paid Loss@12@12@cum@dev.csv").write_text(
            "1,2\n",
            encoding="utf-8",
        )
        (self.sidecars_dir / "Paid Loss.json").write_text(json.dumps({
            "dataset_name": "Paid Loss",
            "dataset_type": "Paid Loss",
            "source_kind": "engine",
        }), encoding="utf-8")
        net_path = self.sidecars_dir / "Net Ultimate.json"
        net_path.write_text(json.dumps({
            "dataset_name": "Net Ultimate",
            "dataset_type": "Net Ultimate",
            "source_kind": "calculated",
            "formula": "\"Paid Loss\" + \"DFM Ultimate\"",
        }), encoding="utf-8")
        (self.sidecars_dir / "Loaded Ultimate.json").write_text(json.dumps({
            "dataset_name": "Loaded Ultimate",
            "dataset_type": "Loaded Ultimate",
            "source_kind": "calculated",
            "formula": "\"Net Ultimate\" * 1.1",
        }), encoding="utf-8")

        original_scan = self.catalog._scan_physical_dataset_files
        with mock.patch.object(
            self.catalog,
            "_scan_physical_dataset_files",
            wraps=original_scan,
        ) as scan:
            self.catalog.refresh_sidecar_graphs_for_rc(self.rc_dir)

        self.assertEqual(scan.call_count, 1)
        payload = json.loads(net_path.read_text(encoding="utf-8"))
        self.assertEqual(
            [item["dataset_name"] for item in payload["precedents"]],
            ["Paid Loss", "DFM Ultimate"],
        )
        self.assertEqual(
            [item["dataset_name"] for item in payload["dependents"]],
            ["Loaded Ultimate"],
        )

    def test_deferred_graph_enrichment_is_completed_by_bulk_refresh(self) -> None:
        payload = {
            "name": "Net Ultimate",
            "dataset_type": "Net Ultimate",
            "category": "Loss",
            "data_format": 1,
            "method_type": "None",
            "method_type_code": 0,
            "origin_length": 12,
            "development_length": 12,
            "origin_count": 1,
            "development_count": 1,
            "origin_labels": ["2026"],
            "development_labels": ["Value"],
            "values": [[123.0]],
            "formula": "\"Paid Loss\" + \"DFM Ultimate\"",
            "user": "tester",
            "created": "2026-01-01T00:00:00",
            "modified": "2026-01-02T00:00:00",
        }

        with self.extractors.defer_sidecar_graph_enrichment():
            self.extractors.write_vector_export(payload, r"Auto\PP", self.rc_dir)

        sidecar_path = self.sidecars_dir / "Net Ultimate.json"
        deferred = json.loads(sidecar_path.read_text(encoding="utf-8"))
        self.assertNotIn("precedents", deferred)
        self.assertNotIn("dependents", deferred)

        self.catalog.refresh_sidecar_graphs_for_rc(self.rc_dir)

        refreshed = json.loads(sidecar_path.read_text(encoding="utf-8"))
        self.assertEqual(
            [item["dataset_name"] for item in refreshed["precedents"]],
            ["Paid Loss", "DFM Ultimate"],
        )
        self.assertEqual(refreshed["dependents"], [])

    def test_deferred_graph_enrichment_scope_restores_after_error(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "stop"):
            with self.extractors.defer_sidecar_graph_enrichment():
                raise RuntimeError("stop")

        with mock.patch.object(
            self.extractors,
            "_apply_sidecar_graph_meta",
            wraps=self.extractors._apply_sidecar_graph_meta,
        ) as apply_graph:
            self.extractors._apply_graph_meta_best_effort(
                {},
                "Paid Loss",
                self.rc_dir,
            )

        self.assertEqual(apply_graph.call_count, 1)

    def _vector_payload(self, name: str, *, formula: str = "") -> dict:
        return {
            "name": name,
            "dataset_type": name,
            "category": "Loss",
            "data_format": 1,
            "method_type": "None",
            "method_type_code": 0,
            "origin_length": 12,
            "development_length": 12,
            "origin_count": 1,
            "development_count": 1,
            "origin_labels": ["2026"],
            "development_labels": ["Value"],
            "values": [[123.0]],
            "formula": formula,
            "user": "tester",
            "created": "2026-01-01T00:00:00",
            "modified": "2026-01-02T00:00:00",
        }

    def test_vector_calculated_only_in_resq_imports_as_editable_input(self) -> None:
        # ResQ derives "Prior Qtr Indicated" from a formula ArcRho does not
        # have; ArcRho's library lists the type as a plain input, and ArcRho's
        # library wins, so the dataset is an ordinary editable input here.
        self.extractors.write_vector_export(
            self._vector_payload("Prior Qtr Indicated", formula='"Current Qtr Indicated - Feb 2026"'),
            r"Auto\PP",
            self.rc_dir,
        )

        payload = json.loads((self.sidecars_dir / "Prior Qtr Indicated.json").read_text(encoding="utf-8"))
        self.assertEqual(payload["source_kind"], "input")
        self.assertFalse(payload["calculated"])
        self.assertEqual(payload["source"], "resq_vector")
        self.assertNotIn("formula", payload)
        self.assertEqual(payload["precedents"], [])

    def test_vector_calculated_in_arcrho_library_stays_calculated(self) -> None:
        # The reverse disagreement: ArcRho computes "Loaded Ultimate" from
        # "Net Ultimate" even though ResQ holds it as a plain vector.
        self.extractors.write_vector_export(
            self._vector_payload("Loaded Ultimate"),
            r"Auto\PP",
            self.rc_dir,
        )

        payload = json.loads((self.sidecars_dir / "Loaded Ultimate.json").read_text(encoding="utf-8"))
        self.assertEqual(payload["source_kind"], "calculated")
        self.assertTrue(payload["calculated"])
        self.assertEqual([entry["dataset_name"] for entry in payload["precedents"]], ["Net Ultimate"])

    def _instance_formula_payload(self, name: str, formula: str) -> dict:
        return {
            "name": name,
            "dataset_type": name,
            "category": "Loss",
            "data_format": 1,
            "method_type": "None",
            "method_type_code": 0,
            # Annual granularity: origin_length is months per period (12), not
            # the row count. The formula link must span origin_count rows (2).
            "origin_length": 12,
            "development_length": 12,
            "origin_count": 2,
            "development_count": 1,
            "origin_labels": ["2025", "2026"],
            "development_labels": ["Value"],
            "values": [[10.0], [20.0]],
            "formula": formula,
            "user": "tester",
            "created": "2026-01-01T00:00:00",
            "modified": "2026-01-02T00:00:00",
        }

    def test_instance_formula_translates_into_a_formula_link_and_edges(self) -> None:
        known = ["Vector A", "Vector B", "Vector C"]
        self.extractors.write_vector_export(
            self._instance_formula_payload("Vector A", ""),
            r"Auto\PP",
            self.rc_dir,
            known_instance_names=known,
        )
        self.extractors.write_vector_export(
            self._instance_formula_payload("Vector C", '"Vector A" * "Vector B" / 1000'),
            r"Auto\PP",
            self.rc_dir,
            known_instance_names=known,
        )

        sidecar = json.loads((self.sidecars_dir / "Vector C.json").read_text(encoding="utf-8"))
        self.assertEqual(sidecar["source_kind"], "input")
        self.assertEqual(
            sidecar["formula_links"],
            [{
                "formula": "=[Vector A][1:2] * [Vector B][1:2] / 1000",
                "target_cells": [
                    {"row": 0, "column": 0, "result_row": 0, "result_column": 0},
                    {"row": 1, "column": 0, "result_row": 1, "result_column": 0},
                ],
            }],
        )
        # The link is written against the display the vector is shown at.
        self.assertEqual(sidecar["linked_period_length"], sidecar["period_length"])
        # The link contributes instance-level precedent edges alongside the
        # type graph, and the RC-wide refresh writes the matching dependents
        # entry on each linked source that exists in the class.
        self.assertEqual(
            [entry["dataset_name"] for entry in sidecar["precedents"]],
            ["Vector A", "Vector B"],
        )
        self.catalog.refresh_sidecar_graphs_for_rc(self.rc_dir)
        source = json.loads((self.sidecars_dir / "Vector A.json").read_text(encoding="utf-8"))
        self.assertIn(
            {"dataset_name": "Vector C"},
            source["dependents"],
        )
        refreshed = json.loads((self.sidecars_dir / "Vector C.json").read_text(encoding="utf-8"))
        self.assertEqual(
            [entry["dataset_name"] for entry in refreshed["precedents"]],
            ["Vector A", "Vector B"],
        )

    def test_instance_formula_falls_back_to_hardcoded_values(self) -> None:
        cases = [
            # A frozen prior-quarter snapshot ArcRho never imports.
            ('"Current Qtr Indicated - Feb 2026"', ["Vector A"]),
            # Text outside the quoted-name / number / operator grammar.
            ('"Vector A" @ 2', ["Vector A"]),
            # A self-reference cannot become a link.
            ('"Vector D" * 2', ["Vector A", "Vector D"]),
            # No caller-supplied inventory means no translation.
            ('"Vector A" * 2', None),
        ]
        for formula, known in cases:
            with self.subTest(formula=formula, known=known):
                self.extractors.write_vector_export(
                    self._instance_formula_payload("Vector D", formula),
                    r"Auto\PP",
                    self.rc_dir,
                    known_instance_names=known,
                )
                sidecar = json.loads(
                    (self.sidecars_dir / "Vector D.json").read_text(encoding="utf-8")
                )
                self.assertNotIn("formula_links", sidecar)
                self.assertEqual(sidecar["source_kind"], "input")

    def test_generated_vector_ignores_resq_formula_metadata(self) -> None:
        self.extractors.write_vector_export({
            "name": "Generated Premium",
            "dataset_type": "Generated Premium",
            "category": "Premium",
            "data_format": 1,
            "method_type": "None",
            "method_type_code": 0,
            "origin_length": 12,
            "development_length": 12,
            "origin_count": 1,
            "development_count": 1,
            "origin_labels": ["2026"],
            "development_labels": ["Value"],
            "values": [[123.0]],
            "formula": '"Some ResQ Source" + 1',
            "user": "tester",
            "created": "2026-01-01T00:00:00",
            "modified": "2026-01-02T00:00:00",
        }, r"Auto\PP", self.rc_dir)

        payload = json.loads((self.sidecars_dir / "Generated Premium.json").read_text(encoding="utf-8"))
        self.assertEqual(payload["source_kind"], "engine")
        self.assertFalse(payload["calculated"])
        self.assertNotIn("formula", payload)
        self.assertEqual(payload["period_length"], 12)
        self.assertNotIn("origin_length", payload)
        self.assertNotIn("development_length", payload)
        self.assertNotIn("development_count", payload)
        self.assertNotIn("cumulative", payload)
        self.assertNotIn("calendar", payload)

    def test_result_selection_output_vector_status_is_persisted(self) -> None:
        class OutputVector:
            Name = "Current Selection"
            OriginCount = 1
            PeriodLength = 12
            MethodType = 4
            Status = 2
            User = "tester"
            Created = "2026-01-01T00:00:00"
            Modified = "2026-01-02T00:00:00"
            Formula = ""
            Notes = "Selected after review.\r\nKeep."
            DatasetType = type(
                "DatasetType",
                (),
                {
                    "Name": "Net Ultimate",
                    "DataFormat": 1,
                    "Category": type("Category", (), {"Name": "Loss"})(),
                },
            )()

            def OriginLabel(self, _index):
                return "2026"

            def ValuesByIndex(self, _index):
                return 123.0

        payload = self.extractors.export_vector(OutputVector())
        payload["precedents"] = ["DFM Ultimate"]
        self.extractors.write_vector_export(payload, r"Auto\PP", self.rc_dir)

        sidecar = json.loads(
            (self.sidecars_dir / "Current Selection.json").read_text(encoding="utf-8")
        )
        self.assertEqual(payload["status"], 2)
        self.assertEqual(sidecar["status"], 2)
        # ResQ Notes ride into the sidecar, the ArcRho owner of dataset notes.
        self.assertEqual(payload["notes"], "Selected after review.\r\nKeep.")
        self.assertEqual(sidecar["notes"], "Selected after review.\r\nKeep.")

        index_path = self.catalog.rebuild_dataset_instance_index(
            "Demo", r"Auto\PP", self.rc_dir
        )
        index = json.loads(index_path.read_text(encoding="utf-8"))
        item = next(row for row in index["files"] if row["name"] == "Current Selection")
        self.assertEqual(item["status"], 2)

    def test_dataset_instance_index_uses_only_reserving_class_metadata(self) -> None:
        config_dir = self.root / "config"
        config_dir.mkdir()
        (config_dir / "username_index.json").write_text(json.dumps({
            "users": [
                {"login_name": "xwei", "full_name": "Wei, Xiao"},
            ],
        }), encoding="utf-8")
        csv_path = self.datasets_dir / "Net Ultimate@12.csv"
        csv_path.write_text("1\n", encoding="utf-8")
        (self.sidecars_dir / "Net Ultimate.json").write_text(json.dumps({
            "dataset_name": "Net Ultimate",
            "dataset_type": "Net Ultimate",
            "source_kind": "engine",
            "data_format": "Vector",
            "formula": "",
            "user": "xwei",
        }), encoding="utf-8")

        index_path = self.catalog.rebuild_dataset_instance_index("Demo", r"Auto\PP", self.rc_dir)
        payload = json.loads(index_path.read_text(encoding="utf-8"))
        rows = {item["name"]: item for item in payload["files"]}

        self.assertEqual(payload["version"], self.module.INDEX_VERSION)
        self.assertNotIn("formula", rows["Net Ultimate"])
        self.assertNotIn("dataset_category", rows["Net Ultimate"])
        self.assertEqual(rows["Net Ultimate"]["user"], "xwei")

    def test_dfm_ultimate_vector_sidecar_uses_period_length(self) -> None:
        contract = importlib.import_module("arcrho_api.dfm_contract")
        method_payload = contract.recalculate_dfm_method({
            "json_format": contract.DFM_JSON_FORMAT,
            "details_tab": {
                "name": "Paid DFM",
                "output_type": "DFM Ultimate",
                "output_dataset": "Ultimate",
                "input_triangle": "Paid Loss",
                "origin_length": 6,
                "development_length": 6,
                "decimal_places": 4,
            },
            "data_tab": {},
            "ratios_tab": {
                "ratio_triangle": {"excluded": [[0]]},
                "average_formulas": {
                    "label": ["Simple - all"],
                    "custom_average_formula_settings": {
                        "average_type": ["custom"],
                        "base": ["simple"],
                        "periods": ["all"],
                        "exclude": [0],
                    },
                    "selected": [[1]],
                    "values": [[1]],
                    "inputs": [[""]],
                },
                "cell_notes": {"ratio_main_table": {}, "ratio_summary_table": {}},
            },
            "results_tab": {},
            "method_metadata": {
                "last_modified": "2026-01-02T00:00:00",
                "data_refreshed": "2026-01-02T00:00:00",
            },
        }, input_snapshot={
            "name": "Paid Loss",
            "origin_labels": ["2026"],
            "development_labels": ["6m", "12m"],
            "values": [[100.0, 123.0]],
            "mask": [[True, True]],
            "data_format": "Triangle",
            "number_format": "#,##0",
            "decimal_places": 0,
        })
        self.extractors.write_dfm_ultimate_vector_export({
            "name": "Ultimate",
            "dataset_type": "DFM Ultimate",
            "category": "Loss",
            "data_format": 1,
            "method_type": "DFM",
            "method_type_code": 7,
            "origin_length": 6,
            "development_length": 6,
            "origin_count": 1,
            "development_count": 1,
            "origin_labels": ["2026"],
            "development_labels": ["Ultimate"],
            "values": [[123.0]],
            "method_name": "Paid DFM",
            "user": "tester",
            "created": "2026-01-01T00:00:00",
            "modified": "2026-01-02T00:00:00",
        }, r"Auto\PP", self.rc_dir, method_payload=method_payload)

        payload = json.loads((self.sidecars_dir / "Ultimate.json").read_text(encoding="utf-8"))
        self.assertEqual(payload["source_kind"], "dfm")
        self.assertEqual(payload["period_length"], 6)
        self.assertNotIn("origin_length", payload)
        self.assertNotIn("development_length", payload)
        self.assertNotIn("development_count", payload)
        self.assertNotIn("cumulative", payload)
        self.assertNotIn("calendar", payload)

    def test_refresh_preserves_result_selection_precedent_strings(self) -> None:
        path = self.sidecars_dir / "Net Ultimate.json"
        path.write_text(json.dumps({
            "dataset_name": "Net Ultimate",
            "dataset_type": "Net Ultimate",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "precedents": ["Paid Loss"],
            "dependents": [],
        }), encoding="utf-8")
        (self.sidecars_dir / "Loaded Ultimate.json").write_text(json.dumps({
            "dataset_name": "Loaded Ultimate",
            "dataset_type": "Loaded Ultimate",
            "source_kind": "calculated",
            "formula": "\"Net Ultimate\" * 1.1",
        }), encoding="utf-8")

        updated = self.module.refresh_sidecar_graphs_for_rc(self.rc_dir)

        payload = json.loads(path.read_text(encoding="utf-8"))
        self.assertGreaterEqual(updated, 1)
        self.assertEqual(payload["precedents"], ["Paid Loss"])
        self.assertEqual(payload["dependents"][0]["dataset_name"], "Loaded Ultimate")

    def test_refresh_preserves_bf_precedents_and_refreshes_method_status(self) -> None:
        (self.sidecars_dir / "Paid Loss.json").write_text(json.dumps({
            "dataset_name": "Paid Loss",
            "dataset_type": "Paid Loss",
            "source_kind": "engine",
            "updated_at": "2026-07-02T00:00:00Z",
        }), encoding="utf-8")
        path = self.sidecars_dir / "Net Ultimate.json"
        precedents = [
            {"dataset_name": "Paid Loss"},
            {"dataset_name": "Selected DFM"},
            {"dataset_name": "Prior Ultimate"},
        ]
        path.write_text(json.dumps({
            "dataset_name": "Net Ultimate",
            "dataset_type": "Net Ultimate",
            "source_kind": "bornhuetter_ferguson",
            "method_type": "Bornhuetter Ferguson",
            "updated_at": "2026-07-01T00:00:00Z",
            "status": 0,
            "precedents": precedents,
            "dependents": [],
        }), encoding="utf-8")
        (self.sidecars_dir / "Loaded Ultimate.json").write_text(json.dumps({
            "dataset_name": "Loaded Ultimate",
            "dataset_type": "Loaded Ultimate",
            "source_kind": "calculated",
            "formula": "\"Net Ultimate\" * 1.1",
        }), encoding="utf-8")

        updated = self.module.refresh_sidecar_graphs_for_rc(self.rc_dir)

        payload = json.loads(path.read_text(encoding="utf-8"))
        self.assertGreaterEqual(updated, 1)
        self.assertEqual(payload["precedents"], precedents)
        self.assertEqual(payload["dependents"][0]["dataset_name"], "Loaded Ultimate")
        self.assertEqual(payload["status"], 2)

    def test_fresh_engine_cache_of_unchanged_data_keeps_method_status_ok(self) -> None:
        # An import rewrites every engine cache at import time, so updated_at is
        # always newer than the imported method's ResQ timestamp. Only
        # source_modified (when the data changed in ResQ) may flip an OK method.
        (self.sidecars_dir / "Paid Loss.json").write_text(json.dumps({
            "dataset_name": "Paid Loss",
            "dataset_type": "Paid Loss",
            "source_kind": "engine",
            "source_modified": "2026-06-15T00:00:00Z",
            "updated_at": "2026-08-03T20:52:00Z",
        }), encoding="utf-8")
        path = self.sidecars_dir / "Net Ultimate.json"
        path.write_text(json.dumps({
            "dataset_name": "Net Ultimate",
            "dataset_type": "Net Ultimate",
            "source_kind": "bornhuetter_ferguson",
            "method_type": "Bornhuetter Ferguson",
            "updated_at": "2026-07-01T00:00:00Z",
            "status": 0,
            "precedents": [{"dataset_name": "Paid Loss"}],
            "dependents": [],
        }), encoding="utf-8")

        self.module.refresh_sidecar_graphs_for_rc(self.rc_dir)

        payload = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(payload["status"], 0)

    def test_engine_cache_with_newer_source_data_flips_method_status(self) -> None:
        (self.sidecars_dir / "Paid Loss.json").write_text(json.dumps({
            "dataset_name": "Paid Loss",
            "dataset_type": "Paid Loss",
            "source_kind": "engine",
            "source_modified": "2026-07-02T00:00:00Z",
            "updated_at": "2026-08-03T20:52:00Z",
        }), encoding="utf-8")
        path = self.sidecars_dir / "Net Ultimate.json"
        path.write_text(json.dumps({
            "dataset_name": "Net Ultimate",
            "dataset_type": "Net Ultimate",
            "source_kind": "bornhuetter_ferguson",
            "method_type": "Bornhuetter Ferguson",
            "updated_at": "2026-07-01T00:00:00Z",
            "status": 0,
            "precedents": [{"dataset_name": "Paid Loss"}],
            "dependents": [],
        }), encoding="utf-8")

        self.module.refresh_sidecar_graphs_for_rc(self.rc_dir)

        payload = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(payload["status"], 2)

    def test_refresh_preserves_imported_needs_review_status_when_precedents_are_current(self) -> None:
        (self.sidecars_dir / "Paid Loss.json").write_text(json.dumps({
            "dataset_name": "Paid Loss",
            "dataset_type": "Paid Loss",
            "source_kind": "engine",
            "updated_at": "2026-07-01T00:00:00Z",
        }), encoding="utf-8")
        path = self.sidecars_dir / "Current Selection.json"
        path.write_text(json.dumps({
            "dataset_name": "Current Selection",
            "dataset_type": "Net Ultimate",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "updated_at": "2026-07-02T00:00:00Z",
            "status": 2,
            "precedents": ["Paid Loss"],
            "dependents": [],
        }), encoding="utf-8")

        self.module.refresh_sidecar_graphs_for_rc(self.rc_dir)

        payload = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(payload["status"], 2)

    def test_refresh_adds_result_selection_to_precedent_dependents(self) -> None:
        source_path = self.sidecars_dir / "DFM Ultimate.json"
        source_path.write_text(json.dumps({
            "dataset_name": "DFM Ultimate",
            "dataset_type": "DFM Ultimate",
            "source_kind": "dfm",
            "method_type": "DFM",
            "updated_at": "2026-07-01T14:51:31Z",
            "precedents": [{"dataset_name": "Paid Loss"}],
            "dependents": [],
        }), encoding="utf-8")
        dependent_path = self.sidecars_dir / "Current Selection.json"
        dependent_path.write_text(json.dumps({
            "dataset_name": "Current Selection",
            "dataset_type": "Loaded Ultimate",
            "source_kind": "result_selection",
            "method_type": "Result Selection",
            "updated_at": "2026-06-18T17:11:12Z",
            "status": 0,
            "precedents": ["DFM Ultimate"],
            "dependents": [],
        }), encoding="utf-8")

        updated = self.module.refresh_sidecar_graphs_for_rc(self.rc_dir)

        source_payload = json.loads(source_path.read_text(encoding="utf-8"))
        dependent_payload = json.loads(dependent_path.read_text(encoding="utf-8"))
        self.assertGreaterEqual(updated, 1)
        self.assertEqual(
            [item["dataset_name"] for item in source_payload["dependents"]],
            ["Current Selection"],
        )
        self.assertEqual(dependent_payload["status"], 2)

    def test_result_selection_vector_metadata_uses_method_tab_origin_labels(self) -> None:
        payload = {
            "origin_labels": ["1", "2"],
            "origin_count": 2,
            "precedents": [],
        }
        result_selection_payload = {
            "_sidecar_notes": "selection note",
            "details_tab": {
                "ratio_basis_datasets": ["Earned Premium"],
            },
            "method_tab": {
                "origin_labels": ["2016", "2017", "2018"],
                "loaded_datasets": [
                    {"name": "Paid Loss"},
                    {"name": "Reported Loss"},
                ],
            },
        }

        self.module._apply_result_selection_vector_metadata(payload, result_selection_payload)

        self.assertEqual(payload["origin_labels"], ["2016", "2017", "2018"])
        self.assertEqual(payload["origin_count"], 3)
        self.assertEqual(payload["precedents"], ["Paid Loss", "Reported Loss", "Earned Premium"])
        self.assertEqual(payload["notes"], "selection note")
        self.assertNotIn("_sidecar_notes", result_selection_payload)

    def test_bornhuetter_ferguson_notes_move_to_output_sidecar_metadata(self) -> None:
        payload = {}
        method_payload = {
            "_sidecar_notes": "BF note",
            "method_tab": {
                "latest_dataset": "Paid Loss",
                "dfm_dataset": "Paid Ultimate",
                "prior_datasets": [{"name": "Prior Ultimate"}],
            },
        }

        self.module._apply_bornhuetter_ferguson_vector_metadata(payload, method_payload)

        self.assertEqual(payload["notes"], "BF note")
        self.assertNotIn("_sidecar_notes", method_payload)
        self.assertEqual(payload["precedents"], ["Paid Loss", "Paid Ultimate", "Prior Ultimate"])

    def test_cape_cod_notes_move_to_output_sidecar_metadata(self) -> None:
        payload = {}
        method_payload = {
            "_sidecar_notes": "Cape Cod note",
            "method_tab": {
                "latest_dataset": "Paid Loss",
                "exposure_dataset": "Earned Premium",
                "prior_ultimate_dataset": "Prior Ultimate",
            },
        }

        self.module._apply_cape_cod_vector_metadata(payload, method_payload)

        self.assertEqual(payload["notes"], "Cape Cod note")
        self.assertNotIn("_sidecar_notes", method_payload)
        self.assertEqual(payload["source_kind"], "cape_cod")
        self.assertEqual(payload["method_type"], "Cape Cod")
        self.assertEqual(payload["method_type_code"], self.module.METHOD_TYPE_CAPE_COD_CODE)
        self.assertEqual(payload["precedents"], ["Paid Loss", "Earned Premium", "Prior Ultimate"])

    def test_result_selection_source_payload_includes_native_origin_length(self) -> None:
        class DatasetType:
            Name = "Paid Loss"
            DataFormat = 1
            Category = type("Category", (), {"Name": "Loss"})()

        class Dataset:
            Name = "Paid Loss"
            MethodType = 0

        Dataset.DatasetType = DatasetType()

        class ResultSelection:
            def Dataset(self, _dataset_index):
                return Dataset()

            def DatasetValues(self, _dataset_index, origin_index, _origin_length):
                return origin_index * 10

            def Weights(self, _dataset_index, origin_index):
                return 1 if origin_index == 1 else 0

        payload = self.module._result_selection_source_payload(ResultSelection(), 1, 2, 12)

        self.assertNotIn("selected", payload)
        self.assertNotIn("value_source", payload)
        self.assertEqual(payload["origin_length"], 12)
        self.assertEqual(payload["source_kind"], "input")
        self.assertEqual(payload["weights"], [1, 0])

    def test_export_result_selection_matches_frontend_method_shape(self) -> None:
        class OutputDatasetType:
            Name = "Selected Ultimate"

        class OutputVector:
            Name = "Selected Ultimate"
            Modified = "2026-01-01T00:00:00"
            Status = 2

        OutputVector.DatasetType = OutputDatasetType()

        class SourceDatasetType:
            Name = "Paid Loss"
            DataFormat = 1
            Category = type("Category", (), {"Name": "Loss"})()

        class SourceDataset:
            Name = "Paid Loss"
            MethodType = 0

        SourceDataset.DatasetType = SourceDatasetType()

        class ResultSelection:
            OriginLength = 12
            OriginCount = 2
            DatasetCount = 1
            Notes = ""

            def OriginLabel(self, origin_index):
                return str(2015 + origin_index)

            def Dataset(self, _dataset_index):
                return SourceDataset()

            def DatasetValues(self, _dataset_index, origin_index, _origin_length):
                return origin_index * 10.123456789

            def Weights(self, _dataset_index, _origin_index):
                return 1.987654321

            def Ultimates(self, origin_index, _origin_length):
                return origin_index * 100.123456789

            def UltimateOverridden(self, *, OriginIndex):
                return OriginIndex == 2

            def RatioBasisDataset(self, dataset_index):
                if dataset_index != 1:
                    raise IndexError(dataset_index)
                return type("RatioBasisDataset", (), {"Name": "Earned Premium"})()

            def RatioBasisValues(self, origin_index, _origin_length):
                return origin_index * 1000.123456789

        ResultSelection.OutputVector = OutputVector()

        payload = self.module.export_result_selection(ResultSelection())

        self.assertEqual(payload["_sidecar_status"], 2)
        self.assertEqual(payload["details_tab"]["ratio_basis_datasets"], ["Earned Premium"])
        self.assertEqual(payload["details_tab"]["active_ratio_basis_dataset"], "Earned Premium")
        self.assertNotIn("ratio_basis_dataset", payload["details_tab"])
        self.assertNotIn("ratio_basis", payload["details_tab"])
        self.assertNotIn("dataset_category", payload["details_tab"])
        self.assertNotIn("output_category", payload["details_tab"])
        self.assertNotIn("sources", payload["method_tab"])
        self.assertEqual(payload["method_tab"]["loaded_datasets"][0]["source_kind"], "input")
        self.assertEqual(payload["method_tab"]["loaded_datasets"][0]["origin_length"], 12)
        self.assertEqual(payload["method_tab"]["loaded_datasets"][0]["values"], [10.123457, 20.246914])
        self.assertEqual(payload["method_tab"]["loaded_datasets"][0]["weights"], [1.987654, 1.987654])
        self.assertEqual(payload["method_tab"]["calculated_ultimate"], [10.123457, 20.246914])
        self.assertEqual(payload["method_tab"]["selected_ultimate"], [10.123457, 200.246914])
        self.assertEqual(payload["method_tab"]["ratio_basis_values"], [{
            "name": "Earned Premium",
            "values": [1000.123457, 2000.246914],
        }])
        self.assertEqual(payload["method_tab"]["ultimate_overrides"], [None, 200.246914])

    def test_write_result_selection_export_uses_simplified_method_filename(self) -> None:
        payload = {
            "json_format": self.module.RS_JSON_FORMAT,
            "details_tab": {"name": "C 91 - Current Qtr Indicated"},
            "method_tab": {},
            "_sidecar_notes": "not method JSON",
        }

        path = self.module.write_result_selection_export(
            payload,
            r"PRNJ - PA\PA\All States\Direct Group\COL",
            self.rc_dir,
        )

        self.assertEqual(path.name, "RS@C 91 - Current Qtr Indicated.json")
        self.assertTrue(path.exists())
        self.assertEqual(path.parent, self.methods_dir)
        self.assertNotIn("_sidecar_notes", json.loads(path.read_text(encoding="utf-8")))

    def test_cleanup_target_reserving_class_dir_removes_existing_target_files(self) -> None:
        nested = self.datasets_dir / "nested"
        nested.mkdir()
        (self.datasets_dir / "old.csv").write_text("1\n", encoding="utf-8")
        (self.methods_dir / "old.json").write_text("{}", encoding="utf-8")
        (nested / "old-sidecar.json").write_text("{}", encoding="utf-8")
        lock_path = self.rc_dir / f".{self.module.INDEX_FILE_NAME}.lock"
        lock_path.write_bytes(b"\0")

        files, dirs = self.module.cleanup_target_reserving_class_dir(self.rc_dir)

        self.assertGreaterEqual(files, 3)
        self.assertGreaterEqual(dirs, 4)
        self.assertTrue(self.rc_dir.exists())
        self.assertEqual(list(self.rc_dir.iterdir()), [lock_path])

    def test_cleanup_target_reserving_class_dir_rejects_project_data_dir(self) -> None:
        with self.assertRaises(ValueError):
            self.module.cleanup_target_reserving_class_dir(self.project_dir / "data")

    def test_cleanup_target_dataset_artifacts_removes_selected_dataset_files(self) -> None:
        files = [
            self.datasets_dir / "Selected@12@12@cum@dev.csv",
            self.datasets_dir / "Selected@3.csv",
            self.sidecars_dir / "Selected.json",
            self.methods_dir / "DFM@Selected Method.json",
            self.methods_dir / "RS@Selected.json",
        ]
        for path in files:
            path.write_text("{}", encoding="utf-8")
        (self.methods_dir / "DFM@Output Lookup.json").write_text(json.dumps({
            "details_tab": {"name": "Output Lookup", "output_dataset": "Selected"},
        }), encoding="utf-8")
        kept = [
            self.datasets_dir / "Other@12@12@cum@dev.csv",
            self.sidecars_dir / "Other.json",
            self.methods_dir / "DFM@Other Method.json",
        ]
        for path in kept:
            path.write_text("{}", encoding="utf-8")

        removed, dirs = self.module.cleanup_target_dataset_artifacts(
            self.rc_dir,
            dataset_names=["Selected"],
            method_names=["Selected Method"],
        )

        self.assertEqual(dirs, 0)
        self.assertEqual(removed, 6)
        for path in files:
            self.assertFalse(path.exists(), path.name)
        self.assertFalse((self.methods_dir / "DFM@Output Lookup.json").exists())
        for path in kept:
            self.assertTrue(path.exists(), path.name)

    def test_selective_sync_cleanup_does_not_match_dataset_name_to_method_filename(self) -> None:
        dataset = self.datasets_dir / "Paid DFM@12.csv"
        sidecar = self.sidecars_dir / "Paid DFM.json"
        unrelated_method = self.methods_dir / "DFM@Paid DFM.json"
        dataset.write_text("1\n", encoding="utf-8")
        sidecar.write_text("{}", encoding="utf-8")
        unrelated_method.write_text(json.dumps({
            "details_tab": {"name": "Paid DFM", "output_dataset": "Paid Ultimate"},
        }), encoding="utf-8")

        removed, _dirs = self.module.cleanup_target_dataset_artifacts(
            self.rc_dir,
            dataset_names=["Paid DFM"],
            match_method_dependencies=False,
        )

        self.assertEqual(removed, 2)
        self.assertFalse(dataset.exists())
        self.assertFalse(sidecar.exists())
        self.assertTrue(unrelated_method.exists())

    def test_merge_preserves_arcrho_only_and_newer_logical_groups(self) -> None:
        live_rc = self.project_dir / "data" / "live"
        staged_rc = self.project_dir / "data" / "stage"
        for rc_dir in (live_rc, staged_rc):
            for folder in ("datasets", "methods", "sidecars"):
                (rc_dir / folder).mkdir(parents=True, exist_ok=True)

        def write_dataset(rc_dir, name, dataset_type, updated_at, value):
            csv_name = f"{name}@12.csv"
            (rc_dir / "datasets" / csv_name).write_text(f"{value}\n", encoding="utf-8")
            (rc_dir / "sidecars" / f"{name}.json").write_text(json.dumps({
                "dataset_name": name,
                "dataset_type": dataset_type,
                "csv_file": csv_name,
                "updated_at": updated_at,
            }), encoding="utf-8")

        write_dataset(
            live_rc,
            "ArcRho Scenario",
            "Net Ultimate",
            "2026-07-01T00:00:00Z",
            "local-only",
        )
        write_dataset(
            live_rc,
            "Paid Loss",
            "Paid Loss",
            "2026-08-03T00:00:00Z",
            "local-newer",
        )
        write_dataset(
            staged_rc,
            "Paid Loss",
            "Paid Loss",
            "2026-08-02T00:00:00Z",
            "resq-older",
        )
        (staged_rc / "datasets" / "Paid Loss@3.csv").write_text(
            "stale-stage-variant\n",
            encoding="utf-8",
        )
        write_dataset(
            live_rc,
            "DFM Ultimate",
            "DFM Ultimate",
            "2026-08-01T00:00:00Z",
            "local-older-dataset",
        )
        write_dataset(
            staged_rc,
            "DFM Ultimate",
            "DFM Ultimate",
            "2026-08-02T00:00:00Z",
            "resq-newer-dataset",
        )
        write_dataset(
            live_rc,
            "Older ResQ Dataset",
            "Paid Loss",
            "2026-08-01T00:00:00Z",
            "local-older",
        )
        write_dataset(
            staged_rc,
            "Older ResQ Dataset",
            "Paid Loss",
            "2026-08-02T00:00:00Z",
            "resq-newer",
        )
        live_method = live_rc / "methods" / "DFM@Selected DFM.json"
        live_method.write_text(json.dumps({
            "json_format": self.module.DFM_JSON_FORMAT,
            "details_tab": {
                "name": "Selected DFM",
                "output_dataset": "DFM Ultimate",
                "output_type": "DFM Ultimate",
            },
            "method_metadata": {"last_modified": "2026-08-04T00:00:00Z"},
        }), encoding="utf-8")
        staged_method = staged_rc / "methods" / live_method.name
        staged_method.write_text(json.dumps({
            "json_format": self.module.DFM_JSON_FORMAT,
            "details_tab": {
                "name": "Selected DFM",
                "output_dataset": "DFM Ultimate",
                "output_type": "DFM Ultimate",
            },
            "method_metadata": {"last_modified": "2026-08-03T00:00:00Z"},
        }), encoding="utf-8")

        snapshot_rc = self.project_dir / "snapshots" / "live"
        snapshot_files = self.module.snapshot_reserving_class_artifacts(live_rc, snapshot_rc)
        result = self.module.merge_preserved_arcrho_artifacts(snapshot_rc, staged_rc)

        self.assertEqual(snapshot_files, 9)
        self.assertEqual(
            set(result["names"]),
            {"ArcRho Scenario", "Paid Loss", "DFM Ultimate"},
        )
        self.assertEqual(result["groups"], 3)
        self.assertEqual(
            (staged_rc / "datasets" / "ArcRho Scenario@12.csv").read_text(encoding="utf-8"),
            "local-only\n",
        )
        self.assertEqual(
            (staged_rc / "datasets" / "Paid Loss@12.csv").read_text(encoding="utf-8"),
            "local-newer\n",
        )
        self.assertFalse((staged_rc / "datasets" / "Paid Loss@3.csv").exists())
        self.assertEqual(
            (staged_rc / "datasets" / "DFM Ultimate@12.csv").read_text(encoding="utf-8"),
            "local-older-dataset\n",
        )
        self.assertEqual(staged_method.read_bytes(), live_method.read_bytes())
        self.assertEqual(
            (staged_rc / "datasets" / "Older ResQ Dataset@12.csv").read_text(encoding="utf-8"),
            "resq-newer\n",
        )

    def test_overwrite_merge_keeps_only_arcrho_only_groups(self) -> None:
        """Overwrite drops the newer-live protection but never ArcRho-only work."""

        live_rc = self.project_dir / "data" / "live-overwrite"
        staged_rc = self.project_dir / "data" / "stage-overwrite"
        for rc_dir in (live_rc, staged_rc):
            for folder in ("datasets", "methods", "sidecars"):
                (rc_dir / folder).mkdir(parents=True, exist_ok=True)

        def write_dataset(rc_dir, name, dataset_type, updated_at, value):
            csv_name = f"{name}@12.csv"
            (rc_dir / "datasets" / csv_name).write_text(f"{value}\n", encoding="utf-8")
            (rc_dir / "sidecars" / f"{name}.json").write_text(json.dumps({
                "dataset_name": name,
                "dataset_type": dataset_type,
                "csv_file": csv_name,
                "updated_at": updated_at,
            }), encoding="utf-8")

        write_dataset(
            live_rc,
            "ArcRho Scenario",
            "Net Ultimate",
            "2026-07-01T00:00:00Z",
            "local-only",
        )
        write_dataset(
            live_rc,
            "Paid Loss",
            "Paid Loss",
            "2026-08-03T00:00:00Z",
            "local-newer",
        )
        write_dataset(
            staged_rc,
            "Paid Loss",
            "Paid Loss",
            "2026-08-02T00:00:00Z",
            "resq-older",
        )

        snapshot_rc = self.project_dir / "snapshots" / "live-overwrite"
        self.module.snapshot_reserving_class_artifacts(live_rc, snapshot_rc)
        result = self.module.merge_preserved_arcrho_artifacts(
            snapshot_rc, staged_rc, overwrite=True
        )

        self.assertEqual(set(result["names"]), {"ArcRho Scenario"})
        self.assertEqual(result["groups"], 1)
        self.assertEqual(
            (staged_rc / "datasets" / "ArcRho Scenario@12.csv").read_text(encoding="utf-8"),
            "local-only\n",
        )
        # The fresh ResQ copy wins even though the live copy was newer.
        self.assertEqual(
            (staged_rc / "datasets" / "Paid Loss@12.csv").read_text(encoding="utf-8"),
            "resq-older\n",
        )

    def test_cleanup_target_flag_defaults_on_and_can_be_disabled(self) -> None:
        self.assertTrue(self.module._parse_args([]).cleanup_target)
        self.assertFalse(self.module._parse_args(["--no-cleanup-target"]).cleanup_target)
        self.assertTrue(self.module._parse_args(["--cleanup-target"]).cleanup_target)

    def test_default_rc_path_list_is_hardcoded_from_resq_path_workbook(self) -> None:
        self.assertEqual(len(self.module.RC_PATH), 17)
        self.assertEqual(self.module.RC_PATH[0], r"PRNJ - PA\PA\NY\Direct Group\BI Total")
        self.assertEqual(self.module.RC_PATH[-1], r"PRNJ - PA\PA\MA\Direct Group\MP+PIP")

    def test_configured_rc_paths_accepts_string_or_list(self) -> None:
        self.assertEqual(self.module._configured_rc_paths(r"Auto\PP"), [r"Auto\PP"])
        self.assertEqual(
            self.module._configured_rc_paths(["", r"Auto\PP", r"Auto\COL", r"auto\pp"]),
            [r"Auto\PP", r"Auto\COL"],
        )

    def test_resq_export_counts_use_triangle_and_vector_total(self) -> None:
        self_module = self.module

        class ResQItem:
            def __init__(self, name, method_type=0):
                self.Name = name
                self.MethodType = method_type

        class ReservingClass:
            def Triangles(self):
                return [ResQItem("Paid Loss"), ResQItem("Reported Loss")]

            def Vectors(self):
                return [
                    ResQItem("Selected Ultimate", self_module.METHOD_TYPE_RESULT_SELECTION_CODE),
                    ResQItem("Manual Ultimate", self_module.METHOD_TYPE_NONE_CODE),
                    ResQItem("DFM Output", self_module.METHOD_TYPE_DFM_CODE),
                ]

            def DFMMethods(self):
                return [ResQItem("Paid DFM"), ResQItem("Reported DFM")]

        original_dfm_names = list(self.module.DFM_NAMES)
        try:
            self.module.DFM_NAMES = []
            counts = self.module.resq_export_dataset_counts(
                ReservingClass(),
                run_triangles=True,
                run_vectors=True,
                run_dfms=True,
            )
        finally:
            self.module.DFM_NAMES = original_dfm_names

        self.assertEqual(counts["triangles"], 2)
        self.assertEqual(counts["vectors"], 3)
        self.assertEqual(counts["dfms"], 2)
        self.assertEqual(counts["methods"], 2)
        self.assertEqual(counts["total"], 5)
        self.assertEqual(counts["dfm_names"], ["Paid DFM", "Reported DFM"])

    def test_resq_triangle_inventory_reuses_com_items_and_method_types(self) -> None:
        self_module = self.module

        class ResQItem:
            def __init__(self, name, method_type):
                self.Name = name
                self.MethodType = method_type

        class TriangleCollection:
            def __init__(self):
                self.items = [
                    ResQItem("Paid Loss", self_module.METHOD_TYPE_NONE_CODE),
                    ResQItem("Adjusted Loss", self_module.METHOD_TYPE_BS_SR_CODE),
                ]
                self.item_calls = 0

            def __iter__(self):
                return iter(self.items)

            def Item(self, _name):
                self.item_calls += 1
                raise AssertionError("cached triangle COM objects should be reused")

        class ReservingClass:
            def __init__(self):
                self.triangles = TriangleCollection()

            def Triangles(self):
                return self.triangles

            def Vectors(self):
                return []

            def DFMMethods(self):
                return []

        reserving_class = ReservingClass()
        events = []
        counts = self.module.resq_export_dataset_counts(
            reserving_class,
            run_triangles=True,
            run_vectors=False,
            run_dfms=False,
            progress_callback=events.append,
        )

        self.assertEqual(counts["triangle_names"], ["Paid Loss", "Adjusted Loss"])
        self.assertEqual(counts["bssr_names"], ["Adjusted Loss"])
        self.assertEqual(reserving_class.triangles.item_calls, 0)
        self.assertEqual([event["event"] for event in events], ["inventory"] * 3)
        self.assertIs(
            counts["triangle_items"]["paid loss"],
            reserving_class.triangles.items[0],
        )

    def test_export_unknown_dataset_type_vector_is_skipped(self) -> None:
        self_module = self.module

        class DatasetType:
            Name = "BF Output"
            DataFormat = 1

        class Vector:
            Name = "BF Output"
            MethodType = 2
            OriginLength = 12
            DevelopmentLength = 12
            OriginCount = 1
            User = ""
            Created = ""
            Modified = ""
            Formula = ""

            def __init__(self):
                self.DatasetType = DatasetType()

            def Value(self, _index):
                return 123.0

            def OriginLabel(self, _index):
                return "2025"

        class VectorCollection:
            def __init__(self):
                self.items = {"BF Output": Vector()}

            def __iter__(self):
                return iter(self.items.values())

            def Item(self, name):
                return self.items[name]

        class ReservingClass:
            def __init__(self):
                self.collection = VectorCollection()

            def Vectors(self):
                return self.collection

        progress_state = {"completed": 0, "total": 1}
        events = []
        written, errors = self_module.export_vectors_for_rc(
            ReservingClass(),
            r"Auto\PP",
            self.rc_dir,
            progress_callback=events.append,
            progress_state=progress_state,
            vector_names=["BF Output"],
            verbose=False,
        )

        self.assertEqual((written, errors), (0, 0))
        self.assertEqual(progress_state, {"completed": 1, "total": 1, "skipped": 1})
        self.assertFalse((self.sidecars_dir / "BF Output.json").exists())
        self.assertFalse((self.datasets_dir / "BF Output@12.csv").exists())
        finished = [event for event in events if event.get("event") == "finish"]
        self.assertEqual(finished[-1]["status"], "skipped")

    def test_export_dfm_method_with_matching_vector_progress_tick(self) -> None:
        self_module = self.module

        class DatasetType:
            Name = "DFM Ultimate"
            DataFormat = 1

        class Vector:
            Name = "Ultimate"
            MethodType = self_module.METHOD_TYPE_DFM_CODE
            OriginLength = 12
            DevelopmentLength = 12
            OriginCount = 1

            def __init__(self):
                self.DatasetType = DatasetType()

        class VectorCollection:
            def __init__(self):
                self.items = {"Ultimate": Vector()}

            def __iter__(self):
                return iter(self.items.values())

            def Item(self, name):
                return self.items[name]

        class Dfm:
            Name = "Paid DFM"

            def __init__(self):
                self.OutputVector = Vector()

        class DfmCollection:
            def __init__(self):
                self.items = {"Paid DFM": Dfm()}

            def __iter__(self):
                return iter(self.items.values())

            def Item(self, name):
                return self.items[name]

        class ReservingClass:
            def __init__(self):
                self.vectors = VectorCollection()
                self.dfms = DfmCollection()

            def Vectors(self):
                return self.vectors

            def DFMMethods(self):
                return self.dfms

        def fake_export_dfm_output(_dfm, _rc_path, rc_dir, **_kwargs):
            (rc_dir / "datasets" / "Ultimate@12.csv").write_text("123\n", encoding="utf-8")
            (rc_dir / "methods" / "DFM@Paid DFM.json").write_text("{}\n", encoding="utf-8")
            (rc_dir / "sidecars" / "Ultimate.json").write_text(json.dumps({
                "source_kind": "dfm",
                "method_name": "Paid DFM",
                "period_length": 12,
            }), encoding="utf-8")
            return "Ultimate", "OK", False

        with mock.patch.object(
            self.module,
            "_export_dfm_output_dataset",
            side_effect=fake_export_dfm_output,
        ):
            progress_state = {"completed": 0, "total": 1}
            method_counts = {"dfms_written": 0}
            events = []

            written, errors = self_module.export_vectors_for_rc(
                ReservingClass(),
                r"Auto\PP",
                self.rc_dir,
                progress_callback=events.append,
                progress_state=progress_state,
                vector_names=["Ultimate"],
                include_dfm_methods=True,
                dfm_names=["Paid DFM"],
                method_counts=method_counts,
                verbose=False,
            )

        self.assertEqual((written, errors), (1, 0))
        self.assertEqual(progress_state, {"completed": 1, "total": 1})
        self.assertEqual(method_counts["dfms_written"], 1)
        self.assertTrue((self.datasets_dir / "Ultimate@12.csv").exists())
        self.assertTrue((self.methods_dir / "DFM@Paid DFM.json").exists())
        sidecar = json.loads((self.sidecars_dir / "Ultimate.json").read_text(encoding="utf-8"))
        self.assertEqual(sidecar["source_kind"], "dfm")
        self.assertEqual(sidecar["method_name"], "Paid DFM")
        self.assertEqual(sidecar["period_length"], 12)
        self.assertNotIn("origin_length", sidecar)
        self.assertNotIn("development_length", sidecar)
        self.assertNotIn("development_count", sidecar)
        self.assertNotIn("cumulative", sidecar)
        self.assertNotIn("calendar", sidecar)
        self.assertFalse([event for event in events if event.get("event") == "method"])
        finished = [event for event in events if event.get("event") == "finish"]
        self.assertEqual(finished[-1]["name"], "Ultimate")
        self.assertEqual(finished[-1]["completed"], 1)

    def test_export_dfms_do_not_advance_shared_dataset_progress(self) -> None:
        class Dfm:
            def __init__(self, name):
                self.Name = name

        class DfmCollection:
            def __init__(self):
                self.items = {name: Dfm(name) for name in ("Paid DFM", "Reported DFM")}

            def __iter__(self):
                return iter(self.items.values())

            def Item(self, name):
                return self.items[name]

        class ReservingClass:
            def __init__(self):
                self.collection = DfmCollection()

            def DFMMethods(self):
                return self.collection

        def fake_export_dfm_output(dfm, _rc_path, _rc_dir, **_kwargs):
            return f"{dfm.Name} Ultimate", "OK", False

        with mock.patch.object(
            self.module,
            "_export_dfm_output_dataset",
            side_effect=fake_export_dfm_output,
        ):
            progress_state = {"completed": 4, "total": 4, "count_methods": False}
            events = []

            written, errors = self.module.export_dfms_for_rc(
                ReservingClass(),
                r"Auto\PP",
                self.rc_dir,
                progress_callback=events.append,
                progress_state=progress_state,
                dfm_names=["Paid DFM", "Reported DFM"],
                verbose=False,
            )

        self.assertEqual((written, errors), (2, 0))
        self.assertEqual(progress_state, {"completed": 4, "total": 4, "count_methods": False})
        finished = [event for event in events if event.get("event") == "method" and event.get("status") == "success"]
        self.assertEqual([event["completed"] for event in finished], [4, 4])
        self.assertEqual([event["total"] for event in finished], [4, 4])
        self.assertEqual(
            [event["dataset_name"] for event in finished],
            ["Paid DFM Ultimate", "Reported DFM Ultimate"],
        )


if __name__ == "__main__":
    unittest.main()

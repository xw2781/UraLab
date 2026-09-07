"""The save rule that keeps a hand-entered dataset's stored shape put.

Step 6 of ``docs/plans/manual_input_period_rollup.md``: the lengths a save
carries are the shape the dataset is displayed at, so an ``input`` sidecar's
stored shape and its CSV survive a display-only save, a save that carries
values at any other shape is refused, and only a dataset whose file holds
nothing is relabelled to the shape asked for.

Step 3 of ``docs/plans/completed/manual_input_stored_length_resq_alignment.md`` relaxes
that refusal on the development axis alone: values entered at a coarser
development view are scattered into the stored cells at their valuation dates.
"""

from __future__ import annotations

import copy
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import numpy as np
import pandas as pd


FRONTEND_ROOT = Path(__file__).resolve().parents[1]
if str(FRONTEND_ROOT) not in sys.path:
    sys.path.insert(0, str(FRONTEND_ROOT))

TEST_TEMP_ROOT = Path(__file__).resolve().parents[2] / "test"
TEST_TEMP_ROOT.mkdir(parents=True, exist_ok=True)

from fastapi import HTTPException

from app_server.services import calculated_dataset_service, dataset_service
from dependent_propagation_workspace_stub import IsolatedPropagationWorkspace


MONTHLY_CSV = "Dataset@1@1@cum@dev.csv"
ANNUAL_CSV = "Dataset@12@12@cum@dev.csv"
ANNUAL_OVER_MONTHLY_CSV = "Dataset@12@1@cum@dev.csv"
MONTHLY_VECTOR_CSV = "Dataset@1.csv"
ANNUAL_VECTOR_CSV = "Dataset@12.csv"


class ManualDatasetStoredShapeSaveTests(unittest.TestCase):
    def setUp(self) -> None:
        self.propagation_workspace = IsolatedPropagationWorkspace().start()
        self.temp = tempfile.TemporaryDirectory(dir=TEST_TEMP_ROOT)
        self.data_dir = self.temp.name
        self.sidecar_path = os.path.join(self.data_dir, "Dataset.json")
        self.existing = {
            "dataset_name": "Dataset",
            "dataset_type": "Input Type",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "input",
            "data_format": "Triangle",
            "origin_length": 1,
            "development_length": 1,
            "stored_origin_length": 1,
            "stored_development_length": 1,
            "cumulative": True,
            "calendar": False,
            "csv_file": MONTHLY_CSV,
        }

    def tearDown(self) -> None:
        self.temp.cleanup()
        self.propagation_workspace.stop()

    def _probe_general_settings(self):
        """Point the save at the ResQ probe project's dates.

        The scatter needs the project's Origin Start Date and Development End
        Date to know what each column is valued at, so the save reads real
        General Settings rather than a patched geometry.
        """
        settings_path = os.path.join(self.data_dir, "general_settings.json")
        with open(settings_path, "w", encoding="utf-8") as fh:
            fh.write(
                '{"origin_start_date":"201701","origin_end_date":"202601",'
                '"development_end_date":"202605"}'
            )
        return patch.object(
            dataset_service.config, "get_general_settings_path", return_value=settings_path
        )

    def _write_stored_csv(self, name: str, values) -> str:
        path = os.path.join(self.data_dir, name)
        pd.DataFrame(values).to_csv(path, header=False, index=False)
        return path

    def _save(
        self,
        *,
        origin_length: int,
        development_length: int,
        values=None,
        stored_development_length: int | None = None,
        data_format: str = "Triangle",
    ):
        written: dict = {}

        def capture_write(path, payload):
            written["path"] = path
            written["payload"] = copy.deepcopy(payload)

        with (
            patch.object(dataset_service, "_get_dataset_sidecar_path", return_value=self.sidecar_path),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value=copy.deepcopy(self.existing)),
            patch.object(dataset_service, "_write_dataset_sidecar_payload", side_effect=capture_write),
            patch.object(dataset_service, "_is_app_calculated_dataset_type", return_value=(False, "")),
            patch.object(
                dataset_service.config,
                "get_project_dataset_cache_dir",
                return_value=self.data_dir,
            ),
            patch.object(dataset_service.dataset_instance_index_service, "rebuild_index"),
            patch.object(
                dataset_service.dataset_sidecar_status_service,
                "refresh_method_statuses_for_dependents",
                return_value=[],
            ),
            patch.object(calculated_dataset_service, "apply_sidecar_graph_fields"),
            patch.object(
                calculated_dataset_service,
                "recalculate_dependents",
                return_value={"ok": True, "steps": []},
            ),
        ):
            result = dataset_service.save_dataset_sidecar(
                "Project",
                "Class",
                "Dataset",
                dataset_type="Input Type",
                source_kind="input",
                data_format=data_format,
                origin_length=origin_length,
                development_length=development_length,
                stored_development_length=stored_development_length,
                values=values,
            )

        return result, written.get("payload")

    def test_display_only_save_keeps_the_stored_shape_and_the_csv(self) -> None:
        monthly_path = self._write_stored_csv(MONTHLY_CSV, [[100.0, 110.0], [120.0, np.nan]])

        result, payload = self._save(origin_length=12, development_length=12)

        self.assertEqual(payload["origin_length"], 12)
        self.assertEqual(payload["development_length"], 12)
        self.assertEqual(payload["stored_origin_length"], 1)
        self.assertEqual(payload["stored_development_length"], 1)
        self.assertEqual(payload["csv_file"], MONTHLY_CSV)
        self.assertEqual(result["origin_length"], 12)
        # The window that just saved is told the stored shape as well as the
        # display one, so its readout and its read-only rule need no reload.
        self.assertEqual(result["stored_origin_length"], 1)
        self.assertEqual(result["stored_development_length"], 1)
        self.assertTrue(os.path.exists(monthly_path))
        self.assertFalse(os.path.exists(os.path.join(self.data_dir, ANNUAL_CSV)))

    def test_values_save_at_a_coarser_origin_is_refused(self) -> None:
        self._write_stored_csv(MONTHLY_CSV, [[100.0, 110.0], [120.0, np.nan]])

        with self.assertRaises(HTTPException) as raised:
            self._save(origin_length=12, development_length=12, values=[[230.0]])

        self.assertEqual(raised.exception.status_code, 400)
        self.assertEqual(
            raised.exception.detail,
            "Values can be entered only at the stored origin period.",
        )
        self.assertFalse(os.path.exists(os.path.join(self.data_dir, ANNUAL_CSV)))

    def test_a_vector_save_at_another_period_is_refused(self) -> None:
        self.existing = {
            "dataset_name": "Dataset",
            "dataset_type": "Input Type",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "input",
            "data_format": "Vector",
            "period_length": 1,
            "stored_period_length": 1,
            "csv_file": MONTHLY_VECTOR_CSV,
        }
        self._write_stored_csv(MONTHLY_VECTOR_CSV, [[100.0], [120.0]])

        with self.assertRaises(HTTPException) as raised:
            self._save(
                origin_length=12,
                development_length=12,
                values=[[230.0]],
                data_format="Vector",
            )

        self.assertEqual(raised.exception.status_code, 422)
        self.assertIn("stores its values at", raised.exception.detail)

    def test_values_at_a_coarser_development_view_land_in_the_stored_cells(self) -> None:
        # The ResQ probe's project: annual origins from 2017-01 valued on
        # 2026-05-31, so a monthly store is 113 columns wide and its annual
        # view is valued at 5, 17, 29, ... 113 months.
        self.existing = {
            **self.existing,
            "origin_length": 12,
            "development_length": 12,
            "stored_origin_length": 12,
            "stored_development_length": 1,
            "csv_file": ANNUAL_OVER_MONTHLY_CSV,
        }
        self._write_stored_csv(ANNUAL_OVER_MONTHLY_CSV, [[100.0]])
        annual = [
            [1000.0 * row + column if column < 10 - row else None for column in range(10)]
            for row in range(10)
        ]

        with self._probe_general_settings():
            result, payload = self._save(
                origin_length=12, development_length=12, values=annual
            )

        self.assertEqual(payload["origin_length"], 12)
        self.assertEqual(payload["development_length"], 12)
        self.assertEqual(payload["stored_origin_length"], 12)
        self.assertEqual(payload["stored_development_length"], 1)
        self.assertEqual(payload["csv_file"], ANNUAL_OVER_MONTHLY_CSV)
        self.assertEqual(result["stored_development_length"], 1)

        written = dataset_service.load_triangle_values(
            os.path.join(self.data_dir, ANNUAL_OVER_MONTHLY_CSV)
        )
        self.assertEqual(written.shape, (10, 113))
        entered = {4 + 12 * column: float(column) for column in range(10)}
        for column, value in enumerate(written.iloc[0].tolist()):
            self.assertEqual(value, entered.get(column, 0.0), f"{column + 1} months")
        newest = written.iloc[9].tolist()
        self.assertEqual(newest[4], 9000.0)
        self.assertTrue(all(bool(pd.isna(value)) for value in newest[5:]))

    def test_values_save_at_the_stored_shape_is_written(self) -> None:
        self._write_stored_csv(MONTHLY_CSV, [[100.0, 110.0], [120.0, np.nan]])

        _, payload = self._save(
            origin_length=1,
            development_length=1,
            values=[[130.0, 140.0], [150.0, None]],
        )

        self.assertEqual(payload["stored_origin_length"], 1)
        self.assertEqual(payload["csv_file"], MONTHLY_CSV)
        written = dataset_service.load_triangle_values(os.path.join(self.data_dir, MONTHLY_CSV))
        self.assertEqual(written.iat[0, 0], 130.0)

    def test_empty_dataset_is_relabelled_and_its_old_csv_deleted(self) -> None:
        monthly_path = self._write_stored_csv(MONTHLY_CSV, [[0.0, 0.0], [0.0, np.nan]])

        result, payload = self._save(
            origin_length=12,
            development_length=12,
            values=[[0.0, 0.0], [0.0, None]],
        )

        self.assertEqual(payload["stored_origin_length"], 12)
        self.assertEqual(payload["stored_development_length"], 12)
        self.assertEqual(payload["csv_file"], ANNUAL_CSV)
        # A relabel moves the stored shape, so the answer carries the new one.
        self.assertEqual(result["stored_origin_length"], 12)
        self.assertEqual(result["stored_development_length"], 12)
        self.assertTrue(os.path.exists(os.path.join(self.data_dir, ANNUAL_CSV)))
        self.assertFalse(os.path.exists(monthly_path))

    def test_empty_dataset_relabel_without_values_rebuilds_the_grid(self) -> None:
        monthly_path = self._write_stored_csv(MONTHLY_CSV, [[0.0, 0.0], [0.0, np.nan]])

        with patch.object(
            dataset_service,
            "_empty_dataset_geometry_from_general_settings",
            return_value=(2, 2, np.array([[True, True], [True, False]])),
        ):
            _, payload = self._save(origin_length=12, development_length=12)

        self.assertEqual(payload["stored_origin_length"], 12)
        self.assertEqual(payload["csv_file"], ANNUAL_CSV)
        self.assertFalse(os.path.exists(monthly_path))
        rebuilt = dataset_service.load_triangle_values(os.path.join(self.data_dir, ANNUAL_CSV))
        self.assertEqual(rebuilt.shape, (2, 2))
        self.assertTrue(bool(pd.isna(rebuilt.iat[1, 1])))

    def test_empty_triangle_can_be_stored_finer_than_it_is_shown(self) -> None:
        monthly_path = self._write_stored_csv(MONTHLY_CSV, [[0.0, 0.0], [0.0, np.nan]])

        with patch.object(
            dataset_service,
            "_empty_dataset_geometry_from_general_settings",
            return_value=(2, 3, np.array([[True, True, True], [True, True, False]])),
        ) as geometry:
            result, payload = self._save(
                origin_length=12,
                development_length=12,
                stored_development_length=1,
            )

        self.assertEqual(payload["origin_length"], 12)
        self.assertEqual(payload["development_length"], 12)
        self.assertEqual(payload["stored_origin_length"], 12)
        self.assertEqual(payload["stored_development_length"], 1)
        self.assertEqual(payload["csv_file"], ANNUAL_OVER_MONTHLY_CSV)
        self.assertEqual(result["stored_origin_length"], 12)
        self.assertEqual(result["stored_development_length"], 1)
        # The empty file that replaces the old one is built at the stored
        # shape, not at the coarser shape the dataset is shown at.
        self.assertEqual(geometry.call_args.args[1:], (12, 1))
        self.assertTrue(os.path.exists(os.path.join(self.data_dir, ANNUAL_OVER_MONTHLY_CSV)))
        self.assertFalse(os.path.exists(monthly_path))

    def test_stored_development_length_must_divide_the_display_length(self) -> None:
        self._write_stored_csv(MONTHLY_CSV, [[0.0, 0.0], [0.0, np.nan]])

        with self.assertRaises(HTTPException) as raised:
            self._save(origin_length=12, development_length=12, stored_development_length=5)

        self.assertEqual(raised.exception.status_code, 400)
        self.assertEqual(
            raised.exception.detail,
            "The stored development length must be a factor of the development length.",
        )

    def test_stored_development_length_cannot_move_once_values_are_stored(self) -> None:
        self._write_stored_csv(MONTHLY_CSV, [[100.0, 110.0], [120.0, np.nan]])

        with self.assertRaises(HTTPException) as raised:
            self._save(origin_length=12, development_length=12, stored_development_length=3)

        self.assertEqual(raised.exception.status_code, 400)
        self.assertEqual(
            raised.exception.detail,
            "The stored development length cannot be changed while the dataset holds values.",
        )

    def test_a_vector_ignores_the_stored_development_length(self) -> None:
        self.existing = {
            "dataset_name": "Dataset",
            "dataset_type": "Input Type",
            "project_name": "Project",
            "reserving_class": "Class",
            "source_kind": "input",
            "data_format": "Vector",
            "period_length": 1,
            "stored_period_length": 1,
            "csv_file": MONTHLY_VECTOR_CSV,
        }
        self._write_stored_csv(MONTHLY_VECTOR_CSV, [[0.0], [0.0]])

        with patch.object(
            dataset_service,
            "_empty_dataset_geometry_from_general_settings",
            return_value=(2, 1, np.array([[True], [True]])),
        ):
            result, payload = self._save(
                origin_length=12,
                development_length=12,
                stored_development_length=1,
                data_format="Vector",
            )

        self.assertEqual(payload["stored_period_length"], 12)
        self.assertNotIn("stored_development_length", payload)
        self.assertEqual(payload["csv_file"], ANNUAL_VECTOR_CSV)
        self.assertEqual(result["stored_period_length"], 12)
        self.assertEqual(result["stored_development_length"], 12)


if __name__ == "__main__":
    unittest.main()

"""Where a triangle's cells stop: the project's own calendar diagonal.

``dataset_service.triangle_grid_shape`` is the one answer to that question. The
Dataset window asks for it when it draws a hand-entered grid that has no file
behind it yet, and a cached load of a hand-entered triangle is masked with the
same geometry, so a figure a file holds past the diagonal is neither shown nor
editable.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd


FRONTEND_ROOT = Path(__file__).resolve().parents[1]
if str(FRONTEND_ROOT) not in sys.path:
    sys.path.insert(0, str(FRONTEND_ROOT))

TEST_TEMP_ROOT = Path(__file__).resolve().parents[2] / "test"
TEST_TEMP_ROOT.mkdir(parents=True, exist_ok=True)

from fastapi import HTTPException

from app_server.services import dataset_service


class TriangleGridShapeTests(unittest.TestCase):
    """Origins from 2017-01 through 2026-12, valued on 2026-05: 113 months."""

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=TEST_TEMP_ROOT)
        self.settings_path = os.path.join(self.temp.name, "general_settings.json")
        with open(self.settings_path, "w", encoding="utf-8") as handle:
            handle.write(
                '{"origin_start_date":"201701","origin_end_date":"202612",'
                '"development_end_date":"202605"}'
            )

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _shape(self, origin_length: int, development_length: int) -> dict:
        with patch.object(
            dataset_service.config,
            "get_general_settings_path",
            return_value=self.settings_path,
        ):
            return dataset_service.triangle_grid_shape(
                "Project", origin_length, development_length
            )

    def test_equal_periods_step_one_column_a_row(self) -> None:
        shape = self._shape(12, 12)

        self.assertEqual(shape["origin_count"], 10)
        self.assertEqual(shape["development_count"], 10)
        self.assertEqual([sum(row) for row in shape["mask"]], list(range(10, 0, -1)))

    def test_a_quarterly_development_of_annual_origins_steps_four_columns(self) -> None:
        # The row below loses a year of quarterly columns, not one column: the
        # cell that ends the row is the one valued on the Development End Date.
        shape = self._shape(12, 3)

        self.assertEqual(shape["origin_count"], 10)
        self.assertEqual(shape["development_count"], 38)
        self.assertEqual([sum(row) for row in shape["mask"]], [38, 34, 30, 26, 22, 18, 14, 10, 6, 2])

    def test_quarterly_origins_of_an_annual_development_keep_their_rows(self) -> None:
        # Every quarter that starts on or before the valuation date is a row
        # with cells, far past the ten columns an annual development has.
        shape = self._shape(3, 12)

        self.assertEqual(shape["origin_count"], 40)
        self.assertEqual(shape["development_count"], 10)
        counts = [sum(row) for row in shape["mask"]]
        self.assertEqual(counts[:4], [10, 10, 9, 9])
        self.assertEqual(counts[37], 1)
        # The last two quarters of 2026 start after the valuation date.
        self.assertEqual(counts[38:], [0, 0])

    def test_a_project_without_dates_is_refused(self) -> None:
        with patch.object(
            dataset_service.config,
            "get_general_settings_path",
            return_value=os.path.join(self.temp.name, "missing.json"),
        ):
            with self.assertRaises(HTTPException) as raised:
                dataset_service.triangle_grid_shape("Project", 12, 12)

        self.assertEqual(raised.exception.status_code, 422)


class HandEnteredLoadMaskTests(unittest.TestCase):
    """A cached load of a hand-entered triangle carries the diagonal with it."""

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=TEST_TEMP_ROOT)
        self.data_dir = self.temp.name
        self.settings_path = os.path.join(self.data_dir, "general_settings.json")
        with open(self.settings_path, "w", encoding="utf-8") as handle:
            handle.write(
                '{"origin_start_date":"201701","origin_end_date":"201812",'
                '"development_end_date":"201812"}'
            )
        self.csv_path = os.path.join(self.data_dir, "Dataset@12@12@cum@dev.csv")
        # Two annual origins valued on 2018-12, and a third figure in the cell
        # past the diagonal that no triangle at this shape has.
        pd.DataFrame([[100.0, 110.0], [120.0, 130.0]]).to_csv(
            self.csv_path, header=False, index=False
        )
        self.sidecar = {
            "dataset_name": "Dataset",
            "dataset_type": "Input Type",
            "source_kind": "input",
            "data_format": "Triangle",
            "origin_length": 12,
            "development_length": 12,
            "stored_origin_length": 12,
            "stored_development_length": 12,
            "origin_labels": ["2017", "2018"],
            "development_labels": ["12", "24"],
            "csv_file": os.path.basename(self.csv_path),
        }

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _load(self) -> dict:
        with (
            patch.object(
                dataset_service, "_get_dataset_sidecar_path", return_value="sidecar.json"
            ),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value=self.sidecar),
            patch.object(
                dataset_service.config,
                "get_project_dataset_cache_dir",
                return_value=self.data_dir,
            ),
            patch.object(
                dataset_service.config,
                "get_general_settings_path",
                return_value=self.settings_path,
            ),
            patch.object(
                dataset_service, "_is_app_calculated_dataset_type", return_value=(False, "")
            ),
        ):
            return dataset_service.load_cached_dataset_values(
                "Project", "Class", "Dataset", csv_file=self.sidecar["csv_file"]
            )

    def test_a_figure_past_the_diagonal_is_masked_out(self) -> None:
        payload = self._load()

        self.assertEqual(payload["mask"], [[True, True], [True, False]])
        # The figure itself is left in the values the load returns; the mask is
        # what the grid, the totals, and the save all read.
        self.assertEqual(payload["values"][1][1], 130.0)

    def test_a_generated_dataset_keeps_the_mask_its_file_gives_it(self) -> None:
        self.sidecar["source_kind"] = "engine"

        self.assertEqual(self._load()["mask"], [[True, True], [True, True]])


if __name__ == "__main__":
    unittest.main()

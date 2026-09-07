from __future__ import annotations

import sys
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch
import json

from fastapi import HTTPException


FRONTEND_ROOT = Path(__file__).resolve().parents[1]
TEST_TEMP_ROOT = FRONTEND_ROOT.parent / "test"
TEST_TEMP_ROOT.mkdir(parents=True, exist_ok=True)
if str(FRONTEND_ROOT) not in sys.path:
    sys.path.insert(0, str(FRONTEND_ROOT))

from app_server import config
from app_server.services import dataset_service

PYTHON_API_SRC = FRONTEND_ROOT.parent / "python-api" / "src"
if str(PYTHON_API_SRC) not in sys.path:
    sys.path.insert(0, str(PYTHON_API_SRC))

from arcrho_api.dataset_index_contract import migrate_legacy_notes_files


class DatasetCacheLoadCandidateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory(dir=str(TEST_TEMP_ROOT))
        self.cache_dir = Path(self.temp_dir.name) / config.DATASET_CACHE_DIR
        self.cache_dir.mkdir()
        self.csv_path = self.cache_dir / "Paid@12.csv"
        self.csv_path.write_text("1\n2\n", encoding="utf-8")
        self.sidecar = {
            "dataset_name": "Paid",
            "dataset_type": "Paid",
            "data_format": "Vector",
            "period_length": 12,
            "csv_file": self.csv_path.name,
            "source_kind": "engine",
        }

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_sidecar_csv_is_checked_before_directory_enumeration(self) -> None:
        with (
            patch.object(
                config,
                "get_project_dataset_cache_dir",
                return_value=str(self.cache_dir),
            ),
            patch.object(dataset_service, "_get_dataset_sidecar_path", return_value="sidecar.json"),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value=self.sidecar),
            patch.object(dataset_service.os, "listdir", side_effect=AssertionError("unexpected listdir")),
            patch.object(dataset_service, "_resolve_origin_labels", return_value=["2020", "2021"]) as resolve_labels,
        ):
            result = dataset_service.load_cached_dataset_values(
                "Example Project",
                "Example RC",
                "Paid",
                origin_length=12,
                development_length=12,
            )

        self.assertEqual(result["csv_file"], self.csv_path.name)
        self.assertEqual(result["values"], [[1], [2]])
        self.assertEqual(result["mask"], [[True], [True]])
        self.assertEqual(result["origin_labels"], ["2020", "2021"])
        resolve_labels.assert_called_once_with(
            result["id"],
            str(self.csv_path),
            "Example Project",
            12,
            2,
        )

    def test_valid_sidecar_origin_labels_keep_the_two_file_fast_path(self) -> None:
        sidecar = {**self.sidecar, "origin_labels": ["2020", "2021"]}
        with (
            patch.object(config, "get_project_dataset_cache_dir", return_value=str(self.cache_dir)),
            patch.object(dataset_service, "_get_dataset_sidecar_path", return_value="sidecar.json"),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value=sidecar),
            patch.object(dataset_service.os, "listdir", side_effect=AssertionError("unexpected listdir")),
            patch.object(dataset_service, "_resolve_origin_labels", side_effect=AssertionError("unexpected header lookup")),
        ):
            result = dataset_service.load_cached_dataset_values(
                "Example Project",
                "Example RC",
                "Paid",
                origin_length=12,
                development_length=12,
            )

        self.assertEqual(result["origin_labels"], ["2020", "2021"])
        self.assertEqual(result["values"], [[1], [2]])

    def test_mismatched_sidecar_origin_labels_use_authoritative_headers(self) -> None:
        sidecar = {**self.sidecar, "origin_labels": ["2020"]}
        with (
            patch.object(config, "get_project_dataset_cache_dir", return_value=str(self.cache_dir)),
            patch.object(dataset_service, "_get_dataset_sidecar_path", return_value="sidecar.json"),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value=sidecar),
            patch.object(dataset_service, "_resolve_origin_labels", return_value=["2020", "2021"]) as resolve_labels,
        ):
            result = dataset_service.load_cached_dataset_values(
                "Example Project",
                "Example RC",
                "Paid",
                origin_length=12,
                development_length=12,
            )

        self.assertEqual(result["origin_labels"], ["2020", "2021"])
        self.assertEqual(len(result["origin_labels"]), len(result["values"]))
        resolve_labels.assert_called_once()

    def test_engine_triangle_hydrates_canonical_development_labels_and_formula(self) -> None:
        csv_path = self.cache_dir / "Ratio@12@12@cum@dev.csv"
        csv_path.write_text("1,2\n3,\n", encoding="utf-8")
        sidecar = {
            "dataset_name": "Ratio",
            "dataset_type": "Ratio",
            "data_format": "Triangle",
            "origin_length": 12,
            "development_length": 12,
            "csv_file": csv_path.name,
            "source_kind": "engine",
            "calendar": False,
            "formula": "",
        }
        with (
            patch.object(config, "get_project_dataset_cache_dir", return_value=str(self.cache_dir)),
            patch.object(dataset_service, "_get_dataset_sidecar_path", return_value="sidecar.json"),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value=sidecar),
            patch.object(dataset_service, "_resolve_origin_labels", return_value=["2025", "2026"]),
            patch.object(
                dataset_service,
                "_resolve_development_labels",
                return_value=["5m", "17m"],
            ) as resolve_development,
            patch.object(
                dataset_service,
                "_is_app_calculated_dataset_type",
                return_value=(False, '"Paid" / "Reported"'),
            ),
        ):
            result = dataset_service.load_cached_dataset_values(
                "Example Project",
                "Example RC",
                "Ratio",
                origin_length=12,
                development_length=12,
            )

        self.assertEqual(result["dev_labels"], ["5m", "17m"])
        self.assertEqual(result["formula"], '"Paid" / "Reported"')
        resolve_development.assert_called_once_with(
            result["id"],
            str(csv_path),
            "Example Project",
            12,
            2,
            calendar=False,
        )

    def test_label_and_formula_hydration_runs_concurrently(self) -> None:
        csv_path = self.cache_dir / "Ratio@12@12@cum@dev.csv"
        csv_path.write_text("1,2\n3,\n", encoding="utf-8")
        sidecar = {
            "dataset_name": "Ratio",
            "dataset_type": "Ratio",
            "data_format": "Triangle",
            "origin_length": 12,
            "development_length": 12,
            "csv_file": csv_path.name,
            "source_kind": "engine",
            "calendar": False,
        }
        # Each hydration path blocks until every other path has started; a
        # sequential implementation would deadlock the barrier and fail fast.
        barrier = threading.Barrier(3, timeout=5)

        def origin_labels(*_args, **_kwargs):
            barrier.wait()
            return ["2025", "2026"]

        def development_labels(*_args, **_kwargs):
            barrier.wait()
            return ["12", "24"]

        def formula_lookup(*_args, **_kwargs):
            barrier.wait()
            return (False, "")

        with (
            patch.object(config, "get_project_dataset_cache_dir", return_value=str(self.cache_dir)),
            patch.object(dataset_service, "_get_dataset_sidecar_path", return_value="sidecar.json"),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value=sidecar),
            patch.object(dataset_service, "_resolve_origin_labels", side_effect=origin_labels),
            patch.object(dataset_service, "_resolve_development_labels", side_effect=development_labels),
            patch.object(dataset_service, "_is_app_calculated_dataset_type", side_effect=formula_lookup),
        ):
            result = dataset_service.load_cached_dataset_values(
                "Example Project",
                "Example RC",
                "Ratio",
                origin_length=12,
                development_length=12,
            )

        self.assertEqual(result["origin_labels"], ["2025", "2026"])
        self.assertEqual(result["dev_labels"], ["12", "24"])

    def test_origin_label_failures_keep_precedence_over_development_failures(self) -> None:
        csv_path = self.cache_dir / "Ratio@12@12@cum@dev.csv"
        csv_path.write_text("1,2\n3,\n", encoding="utf-8")
        sidecar = {
            "dataset_name": "Ratio",
            "dataset_type": "Ratio",
            "data_format": "Triangle",
            "origin_length": 12,
            "development_length": 12,
            "csv_file": csv_path.name,
            "source_kind": "engine",
        }

        def raise_origin(*_args, **_kwargs):
            raise HTTPException(422, "origin failure detail")

        def raise_development(*_args, **_kwargs):
            raise HTTPException(422, "development failure detail")

        with (
            patch.object(config, "get_project_dataset_cache_dir", return_value=str(self.cache_dir)),
            patch.object(dataset_service, "_get_dataset_sidecar_path", return_value="sidecar.json"),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value=sidecar),
            patch.object(dataset_service, "_resolve_origin_labels", side_effect=raise_origin),
            patch.object(dataset_service, "_resolve_development_labels", side_effect=raise_development),
            patch.object(dataset_service, "_is_app_calculated_dataset_type", return_value=(False, "")),
        ):
            with self.assertRaises(HTTPException) as raised:
                dataset_service.load_cached_dataset_values(
                    "Example Project",
                    "Example RC",
                    "Ratio",
                    origin_length=12,
                    development_length=12,
                )

        self.assertEqual(raised.exception.detail, "origin failure detail")

    def test_notes_are_updated_in_the_dataset_sidecar_only(self) -> None:
        sidecar_path = Path(self.temp_dir.name) / "sidecars" / "Paid.json"
        sidecar_path.parent.mkdir()
        sidecar_path.write_text(json.dumps({**self.sidecar, "notes": "before"}), encoding="utf-8")

        with patch.object(dataset_service, "_get_dataset_sidecar_path", return_value=str(sidecar_path)):
            result = dataset_service.save_dataset_notes("Example Project", "Example RC", "Paid", "after")

        self.assertEqual(result["path"], str(sidecar_path))
        self.assertEqual(json.loads(sidecar_path.read_text(encoding="utf-8"))["notes"], "after")
        self.assertEqual(list(sidecar_path.parent.glob("ArcRhoTriNotes@*.json")), [])

    def test_legacy_notes_file_is_migrated_to_the_owning_sidecar(self) -> None:
        rc_dir = Path(self.temp_dir.name) / "reserving-class"
        sidecar_dir = rc_dir / "sidecars"
        sidecar_dir.mkdir(parents=True)
        sidecar_path = sidecar_dir / "Paid.json"
        legacy_path = sidecar_dir / "ArcRhoTriNotes@Paid.json"
        sidecar_path.write_text(json.dumps(self.sidecar), encoding="utf-8")
        legacy_path.write_text(json.dumps({"notes": "legacy note"}), encoding="utf-8")

        self.assertEqual(migrate_legacy_notes_files(rc_dir), 1)
        self.assertEqual(json.loads(sidecar_path.read_text(encoding="utf-8"))["notes"], "legacy note")
        self.assertFalse(legacy_path.exists())


class DatasetCacheLoadDisplayShapeTests(unittest.TestCase):
    """A hand-entered dataset opens at the shape its sidecar shows it at.

    The ResQ import copies a triangle at the shape ResQ stores it at and
    records the shape ResQ showed it at beside it. The window's open reads
    the stored file, so without a roll-up it showed the file's own shape and
    a yearly triangle kept monthly underneath opened as 12/1.
    """

    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory(dir=str(TEST_TEMP_ROOT))
        root = Path(self.temp_dir.name)
        self.cache_dir = root / config.DATASET_CACHE_DIR
        self.cache_dir.mkdir()
        # Two yearly origins valued at the end of the second year, so a yearly
        # view of the monthly development columns is valued at 12 and 24 months.
        settings_path = root / "general_settings.json"
        settings_path.write_text(
            '{"origin_start_date":"202301","origin_end_date":"202412","development_end_date":"202412"}',
            encoding="utf-8",
        )
        self.csv_path = self.cache_dir / "Case@12@1@cum@dev.csv"
        first = ",".join(str(age) for age in range(1, 25))
        second = ",".join(str(age) for age in range(1, 13)) + "," * 12
        self.csv_path.write_text(f"{first}\n{second}\n", encoding="utf-8")
        self.sidecar = {
            "dataset_name": "Case",
            "dataset_type": "Case",
            "data_format": "Triangle",
            "source_kind": "input",
            "csv_file": self.csv_path.name,
            "cumulative": True,
            "calendar": False,
            "origin_length": 12,
            "development_length": 12,
            "stored_origin_length": 12,
            "stored_development_length": 1,
            "origin_labels": ["2023", "2024"],
        }
        self.patches = [
            patch.object(config, "get_project_dataset_cache_dir", return_value=str(self.cache_dir)),
            patch.object(config, "get_general_settings_path", return_value=str(settings_path)),
            patch.object(dataset_service, "_get_dataset_sidecar_path", return_value="sidecar.json"),
            patch.object(dataset_service, "_read_dataset_sidecar", return_value=self.sidecar),
            # The project's headers run one period past the valuation date the
            # view stops at, as the Engine's do; the view takes the ones it has
            # columns for.
            patch.object(dataset_service, "_load_project_header_labels", return_value=["12", "24", "36"]),
        ]
        for item in self.patches:
            item.start()
            self.addCleanup(item.stop)

    def tearDown(self) -> None:
        config.DATASETS.clear()
        config.DATASET_ROLLUPS.clear()
        self.temp_dir.cleanup()

    def _load(self, **kwargs: object) -> dict:
        return dataset_service.load_cached_dataset_values(
            "Example Project",
            "Example RC",
            "Case",
            origin_length=12,
            development_length=12,
            **kwargs,
        )

    def test_the_window_opens_at_the_display_shape_over_the_stored_file(self) -> None:
        result = self._load(at_display_shape=True)

        # The grid is the yearly view, built from the monthly file on the
        # project's own valuation grid, and the lengths describe that view.
        self.assertEqual(result["values"], [[12.0, 24.0], [12.0, None]])
        self.assertEqual((result["origin_length"], result["development_length"]), (12, 12))
        self.assertEqual(
            (result["stored_origin_length"], result["stored_development_length"]), (12, 1)
        )
        self.assertIsNone(result["stored_period_length"])
        self.assertEqual(result["dev_labels"], ["12", "24"])
        self.assertEqual(result["origin_labels"], ["2023", "2024"])
        # The file stays the dataset's data; the view has a handle of its own
        # and is never written beside it.
        self.assertEqual(result["csv_file"], self.csv_path.name)
        self.assertEqual(Path(result["path"]).name, "Case@12@12@cum@dev.csv")
        self.assertFalse(Path(result["path"]).exists())
        self.assertEqual(sorted(p.name for p in self.cache_dir.iterdir()), [self.csv_path.name])
        # The id-addressed grid routes serve the same view.
        with patch.object(dataset_service, "_resolve_origin_labels", return_value=["2023", "2024"]):
            grid = dataset_service.get_dataset(result["id"], "Example Project", 12)
        self.assertIsNotNone(grid)
        self.assertEqual(grid["values"][0], [12.0, 24.0])
        self.assertEqual(grid["mask"], [[True, True], [True, False]])

    def test_a_method_reading_its_input_keeps_the_stored_rows(self) -> None:
        result = self._load()

        self.assertEqual(len(result["values"]), 2)
        self.assertEqual(len(result["values"][0]), 24)
        self.assertEqual(result["values"][0][:3], [1.0, 2.0, 3.0])
        self.assertEqual((result["origin_length"], result["development_length"]), (12, 1))
        self.assertEqual(
            (result["stored_origin_length"], result["stored_development_length"]), (12, 1)
        )
        self.assertEqual(result["path"], str(self.csv_path))
        self.assertEqual(config.DATASET_ROLLUPS, {})

    def test_a_dataset_stored_at_its_display_shape_is_served_as_it_is(self) -> None:
        self.sidecar["stored_development_length"] = 12
        csv_path = self.cache_dir / "Case@12@12@cum@dev.csv"
        csv_path.write_text("12,24\n12,\n", encoding="utf-8")
        self.sidecar["csv_file"] = csv_path.name

        result = self._load(at_display_shape=True)

        self.assertEqual(result["values"], [[12.0, 24.0], [12.0, None]])
        self.assertEqual((result["origin_length"], result["development_length"]), (12, 12))
        self.assertEqual(
            (result["stored_origin_length"], result["stored_development_length"]), (12, 12)
        )
        self.assertEqual(result["path"], str(csv_path))
        self.assertEqual(config.DATASET_ROLLUPS, {})


if __name__ == "__main__":
    unittest.main()

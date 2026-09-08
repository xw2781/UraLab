from __future__ import annotations

import sys
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
FRONTEND_ROOT = REPO_ROOT / "frontend"
PYTHON_API_SRC = REPO_ROOT / "python-api" / "src"
SERVER_COMPONENTS_SRC = REPO_ROOT / "server-components" / "src"
for path in (FRONTEND_ROOT, PYTHON_API_SRC, SERVER_COMPONENTS_SRC):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from app_server.helpers import parse_method_last_modified_timestamp
from app_server.services import dfm_rpc_bridge_service


def _resq_wall_clock_text(moment: datetime) -> str:
    """Match ``resq_client._dfm_last_modified``: local wall clock, no timezone."""
    return moment.replace(tzinfo=None).isoformat()


def _arcrho_save_text(moment: datetime) -> str:
    """Match the DFM save path: ``new Date().toISOString()`` in UTC."""
    utc = moment.astimezone(timezone.utc)
    return utc.isoformat(timespec="milliseconds").replace("+00:00", "Z")


def _meta(raw: str) -> dict:
    return {
        "exists": True,
        "last_modified": raw,
        "last_modified_timestamp": parse_method_last_modified_timestamp(raw),
    }


class ParseMethodLastModifiedTimestampTests(unittest.TestCase):
    def test_timezone_less_value_is_read_as_local_wall_clock(self) -> None:
        local_moment = datetime(2026, 8, 10, 9, 47, 55).astimezone()
        parsed = parse_method_last_modified_timestamp(
            _resq_wall_clock_text(local_moment)
        )
        self.assertAlmostEqual(parsed, local_moment.timestamp(), places=6)

    def test_offset_and_z_values_keep_their_stated_instant(self) -> None:
        moment = datetime(2026, 8, 10, 13, 47, 15, tzinfo=timezone.utc)
        self.assertAlmostEqual(
            parse_method_last_modified_timestamp("2026-08-10T13:47:15.000Z"),
            moment.timestamp(),
            places=6,
        )
        self.assertAlmostEqual(
            parse_method_last_modified_timestamp(
                moment.astimezone(timezone(timedelta(hours=-4))).isoformat()
            ),
            moment.timestamp(),
            places=6,
        )

    def test_unusable_values_return_none(self) -> None:
        for value in (None, True, "", "   ", "not a timestamp", 0, -1.0):
            self.assertIsNone(parse_method_last_modified_timestamp(value))

    def test_numeric_epoch_values_pass_through(self) -> None:
        self.assertEqual(parse_method_last_modified_timestamp(1786887235.5), 1786887235.5)
        self.assertEqual(parse_method_last_modified_timestamp("1786887235.5"), 1786887235.5)


class RpcBridgeComparisonTimezoneTests(unittest.TestCase):
    """A ResQ save 40 seconds after an ArcRho save must read as the newer version.

    Reading the timezone-less ResQ value as UTC skewed it by the machine's UTC
    offset, which put the ``NEW`` seal on the ArcRho card instead.
    """

    def setUp(self) -> None:
        self.local_save = datetime(2026, 8, 10, 9, 47, 15).astimezone()
        self.remote_save = self.local_save + timedelta(seconds=40)

    def test_dfm_compare_state_follows_real_wall_clock_order(self) -> None:
        state = dfm_rpc_bridge_service._compare_state(
            _meta(_arcrho_save_text(self.local_save)),
            _meta(_resq_wall_clock_text(self.remote_save)),
        )
        self.assertEqual(state, "remote_latest")

    def test_dfm_compare_state_reports_local_latest_when_local_is_newer(self) -> None:
        state = dfm_rpc_bridge_service._compare_state(
            _meta(_arcrho_save_text(self.remote_save)),
            _meta(_resq_wall_clock_text(self.local_save)),
        )
        self.assertEqual(state, "local_latest")

    def test_matching_migrated_wall_clock_values_stay_in_sync(self) -> None:
        raw = _resq_wall_clock_text(self.local_save)
        state = dfm_rpc_bridge_service._compare_state(_meta(raw), _meta(raw))
        self.assertEqual(state, "same_time")


if __name__ == "__main__":
    unittest.main()

"""The one persisted timestamp form (docs/plans/completed/persisted_json_contract_v4.md, rule 3)."""

from __future__ import annotations

import os
import re
import sys
import unittest
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))

from arcrho_api.timestamps import (  # noqa: E402
    format_display_timestamp,
    format_persisted_timestamp,
    is_persisted_timestamp,
    normalize_persisted_timestamp,
    persisted_timestamp,
    utc_now_text,
)

PERSISTED_FORM = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$")


class PersistedTimestampTests(unittest.TestCase):
    def test_now_is_utc_with_milliseconds_and_a_z(self) -> None:
        text = utc_now_text()
        self.assertRegex(text, PERSISTED_FORM)
        self.assertTrue(is_persisted_timestamp(text))

    def test_an_aware_value_is_converted_to_utc(self) -> None:
        value = datetime(2026, 8, 19, 10, 5, 30, 500000, tzinfo=timezone(timedelta(hours=-4)))
        self.assertEqual(format_persisted_timestamp(value), "2026-08-19T14:05:30.500Z")

    def test_a_naive_value_is_a_local_wall_clock_reading(self) -> None:
        # ResQ's ``Modified`` reaches the app as a bare ISO string in the
        # machine's own zone, and Python reads such a value the same way.
        naive = datetime(2026, 8, 19, 10, 5, 30, 500000)
        expected = format_persisted_timestamp(naive.astimezone(timezone.utc))
        self.assertEqual(format_persisted_timestamp(naive), expected)
        self.assertEqual(normalize_persisted_timestamp("2026-08-19T10:05:30.500000"), expected)
        self.assertEqual(persisted_timestamp("2026-08-19T10:05:30.500000"), expected)
        self.assertAlmostEqual(
            datetime.fromisoformat(expected[:-1] + "+00:00").timestamp(),
            naive.timestamp(),
        )

    def test_every_older_form_normalizes_to_the_same_text(self) -> None:
        self.assertEqual(normalize_persisted_timestamp("2026-08-19T14:05:30Z"), "2026-08-19T14:05:30.000Z")
        self.assertEqual(normalize_persisted_timestamp("2026-08-19T14:05:30.500000Z"), "2026-08-19T14:05:30.500Z")
        self.assertEqual(normalize_persisted_timestamp("2026-08-19T10:05:30.500-04:00"), "2026-08-19T14:05:30.500Z")
        self.assertEqual(normalize_persisted_timestamp("2026-08-19T14:05:30.500+00:00"), "2026-08-19T14:05:30.500Z")
        self.assertEqual(normalize_persisted_timestamp("2026-08-19T14:05:30.500Z"), "2026-08-19T14:05:30.500Z")

    def test_text_that_is_not_a_time_returns_the_default(self) -> None:
        self.assertEqual(normalize_persisted_timestamp("later"), "")
        self.assertEqual(normalize_persisted_timestamp("later", default="later"), "later")
        self.assertEqual(normalize_persisted_timestamp(None), "")
        self.assertEqual(normalize_persisted_timestamp("   "), "")
        self.assertFalse(is_persisted_timestamp("2026-08-19T14:05:30Z"))
        self.assertFalse(is_persisted_timestamp("later"))

    def test_the_display_form_is_local_time_in_the_list_format(self) -> None:
        # The ResQ window showed 8/13/2026 2:49:34 PM for the value persisted
        # as 18:49:34Z (Eastern Daylight Time); the display form must show the
        # same wall clock ResQ does, whatever zone the machine is in.
        instant = datetime(2026, 8, 13, 18, 49, 34, 302000, tzinfo=timezone.utc)
        local = instant.astimezone()
        hour = local.hour % 12 or 12
        expected = f"{local.month}/{local.day}/{local.year} {hour}:{local:%M:%S} {'PM' if local.hour >= 12 else 'AM'}"
        self.assertEqual(format_display_timestamp("2026-08-13T18:49:34.302Z"), expected)
        self.assertEqual(format_display_timestamp("2026-08-13T18:49:34.302+00:00"), expected)
        self.assertEqual(format_display_timestamp(local.isoformat()), expected)
        self.assertEqual(format_display_timestamp(local.replace(tzinfo=None).isoformat()), expected)
        self.assertRegex(format_display_timestamp("2026-08-13T00:30:00Z"), r"^\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2}:\d{2} [AP]M$")

    def test_the_display_form_keeps_text_that_is_not_a_time(self) -> None:
        self.assertEqual(format_display_timestamp("Not present", default="Not present"), "Not present")
        self.assertEqual(format_display_timestamp(""), "")
        self.assertEqual(format_display_timestamp(None), "")

    def test_the_contract_stamp_keeps_a_token_and_mints_now_for_nothing(self) -> None:
        self.assertEqual(persisted_timestamp("later"), "later")
        self.assertEqual(persisted_timestamp("2026-08-19T14:05:30Z"), "2026-08-19T14:05:30.000Z")
        self.assertRegex(persisted_timestamp(""), PERSISTED_FORM)
        self.assertRegex(persisted_timestamp(None), PERSISTED_FORM)


if __name__ == "__main__":
    unittest.main()

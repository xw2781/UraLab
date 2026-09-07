import sys
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from arcrho_api.triangle_rollup import (
    rollup_factors,
    rollup_reason,
    rollup_triangle,
    scatter_reason,
    scatter_triangle,
)


def _monthly_cumulative(rows: int, columns: int, valuation: int, origin_length: int = 1) -> list:
    """A dev-aligned cumulative triangle whose every cell is 100 x its age in months.

    Row ``i`` starts ``i * origin_length`` months after the anchor and column
    ``j`` is valued ``j + 1`` months after that, so the cell exists only while
    ``i * origin_length + j + 1`` is within ``valuation`` months of the anchor.
    """
    triangle = []
    for row in range(rows):
        triangle.append([
            100.0 * (column + 1) if row * origin_length + column + 1 <= valuation else None
            for column in range(columns)
        ])
    return triangle


class TriangleRollupTests(unittest.TestCase):
    def test_annual_view_of_a_monthly_triangle_follows_the_calendar_diagonal(self) -> None:
        rolled = rollup_triangle(
            _monthly_cumulative(24, 24, 24),
            source_origin_length=1,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=12,
            valuation_months=24,
        )
        self.assertEqual(rolled, [[7800.0, 22200.0], [7800.0, None]])

    def test_incremental_roll_up_matches_the_cumulative_one(self) -> None:
        source = _monthly_cumulative(24, 24, 24)
        incremental = [
            [None if value is None else 100.0 for value in row]
            for row in source
        ]
        rolled = rollup_triangle(
            incremental,
            source_origin_length=1,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=12,
            valuation_months=24,
            cumulative=False,
        )
        self.assertEqual(rolled, [[7800.0, 14400.0], [7800.0, None]])

    def test_origin_and_development_factors_are_independent(self) -> None:
        rolled = rollup_triangle(
            _monthly_cumulative(8, 24, 24, origin_length=3),
            source_origin_length=3,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=12,
            valuation_months=24,
        )
        self.assertEqual(rolled, [[3000.0, 7800.0], [3000.0, None]])

    def test_development_alone_can_coarsen_under_unchanged_origins(self) -> None:
        source = [
            [100.0, 200.0, 300.0, 400.0, 500.0, 600.0, 700.0, 800.0],
            [100.0, 200.0, 300.0, 400.0, None, None, None, None],
        ]
        rolled = rollup_triangle(
            source,
            source_origin_length=1,
            source_development_length=3,
            target_origin_length=1,
            target_development_length=12,
            valuation_months=24,
        )
        self.assertEqual(rolled, [[400.0, 800.0], [400.0, None]])

    def test_development_periods_are_counted_back_from_the_development_end_date(self) -> None:
        # Yearly origins stored monthly and valued 20 months after the anchor:
        # a yearly view is valued at 8 and 20 months of age, the way ResQ
        # labels it, and keeps the latest diagonal as its last column.
        source = _monthly_cumulative(2, 20, 20, origin_length=12)
        yearly = rollup_triangle(
            source,
            source_origin_length=12,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=12,
            valuation_months=20,
        )
        self.assertEqual(yearly, [[800.0, 2000.0], [800.0, None]])
        half_yearly = rollup_triangle(
            source,
            source_origin_length=12,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=6,
            valuation_months=20,
        )
        self.assertEqual(
            half_yearly,
            [[200.0, 800.0, 1400.0, 2000.0], [200.0, 800.0, None, None]],
        )

    def test_an_incremental_first_period_is_the_short_one(self) -> None:
        source = [
            [None if value is None else 100.0 for value in row]
            for row in _monthly_cumulative(2, 20, 20, origin_length=12)
        ]
        rolled = rollup_triangle(
            source,
            source_origin_length=12,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=12,
            valuation_months=20,
            cumulative=False,
        )
        self.assertEqual(rolled, [[800.0, 1200.0], [800.0, None]])

    def test_a_partly_filled_origin_block_is_still_a_row(self) -> None:
        # Five quarterly origins valued 15 months after the anchor: the second
        # yearly row holds the one quarter that has started, and the yearly
        # view is valued at 3 and 15 months of age.
        rolled = rollup_triangle(
            _monthly_cumulative(5, 15, 15, origin_length=3),
            source_origin_length=3,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=12,
            valuation_months=15,
        )
        self.assertEqual(rolled, [[300.0, 4200.0], [300.0, None]])

    def test_monthly_origins_reported_quarterly_cannot_share_a_valuation_date(self) -> None:
        reason = rollup_reason(1, 3, 12, 12)
        self.assertIn("share no valuation date", reason)
        with self.assertRaises(ValueError):
            rollup_factors(1, 3, 12, 12)

    def test_a_target_period_that_is_not_a_whole_multiple_is_refused(self) -> None:
        self.assertEqual(
            rollup_reason(3, 3, 12, 8),
            "requested periods are not whole multiples of the cached periods",
        )
        self.assertEqual(
            rollup_reason(12, 12, 1, 1),
            "local caches can only derive from finer to coarser periods",
        )
        self.assertEqual(rollup_reason(0, 1, 12, 12), "invalid period length")

    def test_a_calendar_triangle_aggregates_by_block(self) -> None:
        source = [
            [100.0 * (column + 1) if column >= row else None for column in range(24)]
            for row in range(24)
        ]
        rolled = rollup_triangle(
            source,
            source_origin_length=1,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=12,
            valuation_months=24,
            calendar=True,
        )
        self.assertEqual(rolled, [[14400.0, 28800.0], [None, 28800.0]])

    def test_a_calendar_triangle_ends_on_a_short_period(self) -> None:
        source = [[100.0 * (column + 1) for column in range(20)]]
        rolled = rollup_triangle(
            source,
            source_origin_length=12,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=12,
            valuation_months=20,
            calendar=True,
        )
        self.assertEqual(rolled, [[1200.0, 2000.0]])

    def test_a_calendar_triangle_may_coarsen_origins_under_finer_development(self) -> None:
        self.assertEqual(rollup_reason(1, 3, 12, 3, calendar=True), "")

    def test_a_valuation_date_is_required(self) -> None:
        with self.assertRaises(ValueError):
            rollup_triangle(
                [[1.0, 2.0]],
                source_origin_length=1,
                source_development_length=1,
                target_origin_length=12,
                target_development_length=12,
                valuation_months=0,
            )


class TriangleScatterTests(unittest.TestCase):
    """Values entered at a coarser development view land in the stored cells.

    Pinned to the ResQ probe recorded in
    ``docs/plans/manual_input_stored_length_resq_alignment.md``: annual origins
    from 2017-01 valued on 2026-05-31, so the project is 113 months long and an
    annual view of a monthly store is valued at 5, 17, 29, ... 113 months.
    """

    valuation_months = 113
    origin_rows = 10

    def _annual_grid(self) -> list:
        """The annual display: cell ``1000 x row + column``, one column narrower each row."""
        return [
            [1000.0 * row + column for column in range(self.origin_rows - row)]
            for row in range(self.origin_rows)
        ]

    def _scatter(self, grid: list, cumulative: bool = True) -> list:
        return scatter_triangle(
            grid,
            source_origin_length=12,
            source_development_length=1,
            target_origin_length=12,
            target_development_length=12,
            valuation_months=self.valuation_months,
            cumulative=cumulative,
        )

    def _filled(self, row_index: int, row: list) -> list:
        width = self.valuation_months - 12 * row_index
        self.assertTrue(
            all(value is None for value in row[width:]),
            f"row {row_index} holds a cell after the valuation date",
        )
        filled = row[:width]
        self.assertTrue(all(value is not None for value in filled))
        return filled

    def test_an_annual_grid_lands_at_the_ages_the_annual_view_reads(self) -> None:
        stored = self._scatter(self._annual_grid())

        self.assertEqual(len(stored), self.origin_rows)
        for row_index, row in enumerate(stored):
            self.assertEqual(len(row), self.valuation_months)
            filled = self._filled(row_index, row)
            entered = {
                4 + 12 * column: 1000.0 * row_index + column
                for column in range(self.origin_rows - row_index)
            }
            for column, value in enumerate(filled):
                self.assertEqual(
                    value,
                    entered.get(column, 0.0),
                    f"stored row {row_index} at {column + 1} months",
                )

    def test_an_incremental_grid_stores_the_running_sums(self) -> None:
        stored = self._scatter(self._annual_grid(), cumulative=False)

        for row_index, row in enumerate(stored):
            filled = self._filled(row_index, row)
            running = 0.0
            expected = {}
            for column in range(self.origin_rows - row_index):
                running += 1000.0 * row_index + column
                # The whole store is cumulative 0 apart from the entered ages,
                # so each running sum shows again as its own negative one month on.
                expected[4 + 12 * column] = running
                expected[5 + 12 * column] = -running
            for column, value in enumerate(filled):
                self.assertEqual(
                    value,
                    expected.get(column, 0.0),
                    f"stored row {row_index} at {column + 1} months",
                )

    def test_the_annual_view_of_the_scattered_store_is_what_was_entered(self) -> None:
        grid = self._annual_grid()
        padded = [row + [None] * (self.origin_rows - len(row)) for row in grid]
        for cumulative in (True, False):
            rolled = rollup_triangle(
                self._scatter(grid, cumulative=cumulative),
                source_origin_length=12,
                source_development_length=1,
                target_origin_length=12,
                target_development_length=12,
                valuation_months=self.valuation_months,
                cumulative=cumulative,
            )
            self.assertEqual(rolled, padded, f"cumulative={cumulative}")

    def test_a_coarser_origin_period_cannot_be_written(self) -> None:
        self.assertEqual(
            scatter_reason(1, 1, 12, 12),
            "values can be entered only at the stored origin period",
        )
        self.assertEqual(scatter_reason(12, 1, 12, 12), "")
        with self.assertRaises(ValueError):
            scatter_triangle(
                [[1.0]],
                source_origin_length=1,
                source_development_length=1,
                target_origin_length=12,
                target_development_length=12,
                valuation_months=self.valuation_months,
            )


if __name__ == "__main__":
    unittest.main()

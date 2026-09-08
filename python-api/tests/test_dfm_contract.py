from __future__ import annotations

import sys
import unittest
from copy import deepcopy
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from arcrho_api.dfm_contract import (  # noqa: E402
    DFM_JSON_FORMAT,
    DfmContractError,
    apply_owned_patch,
    build_dfm_output_sidecar,
    canonical_input_number,
    canonical_number,
    dataset_reference_tokens,
    dfm_dataset_reference_tokens,
    dfm_precedent_names,
    dfm_output_variants,
    method_revisions,
    normalize_dfm_method,
    owned_projection,
    preview_dfm_method,
    recalculate_dfm_method,
    round_half_up,
    source_snapshot_revision,
)


def owned_payload() -> dict:
    return {
        "json_format": DFM_JSON_FORMAT,
        "details_tab": {
            "name": "Paid DFM",
            "output_type": "Paid Ultimate",
            "output_dataset": "Paid Selected",
            "input_triangle": "Paid Loss",
            "origin_length": 12,
            "development_length": 12,
            "decimal_places": 4,
        },
        "data_tab": {},
        "ratios_tab": {
            "ratio_triangle": {"excluded": [[1, 0], [0], []]},
            "average_formulas": {
                "label": ["Volume - all", "Simple - all", "User A", "User B", "Excel Entry"],
                "custom_average_formula_settings": {
                    "average_type": ["custom", "custom", "user_entry", "user_entry", "user_entry"],
                    "base": ["volume", "simple", "simple", "simple", "simple"],
                    "periods": ["all", "all", "all", "all", "all"],
                    "exclude": [0, 0, 0, 0, 0],
                },
                "selected": [
                    [1, 0, 1],
                    [0, 0, 0],
                    [0, 1, 0],
                    [0, 0, 0],
                    [0, 0, 0],
                ],
                "values": [
                    [1, 1, 1],
                    [1, 1, 1],
                    [9, 9, 1],
                    [8, 8, 1],
                    [1.25, 1.3, 1],
                ],
                "inputs": [
                    ["", "", ""],
                    ["", "", ""],
                    ['="User B" * 2', '="User B" * 2', ""],
                    ['="Simple - all" * 1.1', '="Simple - all" * 1.1', ""],
                    ["='[Book.xlsx]Sheet1'!$A$1", "=1.3", ""],
                ],
                "display_inputs": [
                    ["", "", ""],
                    ["", "", ""],
                    ["", "", ""],
                    ["", "", ""],
                    ["=[Premium][2025 Q4]", "", ""],
                ],
            },
            "cell_notes": {
                "ratio_main_table": {"2020": {"(1) 12-24": "Keep"}},
                "ratio_summary_table": {},
            },
        },
        "results_tab": {
            "ratio_basis_dataset": "Earned Premium",
            "ultimate_ratio_decimal_places": 2,
        },
        "method_metadata": {
            "last_modified": "2026-01-01T00:00:00Z",
            "data_refreshed": "2026-01-01T00:00:00Z",
        },
    }


def input_snapshot(*, values: list[list[float | None]] | None = None) -> dict:
    values = values or [[100, 150, 180], [200, 300, None], [400, None, None]]
    return {
        "name": "Paid Loss",
        "origin_labels": ["2020", "2021", "2022"],
        "development_labels": ["12m", "24m", "36m"],
        "values": values,
        "mask": [[item is not None for item in row] for row in values],
        "data_format": "Triangle",
        "number_format": "#,##0",
        "decimal_places": 0,
        "revision": "input:r1",
    }


def basis_snapshot() -> dict:
    return {
        "name": "Earned Premium",
        "origin_labels": ["2022", "2020", "2021"],
        "values": [3000, 1000, 2000],
        "data_format": "Vector",
        "number_format": "$#,##0",
        "decimal_places": 0,
        "revision": "basis:r1",
    }


class DfmExcludeHighLowTests(unittest.TestCase):
    """ResQ stops dropping high/low pairs once two ratios are left.

    The automation help states the rule as excluding pairs "for as long as the
    remaining number of ratios is greater than two", so a column down to two
    ratios averages both instead of excluding itself empty. Measured against
    ResQ on the fake project: an Ex hi/lo row read 1.00457 there where ArcRho
    used to fall back to 1.0.
    """

    @staticmethod
    def _average(ratios: list[float], exclude: int) -> float:
        from arcrho_api.dfm_contract import _calculate_average

        values = [[1.0, ratio] for ratio in ratios]
        mask = [[True, True] for _ in ratios]
        excluded = [[0] for _ in ratios]
        return _calculate_average(
            values, mask, excluded, 0, base="simple", periods="all", extra_exclude=exclude
        )

    def test_two_ratios_keep_both_instead_of_excluding_the_column_empty(self) -> None:
        self.assertAlmostEqual(self._average([1.0, 1.00914], 1), 1.00457)

    def test_one_ratio_is_kept(self) -> None:
        self.assertAlmostEqual(self._average([1.25], 1), 1.25)

    def test_three_ratios_still_drop_the_highest_and_lowest_pair(self) -> None:
        self.assertAlmostEqual(self._average([1.0, 2.0, 9.0], 1), 2.0)

    def test_a_second_pass_stops_before_it_would_leave_two_behind(self) -> None:
        # Four ratios allow one pair only: dropping a second would empty the column.
        self.assertAlmostEqual(self._average([1.0, 2.0, 3.0, 9.0], 2), 2.5)


class ResqMutedRatioAverageTests(unittest.TestCase):
    """Every average row of one ResQ column whose newest origin has no ratio.

    Read off ResQ for column "(1) 8-20": eight ratios for 2017-2024, a muted
    1.0000 at 2025 where the later value is zero, and an empty 2026. The four
    cases are the same column with nothing struck out, with 2024 struck, with
    2021 and 2022 struck, and with 2019-2023 struck. The muted origin takes no
    place in a "last N" window, adds nothing to a sum, and counts toward no
    divisor; treating it as a ratio of zero moved every row of the first case,
    "Simple - 2" from 1.9333 to 0.9596.
    """

    _RATIOS = [1.9232, 1.8949, 1.8139, 1.7265, 1.9000, 2.8251, 1.9474, 1.9191, 0.0]

    def _column(self, struck: set[int]) -> tuple[list, list, list]:
        values = [[1.0, ratio] for ratio in self._RATIOS] + [[None, None]]
        mask = [[True, True] for _ in self._RATIOS] + [[False, False]]
        excluded = [[1 if row in struck else 0] for row in range(len(self._RATIOS))] + [[0]]
        return values, mask, excluded

    def _average(self, struck: set[int], periods: int, exclude: int) -> str:
        from arcrho_api.dfm_contract import _calculate_average, canonical_number

        values, mask, excluded = self._column(struck)
        value = _calculate_average(
            values, mask, excluded, 0, base="simple", periods=periods, extra_exclude=exclude
        )
        return f"{Decimal(str(canonical_number(value))).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)}"

    def _check(self, struck: set[int], expected: dict[tuple[int, int], str]) -> None:
        for (periods, exclude), want in expected.items():
            with self.subTest(periods=periods, exclude=exclude):
                self.assertEqual(self._average(struck, periods, exclude), want)

    def test_nothing_struck_out(self) -> None:
        self._check(set(), {
            (8, 0): "1.9938", (8, 1): "1.8998", (5, 0): "2.0636",
            (5, 1): "1.9222", (3, 0): "2.2305", (2, 0): "1.9333",
        })

    def test_the_newest_ratio_struck_out(self) -> None:
        self._check({7}, {
            (8, 0): "2.0044", (8, 1): "1.8959", (5, 0): "2.0426",
            (5, 1): "1.8871", (3, 0): "2.2242", (2, 0): "2.3863",
        })

    def test_two_middle_ratios_struck_out(self) -> None:
        self._check({4, 5}, {
            (8, 0): "1.8708", (8, 1): "1.8878", (5, 0): "1.8604",
            (5, 1): "1.8760", (3, 0): "1.8643", (2, 0): "1.9333",
        })

    def test_five_struck_out_leaves_three_ratios(self) -> None:
        # Three candidates allow one pair, so an Ex hi/lo row reports the one
        # ratio left standing rather than falling back to 1.0.
        self._check({2, 3, 4, 5, 6}, {
            (8, 0): "1.9124", (8, 1): "1.9191", (5, 0): "1.9124",
            (5, 1): "1.9191", (3, 0): "1.9124", (2, 0): "1.9070",
        })


class DfmContractTests(unittest.TestCase):
    def test_canonical_number_rounds_half_away_from_zero(self) -> None:
        self.assertEqual(canonical_number("1.0000005"), 1.000001)
        self.assertEqual(canonical_number("-1.0000005"), -1.000001)
        self.assertIsNone(canonical_number(float("nan")))

    def test_canonical_input_number_keeps_every_digit(self) -> None:
        self.assertEqual(canonical_input_number("0.00000123456789"), 0.00000123456789)
        self.assertEqual(canonical_input_number("-0.00000123456789"), -0.00000123456789)
        self.assertIsNone(canonical_input_number(float("nan")))
        self.assertIsNone(canonical_input_number(float("inf")))
        self.assertIsNone(canonical_input_number(""))
        self.assertIsNone(canonical_input_number(True))

    def test_a_wide_figure_survives(self) -> None:
        # An observation of any size must stay a number rather than become null.
        self.assertEqual(canonical_input_number(1e21), 1e21)

    def test_the_input_triangle_keeps_the_precision_it_was_read_with(self) -> None:
        # A near-zero "% of" observation divides into a ratio a reader checks at
        # four decimals, so trimming its tail at any fixed decimal place moves
        # that ratio. The stored value is the value the source holds.
        values = [[0.00000123456789, 0.00000234567891, None], [None, None, None], [None, None, None]]
        method = recalculate_dfm_method(
            owned_payload(),
            input_snapshot=input_snapshot(values=values),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="2026-01-02T00:00:00Z",
        )
        self.assertEqual(
            method["data_tab"]["input_data_triangle_values"][0][:2],
            [0.00000123456789, 0.00000234567891],
        )

    def test_a_near_zero_denominator_no_longer_moves_the_ratio(self) -> None:
        # The production case behind this rule: dividing by a rounded copy of a
        # small "% of" figure showed up in the fourth decimal of a large ratio.
        small = 0.00014285714285714287
        values = [[small, small * 1231.8996, None], [None, None, None], [None, None, None]]
        method = recalculate_dfm_method(
            owned_payload(),
            input_snapshot=input_snapshot(values=values),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="2026-01-02T00:00:00Z",
        )
        self.assertAlmostEqual(
            method["ratios_tab"]["ratio_triangle"]["ratio_values"][0][0], 1231.8996, places=4
        )

    def test_a_zero_later_value_holds_no_ratio_at_all(self) -> None:
        values = [[100.0, 0.0, None], [200.0, 260.0, None], [None, None, None]]
        method = recalculate_dfm_method(
            owned_payload(),
            input_snapshot=input_snapshot(values=values),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="2026-01-02T00:00:00Z",
        )
        ratio_values = method["ratios_tab"]["ratio_triangle"]["ratio_values"]
        self.assertEqual(ratio_values[0], [])
        self.assertAlmostEqual(ratio_values[1][0], 1.3)

    def test_a_zero_later_value_is_left_out_of_the_column_average(self) -> None:
        from arcrho_api.dfm_contract import _calculate_average

        values = [[100.0, 0.0], [200.0, 260.0]]
        mask = [[True, True], [True, True]]
        excluded = [[0], [0]]
        for base in ("volume", "simple"):
            with self.subTest(base=base):
                self.assertAlmostEqual(
                    _calculate_average(
                        values, mask, excluded, 0, base=base, periods="all", extra_exclude=0
                    ),
                    1.3,
                )

    def test_a_ratio_divides_unrounded_input_values(self) -> None:
        values = [[0.00000123456789, 0.00000234567891, None], [None, None, None], [None, None, None]]
        method = recalculate_dfm_method(
            owned_payload(),
            input_snapshot=input_snapshot(values=values),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="2026-01-02T00:00:00Z",
        )
        # Six-decimal operands would collapse both figures onto 0.000001 and
        # report a ratio of exactly 1; ten-decimal operands would report
        # 1.899968 rather than the 1.9 the two observations actually make.
        self.assertAlmostEqual(
            method["ratios_tab"]["ratio_triangle"]["ratio_values"][0][0],
            round(0.00000234567891 / 0.00000123456789, 6),
            places=6,
        )

    def test_output_variants_share_canonical_period_aggregation(self) -> None:
        variants = dfm_output_variants({
            "details_tab": {"origin_length": 3},
            "data_tab": {"origin_labels": ["2020 Q1", "2020 Q2", "2020 Q3", "2020 Q4"]},
            "results_tab": {"ultimate_vector": [1, 2, 3, 4]},
        })
        self.assertEqual(variants, {3: [1, 2, 3, 4], 6: [3, 7], 12: [10]})

    def test_recalculation_builds_complete_self_contained_v2(self) -> None:
        method = recalculate_dfm_method(
            owned_payload(),
            input_snapshot=input_snapshot(),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="2026-01-02T00:00:00Z",
        )

        self.assertEqual(method["ratios_tab"]["ratio_triangle"]["development_labels"], [
            "(1) 12-24", "(2) 24-36", "36 - Ult",
        ])
        self.assertEqual(method["results_tab"]["ratio_basis_values"], [1000, 2000, 3000])
        self.assertEqual(method["data_tab"]["data_format"], "Triangle")
        self.assertEqual(method["results_tab"]["ratio_basis_data_format"], "Vector")
        self.assertNotIn("input data triangle csv path", method["data_tab"])
        self.assertNotIn("ultimate vector csv path", method["results_tab"])
        self.assertEqual(method["ratios_tab"]["average_formulas"]["values"][4][0], 1.25)
        self.assertEqual(method["ratios_tab"]["average_formulas"]["values"][4][1], 1.3)
        self.assertEqual(
            method["ratios_tab"]["average_formulas"]["display_inputs"][4][0],
            "=[Premium][2025 Q4]",
        )
        self.assertEqual(method["method_metadata"]["data_refreshed"], "2026-01-02T00:00:00.000Z")
        self.assertEqual(normalize_dfm_method(method), method)

    def test_source_revisions_and_payload_ignore_producer_timestamps(self) -> None:
        first_input = input_snapshot()
        second_input = deepcopy(first_input)
        first_input["revision"] = "frontend:2026-01-01"
        second_input["revision"] = "migration:2026-07-01"
        first_basis = basis_snapshot()
        second_basis = deepcopy(first_basis)
        first_basis["revision"] = "frontend:basis"
        second_basis["revision"] = "migration:basis"

        first = recalculate_dfm_method(
            owned_payload(),
            input_snapshot=first_input,
            ratio_basis_snapshot=first_basis,
            timestamp="same",
        )
        second = recalculate_dfm_method(
            owned_payload(),
            input_snapshot=second_input,
            ratio_basis_snapshot=second_basis,
            timestamp="same",
        )

        self.assertEqual(first, second)
        self.assertEqual(source_snapshot_revision(first_input), source_snapshot_revision(second_input))

    def test_display_inputs_are_backward_compatible_display_metadata(self) -> None:
        method = recalculate_dfm_method(
            owned_payload(), input_snapshot=input_snapshot(), ratio_basis_snapshot=basis_snapshot(), timestamp="same"
        )
        legacy = deepcopy(method)
        legacy["ratios_tab"]["average_formulas"].pop("display_inputs")

        normalized_legacy = normalize_dfm_method(legacy)

        self.assertEqual(
            normalized_legacy["method_metadata"]["owned_revision"],
            method["method_metadata"]["owned_revision"],
        )
        self.assertEqual(
            normalized_legacy["ratios_tab"]["average_formulas"]["display_inputs"],
            [["", "", ""] for _ in method["ratios_tab"]["average_formulas"]["label"]],
        )
        display_patch = deepcopy(method)
        display_patch["ratios_tab"]["average_formulas"]["display_inputs"][4][0] = "=[Premium][2026 Q1]"
        patched = apply_owned_patch(method, display_patch)
        self.assertEqual(
            patched["ratios_tab"]["average_formulas"]["display_inputs"][4][0],
            "=[Premium][2026 Q1]",
        )
        self.assertEqual(
            patched["method_metadata"]["owned_revision"],
            method["method_metadata"]["owned_revision"],
        )

    def test_output_sidecar_projection_is_canonical_and_preserves_owned_sidecar_state(self) -> None:
        method = recalculate_dfm_method(
            owned_payload(), input_snapshot=input_snapshot(), ratio_basis_snapshot=basis_snapshot(), timestamp="same"
        )
        existing = {
            "notes": "Method note",
            "audit_log": [{"event_date": "old", "action": "Insert", "change_info": "", "user": "a"}],
            "dependents": ["Selected Ultimate", {"dataset_name": "Report"}],
            "created": "old",
            "number_format": "$#,##0",
            "show_subtotal": False,
            "producer_only": "must be removed",
        }
        first = build_dfm_output_sidecar(
            method,
            project_name="Demo",
            reserving_class=r"Auto\PP",
            csv_file="Paid Selected@12.csv",
            existing=existing,
            timestamp="new",
            user="tester",
        )
        second = build_dfm_output_sidecar(
            method,
            project_name="Demo",
            reserving_class=r"Auto\PP",
            csv_file="Paid Selected@12.csv",
            existing=deepcopy(existing),
            timestamp="new",
            user="tester",
        )
        self.assertEqual(first, second)
        self.assertNotIn("producer_only", first)
        self.assertEqual(
            first["precedents"],
            [{"dataset_name": "Paid Loss"}, {"dataset_name": "Earned Premium"}],
        )
        self.assertEqual(first["notes"], "Method note")
        self.assertIs(first["show_subtotal"], False)
        self.assertEqual(first["publication_revision"], method["method_metadata"]["publication_revision"])

    def test_dataset_formula_inputs_are_owned_precedents_and_preserve_stored_values(self) -> None:
        payload = owned_payload()
        formulas = payload["ratios_tab"]["average_formulas"]
        formulas["inputs"][2][0] = '=[Accounting Cutoff][-1] * [Growth Adjustment]["2024", "12m"]'
        formulas["inputs"][2][1] = '=[accounting cutoff][1]'
        formulas["display_inputs"][2][0] = "=[Display Metadata Only][2024]"

        method = recalculate_dfm_method(
            payload,
            input_snapshot=input_snapshot(),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="same",
        )

        self.assertEqual(
            dfm_precedent_names(method),
            ["Paid Loss", "Earned Premium", "Accounting Cutoff", "Growth Adjustment"],
        )
        self.assertEqual(method["ratios_tab"]["average_formulas"]["values"][2][:2], [9, 9])
        owned_values = owned_projection(method)["average_formulas"]["owned_values"]
        user_a = next(item for item in owned_values if item["label"] == "User A")
        self.assertEqual(user_a, {"label": "User A", "columns": [0, 1, 2], "values": [9, 9, 1.0]})
        sidecar = build_dfm_output_sidecar(
            method,
            project_name="Demo",
            reserving_class=r"Auto\PP",
            csv_file="Paid Selected@12.csv",
            timestamp="same",
        )
        self.assertEqual(
            sidecar["precedents"],
            [
                {"dataset_name": "Paid Loss"},
                {"dataset_name": "Earned Premium"},
                {"dataset_name": "Accounting Cutoff"},
                {"dataset_name": "Growth Adjustment"},
            ],
        )

    def test_dataset_reference_values_re_evaluate_referenced_formulas(self) -> None:
        payload = owned_payload()
        formulas = payload["ratios_tab"]["average_formulas"]
        formulas["inputs"][2][0] = '="User B" * [Accounting Cutoff][-1]'
        formulas["inputs"][2][1] = "=[Accounting Cutoff][1]"
        method = recalculate_dfm_method(
            payload,
            input_snapshot=input_snapshot(),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="same",
        )
        # Without resolved reference values, the stored evaluations survive.
        self.assertEqual(method["ratios_tab"]["average_formulas"]["values"][2][:2], [9, 9])

        tokens = dfm_dataset_reference_tokens(method)
        self.assertEqual(
            [(token["match"], token["dataset_name"], token["row_idx"], token["col_idx"]) for token in tokens],
            [
                ("[Accounting Cutoff][-1]", "Accounting Cutoff", "-1", None),
                ("[Accounting Cutoff][1]", "Accounting Cutoff", "1", None),
            ],
        )
        self.assertEqual(
            dataset_reference_tokens('=[Quoted]["2024 Q1", \'12, months\']')[0]["col_idx"],
            "'12, months'",
        )

        refreshed = recalculate_dfm_method(
            method,
            dataset_reference_values={
                "[Accounting Cutoff][-1]": 1.02,
                "[Accounting Cutoff][1]": 1.5,
            },
            timestamp="later",
        )
        values = refreshed["ratios_tab"]["average_formulas"]["values"]
        # "User B" col 0 = Simple-all (1.5) * 1.1 = 1.65; User A = 1.65 * 1.02.
        self.assertEqual(values[2][0], 1.683)
        self.assertEqual(values[2][1], 1.5)

        # A partial mapping keeps the stored evaluation for the missing reference.
        partial = recalculate_dfm_method(
            method,
            dataset_reference_values={"[Accounting Cutoff][1]": 1.5},
            timestamp="later",
        )
        partial_values = partial["ratios_tab"]["average_formulas"]["values"]
        self.assertEqual(partial_values[2][0], 9)
        self.assertEqual(partial_values[2][1], 1.5)

    def test_formulas_with_whitespace_after_equals_still_re_evaluate(self) -> None:
        # The UI stores user-entry formulas as "= expr" with a space after the
        # equals sign; stripping the "=" must not leave leading whitespace that
        # makes ast.parse fail and silently keep the stored evaluation.
        payload = owned_payload()
        formulas = payload["ratios_tab"]["average_formulas"]
        formulas["inputs"][2][0] = '= "User B" * [Accounting Cutoff][-1]'
        formulas["inputs"][3][0] = '= "Simple - all" * 1.1'
        method = recalculate_dfm_method(
            payload,
            input_snapshot=input_snapshot(),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="same",
        )
        # The internal formula re-evaluates even without reference values.
        self.assertEqual(method["ratios_tab"]["average_formulas"]["values"][3][0], 1.65)

        refreshed = recalculate_dfm_method(
            method,
            dataset_reference_values={"[Accounting Cutoff][-1]": 1.02},
            timestamp="later",
        )
        values = refreshed["ratios_tab"]["average_formulas"]["values"]
        # User A col 0 = User B (1.65) * resolved cutoff 1.02.
        self.assertEqual(values[2][0], 1.683)

    def test_round_rounds_half_up_on_the_decimal_text(self) -> None:
        self.assertEqual(round_half_up(2.38625, 4), 2.3863)
        self.assertEqual(round_half_up(-2.38625, 4), -2.3863)
        self.assertEqual(round_half_up(1.5), 2.0)
        self.assertEqual(round_half_up(1.35735, 4), 1.3574)

    def test_an_average_row_enters_a_formula_at_the_methods_decimal_places(self) -> None:
        payload = owned_payload()
        payload["details_tab"]["decimal_places"] = 4
        formulas = payload["ratios_tab"]["average_formulas"]
        formulas["inputs"][3][1] = '= "Volume - all" * 1.1'
        # Volume - all in column 1 is 400/300, a repeating ratio, so the row a
        # reader sees at four decimals is not the row stored at six.
        method = recalculate_dfm_method(
            payload,
            input_snapshot=input_snapshot(values=[[100, 300, 400], [200, 430, None], [400, None, None]]),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="same",
        )
        values = method["ratios_tab"]["average_formulas"]["values"]
        self.assertEqual(values[0][1], 1.333333)
        self.assertEqual(values[3][1], canonical_number(1.3333 * 1.1))

        # The same method printed at six decimals multiplies the stored row.
        method["details_tab"]["decimal_places"] = 6
        wider = recalculate_dfm_method(method, timestamp="later")
        wider_values = wider["ratios_tab"]["average_formulas"]["values"]
        self.assertEqual(wider_values[3][1], canonical_number(1.333333 * 1.1))

    def test_round_in_a_formula_fixes_an_operand_before_it_multiplies(self) -> None:
        payload = owned_payload()
        formulas = payload["ratios_tab"]["average_formulas"]
        # Simple - all is 1.5 in column 0; ROUND to whole numbers makes it 2.
        formulas["inputs"][3][0] = '= ROUND("Simple - all", 0) * 1.1'
        formulas["inputs"][3][1] = '= round("Simple - all") * 1.1'
        formulas["inputs"][2][0] = '= "User B" * ROUND([Accounting Cutoff][-1], 2)'
        method = recalculate_dfm_method(
            payload,
            input_snapshot=input_snapshot(),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="same",
        )
        values = method["ratios_tab"]["average_formulas"]["values"]
        self.assertEqual(values[3][0], 2.2)
        self.assertEqual(values[3][1], canonical_number(round_half_up(values[1][1]) * 1.1))

        refreshed = recalculate_dfm_method(
            method,
            dataset_reference_values={"[Accounting Cutoff][-1]": 1.0249},
            timestamp="later",
        )
        # User A col 0 = User B (2.2) * the cutoff rounded to 1.02.
        self.assertEqual(refreshed["ratios_tab"]["average_formulas"]["values"][2][0], 2.244)

        # Any other function name is not a formula the contract evaluates.
        formulas["inputs"][3][0] = '= FLOOR("Simple - all") * 1.1'
        kept = recalculate_dfm_method(
            payload,
            input_snapshot=input_snapshot(),
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="same",
        )
        self.assertEqual(kept["ratios_tab"]["average_formulas"]["values"][3][0], 8)

    def test_upstream_refresh_preserves_owned_projection_and_recalculates_internal_formulas(self) -> None:
        initial = recalculate_dfm_method(
            owned_payload(), input_snapshot=input_snapshot(), ratio_basis_snapshot=basis_snapshot()
        )
        refreshed_snapshot = input_snapshot(values=[[100, 200, 260], [200, 400, None], [400, None, None]])
        refreshed_snapshot["revision"] = "input:r2"
        refreshed = recalculate_dfm_method(initial, input_snapshot=refreshed_snapshot)

        self.assertEqual(owned_projection(refreshed), owned_projection(initial))
        self.assertEqual(
            refreshed["method_metadata"]["owned_revision"],
            initial["method_metadata"]["owned_revision"],
        )
        formulas = refreshed["ratios_tab"]["average_formulas"]["values"]
        self.assertEqual(formulas[3][0], 2.2)
        self.assertEqual(formulas[2][0], 4.4)
        self.assertEqual(formulas[4][0], 1.25)
        self.assertNotEqual(
            refreshed["method_metadata"]["derived_revision"],
            initial["method_metadata"]["derived_revision"],
        )

    def test_unsupported_benchmark_rows_are_frozen_instead_of_recomputed_as_simple(self) -> None:
        payload = owned_payload()
        formulas = payload["ratios_tab"]["average_formulas"]
        formulas["label"].insert(2, "Benchmark")
        settings = formulas["custom_average_formula_settings"]
        settings["average_type"].insert(2, "custom")
        settings["base"].insert(2, "benchmark")
        settings["periods"].insert(2, "all")
        settings["exclude"].insert(2, 0)
        formulas["selected"].insert(2, [0, 0, 0])
        formulas["values"].insert(2, [1.7, 1.6, 1.0])
        formulas["inputs"].insert(2, ["", "", ""])
        initial = recalculate_dfm_method(
            payload, input_snapshot=input_snapshot(), ratio_basis_snapshot=basis_snapshot()
        )
        refreshed = recalculate_dfm_method(
            initial,
            input_snapshot=input_snapshot(values=[[100, 300, 600], [200, 500, None], [400, None, None]]),
        )
        benchmark_row = refreshed["ratios_tab"]["average_formulas"]["label"].index("Benchmark")
        self.assertEqual(
            refreshed["ratios_tab"]["average_formulas"]["values"][benchmark_row],
            [1.7, 1.6, 1.0],
        )
        self.assertEqual(
            refreshed["ratios_tab"]["average_formulas"]["custom_average_formula_settings"]["base"][benchmark_row],
            "benchmark",
        )

    def test_preview_preserves_refresh_timestamp(self) -> None:
        method = recalculate_dfm_method(
            owned_payload(), input_snapshot=input_snapshot(), ratio_basis_snapshot=basis_snapshot(), timestamp="old"
        )
        preview_snapshot = input_snapshot(values=[[100, 160, 180], [200, 300, None], [400, None, None]])
        preview = preview_dfm_method(
            method,
            input_snapshot=preview_snapshot,
            ratio_basis_snapshot=basis_snapshot(),
            timestamp="new",
        )
        self.assertEqual(preview["method_metadata"]["data_refreshed"], "old")

    def test_rejects_geometry_and_ambiguous_or_missing_basis_labels(self) -> None:
        method = recalculate_dfm_method(
            owned_payload(), input_snapshot=input_snapshot(), ratio_basis_snapshot=basis_snapshot()
        )
        changed = input_snapshot()
        changed["development_labels"] = ["12m", "36m"]
        changed["values"] = [[100, 180], [200, None], [400, None]]
        changed["mask"] = [[True, True], [True, False], [True, False]]
        with self.assertRaisesRegex(DfmContractError, "geometry changed"):
            recalculate_dfm_method(method, input_snapshot=changed)

        duplicate = basis_snapshot()
        duplicate["origin_labels"] = ["2020", "2020", "2022"]
        with self.assertRaisesRegex(DfmContractError, "duplicate origin"):
            recalculate_dfm_method(method, ratio_basis_snapshot=duplicate)

        missing = basis_snapshot()
        missing["origin_labels"] = ["2020", "2022"]
        missing["values"] = [1000, 3000]
        with self.assertRaisesRegex(DfmContractError, "missing exact origin"):
            recalculate_dfm_method(method, ratio_basis_snapshot=missing)

    def test_complete_normalization_rejects_stale_revision_metadata(self) -> None:
        method = recalculate_dfm_method(
            owned_payload(), input_snapshot=input_snapshot(), ratio_basis_snapshot=basis_snapshot()
        )
        edited = deepcopy(method)
        edited["results_tab"]["ultimate_vector"][0] = 999
        with self.assertRaisesRegex(DfmContractError, "revision"):
            normalize_dfm_method(edited)
        self.assertEqual(method_revisions(method)["publication_revision"], method["method_metadata"]["publication_revision"])

    def test_publication_revision_includes_period_and_sidecar_formatting(self) -> None:
        method = recalculate_dfm_method(
            owned_payload(), input_snapshot=input_snapshot(), ratio_basis_snapshot=basis_snapshot()
        )
        patch_payload = deepcopy(method)
        patch_payload["details_tab"]["origin_length"] = 6
        changed_period = apply_owned_patch(method, patch_payload)
        self.assertNotEqual(
            changed_period["method_metadata"]["publication_revision"],
            method["method_metadata"]["publication_revision"],
        )

        patch_payload = deepcopy(method)
        patch_payload["details_tab"]["decimal_places"] = 3
        changed_format = apply_owned_patch(method, patch_payload)
        self.assertNotEqual(
            changed_format["method_metadata"]["publication_revision"],
            method["method_metadata"]["publication_revision"],
        )

    def test_owned_exclusion_patch_rebases_by_exact_labels(self) -> None:
        initial = recalculate_dfm_method(
            owned_payload(), input_snapshot=input_snapshot(), ratio_basis_snapshot=basis_snapshot()
        )
        upstream = input_snapshot(values=[[50, 75, 90], [200, 300, None], [100, 150, 180], [400, None, None]])
        upstream["origin_labels"] = ["2019", "2021", "2020", "2022"]
        upstream["revision"] = "input:r2"
        upstream_basis = {
            **basis_snapshot(),
            "origin_labels": ["2019", "2021", "2020", "2022"],
            "values": [500, 2000, 1000, 3000],
            "revision": "basis:r2",
        }
        refreshed = recalculate_dfm_method(
            initial,
            input_snapshot=upstream,
            ratio_basis_snapshot=upstream_basis,
        )
        stale_patch = deepcopy(initial)
        stale_patch["ratios_tab"]["ratio_triangle"]["excluded"][0][0] = 0
        stale_patch["ratios_tab"]["ratio_triangle"]["excluded"][1][0] = 1

        rebased = apply_owned_patch(refreshed, stale_patch, timestamp="save")
        ratio = rebased["ratios_tab"]["ratio_triangle"]
        rows = dict(zip(ratio["origin_labels"], ratio["excluded"]))
        self.assertEqual(rows["2019"], [0, 0])
        self.assertEqual(rows["2020"][0], 0)
        self.assertEqual(rows["2021"][0], 1)

        case_mismatch = basis_snapshot()
        case_mismatch["origin_labels"] = ["2020", "2021", "2022 "]
        with self.assertRaisesRegex(DfmContractError, "missing exact origin"):
            recalculate_dfm_method(initial, ratio_basis_snapshot=case_mismatch)


if __name__ == "__main__":
    unittest.main()

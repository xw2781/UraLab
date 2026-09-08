from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

_MACRO_PATH = (
    Path(__file__).resolve().parents[1]
    / "macros"
    / "apply_growth_and_cutoff_adjustments.py"
)


def load_macro_module():
    spec = importlib.util.spec_from_file_location(
        "apply_growth_and_cutoff_adjustments_under_test",
        _MACRO_PATH,
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load the growth and cutoff adjustment macro.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


MACRO = load_macro_module()

DEV_LABELS = ["(1) 5-17", "(2) 17-29", "(3) 29-41", "(4) 41-53"]
ORIGIN_LABELS = ["2023", "2024", "2025", "2026"]
AVERAGE_LABELS = ["Volume - all", "Simple - 3", "Simple - 2", "User Entry"]
USER_ENTRY_ROW = AVERAGE_LABELS.index("User Entry")


class FakeDfm:
    """The slice of DfmMethod this macro touches."""

    def __init__(
        self,
        *,
        selected_row=1,
        notes="",
        input_triangle="Claim Counts--CWOP",
        output_category="C Claim Count",
        inputs=None,
        origin_labels=None,
    ):
        columns = len(DEV_LABELS)
        selected = [[0] * columns for _ in AVERAGE_LABELS]
        for col in range(columns):
            selected[selected_row][col] = 1
        self.ratios_tab = {
            "average_formulas": {
                "label": list(AVERAGE_LABELS),
                "selected": selected,
                "values": [
                    [2.0, 1.5, 1.2, 1.1],
                    [3.0, 1.6, 1.3, 1.05],
                    [4.0, 1.7, 1.4, 1.02],
                    [9.0, 9.0, 9.0, 9.0],
                ],
                "inputs": inputs or [[""] * columns for _ in AVERAGE_LABELS],
                "display_inputs": [[""] * columns for _ in AVERAGE_LABELS],
            },
            "ratio_triangle": {"development_labels": DEV_LABELS},
            "cell_notes": {"ratio_main_table": {}, "ratio_summary_table": {}},
        }
        self.data_tab = {"origin_labels": list(origin_labels or ORIGIN_LABELS)}
        self.details = {"output_category": output_category, "decimal_places": 4}
        self.input_triangle = input_triangle
        self.name = "C 22 - CWOP DFM w/ Selected LDFs"
        self.project_name = "Project"
        self.reserving_class = "RC"
        self._notes = notes
        self.applied = []

    @property
    def average_formulas(self):
        return self.ratios_tab["average_formulas"]

    @property
    def decimal_places(self):
        return self.details.get("decimal_places")

    @property
    def cell_notes(self):
        return self.ratios_tab["cell_notes"]

    @property
    def notes(self):
        return self._notes

    def set_cell_note(self, row_label, dev_period, note, table="ratio_summary_table"):
        column_label = self.dev_period(dev_period)
        notes_by_row = self.cell_notes.setdefault(table, {})
        row_notes = notes_by_row.setdefault(row_label, {})
        if note:
            row_notes[column_label] = note
        else:
            row_notes.pop(column_label, None)
            if not row_notes:
                notes_by_row.pop(row_label, None)

    def clear_cell_notes_for_development(self, dev_period, table="ratio_summary_table"):
        column_label = self.dev_period(dev_period)
        notes_by_row = self.cell_notes.setdefault(table, {})
        for row_label, row_notes in list(notes_by_row.items()):
            row_notes.pop(column_label, None)
            if not row_notes:
                notes_by_row.pop(row_label, None)

    def _average_col_count(self):
        return len(DEV_LABELS)

    def _ensure_average_label(self, label):
        return AVERAGE_LABELS.index(label)

    def dev_period(self, index):
        return DEV_LABELS[index - 1]

    def set_user_formula(self, formula, value, dev_period):
        self.applied.append((dev_period, formula, value))
        self.average_formulas["inputs"][USER_ENTRY_ROW][dev_period - 1] = formula
        self.average_formulas["values"][USER_ENTRY_ROW][dev_period - 1] = value
        for row in self.average_formulas["selected"]:
            row[dev_period - 1] = 0
        self.average_formulas["selected"][USER_ENTRY_ROW][dev_period - 1] = 1
        return self

    def to_dict(self):
        return {}


def make_resolver(rows_by_name):
    """rows_by_name: {dataset_name: {row_idx: (value, row_label)}}."""

    def resolver(_project, _rc, references):
        results = []
        for reference in references:
            rows = rows_by_name.get(reference["dataset_name"], {})
            entry = rows.get(reference["row_idx"])
            if entry is None:
                results.append(None)
                continue
            results.append({
                "dataset_name": reference["dataset_name"],
                "data_format": "Vector",
                "row_label": entry[1],
                "col_label": "Ultimate",
                "value": entry[0],
            })
        return results, []

    return resolver


ANNUAL_ROWS = {
    "Accounting Cutoff": {"-1": (1.0117, "2026"), "-2": (1.0, "2025"), "-3": (1.0, "2024")},
    "Growth Adjustment--Counts": {"-1": (1.0426, "2026"), "-2": (1.0, "2025"), "-3": (1.0, "2024")},
    "Growth Adjustment--Incurred": {"-1": (1.0438, "2026"), "-2": (1.0, "2025"), "-3": (1.0, "2024")},
    "Growth Adjustment--Paid": {"-1": (1.0456, "2026"), "-2": (1.0, "2025"), "-3": (1.0, "2024")},
}


def plan(dfm, rows=None):
    basis = MACRO.adjustment_basis(dfm)
    return basis, MACRO.plan_adjustments(
        dfm,
        basis,
        resolver=make_resolver(rows or ANNUAL_ROWS),
        project_name="Project",
        reserving_class="RC",
    )


class AdjustmentBasisTests(unittest.TestCase):
    def test_claim_counts_read_the_counts_vector(self):
        basis = MACRO.adjustment_basis(FakeDfm())
        self.assertEqual(basis["kind"], "counts")
        self.assertEqual(basis["growth"], [("*", "Growth Adjustment--Counts")])
        self.assertTrue(basis["cutoff"])

    def test_paid_and_incurred_read_their_own_vectors(self):
        paid = MACRO.adjustment_basis(
            FakeDfm(input_triangle="Gross Loss--Paid", output_category="D Gross Loss")
        )
        self.assertEqual(paid["growth"], [("*", "Growth Adjustment--Paid")])
        incurred = MACRO.adjustment_basis(
            FakeDfm(input_triangle="Net Loss--Incurred", output_category="F Net Loss")
        )
        self.assertEqual(incurred["growth"], [("*", "Growth Adjustment--Incurred")])

    def test_severity_divides_incurred_by_counts_and_drops_the_cutoff(self):
        basis = MACRO.adjustment_basis(
            FakeDfm(
                input_triangle="Severity--Gross Incurred per Reported ex CWOP",
                output_category="H Severity",
            )
        )
        self.assertEqual(
            basis["growth"],
            [("*", "Growth Adjustment--Incurred"), ("/", "Growth Adjustment--Counts")],
        )
        self.assertFalse(basis["cutoff"])

    def test_a_ratio_of_two_bases_is_left_alone(self):
        basis = MACRO.adjustment_basis(
            FakeDfm(input_triangle="CWOP as % of Reported Claims", output_category="C Claim Count")
        )
        self.assertEqual(basis["growth"], [])
        self.assertIn("percentage", basis["reason"])


class FormulaTests(unittest.TestCase):
    def test_the_first_period_carries_the_cutoff_and_the_growth_factor(self):
        dfm = FakeDfm(selected_row=2)
        _basis, result = plan(dfm)
        self.assertEqual(len(result["plans"]), 1)
        first = result["plans"][0]
        self.assertEqual(
            first["formula"],
            '= "Simple - 2" * [Accounting Cutoff][-1] * [Growth Adjustment--Counts][-1]',
        )
        self.assertEqual(
            first["display_formula"],
            '= "Simple - 2" * [Accounting Cutoff][2026] * [Growth Adjustment--Counts][2026]',
        )
        self.assertAlmostEqual(first["value"], round(4.0 * 1.0117 * 1.0426, 6))

    def test_the_average_factor_enters_the_product_rounded_to_four_decimals(self):
        # 1.35735 reads as 1.3574 in the notes, so 1.3574 is what the vectors
        # multiply: 1.3574 * 1.0117 * 1.0426 rather than 1.35735 * ...
        dfm = FakeDfm(selected_row=2)
        dfm.average_formulas["values"][2][0] = 1.35735
        _basis, result = plan(dfm)
        first = result["plans"][0]
        self.assertEqual(first["base_value"], 1.3574)
        self.assertAlmostEqual(first["value"], round(1.3574 * 1.0117 * 1.0426, 6))

    def test_a_period_whose_factors_are_all_one_is_skipped(self):
        dfm = FakeDfm(selected_row=2)
        _basis, result = plan(dfm)
        self.assertEqual([item["col"] for item in result["plans"]], [0])

    def test_a_gap_in_the_middle_still_adjusts_the_later_period(self):
        rows = {
            "Accounting Cutoff": {"-1": (1.0, "2026"), "-2": (1.0, "2025"), "-3": (1.0, "2024")},
            "Growth Adjustment--Counts": {
                "-1": (1.05, "2026"),
                "-2": (1.0, "2025"),
                "-3": (0.98, "2024"),
            },
        }
        dfm = FakeDfm(selected_row=1)
        _basis, result = plan(dfm, rows)
        self.assertEqual([item["col"] for item in result["plans"]], [0, 2])
        self.assertEqual(
            result["plans"][1]["formula"], '= "Simple - 3" * [Growth Adjustment--Counts][-3]'
        )

    def test_only_the_first_three_periods_are_considered(self):
        rows = {
            "Accounting Cutoff": {f"-{n}": (1.0, str(2027 - n)) for n in range(1, 5)},
            "Growth Adjustment--Counts": {f"-{n}": (1.05, str(2027 - n)) for n in range(1, 5)},
        }
        dfm = FakeDfm(selected_row=1)
        _basis, result = plan(dfm, rows)
        self.assertEqual([item["col"] for item in result["plans"]], [0, 1, 2])

    def test_severity_writes_a_division_term(self):
        dfm = FakeDfm(
            selected_row=1,
            input_triangle="Severity--Gross Incurred per Reported ex CWOP",
            output_category="H Severity",
        )
        _basis, result = plan(dfm)
        self.assertEqual(
            result["plans"][0]["formula"],
            '= "Simple - 3" * [Growth Adjustment--Incurred][-1] / [Growth Adjustment--Counts][-1]',
        )
        self.assertAlmostEqual(result["plans"][0]["value"], round(3.0 * 1.0438 / 1.0426, 6))


class BaseRowRecoveryTests(unittest.TestCase):
    NOTES = (
        "Excluded 2020, 2021 LDFs since they are distorted by COVID.\n"
        "\n"
        "For development period (1) 5-17:\n"
        '  ◦ Apply growth adjustments of 1+4.26% = 1.0426;\n'
        '  ◦ Apply accounting cutoff 1+1.17% = 1.0117;\n'
        '  ◦ Selected average factor: "Simple - 2" (2.8539)\n'
        "  ◦ Selected LDF after adjustments: 2.8539 * 1.0426 * 1.0117 = 3.0102"
    )

    def test_the_notes_name_the_row_behind_an_imported_user_entry_value(self):
        self.assertEqual(
            MACRO.base_labels_from_notes(self.NOTES), {"(1) 5-17": "Simple - 2"}
        )
        dfm = FakeDfm(selected_row=USER_ENTRY_ROW, notes=self.NOTES)
        _basis, result = plan(dfm)
        self.assertEqual(result["plans"][0]["base_label"], "Simple - 2")
        self.assertAlmostEqual(result["plans"][0]["value"], round(4.0 * 1.0117 * 1.0426, 6))

    def test_a_user_entry_value_with_no_note_is_left_alone(self):
        dfm = FakeDfm(selected_row=USER_ENTRY_ROW)
        _basis, result = plan(dfm)
        self.assertEqual(result["plans"], [])
        self.assertTrue(any("(1) 5-17" in text for text in result["skipped"]))

    def test_running_twice_rebuilds_from_the_same_base_row(self):
        dfm = FakeDfm(selected_row=2)
        _basis, first = plan(dfm)
        MACRO.apply_adjustments(dfm, first["plans"])
        _basis, second = plan(dfm)
        self.assertEqual(second["plans"][0]["formula"], first["plans"][0]["formula"])
        self.assertEqual(second["plans"][0]["value"], first["plans"][0]["value"])

    def test_a_formula_this_macro_did_not_write_is_not_treated_as_generated(self):
        self.assertIsNone(MACRO.base_label_from_generated_formula('= "Simple - 2" * 1.0426'))
        self.assertIsNone(MACRO.base_label_from_generated_formula("3.33"))
        self.assertEqual(
            MACRO.base_label_from_generated_formula(
                '= "Simple - 2" * [Accounting Cutoff][-1] * [Growth Adjustment--Counts][-1]'
            ),
            "Simple - 2",
        )
        # A formula written before 1.2.0 opened with the bare label.
        self.assertEqual(
            MACRO.base_label_from_generated_formula(
                '= "Simple - 2" * [Accounting Cutoff][-1] * [Growth Adjustment--Counts][-1]'
            ),
            "Simple - 2",
        )


class CellNoteTests(unittest.TestCase):
    def test_applying_an_adjustment_notes_the_original_cell(self):
        dfm = FakeDfm(selected_row=2)
        _basis, result = plan(dfm)
        MACRO.apply_adjustments(dfm, result["plans"])
        self.assertEqual(
            dfm.cell_notes["ratio_summary_table"]["Simple - 2"]["(1) 5-17"],
            MACRO.PRE_ADJUSTMENT_CELL_NOTE,
        )

    def test_reapplying_on_a_different_base_row_clears_the_old_note(self):
        dfm = FakeDfm(selected_row=2)
        _basis, first = plan(dfm)
        MACRO.apply_adjustments(dfm, first["plans"])

        # The actuary re-selects a different average row for the same period.
        for row in dfm.average_formulas["selected"]:
            row[0] = 0
        dfm.average_formulas["selected"][1][0] = 1  # "Simple - 3"

        _basis, second = plan(dfm)
        MACRO.apply_adjustments(dfm, second["plans"])

        notes = dfm.cell_notes["ratio_summary_table"]
        self.assertNotIn("Simple - 2", notes)
        self.assertEqual(notes["Simple - 3"]["(1) 5-17"], MACRO.PRE_ADJUSTMENT_CELL_NOTE)


class OriginGridTests(unittest.TestCase):
    def test_an_annual_method_will_not_read_quarterly_vectors(self):
        quarterly = {
            "Accounting Cutoff": {"-1": (1.0148, "2026 Q2")},
            "Growth Adjustment--Counts": {"-1": (1.03, "2026 Q2")},
        }
        dfm = FakeDfm(selected_row=1)
        _basis, result = plan(dfm, quarterly)
        self.assertEqual(result["plans"], [])
        self.assertIn("2026 Q2", result["grid_mismatch"])

    def test_a_quarterly_method_reads_quarterly_vectors(self):
        quarterly = {
            "Accounting Cutoff": {"-1": (1.0148, "2026 Q2")},
            "Growth Adjustment--Counts": {"-1": (1.03, "2026 Q2")},
        }
        dfm = FakeDfm(
            selected_row=1,
            origin_labels=["2025 Q3", "2025 Q4", "2026 Q1", "2026 Q2"],
        )
        _basis, result = plan(dfm, quarterly)
        self.assertEqual(result["grid_mismatch"], "")
        self.assertEqual(len(result["plans"]), 1)


class RealDfmCellNoteTableTests(unittest.TestCase):
    """The macro leans on the DfmMethod helpers' default table name.

    A production run failed with "Unknown DFM cell-note table:
    'ratio_summary_table'" because the resolver folded underscores to spaces
    before looking the name up in a table that still spelled it with
    underscores, so its own default could never match.
    """

    def test_helper_default_table_names_resolve(self):
        from arcrho_api.dfm import _cell_note_table_name

        for spelled in ("ratio_summary_table", "ratio-summary-table", "Ratio Summary Table", "summary", "average_formulas"):
            self.assertEqual(_cell_note_table_name(spelled), "ratio_summary_table", spelled)
        for spelled in ("ratio_main_table", "ratio main", "main", "ratio"):
            self.assertEqual(_cell_note_table_name(spelled), "ratio_main_table", spelled)

    def test_unknown_table_name_is_still_rejected(self):
        from arcrho_api.dfm import _cell_note_table_name
        from arcrho_api.exceptions import DfmDataError

        with self.assertRaises(DfmDataError):
            _cell_note_table_name("ultimate_vector")


if __name__ == "__main__":
    unittest.main()

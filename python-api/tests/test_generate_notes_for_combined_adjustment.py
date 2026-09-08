from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

_MACRO_PATH = (
    Path(__file__).resolve().parents[1]
    / "macros"
    / "generate_notes_for_combined_adjustment.py"
)


def load_macro_module():
    spec = importlib.util.spec_from_file_location(
        "generate_notes_for_combined_adjustment_under_test",
        _MACRO_PATH,
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load the combined-adjustment notes macro.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


MACRO = load_macro_module()

DEV_LABELS = ["12-24", "24-36", "36-48", "48-60"]


class FakeDfm:
    def __init__(self, formulas, notes=""):
        self.ratios_tab = {
            "average_formulas": formulas,
            "ratio_triangle": {"development_labels": DEV_LABELS},
        }
        self._notes = notes
        self.project_name = "Project"
        self.reserving_class = "RC"
        self.decimal_places = 4

    @property
    def average_formulas(self):
        return self.ratios_tab["average_formulas"]

    def _average_col_count(self):
        selected = self.average_formulas.get("selected") or []
        return max((len(row) for row in selected), default=0)

    def dev_period(self, index):
        return DEV_LABELS[index - 1]

    @property
    def notes(self):
        return self._notes

    def update_notes(self, text):
        self._notes = str(text or "")
        return self

    def add_notes(self, text, append=True, add_space=None):
        text = str(text or "")
        self._notes = f"{self._notes}\n\n{text}" if self._notes and append else text
        return self

    def to_dict(self):
        return {}


def make_resolver(values_by_name):
    """values_by_name: {dataset_name: (value, row_label) or (value, row_label, col_label)}"""

    def resolver(project, rc, references):
        results = []
        errors = []
        for reference in references:
            spec = values_by_name.get(reference["dataset_name"])
            if spec is None:
                results.append(None)
                errors.append(f"[{reference['dataset_name']}]: not found")
                continue
            value, row_label = spec[0], spec[1]
            col_label = spec[2] if len(spec) > 2 else None
            results.append({
                "dataset_name": reference["dataset_name"],
                "data_format": "Triangle" if col_label else "Vector",
                "row_label": row_label,
                "col_label": col_label,
                "value": value,
            })
        return results, errors

    return resolver


def single_column_dfm(formula, user_value, base_label="Simple - 2", base_value=3.0414):
    return FakeDfm({
        "label": [base_label, "User Entry"],
        "selected": [[0], [1]],
        "values": [[base_value], [user_value]],
        "inputs": [[""], [formula]],
    })


def generate(dfm, resolver):
    return MACRO.generate_combined_adjustment_notes(dfm, resolver=resolver)


class ReferenceParsingTests(unittest.TestCase):
    def test_vector_and_triangle_references(self):
        refs = MACRO.find_dataset_references(
            '=[Paid Claims][2024, 12] / [Earned Premium][1] + [Quoted]["2024 Q1", \'12, months\']'
        )
        self.assertEqual(
            [(r["dataset_name"], r["row_idx"], r["col_idx"]) for r in refs],
            [
                ("Paid Claims", "2024", "12"),
                ("Earned Premium", "1", None),
                ("Quoted", '"2024 Q1"', "'12, months'"),
            ],
        )

    def test_excel_workbook_brackets_are_not_references(self):
        self.assertEqual(
            MACRO.find_dataset_references("='C:\\Data\\[Book.xlsx]Sheet'!A1"), []
        )

    def test_split_product_terms(self):
        self.assertEqual(
            MACRO.split_product_terms('"Simple - 2" * [A][-1] / [B][-1] * 1.05'),
            [("*", '"Simple - 2"'), ("*", "[A][-1]"), ("/", "[B][-1]"), ("*", "1.05")],
        )
        # Parenthesized additive groups survive the split but fail classification.
        terms = MACRO.split_product_terms('"Simple - 2" * ([A][-1] + 0.01)')
        self.assertEqual(terms[1], ("*", "([A][-1] + 0.01)"))
        self.assertIsNone(MACRO.classify_term("([A][-1] + 0.01)"))
        self.assertIsNone(MACRO.split_product_terms("[A][-1] + 0.01"))


class NoteGenerationTests(unittest.TestCase):
    def test_two_reference_factors(self):
        dfm = single_column_dfm(
            '= "Simple - 2" * [Accounting Cutoff][-1] * [C 01 - Growth Adjustment][-1]',
            3.0851,
        )
        resolver = make_resolver({
            "Accounting Cutoff": (1.0117, "2026"),
            "C 01 - Growth Adjustment": (1.0026, "2026"),
        })
        result = generate(dfm, resolver)
        self.assertEqual(len(result["note_blocks"]), 1)
        note = result["note_blocks"][0]
        self.assertIn("For development period 12-24:", note)
        self.assertIn("  - Apply accounting cutoff of 1+1.17% = 1.0117;", note)
        self.assertIn("  - Apply growth adjustment of 1+0.26% = 1.0026;", note)
        self.assertNotIn("[Accounting Cutoff] @", note)
        self.assertIn('  - Selected average factor: "Simple - 2" (3.0414)', note)
        self.assertIn(
            "  - Selected LDF after adjustments: 3.0414 * 1.0117 * 1.0026 = 3.0851",
            note,
        )

    def test_a_rounded_base_is_shown_and_multiplied_at_that_precision(self):
        self.assertEqual(
            MACRO.classify_term('ROUND("Simple - 2", 4)'),
            {"kind": "label", "label": "Simple - 2", "round_digits": 4},
        )
        # The cell holds 1.3574 * 0.9949 = 1.350478, and the note's arithmetic
        # reproduces it: 1.3505, not the 1.3504 that 1.35735 * 0.9949 gives.
        dfm = single_column_dfm(
            '= ROUND("Simple - 2", 4) * [Accounting Cutoff][-1]',
            1.350478,
            base_value=1.35735,
        )
        resolver = make_resolver({"Accounting Cutoff": (0.9949, "2026")})
        note = generate(dfm, resolver)["note_blocks"][0]
        self.assertIn('  - Selected average factor: "Simple - 2" (1.3574)', note)
        self.assertIn("  - Selected LDF after adjustments: 1.3574 * 0.9949 = 1.3505", note)

    def test_a_plain_base_is_shown_at_the_methods_decimal_places(self):
        # The formula names no ROUND of its own, so the Ratios tab read the row
        # at the method's four decimals and the cell holds 1.3574 * 0.9949. The
        # note must state the same 1.3574 rather than the stored 1.35735.
        dfm = single_column_dfm(
            '= "Simple - 2" * [Accounting Cutoff][-1]',
            1.350478,
            base_value=1.35735,
        )
        resolver = make_resolver({"Accounting Cutoff": (0.9949, "2026")})
        note = generate(dfm, resolver)["note_blocks"][0]
        self.assertIn('  - Selected average factor: "Simple - 2" (1.3574)', note)
        self.assertIn("  - Selected LDF after adjustments: 1.3574 * 0.9949 = 1.3505", note)

    def test_unity_factor_is_omitted(self):
        dfm = single_column_dfm(
            '= "Simple - 2" * [Accounting Cutoff][-1] * [C 01 - Growth Adjustment][-1]',
            3.0493,
        )
        resolver = make_resolver({
            "Accounting Cutoff": (1.0, "2025"),
            "C 01 - Growth Adjustment": (1.0026, "2026"),
        })
        note = generate(dfm, resolver)["note_blocks"][0]
        self.assertNotIn("accounting cutoff", note)
        self.assertIn("Apply growth adjustment of 1+0.26%", note)
        self.assertIn(
            "  - Selected LDF after adjustments: 3.0414 * 1.0026 = 3.0493", note
        )

    def test_all_factors_unity_produces_no_note(self):
        dfm = single_column_dfm(
            '= "Simple - 2" * [Accounting Cutoff][-1]', 3.0414
        )
        resolver = make_resolver({"Accounting Cutoff": (1.0, "2025")})
        result = generate(dfm, resolver)
        self.assertEqual(result["note_blocks"], [])

    def test_three_factors_with_division_and_number(self):
        value = 3.0414 * 1.0117 / 1.0026 * 1.05
        dfm = single_column_dfm(
            '= "Simple - 2" * [Accounting Cutoff][-1] / [C 01 - Growth Adjustment][-1] * 1.05',
            value,
        )
        resolver = make_resolver({
            "Accounting Cutoff": (1.0117, "2026"),
            "C 01 - Growth Adjustment": (1.0026, "2026"),
        })
        note = generate(dfm, resolver)["note_blocks"][0]
        self.assertIn("Apply accounting cutoff of 1+1.17%", note)
        self.assertIn("Apply growth adjustment of 1/(1+0.26%) = 0.9974;", note)
        self.assertIn("Apply other adjustment of 1+5% = 1.05;", note)
        self.assertIn(f"= {value:.4f}", note)

    def test_negative_adjustment_percent(self):
        dfm = single_column_dfm('= "Simple - 2" * [Growth Adjustment][-1]', 3.0058)
        resolver = make_resolver({"Growth Adjustment": (0.9883, "2026")})
        note = generate(dfm, resolver)["note_blocks"][0]
        self.assertIn("Apply growth adjustment of 1-1.17%", note)

    def test_no_base_label_reference_only(self):
        dfm = single_column_dfm("= [Selected LDF][-1, 2]", 1.5)
        resolver = make_resolver({"Selected LDF": (1.5, "2026", "24")})
        note = generate(dfm, resolver)["note_blocks"][0]
        self.assertIn("Apply selected ldf adjustment of 1+50% = 1.5;", note)
        self.assertIn("Selected LDF after adjustments: 1.5 = 1.5000", note)
        self.assertNotIn("Selected average factor", note)

    def test_complex_formula_falls_back_to_resolved_note(self):
        formula = '= "Simple - 2" * ([Accounting Cutoff][-1] + 0.001)'
        dfm = single_column_dfm(formula, 3.0805)
        resolver = make_resolver({"Accounting Cutoff": (1.0117, "2026")})
        note = generate(dfm, resolver)["note_blocks"][0]
        self.assertIn(f"  - User Entry formula: {formula};", note)
        self.assertIn(
            "  - Resolved references: [Accounting Cutoff] @ 2026 = 1.0117;", note
        )
        self.assertIn("  - Selected LDF after adjustments: 3.0805", note)

    def test_plain_number_entry_is_ignored(self):
        dfm = single_column_dfm("3.05", 3.05)
        result = generate(dfm, make_resolver({}))
        self.assertEqual(result["note_blocks"], [])

    def test_unresolved_reference_falls_back_and_reports_error(self):
        dfm = single_column_dfm('= "Simple - 2" * [Missing][-1]', 3.1)
        result = generate(dfm, make_resolver({}))
        self.assertEqual(len(result["errors"]), 1)
        note = result["note_blocks"][0]
        self.assertIn("User Entry formula", note)

    def test_notes_regeneration_is_idempotent(self):
        dfm = single_column_dfm(
            '= "Simple - 2" * [Accounting Cutoff][-1]', 3.0770
        )
        dfm.update_notes("Keep this actuary comment.")
        resolver = make_resolver({"Accounting Cutoff": (1.0117, "2026")})
        for _ in range(2):
            result = generate(dfm, resolver)
            MACRO.apply_notes(dfm, result["note_blocks"])
        self.assertEqual(dfm.notes.count("For development period 12-24:"), 1)
        self.assertIn("Keep this actuary comment.", dfm.notes)

    def test_clearing_removes_legacy_blocks_with_other_bullet_styles(self):
        dfm = single_column_dfm(
            '= "Simple - 2" * [Accounting Cutoff][-1]', 3.0770
        )
        dfm.update_notes(
            "Excluded 2020, 2021 LDFs since they are distorted by COVID.\n\n"
            "For development period (1) 5-17:\n"
            "  ◦ Apply growth adjustments of 1+4.26% = 1.0426;\n"
            "  ◦ Apply accounting cutoff 1+1.17% = 1.0117;\n"
            "  ◦ Selected average factor: \"Simple - 2\" (2.8539)\n"
            "  ◦ Selected LDF after adjustments: 2.8539 * 1.0426 * 1.0117 = 3.0102"
        )
        resolver = make_resolver({"Accounting Cutoff": (1.0117, "2026")})
        result = generate(dfm, resolver)
        MACRO.apply_notes(dfm, result["note_blocks"])
        self.assertIn("Excluded 2020, 2021 LDFs", dfm.notes)
        self.assertNotIn("◦", dfm.notes)
        self.assertNotIn("1+4.26%", dfm.notes)
        self.assertEqual(dfm.notes.count("For development period"), 1)

    def test_no_formulas_adds_no_adjustment_note(self):
        dfm = FakeDfm({
            "label": ["Simple - 2"],
            "selected": [[1]],
            "values": [[3.0414]],
            "inputs": [[""]],
        })
        outcome = MACRO.run_macro(dfm)
        self.assertTrue(outcome["success"])
        self.assertIn(MACRO.NO_ADJUSTMENT_NOTE, dfm.notes)
        self.assertTrue(outcome["preview"]["has_changes"])

    def test_multiple_columns(self):
        dfm = FakeDfm({
            "label": ["Simple - 2", "User Entry"],
            "selected": [[0, 1], [1, 0]],
            "values": [[3.0414, 1.8], [3.0770, None]],
            "inputs": [["", ""], ['= "Simple - 2" * [Accounting Cutoff][-1]', ""]],
        })
        resolver = make_resolver({"Accounting Cutoff": (1.0117, "2026")})
        result = generate(dfm, resolver)
        self.assertEqual(len(result["note_blocks"]), 1)
        self.assertIn("For development period 12-24:", result["note_blocks"][0])


if __name__ == "__main__":
    unittest.main()

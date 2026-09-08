"""The combined-adjustment formula and note lines shared by the two macros and the ResQ import."""

from __future__ import annotations

from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from arcrho_api import combined_adjustment as shared  # noqa: E402


def _block(period, apply_lines, base_label, base_value, product, value):
    lines = [shared.note_header(period)]
    lines.extend(apply_lines)
    lines.append(shared.base_factor_line(base_label, base_value))
    lines.append(shared.selected_ldf_line(product, value))
    return "\n".join(lines)


class FormulaTests(unittest.TestCase):
    def test_each_period_reads_one_row_further_back(self):
        self.assertEqual(
            shared.adjustment_formula(
                "Simple - 2",
                [("*", "Accounting Cutoff"), ("/", "Growth Adjustment--Counts")],
                2,
            ),
            '= "Simple - 2" * [Accounting Cutoff][-3] / [Growth Adjustment--Counts][-3]',
        )
        self.assertEqual(shared.adjustment_formula("Simple - 2", [], 0), '= "Simple - 2"')

    def test_how_a_note_describes_a_dataset(self):
        self.assertEqual(shared.adjustment_description("Accounting Cutoff"), "accounting cutoff")
        self.assertEqual(shared.adjustment_description("Growth Adjustment--Counts"), "growth adjustment--counts")
        self.assertEqual(shared.adjustment_description("C 01 - Growth Adjustment"), "growth adjustment")
        self.assertEqual(shared.adjustment_description("Selected LDF"), "selected ldf adjustment")
        self.assertEqual(shared.adjustment_description(""), "other adjustment")


class NoteReadingTests(unittest.TestCase):
    def test_the_lines_the_notes_macro_writes_read_back_as_formula_terms(self):
        notes = "Excluded 2020, 2021 LDFs.\n\n" + _block(
            "(1) 5-17",
            [
                shared.apply_line(shared.adjustment_description("Accounting Cutoff"), "1+1.17%", "1.0117"),
                shared.apply_line(
                    shared.adjustment_description("Growth Adjustment--Counts"), "1/(1+4.26%)", "0.9591"
                ),
            ],
            "Simple - 2",
            2.8539,
            "2.8539 * 1.0117 * 0.9591",
            2.7693,
        )
        self.assertEqual(
            shared.parse_adjustment_notes(notes),
            {
                "(1) 5-17": {
                    "base_label": "Simple - 2",
                    "terms": [("*", "Accounting Cutoff"), ("/", "Growth Adjustment--Counts")],
                    "value": 2.7693,
                }
            },
        )

    def test_resq_line_breaks_and_other_bullets_read_the_same(self):
        notes = (
            "For development period 12-24:\r\n"
            "  ◦ Apply growth adjustment--paid of 1-2.1% = 0.979;\r\n"
            '  ◦ Selected average factor: "Volume - all" (1.2000)\r\n'
            "  ◦ Selected LDF after adjustments: 1.2000 * 0.979 = 1.1748\r\n"
        )
        parsed = shared.parse_adjustment_notes(notes)
        self.assertEqual(parsed["12-24"]["terms"], [("*", "Growth Adjustment--Paid")])
        self.assertEqual(parsed["12-24"]["value"], 1.1748)

    def test_a_factor_no_adjustment_dataset_describes_cannot_be_rebuilt(self):
        notes = _block(
            "12-24",
            [
                shared.apply_line("accounting cutoff", "1+1.17%", "1.0117"),
                shared.apply_line("other adjustment", "1+5%", "1.05"),
            ],
            "Simple - 2",
            3.0414,
            "3.0414 * 1.0117 * 1.05",
            3.2307,
        )
        parsed = shared.parse_adjustment_notes(notes)
        self.assertEqual(parsed["12-24"]["base_label"], "Simple - 2")
        self.assertIsNone(parsed["12-24"]["terms"])

    def test_the_legacy_note_style_still_names_the_base_row(self):
        notes = (
            "For development period (1) 5-17:\n"
            "  ◦ Apply growth adjustments of 1+4.26% = 1.0426;\n"
            "  ◦ Apply accounting cutoff 1+1.17% = 1.0117;\n"
            '  ◦ Selected average factor: "Simple - 2" (2.8539)\n'
            "  ◦ Selected LDF after adjustments: 2.8539 * 1.0426 * 1.0117 = 3.0102"
        )
        parsed = shared.parse_adjustment_notes(notes)
        self.assertEqual(parsed["(1) 5-17"]["base_label"], "Simple - 2")
        self.assertIsNone(parsed["(1) 5-17"]["terms"])
        self.assertEqual(parsed["(1) 5-17"]["value"], 3.0102)

    def test_text_outside_a_period_block_is_ignored(self):
        self.assertEqual(shared.parse_adjustment_notes('Selected average factor: "Simple - 2" (2.8539)'), {})
        self.assertEqual(shared.parse_adjustment_notes(""), {})


if __name__ == "__main__":
    unittest.main()

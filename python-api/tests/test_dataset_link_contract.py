"""The dataset-link grammar, canonical text, edges, and evaluator.

``arcrho_api.dataset_link_contract`` owns what a dataset cell link means for
every producer — the app-server services delegate to it, the ResQ migration
translates instance formulas through it, and the Engine's dependent walk
evaluates with it — so these tests pin the grammar, the canonical spellings,
the instance-level edge extraction, and the Excel array arithmetic the browser
mirror (``ui/shared/dataset/dataset_formula.js``) implements token for token.
"""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from arcrho_api.dataset_link_contract import (  # noqa: E402
    DatasetLinkError,
    canonical_dataset_formula,
    canonical_internal_reference,
    compact_sidecar_links,
    evaluate_dataset_formula,
    excel_column_name,
    expand_sidecar_links,
    formula_has_excel_reference,
    formula_reference_tokens,
    link_precedent_names,
    parse_dataset_formula_tree,
    tokenize_dataset_formula,
)
from arcrho_api.io import persisted_json_text  # noqa: E402
from arcrho_api.sidecar_core_contract import finalize_sidecar  # noqa: E402


class CanonicalTextTests(unittest.TestCase):
    def test_internal_reference_normalizes_spacing(self):
        self.assertEqual(
            canonical_internal_reference("  [ C 82 ][ 1 : 6 ]  "),
            "=[C 82][1:6]",
        )
        self.assertEqual(
            canonical_internal_reference("=[Paid Claims][2 , 1:2]"),
            "=[Paid Claims][2, 1:2]",
        )

    def test_formula_round_trips_the_service_spellings(self):
        cases = [
            ("=[C 82][1:6]*2", "=[C 82][1:6] * 2"),
            ("[a][1] + [b][2]", "=[a][1] + [b][2]"),
            ("=-[a][1:3] ^ 2", "=-[a][1:3] ^ 2"),
            (
                "=([C 91 - Current Qtr Indicated][1:7] * [H 01][1:7]) / 1000",
                "=([C 91 - Current Qtr Indicated][1:7] * [H 01][1:7]) / 1000",
            ),
        ]
        for raw, expected in cases:
            with self.subTest(raw=raw):
                self.assertEqual(canonical_dataset_formula(raw), expected)

    def test_rejects_text_outside_the_grammar(self):
        for raw in ("=1 + 1", "=[a][1] %", "=[a]", "=([a][1]"):
            with self.subTest(raw=raw):
                with self.assertRaises(DatasetLinkError):
                    canonical_dataset_formula(raw)


class EdgeExtractionTests(unittest.TestCase):
    def test_link_precedent_names_reads_both_link_kinds_once_each(self):
        names = link_precedent_names(
            [
                {"reference": "=[C 82 - Prior Qtr Selected][1:6]"},
                {"reference": "=[C 82 - Prior Qtr Selected][2]"},
            ],
            [
                {"formula": "=[H 01][1:7] * [C 91][1:7] / 1000"},
                {"formula": "='C:\\\\F\\\\[B.xlsx]S1'!A1:A7 * [H 01][1:7]"},
            ],
        )
        self.assertEqual(names, ["C 82 - Prior Qtr Selected", "H 01", "C 91"])

    def test_excel_references_contribute_no_edge_and_bad_text_is_skipped(self):
        self.assertEqual(
            link_precedent_names(
                [{"reference": "not a reference"}],
                [{"formula": "='C:\\\\F\\\\[B.xlsx]S1'!A1"}],
            ),
            [],
        )

    def test_formula_has_excel_reference(self):
        self.assertTrue(formula_has_excel_reference("='C:\\\\F\\\\[B.xlsx]S1'!A1 + [a][1]"))
        self.assertFalse(formula_has_excel_reference("=[a][1] * 2"))
        self.assertFalse(formula_has_excel_reference("not a formula"))


class EvaluationTests(unittest.TestCase):
    def _evaluate(self, text, matrices):
        tokens = tokenize_dataset_formula(text)
        tree = parse_dataset_formula_tree(tokens)
        return evaluate_dataset_formula(
            tree,
            lambda token: matrices.get(token["canonical"]),
        )

    def test_elementwise_and_scalar_broadcast(self):
        matrices = {
            "[a][1:3]": {"rows": 3, "cols": 1, "values": [[2.0], [4.0], [6.0]]},
            "[b][1:3]": {"rows": 3, "cols": 1, "values": [[1.0], [2.0], [3.0]]},
        }
        result = self._evaluate("=[a][1:3] * [b][1:3] / 2", matrices)
        self.assertEqual(result["values"], [[1.0], [4.0], [9.0]])

    def test_blank_reads_as_zero_like_excel(self):
        matrices = {"[a][1:2]": {"rows": 2, "cols": 1, "values": [[None], [3.0]]}}
        result = self._evaluate("=[a][1:2] + 1", matrices)
        self.assertEqual(result["values"], [[1.0], [4.0]])

    def test_one_row_matrix_stretches_across_the_other(self):
        matrices = {
            "[a][1:2, 1:2]": {"rows": 2, "cols": 2, "values": [[1.0, 2.0], [3.0, 4.0]]},
            "[b][1, 1:2]": {"rows": 1, "cols": 2, "values": [[10.0, 20.0]]},
        }
        result = self._evaluate("=[a][1:2, 1:2] + [b][1, 1:2]", matrices)
        self.assertEqual(result["values"], [[11.0, 22.0], [13.0, 24.0]])

    def test_mismatched_shapes_divide_by_zero_and_unary_minus(self):
        matrices = {
            "[a][1:2]": {"rows": 2, "cols": 1, "values": [[1.0], [2.0]]},
            "[b][1:3]": {"rows": 3, "cols": 1, "values": [[1.0], [2.0], [3.0]]},
        }
        with self.assertRaises(DatasetLinkError):
            self._evaluate("=[a][1:2] + [b][1:3]", matrices)
        with self.assertRaises(DatasetLinkError):
            self._evaluate("=[a][1:2] / 0", matrices)
        result = self._evaluate("=-[a][1:2]", matrices)
        self.assertEqual(result["values"], [[-1.0], [-2.0]])

    def test_reference_with_no_values_fails_by_name(self):
        with self.assertRaises(DatasetLinkError) as ctx:
            self._evaluate("=[a][1:2] * 2", {})
        self.assertIn("[a][1:2]", str(ctx.exception))

    def test_reference_tokens_deduplicate_in_formula_order(self):
        tokens = formula_reference_tokens("=[a][1] + [b][1] + [a][1]")
        self.assertEqual([token["canonical"] for token in tokens], ["[a][1]", "[b][1]"])


class TargetCellBlockTests(unittest.TestCase):
    """The on-disk block form of a link's cells, and its round trip."""

    @staticmethod
    def _triangle_link(rows: int, columns: int) -> dict:
        """A link filling a triangle from an Excel range anchored at B7."""

        return {
            "reference": "='C:\\Book\\[Src.xlsx]Sheet1'!B7:DP126",
            "target_cells": [
                {
                    "row": row,
                    "column": column,
                    "source_cell": f"{excel_column_name(column + 1)}{7 + row}",
                }
                for row in range(rows)
                for column in range(columns - row)
            ],
        }

    def test_a_rectangle_becomes_one_block(self):
        link = {
            "reference": "='C:\\Book\\[Src.xlsx]Sheet1'!B7:D8",
            "target_cells": [
                {"row": row, "column": column, "source_cell": f"{excel_column_name(column + 1)}{7 + row}"}
                for row in range(2)
                for column in range(3)
            ],
        }
        compact = compact_sidecar_links({"external_links": [link]})
        self.assertEqual(compact["external_links"][0]["target_cells"], [[0, 0, 2, 3, "B7"]])
        self.assertEqual(expand_sidecar_links(compact), {"external_links": [link]})

    def test_a_triangle_becomes_one_block_per_row_and_round_trips(self):
        payload = {"external_links": [self._triangle_link(4, 4)]}
        compact = compact_sidecar_links({"external_links": [self._triangle_link(4, 4)]})
        self.assertEqual(
            compact["external_links"][0]["target_cells"],
            [[0, 0, 1, 4, "B7"], [1, 0, 1, 3, "B8"], [2, 0, 1, 2, "B9"], [3, 0, 1, 1, "B10"]],
        )
        self.assertEqual(expand_sidecar_links(compact), payload)

    def test_cells_that_do_not_translate_stay_their_own_blocks(self):
        link = {
            "reference": "='C:\\Book\\[Src.xlsx]Sheet1'!A1:Z99",
            "target_cells": [
                {"row": 0, "column": 0, "source_cell": "A1"},
                {"row": 0, "column": 1, "source_cell": "Z9"},
            ],
        }
        compact = compact_sidecar_links({"external_links": [link]})
        self.assertEqual(
            compact["external_links"][0]["target_cells"],
            [[0, 0, 1, 1, "A1"], [0, 1, 1, 1, "Z9"]],
        )
        self.assertEqual(expand_sidecar_links(compact), {"external_links": [link]})

    def test_internal_and_formula_links_carry_a_numeric_source_pair(self):
        payload = {
            "internal_links": [{
                "reference": "=[Paid][1:3, 1:2]",
                "target_cells": [
                    {"row": row, "column": column, "source_row": row + 4, "source_column": column + 9}
                    for row in range(3)
                    for column in range(2)
                ],
            }],
            "formula_links": [{
                "formula": "=[Paid][1:2] * 2",
                "target_cells": [
                    {"row": 0, "column": 0, "result_row": 0, "result_column": 0},
                    {"row": 5, "column": 7, "result_row": 1, "result_column": 0},
                ],
            }],
        }
        compact = compact_sidecar_links(json.loads(json.dumps(payload)))
        self.assertEqual(compact["internal_links"][0]["target_cells"], [[0, 0, 3, 2, 4, 9]])
        self.assertEqual(
            compact["formula_links"][0]["target_cells"],
            [[0, 0, 1, 1, 0, 0], [5, 7, 1, 1, 1, 0]],
        )
        self.assertEqual(expand_sidecar_links(compact), payload)

    def test_the_write_funnel_stores_blocks_and_is_idempotent(self):
        payload = {"dataset_name": "Paid", "external_links": [self._triangle_link(3, 3)]}
        stored = finalize_sidecar(payload)
        self.assertEqual(
            stored["external_links"][0]["target_cells"],
            [[0, 0, 1, 3, "B7"], [1, 0, 1, 2, "B8"], [2, 0, 1, 1, "B9"]],
        )
        self.assertEqual(finalize_sidecar(stored), stored)

    def test_a_block_is_one_line_of_the_persisted_text(self):
        text = persisted_json_text(compact_sidecar_links({"external_links": [self._triangle_link(40, 40)]}))
        per_cell = persisted_json_text({"external_links": [self._triangle_link(40, 40)]})
        self.assertIn('[0, 0, 1, 40, "B7"],', [line.strip() for line in text.splitlines()])
        # 40 blocks against 820 cells: the point of the whole shape.
        self.assertEqual(sum(1 for line in text.splitlines() if line.strip().startswith("[0")), 1)
        self.assertLess(len(text) * 20, len(per_cell))

    def test_a_reader_refuses_the_retired_per_cell_shape(self):
        with self.assertRaises(DatasetLinkError) as ctx:
            expand_sidecar_links({"external_links": [self._triangle_link(2, 2)]})
        self.assertIn("migrate_dataset_link_blocks.py", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()

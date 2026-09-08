"""The one grammar for ArcRho dataset cell links, and its evaluator.

A manual-input dataset cell can take its value from a link rather than a typed
number: a standalone Excel reference, a standalone ArcRho dataset reference, or
an arithmetic formula whose operands are any mix of the two plus numbers::

    ='C:\\Folder\\[Book.xlsx]Sheet1'!A1:A7
    =[C 82 - Prior Qtr Selected][1:7]
    =([Paid Claims][1:6, 2] + 'C:\\Folder\\[Book.xlsx]Sheet1'!B1:B6) / 1000

This module owns the reference syntax, the formula grammar, the canonical
stored text, and the arithmetic semantics (Excel's array rules: a blank cell
reads as zero, a scalar combines with every cell of a matrix, two matrices
combine cell by cell with a one-row or one-column matrix stretched across the
other's shape). ``frontend/ui/shared/dataset/dataset_formula.js`` is the
browser mirror, token for token, so the text a save sends is the text it reads
back; ``app_server/services/dataset_formula_link_service.py`` and
``dataset_internal_link_service.py`` delegate here and translate
:class:`DatasetLinkError` into their HTTP refusals. The ResQ migration and the
Engine's dependent-propagation walk import this module directly, which is what
lets an imported or auto-refreshed formula round-trip byte for byte with one
the user typed.

Value resolution is deliberately not here: what a dataset reference's
coordinates select is owned by the app-server resolvers, and an Excel operand
is read wherever the workbook is reachable. Callers hand :func:`evaluate` a
``lookup`` that turns each reference token into a matrix.
"""

from __future__ import annotations

import math
import re
from typing import Any, Callable, Dict, Iterable, List, Mapping


class DatasetLinkError(ValueError):
    """Raised when link text is not part of the grammar or cannot evaluate."""


INTERNAL_REFERENCE_SYNTAX_HINT = (
    "Use =[Dataset][row] or =[Dataset][start:end] for a vector, and "
    "=[Dataset][row, col] or =[Dataset][rows, cols] for a triangle."
)

DATASET_FORMULA_SYNTAX_HINT = (
    "Enter an Excel link such as ='C:\\Folder\\[Book.xlsx]Sheet1'!A1:C3, a dataset "
    "link such as =[Dataset][1:6], or a formula that combines them with + - * / ^, "
    "for example =[Dataset][1:6] * 1.05."
)

_EXCEL_TOKEN_RE = re.compile(
    r"'((?:[^']|'')*)'!\$?([A-Z]+)\$?([0-9]+)(?::\$?([A-Z]+)\$?([0-9]+))?",
    re.IGNORECASE,
)
_NUMBER_TOKEN_RE = re.compile(r"(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")
_BINARY_PRECEDENCE = {"+": 1, "-": 1, "*": 2, "/": 2, "^": 3}


# ── Internal dataset references ────────────────────────────────────────────────

def _clean_text(value: Any) -> str:
    return str(value if value is not None else "").strip()


def _split_quote_aware(raw: str, separator: str) -> List[str]:
    parts: List[str] = []
    current = ""
    quote = ""
    for character in str(raw or ""):
        if quote:
            current += character
            if character == quote:
                quote = ""
            continue
        if character in {'"', "'"}:
            quote = character
            current += character
            continue
        if character == separator:
            parts.append(current.strip())
            current = ""
            continue
        current += character
    if quote:
        raise DatasetLinkError("Dataset reference contains an unclosed quote.")
    parts.append(current.strip())
    return parts


def _parse_axis_spec(raw: str, *, axis_name: str) -> Dict[str, Any]:
    endpoints = _split_quote_aware(raw, ":")
    if len(endpoints) > 2 or not endpoints[0] or (len(endpoints) == 2 and not endpoints[1]):
        raise DatasetLinkError(
            f"{axis_name.capitalize()} range must be one index or start:end. "
            + INTERNAL_REFERENCE_SYNTAX_HINT,
        )
    return {
        "start": endpoints[0],
        "end": endpoints[1] if len(endpoints) == 2 else None,
    }


def parse_internal_reference(raw_text: Any) -> Dict[str, Any]:
    """Parse one standalone internal reference; raise when the text is not one."""

    text = _clean_text(raw_text)
    if text.startswith("="):
        text = text[1:].lstrip()
    if not text.startswith("["):
        raise DatasetLinkError(INTERNAL_REFERENCE_SYNTAX_HINT)
    name_end = text.find("]", 1)
    if name_end < 0:
        raise DatasetLinkError("Dataset reference is missing its closing bracket.")
    dataset_name = text[1:name_end].strip()
    if not dataset_name:
        raise DatasetLinkError("Dataset reference name cannot be blank.")
    remainder = text[name_end + 1 :].lstrip()
    if not remainder.startswith("["):
        raise DatasetLinkError(INTERNAL_REFERENCE_SYNTAX_HINT)
    quote = ""
    coordinate_end = -1
    for index in range(1, len(remainder)):
        character = remainder[index]
        if quote:
            if character == quote:
                quote = ""
            continue
        if character in {'"', "'"}:
            quote = character
            continue
        if character == "]":
            coordinate_end = index
            break
    if coordinate_end < 0:
        raise DatasetLinkError("Dataset reference is missing its closing bracket.")
    if remainder[coordinate_end + 1 :].strip():
        raise DatasetLinkError(
            "An internal dataset link must be a single standalone reference. "
            + INTERNAL_REFERENCE_SYNTAX_HINT,
        )
    coordinates = _split_quote_aware(remainder[1:coordinate_end], ",")
    if not coordinates[0]:
        raise DatasetLinkError("Dataset reference row index is required.")
    if len(coordinates) > 2 or (len(coordinates) == 2 and not coordinates[1]):
        raise DatasetLinkError(INTERNAL_REFERENCE_SYNTAX_HINT)
    return {
        "dataset_name": dataset_name,
        "row": _parse_axis_spec(coordinates[0], axis_name="row"),
        "col": _parse_axis_spec(coordinates[1], axis_name="column") if len(coordinates) == 2 else None,
    }


def canonical_internal_reference(raw_text: Any) -> str:
    """Return the normalized stored text for a valid internal reference."""

    parsed = parse_internal_reference(raw_text)

    def axis_text(spec: Mapping[str, Any] | None) -> str:
        if not spec:
            return ""
        start = str(spec["start"])
        end = spec.get("end")
        return f"{start}:{end}" if end is not None else start

    coordinates = axis_text(parsed["row"])
    if parsed["col"] is not None:
        coordinates = f"{coordinates}, {axis_text(parsed['col'])}"
    return f"=[{parsed['dataset_name']}][{coordinates}]"


# ── Formula tokens and grammar ─────────────────────────────────────────────────

def _scan_internal_reference(text: str, start: int) -> int:
    """End index (exclusive) of the ``[name][coords]`` reference at ``start``, or -1."""

    name_end = text.find("]", start + 1)
    if name_end < 0:
        return -1
    cursor = name_end + 1
    while cursor < len(text) and text[cursor].isspace():
        cursor += 1
    if cursor >= len(text) or text[cursor] != "[":
        return -1
    quote = ""
    for index in range(cursor + 1, len(text)):
        character = text[index]
        if quote:
            if character == quote:
                quote = ""
            continue
        if character in {'"', "'"}:
            quote = character
        elif character == "]":
            return index + 1
    return -1


def _canonical_excel_reference(match: "re.Match[str]") -> str:
    source = match.group(1).replace("''", "'")
    open_bracket = source.find("[")
    close_bracket = source.find("]", open_bracket + 1)
    if open_bracket < 0 or close_bracket <= open_bracket + 1 or close_bracket >= len(source) - 1:
        raise DatasetLinkError(
            "Excel reference must be written as 'C:\\Folder\\[Book.xlsx]Sheet1'!A1 "
            "or a range such as !A1:C3.",
        )
    start = f"{match.group(2).upper()}{match.group(3)}"
    end = f"{match.group(4).upper()}{match.group(5)}" if match.group(4) else start
    address = start if end == start else f"{start}:{end}"
    return f"'{source.replace(chr(39), chr(39) * 2)}'!{address}"


def tokenize_dataset_formula(raw_text: Any) -> List[Dict[str, Any]]:
    """Split formula text into typed tokens; raise when a character is not part of the grammar."""

    text = str(raw_text if raw_text is not None else "").strip()
    if text.startswith("="):
        text = text[1:]
    tokens: List[Dict[str, Any]] = []
    cursor = 0
    while cursor < len(text):
        character = text[cursor]
        if character.isspace():
            cursor += 1
            continue
        if character == "[":
            end = _scan_internal_reference(text, cursor)
            if end < 0:
                raise DatasetLinkError(
                    "Dataset reference is missing its coordinates. " + INTERNAL_REFERENCE_SYNTAX_HINT,
                )
            reference = text[cursor:end]
            tokens.append({
                "type": "reference",
                "kind": "internal",
                "text": reference,
                "canonical": canonical_internal_reference(reference)[1:],
            })
            cursor = end
            continue
        if character == "'":
            match = _EXCEL_TOKEN_RE.match(text, cursor)
            if not match:
                raise DatasetLinkError(
                    "Excel reference must be written as 'C:\\Folder\\[Book.xlsx]Sheet1'!A1 "
                    "or a range such as !A1:C3.",
                )
            tokens.append({
                "type": "reference",
                "kind": "excel",
                "text": match.group(0),
                "canonical": _canonical_excel_reference(match),
            })
            cursor = match.end()
            continue
        number = _NUMBER_TOKEN_RE.match(text, cursor)
        if number:
            tokens.append({"type": "number", "text": number.group(0)})
            cursor = number.end()
            continue
        if character in "+-*/^":
            tokens.append({"type": "operator", "text": character})
            cursor += 1
            continue
        if character in "()":
            tokens.append({"type": "paren", "text": character})
            cursor += 1
            continue
        raise DatasetLinkError(f'Unexpected "{character}" in the formula. ' + DATASET_FORMULA_SYNTAX_HINT)
    if not tokens:
        raise DatasetLinkError(DATASET_FORMULA_SYNTAX_HINT)
    return tokens


def parse_dataset_formula_tree(tokens: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Parse tokens into an expression tree, marking each unary sign token.

    The walk both validates the grammar and returns the tree :func:`evaluate`
    visits, so the check and the evaluation can never disagree on what a
    formula means. Node kinds mirror the browser parser: ``number``,
    ``reference``, ``unary`` (minus only), and ``binary``.
    """

    position = 0

    def fail(message: str) -> None:
        raise DatasetLinkError(f"{message} {DATASET_FORMULA_SYNTAX_HINT}")

    def peek() -> Dict[str, Any] | None:
        return tokens[position] if position < len(tokens) else None

    def parse_binary(min_precedence: int) -> Dict[str, Any]:
        nonlocal position
        left = parse_unary()
        while True:
            token = peek()
            if not token or token["type"] != "operator":
                return left
            precedence = _BINARY_PRECEDENCE[token["text"]]
            if precedence < min_precedence:
                return left
            position += 1
            # "^" binds to the right, as it does in Excel.
            right = parse_binary(precedence if token["text"] == "^" else precedence + 1)
            left = {"kind": "binary", "operator": token["text"], "left": left, "right": right}

    def parse_unary() -> Dict[str, Any]:
        nonlocal position
        token = peek()
        if token and token["type"] == "operator" and token["text"] in "+-":
            token["unary"] = True
            position += 1
            operand = parse_unary()
            return {"kind": "unary", "operator": "-", "operand": operand} if token["text"] == "-" else operand
        return parse_primary()

    def parse_primary() -> Dict[str, Any]:
        nonlocal position
        token = peek()
        if token is None:
            fail("The formula ends before its last operand.")
        if token["type"] == "number":
            position += 1
            return {"kind": "number", "value": float(token["text"])}
        if token["type"] == "reference":
            position += 1
            return {"kind": "reference", "token": token}
        if token["type"] == "paren" and token["text"] == "(":
            position += 1
            inner = parse_binary(1)
            closing = peek()
            if not closing or closing["type"] != "paren" or closing["text"] != ")":
                fail("The formula is missing a closing parenthesis.")
            position += 1
            return inner
        fail(f'Unexpected "{token["text"]}" in the formula.')
        raise AssertionError("unreachable")

    tree = parse_binary(1)
    if position < len(tokens):
        fail(f'Unexpected "{tokens[position]["text"]}" in the formula.')
    return tree


def format_dataset_formula(tokens: List[Dict[str, Any]]) -> str:
    out = ""
    for token in tokens:
        if token["type"] == "operator":
            out = f"{out}{token['text']}" if token.get("unary") else f"{out.rstrip()} {token['text']} "
        elif token["type"] == "paren":
            out = f"{out}(" if token["text"] == "(" else f"{out.rstrip()})"
        elif token["type"] == "reference":
            out += token["canonical"]
        else:
            out += token["text"]
    return f"={out.rstrip()}"


def canonical_dataset_formula(raw_text: Any) -> str:
    """Return the normalized stored text for a valid formula; raise otherwise."""

    tokens = tokenize_dataset_formula(raw_text)
    parse_dataset_formula_tree(tokens)
    if not any(token["type"] == "reference" for token in tokens):
        raise DatasetLinkError(
            "A formula needs at least one dataset or Excel reference. " + DATASET_FORMULA_SYNTAX_HINT,
        )
    return format_dataset_formula(tokens)


def formula_reference_tokens(raw_text: Any) -> List[Dict[str, Any]]:
    """Each distinct reference token of a formula once, in formula order."""

    tokens = tokenize_dataset_formula(raw_text)
    references: List[Dict[str, Any]] = []
    seen: set[str] = set()
    for token in tokens:
        if token["type"] != "reference":
            continue
        key = f"{token['kind']}\u001f{token['canonical']}"
        if key in seen:
            continue
        seen.add(key)
        references.append(token)
    return references


# ── Link-derived dependency edges ──────────────────────────────────────────────

def link_precedent_names(
    internal_links: Iterable[Mapping[str, Any]] | None,
    formula_links: Iterable[Mapping[str, Any]] | None,
) -> List[str]:
    """The datasets a sidecar's ArcRho cell links read, once each, in link order.

    These names are the instance-level dependency edges the persisted graph
    carries alongside the dataset-type formula graph: the linked dataset lists
    them as precedents, each named dataset lists it back as a dependent, and
    the dependent-propagation walk re-evaluates the links when one of them is
    refreshed. Excel references contribute no edge — a workbook is not a
    dataset — and unparseable stored text contributes no edge rather than
    failing a read.
    """

    names: List[str] = []
    seen: set[str] = set()

    def add(name: Any) -> None:
        text = _clean_text(name)
        key = re.sub(r"\s+", " ", text).casefold()
        if text and key not in seen:
            seen.add(key)
            names.append(text)

    for link in internal_links or []:
        if not isinstance(link, Mapping):
            continue
        try:
            add(parse_internal_reference(link.get("reference"))["dataset_name"])
        except DatasetLinkError:
            continue
    for link in formula_links or []:
        if not isinstance(link, Mapping):
            continue
        try:
            tokens = formula_reference_tokens(link.get("formula"))
        except DatasetLinkError:
            continue
        for token in tokens:
            if token["kind"] != "internal":
                continue
            try:
                add(parse_internal_reference(token["text"])["dataset_name"])
            except DatasetLinkError:
                continue
    return names


def formula_has_excel_reference(raw_text: Any) -> bool:
    """Whether a formula reads any Excel operand (its soft-failure operands)."""

    try:
        return any(token["kind"] == "excel" for token in formula_reference_tokens(raw_text))
    except DatasetLinkError:
        return False


# ── Evaluation (Excel array rules) ─────────────────────────────────────────────

Matrix = Dict[str, Any]


def scalar_matrix(value: float) -> Matrix:
    return {"rows": 1, "cols": 1, "values": [[value]]}


def _cell_at(matrix: Matrix, row: int, col: int) -> float:
    value = matrix["values"][0 if matrix["rows"] == 1 else row][0 if matrix["cols"] == 1 else col]
    # Excel arithmetic reads a blank cell as zero.
    if value is None or value == "":
        return 0.0
    return float(value)


def _combine(left: Matrix, right: Matrix, operator: str) -> Matrix:
    rows = left["rows"] if left["rows"] == right["rows"] else (
        right["rows"] if left["rows"] == 1 else (left["rows"] if right["rows"] == 1 else -1)
    )
    cols = left["cols"] if left["cols"] == right["cols"] else (
        right["cols"] if left["cols"] == 1 else (left["cols"] if right["cols"] == 1 else -1)
    )
    if rows < 0 or cols < 0:
        raise DatasetLinkError(
            f"Array sizes do not match ({left['rows']}x{left['cols']} and {right['rows']}x{right['cols']}).",
        )
    values: List[List[float]] = []
    for row in range(rows):
        line: List[float] = []
        for col in range(cols):
            a = _cell_at(left, row, col)
            b = _cell_at(right, row, col)
            if operator == "+":
                result = a + b
            elif operator == "-":
                result = a - b
            elif operator == "*":
                result = a * b
            elif operator == "/":
                if b == 0:
                    raise DatasetLinkError("The formula divides by zero.")
                result = a / b
            else:
                result = a ** b
            if not math.isfinite(result):
                raise DatasetLinkError("The formula produced a value that is not a finite number.")
            line.append(result)
        values.append(line)
    return {"rows": rows, "cols": cols, "values": values}


def evaluate_dataset_formula(
    tree: Mapping[str, Any],
    lookup: Callable[[Mapping[str, Any]], Matrix | None],
) -> Matrix:
    """Evaluate a parsed tree. ``lookup(token)`` returns a reference's matrix."""

    def visit(node: Mapping[str, Any]) -> Matrix:
        if node["kind"] == "number":
            return scalar_matrix(node["value"])
        if node["kind"] == "reference":
            matrix = lookup(node["token"])
            if not matrix or not matrix.get("rows", 0) > 0 or not matrix.get("cols", 0) > 0:
                raise DatasetLinkError(f"{node['token']['text']} has no values to calculate with.")
            return matrix
        if node["kind"] == "unary":
            return _combine(scalar_matrix(0.0), visit(node["operand"]), "-")
        return _combine(visit(node["left"]), visit(node["right"]), node["operator"])

    return visit(tree)


# ── Target-cell blocks (the on-disk shape of a link's cells) ───────────────────
#
# A link names every grid cell it fills. In memory each of those cells is its
# own record -- ``{"row", "column", "source_cell"}`` for an Excel link, and the
# numeric source or result pair for the other two kinds -- because every reader
# works cell by cell. Written out one record per cell, a 120x116 Excel range
# costs 6,786 records and roughly 650 KB, which is a file no reviewer can read
# and a read every dataset open pays for over the network.
#
# On disk a link's cells are therefore blocks. A block is one rectangle of the
# grid whose source moves with it cell for cell, written as the flat list::
#
#     [row, column, rows, columns, *source]
#
# ``row``/``column`` are the block's top-left target cell, ``rows``/``columns``
# its extent, and the source tail is the top-left source: one Excel address for
# an external link, and a row/column pair for the other two. The cell ``i``
# rows and ``j`` columns into the block reads the source ``i`` rows and ``j``
# columns into the source. The range above becomes 116 blocks on 116 lines and
# under 9 KB, because :func:`format_json_for_save` renders a list of lists one
# row per line.
#
# ``sidecar_core_contract``'s write funnel applies :func:`compact_sidecar_links`
# and every sidecar reader that consumes link cells applies
# :func:`expand_sidecar_links`, so the block form exists only in the file: the
# HTTP payloads, the browser modules, and every service still work cell by cell.

#: Sidecar field -> the source keys each of its target cells carries, in the
#: order the block's source tail writes them.
LINK_TARGET_FIELDS: Dict[str, tuple] = {
    "external_links": ("source_cell",),
    "internal_links": ("source_row", "source_column"),
    "formula_links": ("result_row", "result_column"),
}

_EXCEL_ADDRESS_RE = re.compile(r"([A-Z]+)([1-9][0-9]*)")


def excel_column_index(letters: Any) -> int:
    """Zero-based column index of an Excel column name (``A`` -> 0)."""

    index = 0
    for character in str(letters or "").upper():
        index = index * 26 + (ord(character) - 64)
    return index - 1


def excel_column_name(index: int) -> str:
    """Excel column name of a zero-based column index (0 -> ``A``)."""

    name = ""
    value = int(index) + 1
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _block_source(target: Mapping[str, Any], field: str) -> tuple:
    """The ``(row, column)`` a target cell reads, whatever the link kind."""

    if field == "external_links":
        match = _EXCEL_ADDRESS_RE.fullmatch(_clean_text(target.get("source_cell")).replace("$", "").upper())
        if not match:
            raise DatasetLinkError(
                "An external link target cell needs a valid Excel source_cell before it can be stored.",
            )
        return int(match.group(2)) - 1, excel_column_index(match.group(1))
    keys = LINK_TARGET_FIELDS[field]
    return int(target[keys[0]]), int(target[keys[1]])


def _source_tail(row: int, column: int, field: str) -> list:
    if field == "external_links":
        return [f"{excel_column_name(column)}{row + 1}"]
    return [row, column]


def compact_target_cells(targets: Iterable[Any], field: str) -> List[Any]:
    """Group one link's target cells into the blocks that go on disk.

    Cells already written as blocks pass through, so the projection is
    idempotent and a payload read from disk can be written back unchanged.
    """

    cells: Dict[tuple, tuple] = {}
    blocks: List[Any] = []
    for target in targets or []:
        if isinstance(target, list):
            blocks.append(list(target))
            continue
        if hasattr(target, "model_dump"):
            target = target.model_dump()
        cells[(int(target["row"]), int(target["column"]))] = _block_source(target, field)
    if blocks and cells:
        raise DatasetLinkError("A link's target cells must be all blocks or all single cells.")
    if blocks:
        return blocks

    remaining = dict(cells)
    for key in sorted(cells):
        if key not in remaining:
            continue
        row, column = key
        source_row, source_column = remaining.pop(key)
        columns = 1
        while remaining.get((row, column + columns)) == (source_row, source_column + columns):
            remaining.pop((row, column + columns))
            columns += 1
        rows = 1
        while all(
            remaining.get((row + rows, column + offset)) == (source_row + rows, source_column + offset)
            for offset in range(columns)
        ):
            for offset in range(columns):
                remaining.pop((row + rows, column + offset))
            rows += 1
        blocks.append([row, column, rows, columns] + _source_tail(source_row, source_column, field))
    return blocks


def expand_target_cells(blocks: Iterable[Any], field: str) -> List[Dict[str, Any]]:
    """Return one link's stored blocks as the cell records every reader uses."""

    keys = LINK_TARGET_FIELDS[field]
    targets: List[Dict[str, Any]] = []
    for block in blocks or []:
        if not isinstance(block, list):
            raise DatasetLinkError(
                "This dataset's cell links are stored one entry per cell, which this build no "
                "longer reads. Convert the workspace with tools/migrate_dataset_link_blocks.py.",
            )
        row, column, rows, columns = (int(value) for value in block[:4])
        source_row, source_column = (
            _block_source({keys[0]: block[4]}, field)
            if field == "external_links"
            else (int(block[4]), int(block[5]))
        )
        for offset_row in range(rows):
            for offset_column in range(columns):
                target = {"row": row + offset_row, "column": column + offset_column}
                target.update(
                    dict(zip(keys, _source_tail(source_row + offset_row, source_column + offset_column, field)))
                )
                targets.append(target)
    return targets


def _map_sidecar_links(payload: Any, convert: Callable[[Any, str], List[Any]]) -> Any:
    """Rewrite every link field of *payload* in place through *convert*."""

    if not isinstance(payload, dict):
        return payload
    for field in LINK_TARGET_FIELDS:
        links = payload.get(field)
        if not isinstance(links, list) or not links:
            continue
        payload[field] = [
            {**link, "target_cells": convert(link.get("target_cells"), field)}
            if isinstance(link, Mapping) and link.get("target_cells") is not None
            else link
            for link in links
        ]
    return payload


def compact_sidecar_links(payload: Any) -> Any:
    """Put a sidecar payload's link cells into their on-disk block form."""

    return _map_sidecar_links(payload, compact_target_cells)


def expand_sidecar_links(payload: Any) -> Any:
    """Put a sidecar payload's stored link blocks back into cell records."""

    return _map_sidecar_links(payload, expand_target_cells)

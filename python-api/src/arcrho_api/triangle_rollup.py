"""Roll a triangle up from a finer origin/development period to a coarser one,
and scatter values entered at the coarser one back into the finer store.

This module is the single owner of that arithmetic. The bundled app server
derives a coarser cached view of a stored triangle with it, and the Engine
bundle carries the same module for the same reason.

Geometry
--------
Every ArcRho triangle is anchored on the project's Origin Start Date and
valued on its Development End Date: row ``i`` covers the months starting
``origin_start_month + i * origin_length``, whatever the period length, and
the cell mask is built from the same two dates
(``dataset_service._empty_dataset_geometry_from_general_settings``). A coarse
origin period therefore always begins where one of the finer origin periods
begins, and the two grids line up without an offset.

Development periods are counted back from the valuation date, not forward
from the origin. ``valuation_months`` is the number of months from the anchor
through the Development End Date; the newest cell of every row is valued
there, and each earlier column is one development period before it. With 116
such months a yearly triangle is valued at 8, 20, 32, ... 116 months of age,
the way ResQ labels it, and the first development period is the short one. A
monthly triangle shown yearly therefore reads the 8th, 20th, 32nd, ... stored
column of each row, and a coarse cell is blank exactly where a triangle
created at the coarse shape would have no cell.

A development-aligned triangle (the ``dev`` cache variant) holds, in row ``i``
and column ``j``, the figure for origin ``i`` at the ``j``-th valuation date
of that row. Two cells in the same column carry two different valuation
dates, one origin period apart. A coarse cell has a single valuation date, so
its parts are read along the calendar diagonal of the finer triangle: every
finer row of the block contributes the cell it holds at that date.

Writing at a coarser development view
-------------------------------------
``scatter_triangle`` is the inverse of that read, and follows ResQ: a value
entered in a coarse cell is the row's cumulative figure at that cell's
valuation date, so it lands in the one stored cell valued at the same date
and every other stored cell of the triangle becomes cumulative 0. The origin
axis is never relaxed -- a coarse origin row has no single valuation date to
write to -- so the two grids share their rows.

A calendar-aligned triangle (the ``cal`` variant) has already been reshaped so
that a column is a calendar period counted forward from the anchor. Its
columns share one valuation date down the whole triangle, so rolling it up is
a plain block aggregation with no diagonal shift, and there it is the last
period that may be short.

Blank cells
-----------
For a development-aligned triangle every finer cell a coarse cell needs shares
one valuation date, so a well-formed triangle either holds all of them or none
of them. A coarse cell whose parts are not all present is therefore blank
rather than a partial sum. A calendar-aligned triangle is different: a finer
origin period that starts later than a calendar column has no cell there and
contributed nothing, so blanks are read as zero and the coarse cell is blank
only when the whole block is.
"""
from __future__ import annotations

from typing import Any, List, Sequence

Triangle = Sequence[Sequence[Any]]

__all__ = [
    "rollup_reason",
    "rollup_factors",
    "rollup_triangle",
    "scatter_reason",
    "scatter_triangle",
]


def _positive_int(value: Any) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return 0
    return number if number > 0 else 0


def rollup_reason(
    source_origin_length: Any,
    source_development_length: Any,
    target_origin_length: Any,
    target_development_length: Any,
    *,
    calendar: bool = False,
) -> str:
    """Return an empty string when the roll-up is possible, else why it is not."""
    source_origin = _positive_int(source_origin_length)
    source_development = _positive_int(source_development_length)
    target_origin = _positive_int(target_origin_length)
    target_development = _positive_int(target_development_length)
    if not (source_origin and source_development and target_origin and target_development):
        return "invalid period length"
    if target_origin < source_origin or target_development < source_development:
        return "local caches can only derive from finer to coarser periods"
    if target_origin % source_origin or target_development % source_development:
        return "requested periods are not whole multiples of the cached periods"
    if not calendar and target_origin // source_origin > 1 and source_origin % source_development:
        return (
            f"origin periods of {source_origin} months are not a whole number of "
            f"{source_development}-month development periods, so the rows of a block "
            "share no valuation date"
        )
    return ""


def rollup_factors(
    source_origin_length: Any,
    source_development_length: Any,
    target_origin_length: Any,
    target_development_length: Any,
    *,
    calendar: bool = False,
) -> tuple[int, int]:
    """How many finer origin rows and development columns make one coarse cell."""
    reason = rollup_reason(
        source_origin_length,
        source_development_length,
        target_origin_length,
        target_development_length,
        calendar=calendar,
    )
    if reason:
        raise ValueError(reason)
    return (
        _positive_int(target_origin_length) // _positive_int(source_origin_length),
        _positive_int(target_development_length) // _positive_int(source_development_length),
    )


def _cell(rows: List[List[Any]], row_index: int, column_index: int) -> float | None:
    if row_index >= len(rows):
        return None
    row = rows[row_index]
    if column_index >= len(row):
        return None
    value = row[column_index]
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if number != number:
        return None
    return number


def _valued_at(
    row: int,
    column: int,
    origin_length: int,
    development_length: int,
    last_month: int,
) -> int:
    """The month, counted from the anchor, a development-aligned cell is valued at.

    ``last_month`` is the Development End Date's offset from the anchor. The
    row's newest cell is valued there, so its columns are phased back from it.
    """
    start = row * origin_length
    return start + column * development_length + (last_month - start) % development_length


def _rollup_development(
    rows: List[List[Any]],
    target_rows: int,
    target_columns: int,
    source_origin: int,
    source_development: int,
    target_origin: int,
    target_development: int,
    last_month: int,
    cumulative: bool,
) -> List[List[float | None]]:
    origin_factor = target_origin // source_origin
    development_factor = target_development // source_development
    values: List[List[float | None]] = []
    for block in range(target_rows):
        row_values: List[float | None] = []
        for column in range(target_columns):
            valued_at = _valued_at(block, column, target_origin, target_development, last_month)
            if valued_at > last_month:
                row_values.append(None)
                continue
            total = 0.0
            complete = True
            for offset in range(origin_factor):
                row_index = block * origin_factor + offset
                if row_index >= len(rows):
                    # Origins after the last stored row are not part of the project.
                    break
                first_valued_at = _valued_at(row_index, 0, source_origin, source_development, last_month)
                last = (valued_at - first_valued_at) // source_development
                if last < 0:
                    # This finer origin period had not started at the coarse
                    # valuation date, so it contributed nothing.
                    continue
                first = last if cumulative else max(last - development_factor + 1, 0)
                for column_index in range(first, last + 1):
                    cell = _cell(rows, row_index, column_index)
                    if cell is None:
                        complete = False
                        break
                    total += cell
                if not complete:
                    break
            row_values.append(total if complete else None)
        values.append(row_values)
    return values


def _rollup_calendar(
    rows: List[List[Any]],
    target_rows: int,
    target_columns: int,
    origin_factor: int,
    development_factor: int,
    last_source_column: int,
    cumulative: bool,
) -> List[List[float | None]]:
    values: List[List[float | None]] = []
    for block in range(target_rows):
        row_values: List[float | None] = []
        for column in range(target_columns):
            last = min((column + 1) * development_factor - 1, last_source_column)
            if cumulative:
                columns = [last]
            else:
                columns = list(range(column * development_factor, last + 1))
            total = 0.0
            seen = False
            for offset in range(origin_factor):
                for column_index in columns:
                    cell = _cell(rows, block * origin_factor + offset, column_index)
                    if cell is not None:
                        total += cell
                        seen = True
            row_values.append(total if seen else None)
        values.append(row_values)
    return values


def rollup_triangle(
    values: Triangle,
    *,
    source_origin_length: Any,
    source_development_length: Any,
    target_origin_length: Any,
    target_development_length: Any,
    valuation_months: Any,
    cumulative: bool = True,
    calendar: bool = False,
) -> List[List[float | None]]:
    """Aggregate ``values`` to the coarser target periods.

    ``values`` is a development-aligned triangle unless ``calendar`` is set.
    ``valuation_months`` counts the months from the project's Origin Start
    Date through its Development End Date. The result has the rows and
    columns a triangle created at the target shape would have: an origin
    block that the stored rows only partly fill is still a row, and the
    development period the valuation date falls in is still a column.
    """
    origin_factor, development_factor = rollup_factors(
        source_origin_length,
        source_development_length,
        target_origin_length,
        target_development_length,
        calendar=calendar,
    )
    months = _positive_int(valuation_months)
    if not months:
        raise ValueError("invalid valuation date")
    last_month = months - 1
    rows = [list(row) for row in (values or [])]
    if not rows:
        raise ValueError("cached triangle holds no rows")
    target_rows = -(-len(rows) // origin_factor)
    if calendar:
        # Calendar columns run forward from the anchor, so the stored width
        # already says how many coarse periods there are; a vector is one.
        source_columns = max((len(row) for row in rows), default=0)
        return _rollup_calendar(
            rows,
            target_rows,
            -(-source_columns // development_factor),
            origin_factor,
            development_factor,
            source_columns - 1,
            cumulative,
        )
    return _rollup_development(
        rows,
        target_rows,
        last_month // _positive_int(target_development_length) + 1,
        _positive_int(source_origin_length),
        _positive_int(source_development_length),
        _positive_int(target_origin_length),
        _positive_int(target_development_length),
        last_month,
        cumulative,
    )


def scatter_reason(
    source_origin_length: Any,
    source_development_length: Any,
    target_origin_length: Any,
    target_development_length: Any,
) -> str:
    """Return an empty string when values at the target shape can be scattered."""
    if _positive_int(target_origin_length) != _positive_int(source_origin_length):
        return "values can be entered only at the stored origin period"
    return rollup_reason(
        source_origin_length,
        source_development_length,
        target_origin_length,
        target_development_length,
    )


def scatter_triangle(
    values: Triangle,
    *,
    source_origin_length: Any,
    source_development_length: Any,
    target_origin_length: Any,
    target_development_length: Any,
    valuation_months: Any,
    cumulative: bool = True,
) -> List[List[float | None]]:
    """Write ``values``, entered at the target shape, into the source shape.

    The inverse of :func:`rollup_triangle` on the development axis: each
    entered cell is the row's cumulative figure at its valuation date, so it
    lands in the stored cell valued at the same date and every other stored
    cell becomes cumulative 0. An incremental grid is accumulated along the
    row first, and the result is returned in the same convention it arrived
    in, which puts each figure at its age and its negative at the next stored
    age. The rows are the rows of ``values``; the columns are the ones a
    triangle created at the source shape would have.
    """
    reason = scatter_reason(
        source_origin_length,
        source_development_length,
        target_origin_length,
        target_development_length,
    )
    if reason:
        raise ValueError(reason)
    months = _positive_int(valuation_months)
    if not months:
        raise ValueError("invalid valuation date")
    last_month = months - 1
    origin_length = _positive_int(source_origin_length)
    source_development = _positive_int(source_development_length)
    target_development = _positive_int(target_development_length)
    rows = [list(row) for row in (values or [])]
    if not rows:
        raise ValueError("entered triangle holds no rows")
    stored_columns = last_month // source_development + 1
    scattered: List[List[float | None]] = []
    for row_index, row in enumerate(rows):
        first_valued_at = _valued_at(row_index, 0, origin_length, source_development, last_month)
        entered_at: dict[int, float] = {}
        running = 0.0
        for column in range(len(row)):
            valued_at = _valued_at(row_index, column, origin_length, target_development, last_month)
            if valued_at > last_month:
                break
            entered = _cell(rows, row_index, column)
            if entered is None:
                continue
            running = entered if cumulative else running + entered
            entered_at[(valued_at - first_valued_at) // source_development] = running
        scattered_row: List[float | None] = []
        previous = 0.0
        for column in range(stored_columns):
            if _valued_at(row_index, column, origin_length, source_development, last_month) > last_month:
                scattered_row.append(None)
                continue
            value = entered_at.get(column, 0.0)
            scattered_row.append(value if cumulative else value - previous)
            previous = value
        scattered.append(scattered_row)
    return scattered

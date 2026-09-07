"""Probe ResQ's stored-length, editing and presentation rules against a live project.

Creates throwaway triangles and vectors named ``ArcRho probe ...`` of a
non-unique, non-calculated dataset type in one reserving class, exercises every
interaction case recorded in
docs/reference/resq_stored_and_display_lengths.md, and deletes everything it
created. Nothing that already exists in the class is touched. Server PC only
(ResQ COM), with a Python that has pywin32::

    py -3.10 tools/resq_stored_length_probe.py
    py -3.10 tools/resq_stored_length_probe.py --only A B      # one or more case groups
    py -3.10 tools/resq_stored_length_probe.py --keep          # leave the probe objects for a GUI look
    py -3.10 tools/resq_stored_length_probe.py --json temp/probe.json

The case identifiers printed here (A1, B2, ...) are the ones the reference
document cites, so a rule can always be traced back to the run that established
it.

Early binding throughout, and every property put goes through IDispatch so a
refused set raises ResQ's own error text instead of pywin32 quietly creating a
Python attribute on the wrapper. The probe writes and saves inside the fake
project only; ``PROJECT`` must never name a production project.
"""
from __future__ import annotations

import argparse
import datetime
import json
import sys
import traceback

import pythoncom
import pywintypes
from win32com.client import gencache

PROJECT = "NJ_Annual_Prod_202605_Fake"
RC_PATH = r"HPPREF\HO+DF\NJ\Legacy\HOL"
TRIANGLE_TYPE = "Net Loss - ad hoc"          # non-unique, not calculated, triangle-only
VECTOR_TYPE = "F 00 - Ultimate Net Loss "    # non-unique, not calculated, origin vector
PREFIX = "ArcRho probe "
CONFIG = r"E:/ArcRho Server/config/config.json"

# The 10x10 annual cumulative triangle of the GUI walk-through, rounded to whole
# numbers. Rows are origin years 2017..2026 of the fake project.
ANNUAL_10x10 = [
    [1357, 1385, 1412, 1440, 1469, 1499, 1529, 1559, 1590, 1622],
    [1493, 1523, 1553, 1585, 1616, 1649, 1682, 1715, 1749],
    [1642, 1675, 1709, 1743, 1778, 1813, 1850, 1887],
    [1807, 1843, 1880, 1917, 1956, 1995, 2035],
    [1987, 2027, 2068, 2109, 2151, 2194],
    [2186, 2230, 2274, 2320, 2366],
    [2405, 2453, 2502, 2552],
    [2645, 2698, 2752],
    [2910, 2968],
    [3201],
]

OBSERVED: dict[str, object] = {}


# ----- COM helpers ----------------------------------------------------------------

def connect():
    cfg = json.load(open(CONFIG))["resq"]
    app = gencache.EnsureDispatch("ResQ3Automation.ResQApplication")
    app.ConnectByName(cfg["connection_name"], cfg["user_name"], cfg["password"])
    return app


def err_text(exc):
    if isinstance(exc, pywintypes.com_error):
        args = exc.args
        return str(args[2][2]) if len(args) > 2 and args[2] and len(args[2]) > 2 else str(args)
    return f"{type(exc).__name__}: {exc}"


def put(obj, name, value, record=None):
    """Property put through Invoke; prints and returns whether ResQ accepted it."""
    try:
        dispid = obj._oleobj_.GetIDsOfNames(name)
        obj._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, value)
        print(f"  put {name}={value!r}: ok")
        if record is not None:
            record.append({"put": name, "value": value, "accepted": True})
        return True
    except Exception as exc:  # noqa: BLE001
        text = err_text(exc)
        print(f"  put {name}={value!r}: REFUSED {text}")
        if record is not None:
            record.append({"put": name, "value": value, "accepted": False, "error": text})
        return False


def call(label, fn, record=None):
    """Invoke an action method and report whether ResQ accepted it."""
    try:
        fn()
        print(f"  {label}: ok")
        if record is not None:
            record.append({"call": label, "accepted": True})
        return True
    except Exception as exc:  # noqa: BLE001
        text = err_text(exc)
        print(f"  {label}: REFUSED {text}")
        if record is not None:
            record.append({"call": label, "accepted": False, "error": text})
        return False


def section(title):
    print("\n" + "=" * 8, title)


# ----- triangle helpers ------------------------------------------------------------

def dev_labels(t):
    return [str(t.DevelopmentLabel(j)) for j in range(1, int(t.DevelopmentCountByIndex(1)) + 1)]


def origin_labels(t):
    return [str(t.OriginLabel(i)) for i in range(1, int(t.OriginCount) + 1)]


def row_widths(t):
    return [int(t.DevelopmentCountByIndex(i)) for i in range(1, int(t.OriginCount) + 1)]


def grid(t):
    return [[float(t.ValuesByIndex(i, j)) for j in range(1, int(t.DevelopmentCountByIndex(i)) + 1)]
            for i in range(1, int(t.OriginCount) + 1)]


def shape(t):
    return {
        "display": [int(t.OriginLength), int(t.DevelopmentLength)],
        "stored": [int(t.StoredOriginLength), int(t.StoredDevelopmentLength)],
        "cumulative": bool(t.Cumulative),
        "calendarised": bool(t.Calendarised),
        "transposed": bool(t.Transposed),
        "origin_count": int(t.OriginCount),
    }


def snapshot(t):
    s = shape(t)
    s["development_labels"] = dev_labels(t)
    s["origin_labels"] = origin_labels(t)
    s["row_widths"] = row_widths(t)
    s["values"] = grid(t)
    return s


def nonzero(t):
    labels = dev_labels(t)
    return [(i + 1, j + 1, labels[j], v) for i, row in enumerate(grid(t)) for j, v in enumerate(row) if v]


def show(t, label=""):
    print(f"  {label}O{t.OriginLength}/D{t.DevelopmentLength} stored O{t.StoredOriginLength}/D{t.StoredDevelopmentLength} "
          f"cum={t.Cumulative} rows={t.OriginCount} widths={row_widths(t)[:4]}..{row_widths(t)[-2:]} "
          f"dev={dev_labels(t)[:4]}..{dev_labels(t)[-2:]}")


def print_grid(t, title, max_cols=12):
    print(f"  -- {title}: O{t.OriginLength}/D{t.DevelopmentLength} stored O{t.StoredOriginLength}/D{t.StoredDevelopmentLength} cumulative={t.Cumulative}")
    labels = dev_labels(t)
    print("          " + " ".join(f"{l:>9}" for l in labels[:max_cols]))
    for i, row in enumerate(grid(t), start=1):
        print(f"  {str(t.OriginLabel(i)):>8} " + " ".join(f"{v:9.0f}" for v in row[:max_cols]))


def fill(t, fn):
    n = 0
    for i in range(1, int(t.OriginCount) + 1):
        for j in range(1, int(t.DevelopmentCountByIndex(i)) + 1):
            t.SetValuesByIndex(i, j, float(fn(i, j)))
            n += 1
    return n


def write_rows(t, rows):
    """Write a ragged list of rows into the current display, row 1 first."""
    n = 0
    for i, row in enumerate(rows, start=1):
        for j, value in enumerate(row, start=1):
            t.SetValuesByIndex(i, j, float(value))
            n += 1
    return n


def discard(item):
    """Drop a probe object. A triangle that was never saved does not exist in the
    database, so ResQ answers `Delete` on it with an object-version clash; that is
    nothing to report, and the cleanup pass sweeps up anything that was saved."""
    try:
        item.Delete()
    except Exception:  # noqa: BLE001
        pass


def changed_cells(t, before, after):
    labels = dev_labels(t)
    return [(i + 1, j + 1, labels[j], before[i][j], after[i][j])
            for i in range(len(before)) for j in range(len(before[i])) if before[i][j] != after[i][j]]


class Probe:
    def __init__(self, app):
        self.app = app
        self.project = app.Projects().Item(PROJECT)
        self.rc = self.project.ReservingClasses().Item(RC_PATH)

    # -- fixtures ---------------------------------------------------------------

    def new_triangle(self, name):
        t = self.rc.Triangles().Add()
        t.DatasetType = self.project.DatasetTypes().Item(TRIANGLE_TYPE)
        t.Name = t.UniqueName(PREFIX + name)
        return t

    def find_triangle(self, name):
        tris = self.rc.Triangles()
        for i in range(1, tris.Count + 1):
            if str(tris.Item(i).Name) == name:
                return tris.Item(i)
        return None

    def empty_store(self, name, origin=12, display_dev=12, stored_dev=1, cumulative=True):
        """A saved but empty triangle at the requested display and stored shape."""
        t = self.new_triangle(name)
        put(t, "OriginLength", origin)
        put(t, "DevelopmentLength", display_dev)
        if stored_dev != display_dev:
            put(t, "StoredDevelopmentLength", stored_dev)
        put(t, "Cumulative", cumulative)
        t.Save()
        return t

    def monthly_store(self, name):
        """Saved triangle stored O12/D1, shown O12/D12, cumulative, filled with 100000*row + age."""
        t = self.new_triangle(name)
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 1); put(t, "Cumulative", True); t.Save()
        n = fill(t, lambda i, m: 100000 * i + m); t.Save()
        put(t, "DevelopmentLength", 12); show(t, f"filled {n} monthly cells, now ")
        return t

    # ===== group A: creating and shaping an empty triangle =========================

    def survey(self):
        p = self.project
        print(f"  project {p.Name}: origins {p.OriginStartDate:%Y-%m-%d}..{p.OriginEndDate:%Y-%m-%d}, "
              f"Development End Date {p.DevelopmentEndDate:%Y-%m-%d}, O{p.OriginLength}/D{p.DevelopmentLength}, {p.OriginCount} origins")
        dt = p.DatasetTypes().Item(TRIANGLE_TYPE)
        print(f"  type {dt.Name!r}: Unique={dt.Unique} Calculated={dt.Calculated}")
        OBSERVED["project"] = {
            "name": str(p.Name),
            "origin_start": f"{p.OriginStartDate:%Y-%m-%d}",
            "development_end": f"{p.DevelopmentEndDate:%Y-%m-%d}",
            "origin_length": int(p.OriginLength),
            "development_length": int(p.DevelopmentLength),
            "origin_count": int(p.OriginCount),
        }

    def a1_defaults(self):
        """A1 — what a freshly added triangle reports before anything is set."""
        t = self.new_triangle("A1 defaults")
        show(t, "after Add: ")
        OBSERVED["A1"] = shape(t)
        discard(t)

    def a2_put_order(self):
        """A2 — a same-value display put does nothing; a changing origin put resets development."""
        rec = []
        t = self.new_triangle("A2 put order")
        print("  -- the same value the triangle already carries")
        put(t, "DevelopmentLength", 12, rec); show(t, "after D12: ")
        put(t, "OriginLength", 12, rec); show(t, "after O12: ")
        same_value = shape(t)
        print("  -- a real origin change on a never-saved triangle")
        put(t, "DevelopmentLength", 6, rec); show(t, "after D6: ")
        before_origin_change = shape(t)
        put(t, "OriginLength", 24, rec); show(t, "after O24: ")
        after_origin_change = shape(t)
        OBSERVED["A2"] = {"same_value": same_value, "before_origin_change": before_origin_change,
                          "after_origin_change": after_origin_change, "calls": rec}
        discard(t)

    def a3_display_put_moves_the_store(self):
        """A3 — a display put moves the store while empty, including a put of the same value."""
        rec = []
        t = self.new_triangle("A3 display put")
        put(t, "OriginLength", 12, rec); put(t, "DevelopmentLength", 12, rec)
        put(t, "StoredDevelopmentLength", 1, rec); show(t, "stored lowered to 1: ")
        stored_after_lower = shape(t)
        print("  -- putting the SAME display development length back")
        put(t, "DevelopmentLength", 12, rec); show(t, "after D12 again: ")
        same_value_put = shape(t)
        print("  -- putting a DIFFERENT display development length")
        put(t, "StoredDevelopmentLength", 1, rec)
        put(t, "DevelopmentLength", 6, rec); show(t, "after D6: ")
        different_value_put = shape(t)
        print("  -- a put with no multiple check: display 4 over a stored 3")
        put(t, "DevelopmentLength", 12, rec); put(t, "StoredDevelopmentLength", 3, rec)
        put(t, "DevelopmentLength", 4, rec); show(t, "after D4 over stored 3: ")
        no_multiple_check = shape(t)
        print("  -- the same on the origin axis")
        put(t, "OriginLength", 6, rec); show(t, "after O6: ")
        origin_put = shape(t)
        OBSERVED["A3"] = {
            "stored_after_lower": stored_after_lower,
            "same_value_put": same_value_put,
            "different_value_put": different_value_put,
            "no_multiple_check": no_multiple_check,
            "origin_put": origin_put,
            "calls": rec,
        }
        discard(t)

    def a4_stored_development_factors(self):
        """A4 — StoredDevelopmentLength accepts factors of the display and refuses the rest."""
        rec = []
        t = self.new_triangle("A4 stored factors")
        put(t, "OriginLength", 12, rec); put(t, "DevelopmentLength", 12, rec)
        for value in (1, 2, 3, 4, 6, 12, 5, 7, 8, 24, 0):
            put(t, "StoredDevelopmentLength", value, rec)
            put(t, "DevelopmentLength", 12)
        OBSERVED["A4"] = {"calls": rec}
        discard(t)

    def a5_stored_origin_has_no_setter(self):
        """A5 — StoredOriginLength cannot be set at all."""
        rec = []
        t = self.new_triangle("A5 stored origin")
        put(t, "OriginLength", 12, rec)
        put(t, "StoredOriginLength", 1, rec)
        put(t, "StoredOriginLength", 12, rec)
        OBSERVED["A5"] = {"calls": rec}
        discard(t)

    def a6_display_development_divides_origin(self):
        """A6 — the display development length must be a factor of the display origin length."""
        rec = []
        t = self.new_triangle("A6 display factors")
        put(t, "OriginLength", 12, rec)
        for value in (1, 2, 3, 4, 6, 12, 5, 7, 24):
            put(t, "DevelopmentLength", value, rec)
        put(t, "OriginLength", 6, rec)
        for value in (12, 6, 3):
            put(t, "DevelopmentLength", value, rec)
        OBSERVED["A6"] = {"calls": rec}
        discard(t)

    def a7_empty_save_persists(self):
        """A7 — a Save of an empty triangle records the stored pair; reread it in this session."""
        t = self.empty_store("A7 empty save")
        show(t, "saved empty: ")
        name = str(t.Name)
        # UnloadChildren frees the objects behind the wrappers already handed
        # out, so every reading of the old one has to be taken first: touching
        # it afterwards access-violates inside ResQ3Automation.dll.
        saved = shape(t)
        self.rc.UnloadChildren()
        again = self.find_triangle(name)
        show(again, "after UnloadChildren: ")
        OBSERVED["A7"] = {"saved": saved, "reloaded": shape(again), "name": name}
        discard(again)

    def a8_origin_change_on_a_saved_empty_triangle(self):
        """A8 — an origin change on a saved but empty triangle, and what it does to development."""
        rec = []
        t = self.empty_store("A8 origin change")
        show(t, "saved empty at 12/12 over a store of 1: ")
        put(t, "OriginLength", 6, rec); show(t, "after O6: ")
        after_origin = shape(t)
        put(t, "OriginLength", 12, rec); put(t, "DevelopmentLength", 12, rec)
        put(t, "StoredDevelopmentLength", 1, rec)
        print("  -- Cumulative is not part of the shape")
        put(t, "Cumulative", False, rec); show(t, "after Cumulative=False: ")
        after_cumulative = shape(t)
        put(t, "Cumulative", True, rec); t.Save()
        OBSERVED["A8"] = {"after_origin": after_origin, "after_cumulative": after_cumulative, "calls": rec}
        discard(t)

    def a9_stored_grid_labels(self):
        """A9 - the stored grid runs forward from the origin start, with a short last period."""
        table = {}
        for length in (1, 2, 3, 4, 6, 12):
            t = self.new_triangle(f"A9 stored {length}")
            put(t, "OriginLength", 12)
            put(t, "DevelopmentLength", length)
            labels = dev_labels(t)
            table[length] = {"columns": len(labels), "first": labels[:3], "last": labels[-2:],
                             "stored": [int(t.StoredOriginLength), int(t.StoredDevelopmentLength)]}
            print(f"  store {length}: {len(labels)} columns {labels[:3]} .. {labels[-2:]}")
            discard(t)
        OBSERVED["A9"] = table

    # ===== group B: entering data ==================================================

    def b1_write_at_the_stored_shape(self):
        """B1 — the baseline: display equals the store."""
        t = self.new_triangle("B1 write at store")
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 12); put(t, "Cumulative", True); t.Save()
        n = write_rows(t, ANNUAL_10x10); t.Save()
        print(f"  wrote {n} cells at O12/D12 stored O12/D12")
        print_grid(t, "read back", max_cols=10)
        OBSERVED["B1"] = snapshot(t)
        discard(t)

    def b2_annual_paste_into_a_monthly_store(self):
        """B2 — the production case: a 10x10 annual paste into a triangle stored at development 1."""
        t = self.empty_store("B2 annual paste")
        show(t, "empty, stored at 1: ")
        n = write_rows(t, ANNUAL_10x10)
        print(f"  wrote {n} cells at the D12 display")
        before_save = shape(t)
        t.Save()
        after_save = shape(t)
        print(f"  stored pair before Save {before_save['stored']}, after Save {after_save['stored']}")
        print_grid(t, "read back at D12", max_cols=10)
        views = {}
        for d in (1, 2, 3, 4, 6, 12):
            put(t, "DevelopmentLength", d)
            views[f"D{d}"] = snapshot(t)
            row1 = views[f"D{d}"]["values"][0]
            labels = views[f"D{d}"]["development_labels"]
            hits = [(labels[j], v) for j, v in enumerate(row1) if v]
            print(f"  D{d}: {len(labels)} columns, row 1 non-zero at {hits[:4]}{' ...' if len(hits) > 4 else ''}")
        put(t, "DevelopmentLength", 1)
        print(f"  stored cells holding a value: {len(nonzero(t))} of {sum(row_widths(t))}")
        put(t, "Cumulative", False)
        row1 = grid(t)[0]
        labels = dev_labels(t)
        print(f"  D1 incremental row 1 non-zero: {[(labels[j], v) for j, v in enumerate(row1) if v][:6]}")
        incremental_d1 = snapshot(t)
        put(t, "Cumulative", True); put(t, "DevelopmentLength", 12); t.Save()
        OBSERVED["B2"] = {
            "before_save": before_save,
            "after_save": after_save,
            "views": views,
            "incremental_d1": incremental_d1,
        }
        discard(t)

    def b3_coarse_write_over_a_filled_store(self):
        """B3 — one coarse cell written over a filled monthly store rebuilds the whole triangle."""
        t = self.monthly_store("B3 coarse over filled")
        put(t, "DevelopmentLength", 1); before = grid(t); put(t, "DevelopmentLength", 12)
        t.SetValuesByIndex(2, 2, 999999.0); print("  SetValuesByIndex(2,2)=999999 at D12, no Save yet")
        put(t, "DevelopmentLength", 1); after = grid(t)
        changed = changed_cells(t, before, after)
        kept = sum(1 for i, row in enumerate(before) for j, v in enumerate(row) if v and after[i][j] == v)
        zeroed = sum(1 for i, row in enumerate(before) for j, v in enumerate(row) if v and after[i][j] == 0)
        print(f"  stored cells changed before Save: {len(changed)} in rows {sorted({c[0] for c in changed})}; kept={kept} zeroed={zeroed}")
        print("  row 2:", [(j, l, b, a) for i, j, l, b, a in changed if i == 2][:18])
        OBSERVED["B3"] = {"changed": len(changed), "kept": kept, "zeroed": zeroed,
                          "rows_touched": sorted({c[0] for c in changed})}
        put(t, "DevelopmentLength", 12); t.Save()
        discard(t)

    def b4_partial_coarse_write(self):
        """B4 — writing only some display cells still clears the stored cells between them."""
        t = self.empty_store("B4 partial write")
        t.SetValuesByIndex(1, 1, 111.0)
        t.SetValuesByIndex(1, 3, 333.0)
        t.Save()
        put(t, "DevelopmentLength", 1)
        print("  stored non-zero after writing display columns 1 and 3:", nonzero(t)[:8])
        OBSERVED["B4"] = {"stored_nonzero": [(i, j, l, v) for i, j, l, v in nonzero(t)]}
        put(t, "DevelopmentLength", 12); t.Save()
        discard(t)

    def b5_incremental_coarse_write(self):
        """B5 — an incremental display stores the running sum at each display age."""
        t = self.empty_store("B5 incremental write", cumulative=False)
        write_rows(t, [[100, 10, 1], [200, 20]])
        t.Save()
        print_grid(t, "read back incremental", max_cols=6)
        put(t, "Cumulative", True); print_grid(t, "read back cumulative", max_cols=6)
        put(t, "DevelopmentLength", 1)
        print("  stored non-zero (cumulative):", [(i, l, v) for i, j, l, v in nonzero(t)][:8])
        OBSERVED["B5"] = {"stored_nonzero": [(i, j, l, v) for i, j, l, v in nonzero(t)]}
        put(t, "DevelopmentLength", 12); put(t, "Cumulative", False); t.Save()
        discard(t)

    def b6_coarse_origin_write_refused(self):
        """B6 — a write at an origin display coarser than the store is refused."""
        rec = []
        t = self.new_triangle("B6 coarse origin")
        put(t, "OriginLength", 1, rec); put(t, "DevelopmentLength", 1, rec); put(t, "Cumulative", True, rec)
        t.Save()
        fill(t, lambda i, j: 1000 * i + j); t.Save()
        put(t, "OriginLength", 12, rec); put(t, "DevelopmentLength", 12, rec)
        call("SetValuesByIndex(2,2) at O12 over an O1 store", lambda: t.SetValuesByIndex(2, 2, 7777.0), rec)
        OBSERVED["B6"] = {"calls": rec}
        put(t, "OriginLength", 1, rec); put(t, "DevelopmentLength", 1, rec); t.Save()
        discard(t)

    def b7_setvalues_by_age(self):
        """B7 — SetValues(originDate, ageMonths, value) writes the display column holding that age."""
        rec = []
        t = self.empty_store("B7 SetValues by age")
        d = datetime.datetime(2017, 1, 1)
        for m, v in ((17, 55555.0), (10, 44444.0), (5, 33333.0)):
            call(f"SetValues(2017-01-01, {m}, {v})", lambda m=m, v=v: t.SetValues(d, m, v), rec)
        t.Save(); print("  D12 row 1:", grid(t)[0][:3])
        reads = {m: float(t.Values(d, m)) for m in (1, 5, 6, 10, 17, 18, 113)}
        print("  Values(2017-01-01, m) reads:", reads)
        put(t, "DevelopmentLength", 1)
        stored = nonzero(t)
        print("  stored cells:", [(l, v) for i, j, l, v in stored if i == 1])
        OBSERVED["B7"] = {"calls": rec, "display_reads_by_age": reads,
                          "stored_nonzero": [(i, j, l, v) for i, j, l, v in stored]}
        put(t, "DevelopmentLength", 12); t.Save()
        discard(t)

    def b8_zeros_are_empty(self):
        """B8 — an all-zero save leaves the triangle empty; a value locks it; ClearData unlocks it."""
        rec = []
        t = self.empty_store("B8 emptiness", stored_dev=12)
        print("  -- unsaved value"); t.SetValuesByIndex(1, 1, 5.0); put(t, "StoredDevelopmentLength", 3, rec)
        print("  -- saved value"); t.Save(); put(t, "StoredDevelopmentLength", 3, rec)
        print("  -- explicit zeros everywhere, saved"); fill(t, lambda i, j: 0.0); t.Save()
        put(t, "StoredDevelopmentLength", 3, rec)
        print("  -- value saved, then ClearData without Save"); put(t, "StoredDevelopmentLength", 12)
        t.SetValuesByIndex(1, 1, 7.0); t.Save()
        call("ClearData", t.ClearData, rec)
        put(t, "StoredDevelopmentLength", 3, rec)
        OBSERVED["B8"] = {"calls": rec}
        put(t, "StoredDevelopmentLength", 12); t.Save()
        discard(t)

    def b9_cleardata_frees_the_origin_store(self):
        """B9 — after ClearData an OriginLength put moves the stored origin length again."""
        rec = []
        t = self.new_triangle("B9 ClearData origin")
        put(t, "OriginLength", 1, rec); put(t, "DevelopmentLength", 1, rec); t.Save()
        fill(t, lambda i, j: 1.0); t.Save(); show(t, "filled at O1/D1: ")
        call("ClearData", t.ClearData, rec)
        after_clear = shape(t)
        print(f"  straight after ClearData: display {after_clear['display']} stored {after_clear['stored']}")
        put(t, "OriginLength", 12, rec); show(t, "after ClearData + a real O12 change: ")
        after_origin_change = shape(t)
        put(t, "StoredOriginLength", 12, rec)
        OBSERVED["B9"] = {"after_clear": after_clear, "after_origin_change": after_origin_change,
                          "after": shape(t), "calls": rec}
        t.Save()
        discard(t)

    # ===== group C: what data locks ================================================

    def c1_locked_after_data(self):
        """C1 — with saved values the store is fixed and the display must be a whole multiple."""
        rec = []
        t = self.empty_store("C1 locked")
        write_rows(t, ANNUAL_10x10); t.Save()
        put(t, "StoredDevelopmentLength", 12, rec)
        put(t, "StoredDevelopmentLength", 1, rec)
        for d in (1, 2, 3, 4, 6, 12, 5, 7, 24):
            put(t, "DevelopmentLength", d, rec)
        put(t, "DevelopmentLength", 12)
        for o in (1, 6, 12, 24, 36):
            put(t, "OriginLength", o, rec)
            if int(t.OriginLength) != 12:
                put(t, "OriginLength", 12)
                put(t, "DevelopmentLength", 12)
        OBSERVED["C1"] = {"calls": rec}
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 12); t.Save()
        discard(t)

    def c2_coarse_origin_display_reads(self):
        """C2 — a coarser origin display reads the calendar diagonal over the finer rows."""
        t = self.empty_store("C2 coarse origin read")
        write_rows(t, ANNUAL_10x10); t.Save()
        put(t, "OriginLength", 24)
        print_grid(t, "O24/D12", max_cols=10)
        o24 = snapshot(t)
        put(t, "DevelopmentLength", 24)
        print_grid(t, "O24/D24", max_cols=6)
        o24d24 = snapshot(t)
        put(t, "DevelopmentLength", 12); put(t, "OriginLength", 12); t.Save()
        OBSERVED["C2"] = {"O24_D12": o24, "O24_D24": o24d24}
        discard(t)

    def c3_stored_shape_survives_a_reconnect(self):
        """C3 — the stored pair, the display pair and Cumulative all persist across a reconnect."""
        t = self.empty_store("C3 reconnect")
        write_rows(t, ANNUAL_10x10)
        put(t, "DevelopmentLength", 12)
        t.Save()
        name = str(t.Name)
        saved = shape(t)
        self.rc.UnloadChildren()
        again = self.find_triangle(name)
        reloaded = snapshot(again)
        print(f"  saved   {saved}")
        print(f"  reloaded display {reloaded['display']} stored {reloaded['stored']} cum={reloaded['cumulative']}")
        OBSERVED["C3"] = {"saved": saved, "reloaded": reloaded, "name": name}
        discard(again)

    # ===== group D: reading and presentation =======================================

    def d1_label_arithmetic(self):
        """D1 — development labels and row widths at every legal display length."""
        t = self.empty_store("D1 labels")
        write_rows(t, ANNUAL_10x10); t.Save()
        table = {}
        for d in (1, 2, 3, 4, 6, 12):
            put(t, "DevelopmentLength", d)
            labels = dev_labels(t)
            widths = row_widths(t)
            table[f"D{d}"] = {"labels": labels, "row_widths": widths}
            print(f"  D{d}: {len(labels)} columns {labels[:4]} .. {labels[-2:]}  widths {widths}")
        put(t, "DevelopmentLength", 12)
        origins = {}
        for o in (12, 24):
            put(t, "OriginLength", o)
            origins[f"O{o}"] = {"labels": origin_labels(t), "row_widths": row_widths(t)}
            print(f"  O{o}: rows {origin_labels(t)}")
            put(t, "OriginLength", 12); put(t, "DevelopmentLength", 12)
        dates = {}
        for j in (1, 2, 10):
            dates[j] = {"label": str(t.DevelopmentLabel(j)),
                        "development_date": f"{t.GetDevelopmentDate(1, j):%Y-%m-%d}"}
        for i in (1, 10):
            dates[f"origin{i}"] = {"label": str(t.OriginLabel(i)),
                                   "origin_date": f"{t.GetOriginDate(i):%Y-%m-%d}"}
        print("  dates:", dates)
        OBSERVED["D1"] = {"development": table, "origin": origins, "dates": dates}
        t.Save()
        discard(t)

    def d2_calendarised_and_transposed(self):
        """D2 — the Calendar radio relabels the columns; Transposed is display-only to the API."""
        rec = []
        t = self.empty_store("D2 calendarised")
        write_rows(t, ANNUAL_10x10); t.Save()
        development = snapshot(t)
        if put(t, "Calendarised", True, rec):
            calendarised = snapshot(t)
            print("  calendarised labels:", calendarised["development_labels"])
            print("  values identical to the development view:",
                  calendarised["values"] == development["values"])
            put(t, "Calendarised", False, rec)
        else:
            calendarised = None
        if put(t, "Transposed", True, rec):
            transposed = snapshot(t)
            print(f"  transposed: origin_count={transposed['origin_count']} "
                  f"row_widths={transposed['row_widths'][:3]} first labels "
                  f"{transposed['origin_labels'][:2]} / {transposed['development_labels'][:2]}")
            put(t, "Transposed", False, rec)
        else:
            transposed = None
        OBSERVED["D2"] = {"development": development, "calendarised": calendarised,
                          "transposed": transposed, "calls": rec}
        t.Save()
        discard(t)

    def d3_leading_diagonal(self):
        """D3 — LeadingDiagonal reads the newest cell of each row at the current display."""
        t = self.empty_store("D3 diagonal")
        write_rows(t, ANNUAL_10x10); t.Save()
        by_display = {}
        for d in (1, 2, 3, 4, 6, 12):
            put(t, "DevelopmentLength", d)
            by_display[f"D{d}"] = [float(t.LeadingDiagonalByIndex(i)) for i in range(1, int(t.OriginCount) + 1)]
            print(f"  D{d}: {by_display[f'D{d}']}")
        put(t, "DevelopmentLength", 12)
        OBSERVED["D3"] = by_display
        t.Save()
        discard(t)

    def d4_monthly_origin_rollup(self):
        """D4 — an annual view of a monthly-origin store is a calendar diagonal (regression fixture)."""
        t = self.new_triangle("D4 monthly origin")
        put(t, "OriginLength", 1); put(t, "DevelopmentLength", 1); put(t, "Cumulative", True); t.Save(); show(t)
        n = fill(t, lambda k, d: 1000 * k + d); t.Save(); print(f"  filled {n} cells at O1/D1")
        before = grid(t)
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 12); show(t)
        print_grid(t, "annual view", max_cols=10)
        g = grid(t)

        def expected(y, a):   # each origin month of year y is a-m months old at the column's calendar end
            k0 = 12 * (y - 1) + 1
            return sum(1000 * (k0 + m) + (a - m) for m in range(12) if a - m >= 1 and k0 + m <= 113)
        ages = [int(l[:-1]) for l in dev_labels(t)]
        mism = [(y, ages[j], g[y - 1][j]) for y in range(1, 11) for j in range(11 - y)
                if abs(g[y - 1][j] - expected(y, ages[j])) > 1e-6]
        print(f"  calendar-diagonal origin roll-up mismatches: {len(mism)} {mism[:3]}")
        put(t, "OriginLength", 1); put(t, "DevelopmentLength", 1)
        print("  stored cells changed by the coarse view:", len(changed_cells(t, before, grid(t))))
        OBSERVED["D4"] = {"annual_grid": g, "mismatches": len(mism)}
        t.Save()
        discard(t)

    def d5_development_rollup_fixture(self):
        """D5 — a 12/1 store filled with 100000*row + age read at D12 (regression fixture)."""
        t = self.monthly_store("D5 development rollup")
        print_grid(t, "annual view of the monthly store", max_cols=10)
        OBSERVED["D5"] = snapshot(t)
        t.Save()
        discard(t)

    def d6_coarse_display_over_a_coarse_store(self):
        """D6 - a coarse display over a store other than 1 ends on the store's newest column."""
        table = {}
        for stored, displays in ((2, (4, 6, 12)), (3, (6, 12)), (6, (12,))):
            t = self.empty_store(f"D6 store {stored}", display_dev=12, stored_dev=stored)
            fill(t, lambda i, j: 0.0)
            put(t, "DevelopmentLength", stored)
            n = fill(t, lambda i, j: 1000 * i + j)
            t.Save()
            print(f"  store {stored}: wrote {n} stored cells; labels {dev_labels(t)[:3]} .. {dev_labels(t)[-2:]}")
            table[stored] = {"stored_labels": dev_labels(t), "displays": {}}
            for display in displays:
                put(t, "DevelopmentLength", display)
                labels = dev_labels(t)
                table[stored]["displays"][display] = {"labels": labels, "row_widths": row_widths(t)}
                print(f"    shown at {display}: {len(labels)} columns {labels[:3]} .. {labels[-2:]}")
            put(t, "DevelopmentLength", 12); t.Save()
            discard(t)
        OBSERVED["D6"] = table

    def d7_incremental_coarse_read(self):
        """D7 - the same store read at a coarse display, cumulative and incremental."""
        t = self.empty_store("D7 incremental read")
        write_rows(t, ANNUAL_10x10); t.Save()
        cumulative = snapshot(t)
        put(t, "Cumulative", False)
        incremental = snapshot(t)
        print_grid(t, "incremental at D12", max_cols=10)
        put(t, "DevelopmentLength", 1)
        fine = snapshot(t)
        labels = fine["development_labels"]
        row1 = [(labels[j], v) for j, v in enumerate(fine["values"][0]) if v]
        print(f"  incremental at D1, row 1: {row1[:6]}")
        put(t, "Cumulative", True); put(t, "DevelopmentLength", 12); t.Save()
        # A coarse incremental column is the difference of the cumulative view,
        # not a sum of the stored increments in its block.
        differences = [cumulative["values"][0][0]] + [
            round(cumulative["values"][0][j] - cumulative["values"][0][j - 1], 6)
            for j in range(1, len(ANNUAL_10x10[0]))
        ]
        matches = [round(v, 6) for v in incremental["values"][0]] == differences
        print(f"  coarse incremental row 1 is the difference of the cumulative view: {matches}")
        OBSERVED["D7"] = {"cumulative": cumulative, "incremental": incremental,
                          "incremental_at_the_store": fine,
                          "coarse_incremental_is_the_cumulative_difference": matches}
        discard(t)

    # ===== group E: vectors ========================================================

    def e1_origin_vector(self):
        """E1 — a vector's stored period follows its display while empty and is strict afterwards."""
        rec = []
        v = self.rc.Vectors().Add()
        v.DatasetType = self.project.DatasetTypes().Item(VECTOR_TYPE)
        v.Name = v.UniqueName(PREFIX + "E1 vector")
        print(f"  defaults: P{v.PeriodLength} stored P{v.StoredPeriodLength} count={v.Count}")
        put(v, "StoredPeriodLength", 1, rec)
        put(v, "PeriodLength", 1, rec); v.Save()
        print(f"  saved: P{v.PeriodLength} stored P{v.StoredPeriodLength} count={v.Count}")
        for i in range(1, int(v.Count) + 1):
            v.SetValuesByIndex(i, float(1000 + i))
        v.Save()
        put(v, "PeriodLength", 12, rec)
        coarse = [float(v.ValuesByIndex(i)) for i in range(1, 4)]
        print("  P12 values (sum of 12 months?):", coarse)
        call("SetValuesByIndex(2) at the coarse period", lambda: v.SetValuesByIndex(2, 777777.0), rec)
        put(v, "StoredPeriodLength", 12, rec)
        OBSERVED["E1"] = {"coarse_reads": coarse, "calls": rec,
                          "labels": [str(v.PeriodLabel(i)) for i in range(1, min(int(v.Count), 4) + 1)]}
        put(v, "PeriodLength", 1, rec); v.Save()
        discard(v)

    # ===== group F: the export write sequence ======================================

    def f1_cleardata_resync(self):
        """F1 — after ClearData, does the store follow the display on its own or only on a put?"""
        rec = []
        # The store has to differ from the display before ClearData, or a store
        # that never moved cannot be told from one that moved and came back.
        t = self.empty_store("F1 ClearData resync", stored_dev=1)
        write_rows(t, ANNUAL_10x10); t.Save(); show(t, "saved at display 12 over a store of 1: ")
        before_clear = shape(t)
        call("ClearData", t.ClearData, rec)
        after_clear = shape(t)
        print(f"  straight after ClearData: display {after_clear['display']} stored {after_clear['stored']}")
        put(t, "DevelopmentLength", 12, rec)
        after_same_display_put = shape(t)
        print(f"  after a same-value DevelopmentLength put: stored {after_same_display_put['stored']}")
        put(t, "StoredDevelopmentLength", 12, rec)
        after_stored_put = shape(t)
        print(f"  after StoredDevelopmentLength=12: stored {after_stored_put['stored']}")
        put(t, "StoredDevelopmentLength", 1, rec)
        OBSERVED["F1"] = {"before_clear": before_clear, "after_clear": after_clear,
                          "after_same_display_put": after_same_display_put,
                          "after_stored_put": after_stored_put, "calls": rec}
        t.Save()
        discard(t)

    def f2_export_write_sequence(self):
        """F2 — the export macro's sequence over a triangle already stored at 12/12."""
        rec = []
        t = self.empty_store("F2 export sequence", stored_dev=12)
        write_rows(t, ANNUAL_10x10); t.Save(); show(t, "target as ArcRho finds it: ")
        # ArcRho's own CSV after the same paste: a 12/1 store whose only non-zero
        # cells are at ages 5, 17, ... 113, exactly as case B2 showed ResQ storing
        # them, with cumulative 0 in between.
        stored_rows = []
        for i, row in enumerate(ANNUAL_10x10):
            cells = [0.0] * (113 - 12 * i)
            for j, value in enumerate(row):
                cells[(5 + 12 * j) - 1] = float(value)
            stored_rows.append(cells)
        call("ClearData", t.ClearData, rec)
        put(t, "DevelopmentLength", 12, rec)
        put(t, "StoredDevelopmentLength", 1, rec)
        put(t, "DevelopmentLength", 1, rec)
        show(t, "shown at the stored pair: ")
        n = write_rows(t, stored_rows)
        put(t, "DevelopmentLength", 12, rec)
        t.Save()
        print(f"  wrote {n} stored cells, then put the display back and saved")
        final = snapshot(t)
        print(f"  final display {final['display']} stored {final['stored']}")
        print_grid(t, "read back at D12", max_cols=10)
        matches = final["values"] == [[float(v) for v in row] for row in ANNUAL_10x10]
        print(f"  annual view matches ArcRho's grid: {matches}")
        OBSERVED["F2"] = {"final": final, "matches_arcrho_grid": matches, "calls": rec}
        discard(t)

    def _arcrho_stored_rows(self):
        """ArcRho's own 12/1 CSV for the annual grid: values at ages 5, 17 ... 113."""
        rows = []
        for i, row in enumerate(ANNUAL_10x10):
            cells = [0.0] * (113 - 12 * i)
            for j, value in enumerate(row):
                cells[(5 + 12 * j) - 1] = float(value)
            rows.append(cells)
        return rows

    def _restate(self, label, resq_stored, arcrho_display, arcrho_stored, reopen, write_grid):
        """Run the export's restating sequence over a triangle that already holds data.

        *resq_stored* is the shape ResQ holds the dataset at, *arcrho_display*
        and *arcrho_stored* the pair the sidecar asks for. With *reopen* the
        emptied triangle is saved and read again before the shape is moved,
        which is what the ResQ GUI asks a person to do.
        """
        rec = []
        origin, development = arcrho_display
        stored_origin, stored_development = arcrho_stored
        print(f"  -- {label}")
        t = self.new_triangle(f"F3 {label}")
        name = str(t.Name)
        put(t, "OriginLength", resq_stored[0]); put(t, "DevelopmentLength", resq_stored[1])
        put(t, "Cumulative", True); t.Save()
        fill(t, lambda i, j: 1.0); t.Save()
        show(t, "as ResQ holds it: ")
        before = shape(t)

        call("ClearData", t.ClearData, rec)
        after_clear = shape(t)
        print(f"    after ClearData: display {after_clear['display']} stored {after_clear['stored']}")
        if reopen:
            call("Save the emptied triangle", t.Save, rec)
            self.rc.UnloadChildren()
            t = self.find_triangle(name)
            if t is None:
                print("    REOPEN FAILED: triangle not found after UnloadChildren")
                return {"reopen_failed": True, "calls": rec}
            print(f"    reopened: display {shape(t)['display']} stored {shape(t)['stored']}")
            rec.append({"reopened": shape(t)})

        # The macro's own order: display pair, stored development length, then
        # show the triangle at the stored pair to write by index.
        put(t, "OriginLength", origin, rec)
        put(t, "DevelopmentLength", development, rec)
        after_display = shape(t)
        print(f"    after the display pair: display {after_display['display']} stored {after_display['stored']}")
        if stored_development != development:
            put(t, "StoredDevelopmentLength", stored_development, rec)
        if stored_development != development:
            put(t, "DevelopmentLength", stored_development, rec)
        if stored_origin != origin:
            put(t, "OriginLength", stored_origin, rec)
        at_store = shape(t)
        print(f"    shown at the stored pair: display {at_store['display']} stored {at_store['stored']}")

        wrote = 0
        if write_grid and at_store["stored"] == list(arcrho_stored):
            wrote = write_rows(t, self._arcrho_stored_rows())
        elif at_store["stored"] == list(arcrho_stored):
            wrote = write_rows(t, [[1.0, 2.0, 3.0], [4.0, 5.0]])
        else:
            print("    not at the stored shape ArcRho wants; no values written")
        put(t, "OriginLength", origin, rec)
        put(t, "DevelopmentLength", development, rec)
        call("Save", t.Save, rec)
        final = snapshot(t)
        matches = None
        if write_grid:
            matches = final["values"] == [[float(v) for v in row] for row in ANNUAL_10x10]
        print(f"    wrote {wrote} cells; final display {final['display']} stored {final['stored']}; "
              f"annual view matches ArcRho: {matches}")
        result = {
            "before": before, "after_clear": after_clear, "after_display": after_display,
            "at_store": at_store, "final_display": final["display"], "final_stored": final["stored"],
            "cells_written": wrote, "matches_arcrho_grid": matches, "calls": rec,
        }
        discard(t)
        return result

    def f3_restating_a_filled_triangle(self):
        """F3 - moving the stored lengths of a triangle that already holds data.

        The export's real case: ResQ holds the dataset stored at one shape and
        ArcRho stores it at another, so the store has to move before the values
        go in. The stored origin length moves in both directions, and the
        ClearData + Save + reload the ResQ GUI asks for is measured beside a
        plain ClearData.
        """
        OBSERVED["F3"] = {
            "origin store 1 to 12, clear only":
                self._restate("origin up, clear only", (1, 1), (12, 12), (12, 1), False, True),
            "origin store 1 to 12, clear save reload":
                self._restate("origin up, clear save reload", (1, 1), (12, 12), (12, 1), True, True),
            "origin store 12 to 1, clear save reload":
                self._restate("origin down", (12, 12), (12, 12), (1, 1), True, False),
        }

    # ===== cleanup =================================================================

    def delete_probe_objects(self):
        self.rc.UnloadChildren()
        for coll in (self.rc.Triangles(), self.rc.Vectors()):
            for item in [coll.Item(i) for i in range(1, coll.Count + 1)]:
                if str(item.Name).startswith(PREFIX):
                    item.Delete(); print("  deleted", item.Name)
        self.rc.UnloadChildren()
        print(f"  left: {self.rc.Triangles().Count} triangles, {self.rc.Vectors().Count} vectors")


CASES = [
    ("A", "A0 survey", "survey"),
    ("A", "A1 defaults of a new triangle", "a1_defaults"),
    ("A", "A2 put order on a never-saved triangle", "a2_put_order"),
    ("A", "A3 a display put moves the store while empty", "a3_display_put_moves_the_store"),
    ("A", "A4 StoredDevelopmentLength must be a factor of the display", "a4_stored_development_factors"),
    ("A", "A5 StoredOriginLength has no setter", "a5_stored_origin_has_no_setter"),
    ("A", "A6 the display development length must divide the display origin length", "a6_display_development_divides_origin"),
    ("A", "A7 an empty save records the stored pair", "a7_empty_save_persists"),
    ("A", "A8 an origin change on a saved empty triangle", "a8_origin_change_on_a_saved_empty_triangle"),
    ("A", "A9 the stored grid at each stored length", "a9_stored_grid_labels"),
    ("B", "B1 write at the stored shape", "b1_write_at_the_stored_shape"),
    ("B", "B2 a 10x10 annual paste into a monthly store", "b2_annual_paste_into_a_monthly_store"),
    ("B", "B3 one coarse cell over a filled store", "b3_coarse_write_over_a_filled_store"),
    ("B", "B4 a partial coarse write", "b4_partial_coarse_write"),
    ("B", "B5 an incremental coarse write", "b5_incremental_coarse_write"),
    ("B", "B6 a coarse origin write is refused", "b6_coarse_origin_write_refused"),
    ("B", "B7 SetValues by age at a coarse display", "b7_setvalues_by_age"),
    ("B", "B8 what counts as empty", "b8_zeros_are_empty"),
    ("B", "B9 ClearData frees the stored origin length", "b9_cleardata_frees_the_origin_store"),
    ("C", "C1 what saved data locks", "c1_locked_after_data"),
    ("C", "C2 a coarser origin display reads the calendar diagonal", "c2_coarse_origin_display_reads"),
    ("C", "C3 the shape survives a reconnect", "c3_stored_shape_survives_a_reconnect"),
    ("D", "D1 label arithmetic at every display length", "d1_label_arithmetic"),
    ("D", "D2 Calendarised and Transposed", "d2_calendarised_and_transposed"),
    ("D", "D3 the leading diagonal", "d3_leading_diagonal"),
    ("D", "D4 an annual view of a monthly-origin store", "d4_monthly_origin_rollup"),
    ("D", "D5 an annual view of a monthly-development store", "d5_development_rollup_fixture"),
    ("D", "D6 a coarse display over a coarse store", "d6_coarse_display_over_a_coarse_store"),
    ("D", "D7 a coarse incremental read", "d7_incremental_coarse_read"),
    ("E", "E1 origin vectors", "e1_origin_vector"),
    ("F", "F1 what ClearData frees", "f1_cleardata_resync"),
    ("F", "F2 the export write sequence", "f2_export_write_sequence"),
    ("F", "F3 restating a triangle that already holds data", "f3_restating_a_filled_triangle"),
]


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--only", nargs="+", metavar="GROUP", help="run only these case groups (A B C D E F) or case ids (A3 B2)")
    parser.add_argument("--keep", action="store_true", help="leave the probe objects in ResQ (delete them by hand later)")
    parser.add_argument("--json", metavar="PATH", help="write the observations to this file")
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    wanted = {value.upper() for value in (args.only or [])}
    app = connect()
    probe = None
    try:
        probe = Probe(app)
        for group, title, attr in CASES:
            case_id = title.split()[0]
            if wanted and group not in wanted and case_id not in wanted:
                continue
            section(title)
            try:
                getattr(probe, attr)()
            except Exception:  # noqa: BLE001
                traceback.print_exc()
        if not args.keep:
            section("cleanup")
            try:
                probe.delete_probe_objects()
            except Exception:  # noqa: BLE001
                traceback.print_exc()
    finally:
        if probe is not None:
            try:
                probe.rc.UnloadChildren()
            except Exception:  # noqa: BLE001
                pass
        app.Disconnect()
    if args.json:
        with open(args.json, "w", encoding="utf-8") as handle:
            json.dump(OBSERVED, handle, indent=2, default=str)
        print(f"\nobservations written to {args.json}")


if __name__ == "__main__":
    main()

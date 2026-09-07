"""Probe ResQ's stored-length and coarse-write rules against a live project.

Creates throwaway triangles and a vector named ``ArcRho probe ...`` of a
non-unique, non-calculated dataset type in one reserving class, exercises the
rules recorded in docs/plans/manual_input_stored_length_resq_alignment.md, and
deletes everything it created. Nothing that already exists in the class is
touched. Server PC only (ResQ COM), with a Python that has pywin32::

    py -3.10 tools/resq_stored_length_probe.py
    py -3.10 tools/resq_stored_length_probe.py --keep   # leave the probe objects for a GUI look

Early binding throughout, and every property put goes through IDispatch so a
refused set raises ResQ's own error text instead of pywin32 quietly creating a
Python attribute on the wrapper.
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


# ----- COM helpers ----------------------------------------------------------------

def connect():
    cfg = json.load(open(CONFIG))["resq"]
    app = gencache.EnsureDispatch("ResQ3Automation.ResQApplication")
    app.ConnectByName(cfg["connection_name"], cfg["user_name"], cfg["password"])
    return app


def err_text(exc):
    if isinstance(exc, pywintypes.com_error):
        args = exc.args
        desc = str(args[2][2]) if len(args) > 2 and args[2] and len(args[2]) > 2 else ""
        return f"{args[1]!r} {desc!r}"
    return f"{type(exc).__name__}: {exc}"


def put(obj, name, value):
    """Property put through Invoke; prints and returns whether ResQ accepted it."""
    try:
        dispid = obj._oleobj_.GetIDsOfNames(name)
        obj._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, value)
        print(f"  put {name}={value!r}: ok")
        return True
    except Exception as exc:  # noqa: BLE001
        print(f"  put {name}={value!r}: REFUSED {err_text(exc)}")
        return False


def section(title):
    print("\n" + "=" * 8, title)


def run(title, fn):
    section(title)
    try:
        fn()
    except Exception:  # noqa: BLE001
        traceback.print_exc()


# ----- triangle helpers ------------------------------------------------------------

def dev_labels(t):
    return [str(t.DevelopmentLabel(j)) for j in range(1, int(t.DevelopmentCountByIndex(1)) + 1)]


def row_widths(t):
    return [int(t.DevelopmentCountByIndex(i)) for i in range(1, int(t.OriginCount) + 1)]


def grid(t):
    return [[float(t.ValuesByIndex(i, j)) for j in range(1, int(t.DevelopmentCountByIndex(i)) + 1)]
            for i in range(1, int(t.OriginCount) + 1)]


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


def changed_cells(t, before, after):
    labels = dev_labels(t)
    return [(i + 1, j + 1, labels[j], before[i][j], after[i][j])
            for i in range(len(before)) for j in range(len(before[i])) if before[i][j] != after[i][j]]


class Probe:
    def __init__(self, app):
        self.app = app
        self.project = app.Projects().Item(PROJECT)
        self.rc = self.project.ReservingClasses().Item(RC_PATH)

    def new_triangle(self, name):
        t = self.rc.Triangles().Add()
        t.DatasetType = self.project.DatasetTypes().Item(TRIANGLE_TYPE)
        t.Name = t.UniqueName(PREFIX + name)
        return t

    def monthly_store(self, name):
        """Saved triangle stored O12/D1, shown O12/D12, cumulative, filled with 100000*row + age."""
        t = self.new_triangle(name)
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 1); put(t, "Cumulative", True); t.Save()
        n = fill(t, lambda i, m: 100000 * i + m); t.Save()
        put(t, "DevelopmentLength", 12); show(t, f"filled {n} monthly cells, now ")
        return t

    # -- tests ------------------------------------------------------------------

    def survey(self):
        p = self.project
        print(f"  project {p.Name}: origins {p.OriginStartDate:%Y-%m-%d}..{p.OriginEndDate:%Y-%m-%d}, "
              f"Development End Date {p.DevelopmentEndDate:%Y-%m-%d}, O{p.OriginLength}/D{p.DevelopmentLength}, {p.OriginCount} origins")
        dt = p.DatasetTypes().Item(TRIANGLE_TYPE)
        print(f"  type {dt.Name!r}: Unique={dt.Unique} Calculated={dt.Calculated}")

    def t1_rules(self):
        t = self.new_triangle("T1 rules")
        show(t, "defaults after Add: ")
        print("  -- a display put on an empty triangle moves the stored length with it (and OriginLength resets development to 1 before the first save)")
        put(t, "OriginLength", 6); show(t)
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 12); show(t)
        print("  -- StoredDevelopmentLength must be a factor of the display length")
        put(t, "StoredDevelopmentLength", 1); show(t)
        put(t, "StoredDevelopmentLength", 5)
        put(t, "StoredDevelopmentLength", 3); show(t)
        print("  -- a display put on an empty triangle is never refused; it resyncs the store")
        put(t, "DevelopmentLength", 4); show(t)
        put(t, "StoredOriginLength", 6)
        put(t, "DevelopmentLength", 12); put(t, "StoredDevelopmentLength", 1); t.Save(); show(t, "saved: ")
        print("  -- labels follow the stored grid, grouped from the newest stored cell")
        for d in (1, 3, 12):
            put(t, "DevelopmentLength", d); print(f"     D{d}: {dev_labels(t)[:5]} .. {dev_labels(t)[-2:]} stored D{t.StoredDevelopmentLength}")
        put(t, "DevelopmentLength", 12); put(t, "StoredDevelopmentLength", 1); t.Save(); show(t, "final: ")

    def t2_coarse_write_empty(self):
        t = self.new_triangle("T2 coarse write empty")
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 12); put(t, "StoredDevelopmentLength", 1); put(t, "Cumulative", True); t.Save()
        n = fill(t, lambda i, j: 1000 * i + j); t.Save(); print(f"  wrote {n} annual cells at D12 into the D1 store")
        print_grid(t, "read back at D12")
        put(t, "DevelopmentLength", 1)
        nz = nonzero(t); print(f"  non-zero stored cells (cumulative): {len(nz)}; row 1: {[(j, l, v) for i, j, l, v in nz if i == 1][:5]}")
        put(t, "Cumulative", False)
        nz = nonzero(t); print(f"  non-zero stored cells (incremental): {len(nz)}; row 1: {[(j, l, v) for i, j, l, v in nz if i == 1][:6]}")
        put(t, "Cumulative", True); put(t, "DevelopmentLength", 12); t.Save()

    def t3_coarse_cell_over_filled(self):
        t = self.monthly_store("T3 one coarse cell over a filled store")
        put(t, "DevelopmentLength", 1); before = grid(t); put(t, "DevelopmentLength", 12)
        t.SetValuesByIndex(2, 2, 999999.0); print("  SetValuesByIndex(2,2)=999999 at D12, no Save yet")
        put(t, "DevelopmentLength", 1); changed = changed_cells(t, before, grid(t))
        print(f"  stored cells changed before Save: {len(changed)} in rows {sorted({c[0] for c in changed})}")
        print("  row 2:", [(j, l, b, a) for i, j, l, b, a in changed if i == 2][:18])
        put(t, "DevelopmentLength", 12); t.Save()

    def t4_incremental_coarse_write(self):
        t = self.new_triangle("T4 incremental coarse write")
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 12); put(t, "StoredDevelopmentLength", 1); put(t, "Cumulative", False); t.Save()
        fill(t, lambda i, j: 1000 * i + j); t.Save()
        print_grid(t, "read back incremental")
        put(t, "Cumulative", True); print_grid(t, "read back cumulative")
        put(t, "DevelopmentLength", 1)
        nz = nonzero(t); print(f"  non-zero stored cells (cumulative): {len(nz)}; row 1: {[(j, l, v) for i, j, l, v in nz if i == 1][:4]}")
        put(t, "DevelopmentLength", 12); t.Save()

    def t5_emptiness(self):
        t = self.new_triangle("T5 emptiness")
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 12); put(t, "StoredDevelopmentLength", 1); put(t, "Cumulative", True); t.Save()
        print("  -- unsaved value"); t.SetValuesByIndex(1, 1, 5.0); put(t, "StoredDevelopmentLength", 3)
        print("  -- saved value"); t.Save(); put(t, "StoredDevelopmentLength", 3)
        print("  -- explicit zeros everywhere, saved"); fill(t, lambda i, j: 0.0); t.Save(); put(t, "StoredDevelopmentLength", 3)
        print("  -- value saved, then ClearData without Save"); put(t, "StoredDevelopmentLength", 1)
        t.SetValuesByIndex(1, 1, 7.0); t.Save(); t.ClearData(); put(t, "StoredDevelopmentLength", 3)
        put(t, "StoredDevelopmentLength", 1); t.Save()

    def t6_monthly_origin(self):
        t = self.new_triangle("T6 monthly origin")
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
        mism = [(y, ages[j], g[y - 1][j]) for y in range(1, 11) for j in range(11 - y) if abs(g[y - 1][j] - expected(y, ages[j])) > 1e-6]
        print(f"  calendar-diagonal origin roll-up mismatches: {len(mism)} {mism[:3]}")
        try:
            t.SetValuesByIndex(2, 2, 7777777.0); t.Save(); print("  write at the coarse origin display: accepted")
        except Exception as exc:  # noqa: BLE001
            print("  write at the coarse origin display: REFUSED", err_text(exc))
        put(t, "OriginLength", 1); put(t, "DevelopmentLength", 1)
        print("  stored cells changed:", len(changed_cells(t, before, grid(t))))
        put(t, "OriginLength", 12); show(t, "O12/D1 view of the O1/D1 store: "); put(t, "OriginLength", 1); t.Save()

    def t9_setvalues_by_age(self):
        t = self.new_triangle("T9 SetValues by age")
        put(t, "OriginLength", 12); put(t, "DevelopmentLength", 12); put(t, "StoredDevelopmentLength", 1); put(t, "Cumulative", True); t.Save()
        d = datetime.datetime(2017, 1, 1)
        for m, v in ((17, 55555.0), (10, 44444.0), (5, 33333.0)):
            t.SetValues(d, m, v); print(f"  SetValues(2017-01-01, {m}, {v}) at D12: ok")
        t.Save(); print("  D12 row 1:", grid(t)[0][:3])
        put(t, "DevelopmentLength", 1); print("  stored cells:", nonzero(t))
        put(t, "DevelopmentLength", 12); t.Save()

    def v2_origin_vector(self):
        v = self.rc.Vectors().Add()
        v.DatasetType = self.project.DatasetTypes().Item(VECTOR_TYPE)
        v.Name = v.UniqueName(PREFIX + "V2")
        print(f"  defaults: P{v.PeriodLength} stored P{v.StoredPeriodLength} count={v.Count}")
        put(v, "StoredPeriodLength", 1)
        put(v, "PeriodLength", 1); v.Save(); print(f"  saved: P{v.PeriodLength} stored P{v.StoredPeriodLength} count={v.Count}")
        for i in range(1, int(v.Count) + 1):
            v.SetValuesByIndex(i, float(1000 + i))
        v.Save()
        put(v, "PeriodLength", 12); print("  P12 values (sum of 12 months?):", [float(v.ValuesByIndex(i)) for i in range(1, 4)])
        try:
            v.SetValuesByIndex(2, 777777.0); v.Save(); print("  write at the coarse period display: accepted")
        except Exception as exc:  # noqa: BLE001
            print("  write at the coarse period display: REFUSED", err_text(exc))
        put(v, "PeriodLength", 1); v.Save()

    # -- cleanup ----------------------------------------------------------------

    def delete_probe_objects(self):
        self.rc.UnloadChildren()
        for coll in (self.rc.Triangles(), self.rc.Vectors()):
            for item in [coll.Item(i) for i in range(1, coll.Count + 1)]:
                if str(item.Name).startswith(PREFIX):
                    item.Delete(); print("  deleted", item.Name)
        self.rc.UnloadChildren()
        print(f"  left: {self.rc.Triangles().Count} triangles, {self.rc.Vectors().Count} vectors")


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--keep", action="store_true", help="leave the probe objects in ResQ (delete them by hand later)")
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    app = connect()
    try:
        probe = Probe(app)
        run("survey", probe.survey)
        run("T1 stored-length rules on an empty triangle", probe.t1_rules)
        run("T2 coarse development write into an empty monthly store", probe.t2_coarse_write_empty)
        run("T3 one coarse cell over a filled monthly store", probe.t3_coarse_cell_over_filled)
        run("T4 incremental coarse write", probe.t4_incremental_coarse_write)
        run("T5 what counts as empty", probe.t5_emptiness)
        run("T6 monthly origin store", probe.t6_monthly_origin)
        run("T9 SetValues(date, age) at a coarse display", probe.t9_setvalues_by_age)
        run("V2 origin vector", probe.v2_origin_vector)
        if not args.keep:
            run("cleanup", probe.delete_probe_objects)
    finally:
        try:
            probe.rc.UnloadChildren()
        except Exception:  # noqa: BLE001
            pass
        app.Disconnect()


if __name__ == "__main__":
    main()

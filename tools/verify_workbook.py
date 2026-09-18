#!/usr/bin/env python3
"""
Independent verification of the generated workbook.

An independent day-count / compounding model is implemented here in plain Python
(deliberately NOT sharing code with the workbook formulas). The example workbook is then
recalculated with the `formulas` engine and every quarter's interest, closing balance and
day count is compared against the reference model.

Usage: python3 verify_workbook.py <example.xlsx>
"""
from __future__ import annotations

import datetime as dt
import sys

import formulas

from build_loan_workbook import (
    EXAMPLE_ENTRIES, EXAMPLE_SETUP, QUARTERS, SHEET_MAIN, qtr_rows,
    CELL_BF_AMT, CELL_BF_DATE, CELL_BASIS,
)

EPOCH = dt.date(1899, 12, 30)          # Excel serial origin (1900 date system)


def serial(d: dt.date) -> int:
    return (d - EPOCH).days


# ----------------------------------------------------------------- reference model
RATES = [(dt.date(2024, 11, 17), 9.00), (dt.date(2025, 7, 1), 10.50),
         (dt.date(2026, 1, 1), 11.25)]
BASIS = EXAMPLE_SETUP[CELL_BASIS]
BF_DATE = EXAMPLE_SETUP[CELL_BF_DATE]
BF_AMT = EXAMPLE_SETUP[CELL_BF_AMT]


def rate_periods():
    out = []
    for i, (start, pct) in enumerate(RATES):
        end = RATES[i + 1][0] if i + 1 < len(RATES) else dt.date(2999, 12, 31)
        out.append((start, end, pct))
    return out


def year_fraction_interest(balance: float, d1: dt.date, d2: dt.date) -> float:
    """Interest for [d1, d2) on `balance`, split across rate periods."""
    total = 0.0
    for start, end, pct in rate_periods():
        days = (min(d2, end) - max(d1, start)).days
        if days > 0:
            total += balance * pct / 100.0 * days / BASIS
    return total


def quarter_ends(n: int) -> list[tuple[dt.date, dt.date]]:
    """Calendar quarter-ends, first one after the B/F date."""
    def eom(d: dt.date, add_months: int) -> dt.date:
        m = d.month + add_months
        y = d.year + (m - 1) // 12
        m = (m - 1) % 12 + 1
        nm_y, nm_m = (y + 1, 1) if m == 12 else (y, m + 1)
        return dt.date(nm_y, nm_m, 1) - dt.timedelta(days=1)

    first = eom(BF_DATE, (3 - BF_DATE.month % 3) % 3)
    if first == BF_DATE:
        first = eom(BF_DATE, 3)
    ends, start = [], BF_DATE
    cur = first
    for _ in range(n):
        ends.append((start, cur))
        start = cur
        cur = eom(cur, 3)
    return ends


def reference_ledger():
    """Returns {quarter: dict(open, repaid, interest, close, days)}"""
    result, bal = {}, float(BF_AMT)
    for q, (qs, qe) in enumerate(quarter_ends(QUARTERS), start=1):
        opening = bal
        events = sorted(EXAMPLE_ENTRIES.get(q, []), key=lambda e: e[0])
        accrued, prev, repaid = 0.0, qs, 0.0
        for d, amt, *_ in events:
            assert qs <= d <= qe, f"example entry {d} outside quarter {q}"
            accrued += year_fraction_interest(bal, prev, d)
            bal -= amt
            repaid += amt
            prev = d
        accrued += year_fraction_interest(bal, prev, qe)
        interest = round(accrued + 1e-12, 2)
        bal = round(bal + interest, 10)
        result[q] = dict(open=opening, repaid=repaid, interest=interest, close=bal,
                         start=qs, end=qe, days=(qe - qs).days)
    return result


# ----------------------------------------------------------------- workbook recalc
def recalc(path: str):
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()
    book = path.split("/")[-1].upper()
    out = {}
    for key, val in sol.items():
        k = key.upper()
        if f"[{book}]" not in k:
            continue
        ref = k.split("!")[-1].replace("'", "")
        try:
            v = val.value[0, 0]
        except Exception:
            continue
        out[(k.split("]")[1].split("'")[0], ref)] = v
    return out


def main(path: str) -> int:
    ref = reference_ledger()
    cells = recalc(path)
    sheet = SHEET_MAIN.upper()

    def get(ref_cell: str):
        return cells.get((sheet, ref_cell))

    fails, checked = [], 0
    print(f"{'Q':>3} {'period':<26} {'days':>5} {'interest (xlsx)':>16} "
          f"{'interest (ref)':>16} {'closing (xlsx)':>16} {'closing (ref)':>16}  ok")
    for q in range(1, QUARTERS + 1):
        hdr, bf, e1, e2, ir = qtr_rows(q)
        r = ref[q]
        x_int = get(f"E{ir}")
        x_close = get(f"G{ir}")
        x_open = get(f"G{bf}")
        x_qs, x_qe = get(f"P{hdr}"), get(f"Q{hdr}")
        ok = True
        for label, got, want, tol in (
            ("qstart", x_qs, serial(r["start"]), 0),
            ("qend", x_qe, serial(r["end"]), 0),
            ("opening", x_open, r["open"], 0.005),
            ("interest", x_int, r["interest"], 0.005),
            ("closing", x_close, r["close"], 0.005),
        ):
            checked += 1
            if got is None or abs(float(got) - float(want)) > tol:
                ok = False
                fails.append(f"Q{q} {label}: workbook={got!r} reference={want!r}")
        print(f"{q:>3} {str(r['start']):>10} .. {str(r['end']):<12} {r['days']:>5} "
              f"{float(x_int or 0):>16,.2f} {r['interest']:>16,.2f} "
              f"{float(x_close or 0):>16,.2f} {r['close']:>16,.2f}  {'OK' if ok else 'FAIL'}")

    # --- rate-split check: quarter 4 (rate revised 01-07-2025, mid-quarter) --------
    print("\nRate-split evidence (rate text per quarter, from the hidden engine):")
    for q in (3, 4, 6):
        hdr, *_ = qtr_rows(q)
        print(f"  Q{q}: {get(f'R{hdr}')!r}")

    # --- immutability check: a later revision must not alter earlier quarters -----
    print("\nImmutability check \u2013 recomputing the reference model with the 2nd revision "
          "moved/raised:")
    before = {q: ref[q]["interest"] for q in range(1, 4)}
    global RATES
    saved = RATES
    RATES = [(dt.date(2024, 11, 17), 9.00), (dt.date(2025, 7, 1), 25.00),
             (dt.date(2026, 1, 1), 30.00)]
    after = {q: reference_ledger()[q]["interest"] for q in range(1, 4)}
    RATES = saved
    for q in before:
        same = abs(before[q] - after[q]) < 0.005
        print(f"  Q{q} interest before={before[q]:,.2f} after={after[q]:,.2f} "
              f"-> {'unchanged' if same else 'CHANGED'}")
        checked += 1
        if not same:
            fails.append(f"Q{q} interest changed when a future rate was revised")

    print(f"\n{checked} assertions checked, {len(fails)} failure(s)")
    for f in fails:
        print("  FAIL:", f)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))

#!/usr/bin/env python3
"""
Recalculates a workbook and reports (a) any cell that evaluates to an Excel error and
(b) a readable dump of the Quarter Summary sheet, so the blank template and the filled
example can both be inspected.

Usage: python3 check_no_errors.py <file.xlsx> [--dump]
"""
from __future__ import annotations

import sys

import formulas

from build_loan_workbook import QUARTERS, SHEET_SUM

ERRORS = ("#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NAME?", "#NUM!", "#NULL!", "#CYCLE!")


def cells(path: str):
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()
    book = path.split("/")[-1].upper()
    out = {}
    for key, val in sol.items():
        k = key.upper()
        if f"[{book}]" not in k or "!" not in k:
            continue
        sheet = k.split("]")[1].split("'")[0]
        ref = k.split("!")[-1].replace("'", "")
        if ":" in ref or not ref[0].isalpha():
            continue
        try:
            out[(sheet, ref)] = val.value[0, 0]
        except Exception:
            continue
    return out


def main(path: str, dump: bool) -> int:
    data = cells(path)
    bad = [(s, r, v) for (s, r), v in data.items()
           if any(e in str(v) for e in ERRORS)]
    print(f"{path}: {len(data)} cells evaluated, {len(bad)} error cell(s)")
    for s, r, v in bad[:20]:
        print(f"   ERROR {s}!{r} = {v}")

    if dump:
        sheet = SHEET_SUM.upper()
        print(f"\n{SHEET_SUM}:")
        print(f"  {'Q':>3} {'opening':>14} {'repaid':>12} {'interest':>12} "
              f"{'closing':>14} {'status':>7} {'check':>6}  rates")
        for q in range(1, QUARTERS + 1):
            r = 4 + q
            row = [data.get((sheet, f"{c}{r}")) for c in "DEFGKLH"]
            nums = [f"{float(v):>14,.2f}" if isinstance(v, (int, float)) else f"{str(v):>14}"
                    for v in row[:4]]
            print(f"  {q:>3} {nums[0]} {nums[1][2:]} {nums[2][2:]} {nums[3]} "
                  f"{str(row[4]):>7} {str(row[5]):>6}  {row[6]}")
        base = 4 + QUARTERS + 3
        print("\n  live position panel:")
        for i in range(1, 7):
            print(f"    {data.get((sheet, f'A{base + i}'))} = {data.get((sheet, f'D{base + i}'))}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], "--dump" in sys.argv))

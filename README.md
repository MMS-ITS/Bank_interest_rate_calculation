# Bank loan – quarterly compound interest statement (Excel)

A protected, self-calculating Excel workbook that keeps a loan account in bank-statement
form: balance brought forward, repayments entered by date and amount, interest accrued on
the daily outstanding balance, capitalised at every quarter-end and carried forward to the
next quarter. Up to three interest rates are supported, each with its own effective date,
so a rate revision can never recompute a period that closed before it took effect.

## Files

| File | What it is |
| --- | --- |
| `Bank_Loan_Quarterly_Compound_Interest_TEMPLATE.xlsx` | Empty, ready to use. Fill the yellow cells. |
| `Bank_Loan_Quarterly_Compound_Interest_EXAMPLE.xlsx` | Same workbook pre-filled with a worked scenario (see below) so the behaviour can be checked. |
| `tools/build_loan_workbook.py` | Generator. Rebuilds either file from scratch. |
| `tools/verify_workbook.py` | Recalculates the workbook and compares every quarter against an independent Python model of the same loan. |
| `tools/audit_workbook.py` | Proves the locking claims and that past quarters survive a rate revision. |
| `tools/check_no_errors.py` | Recalculates a workbook and reports any cell that evaluates to an Excel error. |

Workbook sheets: **Loan Statement** (the ledger), **Quarter Summary**, **How to Use**, and a
very-hidden protected **Engine** sheet holding the rate periods and quarter calendar.

## What you type, and what the file does

Yellow, gold-bordered cells are the only cells that accept input:

* **Account particulars** – A/C no., borrower, facility, address, sanctioned limit.
* **Loan basis** – balance B/F date and amount, day basis (365 / 366 / 360), quarter-end
  convention (calendar quarter-ends, or three months at a time from the B/F date), and the
  **Locked / reconciled up to** date.
* **Rate register** – slot 1 is the original rate (always effective from the B/F date);
  slots 2 and 3 are revisions, each with an effective-from date.
* **Ledger** – per repayment: date, cheque/instrument, particulars, credit amount. Optional
  *Balance per Bank* for reconciliation.

Everything else is a locked formula: quarter-end dates, day counts, interest per period,
accrued interest, capitalisation at quarter-end, closing balance, the balance forward row
of the next quarter, variance against the bank, lock status, and all totals.

24 quarter blocks are generated, each with 10 repayment rows. Unused rows and quarters that
have not started yet stay blank.

## Interest convention

```
interest for a period = balance outstanding × rate in force × actual days ÷ basis
```

Interest is accrued segment by segment between events (a repayment, a rate change, a
quarter-end), accumulated in the *Accrued interest* column, then debited and added to the
balance on the quarter-end row. The capitalised figure is rounded to 2 decimals at that
point only – as a bank ledger does – and becomes the next quarter's opening balance, so
interest compounds quarterly. A repayment reduces the balance from its value date.

## How earlier periods are protected against a later rate change

1. **Effective-date arithmetic.** Each accrual segment is split across the rate periods it
   overlaps: `MAX(0, MIN(segment_end, rate_end) − MAX(segment_start, rate_start))` days at
   each rate. A rate effective 01-07-2025 has zero overlap with any period that ended on
   30-06-2025, so earlier quarters are arithmetically unreachable by it.
2. **Hard cut-off.** The *Locked / reconciled up to* date is enforced by data validation:
   no transaction may be dated on or before it, and no rate revision may be dated on or
   before it. Those rows turn grey and read `LOCKED`.
3. **Permanent audit trail.** The title line states every rate with its effective date, the
   quarter bar prints the rates and day counts used inside that quarter (e.g. `9.00% for
   1 day  10.50% for 91 days`), and the Quarter Summary records them per quarter.
4. **Tamper detection.** Fill *Balance per Bank* on reconciled rows; the Variance column and
   the Quarter Summary's *Arithmetic check* turn red the moment any historical figure stops
   agreeing with the bank or with the ledger.

## Why the programmable cells cannot be disturbed

* All four sheets are protected with a password (`MMS-loan-2026`); only the yellow cells are
  unlocked. An audit confirms **exactly 1278 typeable cells and not one formula among them**.
* Inserting, deleting or sorting rows and columns is blocked, and the workbook structure is
  locked, so sheets cannot be added, deleted or unhidden.
* The calculation engine sits on a very-hidden, protected sheet.
* Data validation rejects: a date outside its quarter block, a date out of sequence, a date
  in a locked period, an amount typed without a date, a rate outside 0–100%, and a
  back-dated or out-of-order rate revision. Each input carries a hover prompt.

### Known limit

Plain `.xlsx` has no "write-once" cell: while a period is still open, the original rate
percentage remains editable by whoever has the file (which is also how you correct a typo).
Two answers are provided: the reconciliation columns make any such change visible
immediately, and `--freeze-setup` produces a build where the opening position and the
original rate are written as **locked** cells, leaving only future repayments, revisions and
the lock date editable.

## Printing (A4)

Both worksheets are set up for A4, so `Ctrl+P` is all that is needed:

* **Loan Statement** – A4 landscape, scaled to exactly one page wide, column widths tuned so
  the 14 columns fit at roughly 90% scale (still legible). The column-heading row repeats on
  every page, manual page breaks keep each quarter block whole on a page (two blocks on
  page 1 below the setup panel, three per page after that), and the footer carries
  `Page n of m`.
* **Quarter Summary** – A4 landscape, one page wide, heading row repeated.
* **How to Use** – A4 portrait, one page wide.

## No defined names, by design

Excel 365 rejected the first build with *"Removed Records: Named range from
/xl/workbook.xml"* followed by the loss of formulas on the two sheets that referenced those
names. The workbook therefore carries **no workbook-level defined names at all**: the
generator still writes formulas with readable symbolic names (`Rate_2`, `Locked_To`, …) but
`resolve()` expands each one into a direct cell reference (`'Engine'!$B$9`, `$H$10`, …) as
the cell is written, for worksheet formulas and data-validation rules alike. Emoji were also
removed from formula strings, since non-BMP characters are a needless risk inside formulas.
The only remaining entries are the two standard `_xlnm.Print_Titles` records that Excel
itself writes for repeating print headings.

## Rebuilding and verifying

```bash
pip install openpyxl formulas

python3 tools/build_loan_workbook.py Bank_Loan_Quarterly_Compound_Interest_TEMPLATE.xlsx
python3 tools/build_loan_workbook.py Bank_Loan_Quarterly_Compound_Interest_EXAMPLE.xlsx --example
python3 tools/build_loan_workbook.py frozen.xlsx --example --freeze-setup   # locked opening position

cd tools
python3 verify_workbook.py  ../Bank_Loan_Quarterly_Compound_Interest_EXAMPLE.xlsx
python3 audit_workbook.py   ../Bank_Loan_Quarterly_Compound_Interest_TEMPLATE.xlsx \
                            ../Bank_Loan_Quarterly_Compound_Interest_EXAMPLE.xlsx
python3 check_no_errors.py  ../Bank_Loan_Quarterly_Compound_Interest_TEMPLATE.xlsx
```

Verification status of the committed files:

* 24 of 24 quarters match an independently written day-count/compounding model to the paisa,
  including the two quarters in which a rate revision lands mid-quarter (123 assertions, 0
  failures).
* Raising the second rate from 10.50% to 25.00%, and moving its effective date, leaves every
  quarter that closed earlier bit-identical while later quarters pick the new rate up.
* 0 error cells out of 4,716 recalculated cells in the empty template and 4,737 in the
  example.
* Validated with the Open XML SDK schema validator (`FileFormatVersions.Microsoft365`):
  the only findings are 7 font child-order notices that openpyxl emits for every file it
  writes — a trivial two-cell openpyxl workbook produces the same two notices, and Excel
  accepts them (its repair log listed no style records).

## Worked example in the example file

Opening balance 5,681,000.00 on 17-11-2024 at 9.00% p.a., revised to 10.50% w.e.f.
01-07-2025 and 11.25% w.e.f. 01-01-2026, calendar quarter-ends, actual/365, seven
repayments, locked and reconciled up to 31-03-2025.

| Qtr | Period | Opening | Repaid | Interest | Closing | Rates used |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | 17-11-2024 – 31-12-2024 | 5,681,000.00 | 500,000.00 | 59,662.36 | 5,240,662.36 | 9.00% / 44 days |
| 2 | 31-12-2024 – 31-03-2025 | 5,240,662.36 | 500,000.00 | 109,592.78 | 4,850,255.14 | 9.00% / 90 days |
| 3 | 31-03-2025 – 30-06-2025 | 4,850,255.14 | 400,000.00 | 104,294.77 | 4,554,549.91 | 9.00% / 91 days |
| 4 | 30-06-2025 – 30-09-2025 | 4,554,549.91 | 250,000.00 | 116,684.61 | 4,421,234.52 | 9.00% / 1 day + 10.50% / 91 days |
| 5 | 30-09-2025 – 31-12-2025 | 4,421,234.52 | 350,000.00 | 111,372.95 | 4,182,607.47 | 10.50% / 92 days |
| 6 | 31-12-2025 – 31-03-2026 | 4,182,607.47 | 300,000.00 | 112,332.28 | 3,994,939.75 | 10.50% / 1 day + 11.25% / 89 days |
| 7 | 31-03-2026 – 30-06-2026 | 3,994,939.75 | 250,000.00 | 108,505.33 | 3,853,445.08 | 11.25% / 91 days |

## Options available on request

More than three rate slots; more repayment rows per quarter or a longer tenure; repayments
applied to outstanding interest before principal instead of quarter-end capitalisation;
penal/overdue rate on arrears; a macro-enabled (`.xlsm`) build that freezes each quarter
automatically on save.

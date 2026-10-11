---
name: gst-quick-method
description: Maintain a Canadian GST/HST Quick Method (GST34) workbook for a quarterly filer — add a quarter's invoices from a pasted screenshot of an accounting income report (e.g. QuickBooks), record capital purchases that earn an input tax credit (from a pasted receipt or invoice), and produce the exact numbers to type into the CRA GST/HST return screen, all as Excel formulas, independently cross-checked. Needs only uv (plus Excel to open the workbook); works on Windows, macOS, Linux and WSL. Use whenever the user says "add Q3 to the GST", "here are this quarter's invoices", pastes an income report screenshot, asks "what do I owe for GST", "fill the GST return", "add a capital purchase / I bought a laptop", pastes a receipt or invoice for equipment, or starts a new fiscal year's GST workbook.
---

# gst-quick-method: quarterly GST return workbook

Accuracy is the whole point. Never guess a number; if a digit in a screenshot is unreadable, ask.

## Assumptions and requirements

**Not tax or accounting advice.** Not tax or accounting advice. This tool is provided as is, with no warranty. You are responsible for everything you file with CRA. Before filing, check the numbers against your own records, and confirm that the Quick Method applies to you and that the rates on the Settings sheet are current for your province and type of business. Rates and rules change. Not affiliated with or endorsed by the Canada Revenue Agency. MIT licensed (see `LICENSE` at the repository root).

Whenever you give the user numbers to file, end with a one-line reminder to check them against their own
records before submitting, and to confirm the rates on the Settings sheet are current.

- Canadian GST/HST registrant on the **Quick Method**, filing **quarterly**. Defaults: 5% GST
  (GST-only province), 3.6% remittance rate (services), 1% credit on the first $30,000 of
  GST-included sales per fiscal year. Rates live on the workbook's Settings sheet; check CRA's
  current Quick Method rates for your province and business type before relying on them.
- Invoice amounts are entered **before GST** (what an income-account report such as QuickBooks'
  shows). Check once: report amount × 1.05 should equal the invoice total.
- Requires [uv](https://docs.astral.sh/uv/) (it installs the Python dependencies) and Microsoft
  Excel to open and use the workbook. Works on Windows, macOS, Linux and WSL.
- Run the tool as `uv run scripts/gst.py ...` (paths relative to this skill's directory; use the
  full path from elsewhere). This works the same from PowerShell, cmd, Git Bash, zsh or bash.
- `check` evaluates every formula with a pure-Python engine; when Excel is reachable (native Windows
  or WSL) it also recalculates in real Excel and scans for Excel warning flags.
- Workbook location: env `GST_WORKBOOK`, or `~/.config/gst-quick-method/config.json`:
  `{"workbook": "/path/to/{year}/GST Returns {year} - Quick Method.xlsx"}` (`{year}` = fiscal year).
  `uv run scripts/gst.py path [year]` prints the resolved path. If nothing is configured, ask the user
  where to keep it and write the config file.

## Rules

- Only ever change the workbook through `scripts/gst.py`; never touch other files in that folder.
- Workbooks, screenshots and row data are financial data: keep them outside any git repository.
  Put temporary JSON in a scratch directory and delete it afterwards.
- The tool backs up the workbook to `_backup/` beside it before every change, and refuses to
  write while Excel has it open (`~$` lock file) — ask the user to close it.

## The math (do not change without telling the user)

- Invoice GST = ROUND(amount excl. GST × 5%, 2); total = amount + GST.
- "Sales and other revenues" (101) = ROUND(quarter's GST-included sales, 0).
- "GST/HST Amounts Collected or that Became Collectible" (103) = line 101 × remittance rate.
- "Input Tax Credits Adjustments" (107) = 1% Quick Method credit, limited to the first $30,000 of
  the fiscal year (carries across quarters automatically).
- "Input Tax Credits" (106) = GST paid on capital purchases used primarily (>50%) in the business.
- "Net tax" (109) = 103 + 104 − 106 − 107. Positive → amount to pay (115); negative → "Refund claimed" (114).

All of it is Excel formulas in the workbook; the tool only writes input cells.

## Workflow: add a quarter from a screenshot

1. Read the screenshot (an income report: Date, Num, Name, Memo, Amount).
2. Transcribe to JSON in a scratch directory, one object per row:
   `{"date":"2026-07-31","invoice":123,"customer":"Client Co","memo":"Consulting Services ABC","project":"ABC","amount":"1000.00"}`
   `project` is optional (e.g. a short code from the memo). Amounts as strings with 2 decimals,
   no `$` or commas. Dates `YYYY-MM-DD` (or `MM/DD/YYYY`).
3. **Verify the transcription before writing**: look at the image again and compare row by row
   (count, every invoice #, every date, every amount digit). If the report shows a total row, the
   sum must equal it.
4. Dry run, show the user the printed rows and total:
   `uv run scripts/gst.py add-sales "<workbook>" <rows.json> --dry-run`
5. Write: same command without `--dry-run`. Already-present invoice numbers are skipped, so a full
   year-to-date screenshot is fine — and if a skipped invoice's date or amount disagrees with the
   workbook, the tool stops with CONFLICT and writes nothing (resolve with the user).
   Rows are re-sorted by date. Dates outside the fiscal year are rejected.
6. `uv run scripts/gst.py check "<workbook>"` — evaluates every formula (Python engine, plus real Excel
   when reachable), cross-checks every line against an independent calculation, and scans for
   Excel warning triangles when Excel is used. Must say `OK`.
7. Report the quarter's CRA fields **in the order of the CRA screen** (labels verbatim):
   Payment Period Start Date, Payment Period End Date, Sales and other revenues, Instalment & Annual
   Filer Payments, Rebates, GST/HST due on purchases of real property or purchases of Emission allowances,
   Other GST/HST to be Self-Assessed, Total GST/HST new housing rebates included in line 108, Net tax,
   Refund claimed, GST/HST Amounts Collected or that Became Collectible, GST/HST Adjustments,
   Input Tax Credits, Input Tax Credits Adjustments — plus the amount to pay and the due date
   (one month after the quarter ends). Dates on the CRA screen are DD/MM/YYYY.
   Also give the YEAR TO DATE panel (Summary!I:J): gross income (incl. GST), net income
   (excl. GST), GST charged, invoice count, latest invoice date.

## Capital purchases (ITC refund)

Computers, equipment, furniture, etc. Not operating expenses like software subscriptions, supplies
or meals (the Quick Method rate already accounts for those).

The user may paste a receipt, an invoice, or an asset/expense report from their accounting software.

1. Read the image and transcribe, one row per capital item:
   `{"date":"2026-08-15","vendor":"Store","description":"Laptop","cost":"2000.00","gst":"100.00","business":"Y"}`
   - `cost` = price before taxes. `gst` = the **GST** printed on the receipt, never PST, QST or the
     tax total (in BC the receipt usually shows GST 5% and PST 7% separately). In HST provinces, use the HST.
   - If one receipt mixes a capital item with supplies or services, include only the capital item, and
     take its GST from the receipt's line detail. If the receipt shows only a total GST, ask how it splits.
   - If no GST is shown (e.g. a private sale or a non-registered seller), there is no ITC. Don't add the row.
2. **Ask** whether each item is used more than 50% for the business (`business` "Y"/"N"). A receipt
   can't tell you this; don't assume it.
3. Verify the transcription against the image (date, amounts, the GST figure, not the total), then dry run:
   `uv run scripts/gst.py add-capital "<workbook>" <rows.json> --dry-run`
4. Write (same command without `--dry-run`), then `check`. Report the quarter's new "Input Tax Credits"
   (line 106) and net tax. On the Capital sheet, orange means the GST is above 5% of the cost, which
   usually means PST got mixed in.

## After filing

Record it (after confirming the amount equals the quarter's amount to pay):
`uv run scripts/gst.py record-payment "<workbook>" --quarter 3 --date YYYY-MM-DD --amount 1234.56 [--confirmation N]`
Business name / GST account number for the Summary header:
`uv run scripts/gst.py set-business "<workbook>" --name "..." --bn "123456789RT0001"`

## New fiscal year

`uv run scripts/gst.py build "$(uv run scripts/gst.py path <yyyy>)" --fy-start <yyyy>-01-01 [--business NAME] [--bn NUMBER]`
(refuses to overwrite; the fiscal year may start on any date).

## After changing the template (scripts/gst.py layout)

`uv run scripts/gst.py rebuild "<workbook>"` rebuilds on the current template and carries every input across
(settings, invoices, capital rows, typed Summary inputs such as payments). Then run `check`.

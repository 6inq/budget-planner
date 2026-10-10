# Changelog

## 1.0.1

### Added
- **Document upload during the questionnaire.**
  - Share pay stubs, bills, or bank and credit card statements (PDF, photo, or CSV), and Claude extracts your pay schedule, take-home pay, bills, due dates, and spending.
  - Everything it found is shown in a confirmation table before it goes into the workbook.
- **`import_transactions.py`**: exact per-merchant, per-month totals from bank or card CSV exports.
  - Handles bank preamble lines, debit/credit columns, transaction-type columns, day-first dates, semicolon files, and masked card numbers.
  - Flags recurring bills.
- **Update without losing edits.**
  - `read_workbook.py` reads an existing workbook back into a config, including values typed directly in Excel and per-paycheck amounts and notes.
  - Rebuilds keep a `(backup)` copy of the previous version via `build_budget.py --backup`.
- Questionnaire choices are saved inside the workbook (no financial data), so updates keep the same look.
- A small, muted "Budget Planner - Claude Skill by 6inq" credit line on every sheet (linked to this repo), in the print footer, and in the file properties.
- **Privacy rules in the skill**: document data is used only for the user's workbook, never sent elsewhere, and sensitive identifiers are never copied. A test enforces that the scripts contain no networking code.

## 1.0.0

- First release: Personal and Business budgets, an annual overview dashboard, four chart styles, light and dark mode, pay and payroll calendars, and savings goal and debt trackers.

# Changelog

## 1.0.3

### Added
- **Loan Payoff Plan sheet** (optional, `features.payoff_plan`), for paying down a car or home loan faster, including one worth more than the car or home.
  - `loan_type` `car` (72-month schedule) or `house` (360-month schedule, home wording, and a home value that can rise).
  - The questionnaire asks whether to add a Payoff Plan and, if so, for the loan's details.
  - Inputs: loan balance, APR, regular payment, vehicle value and monthly value drop, take-home per paycheck, paychecks per month, monthly bills, savings %, and a buffer for surprises.
  - Works out a monthly extra-payment goal from what is left after bills, savings and the buffer.
  - A 72-month schedule shows interest, the extra payment, the balance, the vehicle's value, how far underwater the loan is, and when it is above water and paid off, next to the same dates with regular payments only.
  - An Extra Paid column logs what was really paid; blank months use the goal.
  - The schedule ends at the payoff month and grows or shrinks as payments change; the regular-payments-only dates still cover the full 72 months.
  - An optional target (how far underwater is acceptable, for example before trading in) shows the month it is reached and highlights that row.
  - Inputs and Extra Paid amounts are kept by `read_workbook.py`.

## 1.0.2

### Added
- **Paid tracking on the Expenses sheet.**
  - A "Paid" column for each month, right of the Annual Total, with a dropdown on every bill.
  - Marking a month Paid turns that amount and its Paid cell green. Totals and the Dashboard do not change.
  - New PAID and LEFT TO PAY rows under TOTAL show what is settled and what is still due, per month and for the year.
  - Paid marks are saved in the config (`expenses[].paid`, 12 true/false values) and kept by `read_workbook.py`, so updates never lose them.

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

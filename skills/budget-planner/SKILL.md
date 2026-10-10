---
name: budget-planner
description: Builds a professional, formula-driven Excel budget workbook (.xlsx) for Personal or Business finances. The first sheet is an annual overview dashboard (Money In / Money Out / Net / Savings Rate cards, or Revenue / Expenses / Net Profit / Margin for business, a monthly cash-flow chart, a spending-by-category chart, and a month-by-month breakdown), followed by Income or Revenue, a Paycheck or Payroll calendar, Expenses, and optional goal and debt trackers, in light or dark mode. Runs a short guided questionnaire first. Use when the user says "Start Budgeting", "budget planner", "make me a budget", "business budget", "build a budget spreadsheet", "track my bills, expenses and income in Excel", or similar.
---

# Budget Planner

Build a polished Excel budget through a short guided questionnaire, then generate it with `scripts/build_budget.py`. Every total in the workbook is a live Excel formula, so the user can keep editing it and the Dashboard updates itself.

## Step 0: Prerequisites

Run `python -c "import openpyxl"`. If it fails, run `python -m pip install openpyxl` (use `python3` if `python` is missing).

## Step 1: Questionnaire

Ask in rounds. If a structured question tool is available (for example AskUserQuestion), use it with the options below; it allows up to 4 questions per call. Otherwise send ONE message per round with numbered questions and numbered options, and tell the user they can answer like `1, 2, 2, 1`. Skip anything the user already told you.

**Round 1**
1. "Is this budget for Personal or Business finances?" Personal / Business
2. "For your dashboard's spending breakdown, what do you prefer?" Pie Chart / Doughnut Chart / Bar Graph / Column Graph
3. "Light Mode or Dark Mode?" Light Mode / Dark Mode
4. "Pick an accent color." Violet / Teal / Rose / Gold

**Round 2**
1. Pay frequency.
   - Personal: "How often do you get paid?"
   - Business: "How often do you run payroll?"
   - Options: Weekly / Bi-weekly (every 2 weeks) / Semi-monthly (twice a month) / Monthly. If the user answers "irregular" or "no payroll", use frequency `none`.
2. "Which budget period?" January - December {current year} / January - December {next year} / Starting this month (next 12 months)
3. "Which currency?" $ (USD) / EUR / GBP / Other
4. "Add any extra trackers?" (multi-select, optional)
   - Personal: Savings Goals / Debt Tracker
   - Business: Reserve Goals / Loans & Credit

**Round 3**: the payday detail (only if it applies), then ALWAYS end with the fill-out question.
- Semi-monthly: "Which two days do you get paid?" (multi-select, exactly two) 1st of the month / 15th of the month / Last day of the month. If the user picks one or three, ask again.
- Weekly or Bi-weekly: "What is the date of your first payday in this budget?" Offer the first two Fridays of the budget's first month as options; accept any date.
- Monthly: "Which day do you get paid?" 1st / 15th / Last day of the month / Other (a day from 1 to 28).
- LAST question: "Do you want to fill out the sheet, or me?" I'll fill it out myself / You fill it out for me

## Step 2: Build

Write the answers to `budget-config.json` in the current working directory. `scripts/example_config.json` (personal) and `scripts/example_business_config.json` (business) show every field. Use these mappings:

| Answer | Config field and value |
|---|---|
| Personal / Business | `budget_type`: `personal` / `business` |
| Pie Chart / Doughnut Chart / Bar Graph / Column Graph | `category_chart`: `pie` / `doughnut` / `bar` / `column` |
| Light Mode / Dark Mode | `mode`: `light` / `dark` |
| Accent color | `accent`: `violet` / `teal` / `rose` / `gold` |
| Pay frequency | `pay.frequency`: `weekly` / `biweekly` / `semimonthly` / `monthly` / `none` |
| Semi-monthly paydays | `pay.semimonthly_days`: two of `"1"`, `"15"`, `"last"` |
| Weekly / Bi-weekly first payday | `pay.first_payday`: `"YYYY-MM-DD"` |
| Monthly payday | `pay.monthly_day`: `"1"`, `"15"`, `"last"` or a day 1-28 |
| Budget period | `year` and `start_month` ("Starting this month" sets `start_month` to the current month) |
| Currency | `currency_symbol`: `$`, `€` or `£` |
| Trackers | `features.savings_goals`, `features.debt_tracker` (the business names map to the same two flags) |

Leave `income`, `expenses`, `savings_goals` and `debts` empty for now. Then run:

```
python "<skill dir>/scripts/build_budget.py" --config budget-config.json
```

The default file name is `Budget Planner 2026.xlsx` or `Business Budget 2026.xlsx`; pass `--out` to choose another. The script prints a JSON summary. If it prints `ERROR:`, fix the config and run again. Never overwrite a workbook from a different budget without asking.

## Step 3: Finish

**If the user chose "I'll fill it out myself"**: give the file path, then end with instructions telling them where to enter their numbers:
- Personal:
  - **Paychecks** sheet: your take-home pay on each payday (every payday is already listed).
  - **Income** sheet: any other income per month (side jobs, bonuses, interest).
  - **Expenses** sheet: your bills and expenses per month. Pick a Category and a Type (Bill or Expense) from the dropdowns; starter rows are already there.
- Business:
  - **Revenue** sheet: expected revenue per month for each stream.
  - **Payroll** sheet: the total cost of each payroll run.
  - **Expenses** sheet: costs per month, marked Fixed or Variable; starter rows are already there.
- Goal and debt sheets, if added: targets, balances, APR and minimum payments.
- The Dashboard updates automatically. Highlighted months on it are 3-paycheck (or 3-payroll) months.

**If the user chose "You fill it out for me"**: give the file path, then end with a message asking them to send their numbers in chat:
- Personal:
  - Take-home pay per paycheck, plus any other income.
  - Bills: name, amount, how often, and due day (rent, utilities, phone, insurance, subscriptions, loans).
  - Expenses: typical monthly amounts (groceries, gas, dining out, entertainment, personal).
- Business:
  - Revenue streams and expected monthly amounts.
  - Payroll cost per run.
  - Fixed costs (rent, software, insurance, bookkeeping) and variable costs (inventory, ads, contractors, fees).
- Goals (target, saved so far, target date) and debts or loans (balance, APR, minimum payment), if those trackers were added.

When they reply, write the data into `budget-config.json` and re-run the script to the same file:
- Fixed recurring costs get type `Bill` (personal) or `Fixed` (business). Variable spending gets `Expense` or `Variable`.
- Pick the closest category from the profile's list; a new category name is fine if nothing fits.
- Amounts that change by month go in a `months` list of 12 numbers.
- Put the pay amount in `pay.amount`.
- Never invent amounts. Ask if something is unclear.

Then report the totals from the printed summary: Money In, Money Out, Net and Savings Rate (business: Revenue, Expenses, Net Profit and Margin).

## Reference: summary JSON

- `annual_income`: money in, or revenue for business.
- `annual_fixed`: bills, or fixed costs including payroll for business.
- `annual_variable`: variable spending.
- `annual_out`, `net`, `savings_rate`: totals for the year.
- `paychecks`: the number of pay or payroll dates.
- `monthly`: a list of 12 entries, each with `in`, `out` and `net`.

---
name: budget-planner
description: Builds and updates a professional, formula-driven Excel budget workbook (.xlsx) for Personal or Business finances. The first sheet is an annual overview dashboard (Money In / Money Out / Net / Savings Rate cards, or Revenue / Expenses / Net Profit / Margin for business, a monthly cash-flow chart, a spending-by-category chart, and a month-by-month breakdown), followed by Income or Revenue, a Paycheck or Payroll calendar, Expenses, and optional goal and debt trackers, in light or dark mode. Runs a short guided questionnaire and can pull numbers from uploaded pay stubs, bills and bank or card statements (PDF, image or CSV). Use when the user says "Start Budgeting", "budget planner", "make me a budget", "business budget", "update my budget", "build a budget spreadsheet", "track my bills, expenses and income in Excel", or uploads financial documents for a budget.
---

# Budget Planner

Build a polished Excel budget through a short guided questionnaire, optionally filled from the user's own documents, then generate it with `scripts/build_budget.py`. Every total in the workbook is a live Excel formula, so the user can keep editing it and the Dashboard updates itself.

## Privacy rules (always apply)

Users may share pay stubs, bills and bank statements. Treat them as confidential:

1. Use document contents ONLY to build or update this user's workbook. Never send them, or anything taken from them, to any other tool, website, search, API or person. Do not web-search payee names or amounts.
2. Extract only what the budget needs: dates, amounts, pay frequency, payee or merchant names, due dates, balances, APRs and minimum payments.
   - Never copy account numbers, card numbers, routing numbers, Social Security or tax ID numbers, or addresses into the config, the workbook or your replies.
   - If you need to refer to an account, use the last 4 digits at most.
3. The scripts run locally and make no network connections. Keep their inputs and outputs in the user's working directory; do not copy documents anywhere else.
4. Do not repeat sensitive details back in chat beyond the confirmation summary described below.
5. When you finish, remind the user that `budget-config.json` holds their numbers too. They can delete it, but it is needed for future updates unless the workbook is read back (see Updating).
6. If the user asks, say plainly: the files they share are processed by Claude in this conversation, and this skill does not upload, store or share them anywhere else.

## Step 0: Prerequisites

Run `python -c "import openpyxl"`. If it fails, run `python -m pip install openpyxl` (use `python3` if `python` is missing).

If the user already has a workbook made by this skill and wants to change or add to it, skip to **Updating an existing workbook**.

## Step 1: Questionnaire

Ask in rounds. If a structured question tool is available (for example AskUserQuestion), use it with the options below; it allows up to 4 questions per call. Otherwise send ONE message per round with numbered questions and numbered options, and tell the user they can answer like `1, 2, 2, 1`. Skip anything the user already told you or that their documents already answer.

**Round 1**
1. "Is this budget for Personal or Business finances?" Personal / Business
2. "For your dashboard's spending breakdown, what do you prefer?" Pie Chart / Doughnut Chart / Bar Graph / Column Graph
3. "Light Mode or Dark Mode?" Light Mode / Dark Mode
4. "Pick an accent color." Violet / Teal / Rose / Gold

**Round 2**
1. "Do you have documents I can pull your numbers from? For example pay stubs, bills, or bank and credit card statements (PDF, photo or CSV). They are only used to build your spreadsheet." Yes, I'll upload them / No, skip this
2. "Which budget period?" January - December {current year} / January - December {next year} / Starting this month (next 12 months)
3. "Which currency?" $ (USD) / EUR / GBP / Other
4. "Add any extra trackers?" (multi-select, optional)
   - Personal: Savings Goals / Debt Tracker
   - Business: Reserve Goals / Loans & Credit

If they said yes to documents, run **Step 1b** before Round 3.

**Round 3**: ask only what is still unknown, then ALWAYS end with the fill-out question.
1. Pay frequency, unless a pay stub already showed it.
   - Personal: "How often do you get paid?"
   - Business: "How often do you run payroll?"
   - Options: Weekly / Bi-weekly (every 2 weeks) / Semi-monthly (twice a month) / Monthly. If the user answers "irregular" or "no payroll", use frequency `none`.
2. The payday detail, if it applies:
   - Semi-monthly: "Which two days do you get paid?" (multi-select, exactly two) 1st of the month / 15th of the month / Last day of the month. If the user picks one or three, ask again.
   - Weekly or Bi-weekly: "What is the date of your first payday in this budget?" Offer the first two Fridays of the budget's first month as options (or the date implied by a pay stub); accept any date.
   - Monthly: "Which day do you get paid?" 1st / 15th / Last day of the month / Other (a day from 1 to 28).
3. Payoff Plan: "Do you want a Payoff Plan sheet to pay off a car or home loan faster?" Yes, a car loan / Yes, a home loan / No thanks.
   - If yes, ask for the loan details in one message (they can also upload the lender statement):
     - Loan balance, interest rate (APR), and regular monthly payment (for a home loan, principal + interest only, not escrow).
     - What the car or home is worth today (for example a Kelley Blue Book or Zillow estimate).
     - When extra payments should start, and optionally how far underwater they are willing to be (the target).
     - How much they want to save each month (% of take-home) and a small monthly buffer for surprises.
   - Never guess loan numbers. If they do not have them yet, build the sheet with those cells blank and tell them which to fill.
4. LAST question: "Do you want to fill out the sheet, or me?" I'll fill it out myself / You fill it out for me. If documents already covered most items, phrase the options as "I'll fill in the rest myself" / "You fill in the rest for me".

## Step 1b: Reading the user's documents

Ask the user to attach the files in chat; in Claude Code they can also give file paths or a folder. Then read each one:

**Pay stubs** give you:
- Net (take-home) pay. For business payroll registers, use the total employer cost per run.
- The pay date.
- The pay frequency, from the pay-period length: 7 days = weekly, 14 = bi-weekly, two fixed dates a month = semi-monthly, one = monthly.
- Use the pay date to set `first_payday`, or the semi-monthly days.
- If several stubs show different amounts, use the most recent as `pay.amount` and mention the difference to the user.

**Bills** (utilities, phone, insurance, rent, subscriptions, loan statements) give you: payee, amount due, due date (day of month), and how often it repeats. These are fixed costs: type `Bill` (personal) or `Fixed` (business). For loans and cards with the Debt Tracker on, also take the balance, APR and minimum payment.

**Bank or credit card statements**:
- For a CSV export, run the summarizer so totals are exact:
  ```
  python "<skill dir>/scripts/import_transactions.py" statement.csv [more.csv ...] --start YYYY-MM
  ```
  - `--start` is the budget's first month.
  - Add `--invert` if purchases show as positive numbers, which is common for credit card exports.
  - Add `--day-first` if dates are written as day/month.
  - The script prints merchants grouped under `money_out` and `money_in`, each with exact monthly totals, `average_per_month`, `recurring` (a steady repeating charge) and `typical_day`. It reads only the date, description and amount columns.
- For a PDF or photo statement, read it directly and total the items carefully. If it is long, suggest the user also download the CSV export from their bank for exact totals.
- Map merchants to the closest profile category:
  - Recurring merchants become bills or fixed costs, with `due_day` taken from `typical_day`.
  - Everything else becomes expenses or variable costs.
  - Skip transfers between the user's own accounts and credit card payments; they are not spending.
- Where the statement covers budget months, put the actual amounts in those months (`months`, 12 numbers). For the remaining months use the monthly average, and say you projected them.

Then show ONE compact confirmation table: Item | Amount | How often | Category | Type | Source document. Include anything uncertain, marked with a "?". Ask the user to confirm or correct it before building. Never invent amounts; if a document is unreadable or a number is unclear, ask.

## Step 2: Build

Write the answers and any confirmed document data to `budget-config.json` in the current working directory. `scripts/example_config.json` (personal) and `scripts/example_business_config.json` (business) show every field. Use these mappings:

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
| Payoff Plan: car loan / home loan | `features.payoff_plan: true` and `payoff_plan.loan_type`: `car` / `house`, plus the loan details (see Loan payoff plan) |

Then run:

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
  - Paid columns (right of the Annual Total): pick "Paid" under a month once that bill is paid. It turns green, and the PAID and LEFT TO PAY rows under TOTAL show what is still due.
- Business:
  - **Revenue** sheet: expected revenue per month for each stream.
  - **Payroll** sheet: the total cost of each payroll run.
  - **Expenses** sheet: costs per month, marked Fixed or Variable; starter rows are already there.
- Goal and debt sheets, if added: targets, balances, APR and minimum payments.
- The Dashboard updates automatically. Highlighted months on it are 3-paycheck (or 3-payroll) months.
- They can also come back any time, upload new documents, and ask you to update the workbook.

**If the user chose "You fill it out for me"**: give the file path, then end with a message asking them to send their numbers in chat (or upload more documents):
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

## Updating an existing workbook

Use this whenever the user wants to change, add to or refresh a workbook made by this skill, for example "add my new electric bill" or "here is March's statement". The workbook may contain edits they typed directly in Excel, so ALWAYS read it back first instead of rebuilding from an old config:

1. Ask the user to close the workbook in Excel.
2. Read it back:
   ```
   python "<skill dir>/scripts/read_workbook.py" --workbook "<their file>.xlsx" --out budget-config.json
   ```
   - This recovers their original choices (saved inside the workbook) and every value they entered: income, expenses, each paycheck amount and note, goals and debts.
   - If it reports that a cell has a formula with no saved result, ask them to open the file in Excel, save it, and close it again.
   - If it reports a cell that should be a number, tell them which cell and ask what it should be.
3. Merge in the new information: use Step 1b for any documents, or apply the changes they asked for. Show the confirmation table for anything taken from documents.
   - When the user says a bill is paid, set that month to `true` in the item's `paid` list (12 true/false values, in budget order). Keep the amount; never zero it out to show it was paid.
4. Rebuild to the same file, keeping a backup copy of the previous version:
   ```
   python "<skill dir>/scripts/build_budget.py" --config budget-config.json --out "<their file>.xlsx" --backup
   ```
5. Tell them what changed (new or updated rows and the new annual totals), and that the previous version was saved as `<name> (backup).xlsx`.

If `read_workbook.py` reports no saved settings, the file was made by another tool or an older version. Offer to build a fresh workbook through the questionnaire, then enter their existing numbers into it.

## Loan payoff plan (on request)

When the user wants to pay a car or home loan down faster (from the questionnaire or later), or get above water on a car or home worth less than they owe, turn on `features.payoff_plan` and fill `payoff_plan`:
- `loan_type`: `car` (72-month schedule) or `house` (360-month schedule, for a mortgage).
- From their lender: `balance`, `apr`, `payment`. Ask for these; never guess them. If they want the plan before they have them, use a clearly stated placeholder and tell them which cell to update.
- `asset_value` (what the car or home is worth) and `value_drop` (its expected change per month; an estimate, say so). A car usually loses value; a home may gain it, so use a negative `value_drop` for a rising home value.
- For a home loan, `payment` is principal + interest only. Escrow (taxes and insurance) is not part of the loan; count it in `monthly_bills`.
- `take_home` per paycheck and `paychecks_per_month` (defaults from the pay frequency).
- `monthly_bills`: their regular monthly bills, not counting this loan or savings.
- `savings_percent`, `buffer` (a monthly cushion for surprises), and `start` (`YYYY-MM-DD`, when extra payments begin).
- Optional `target_gap`: the most they are willing to still be underwater by (for example before trading in). The sheet shows the month it is reached and highlights that row.

The sheet shows the monthly extra-payment goal and when the loan will be above water and paid off, with and without the extra. Report those dates and the interest saved. When the user says what they actually paid extra in a month, put it in `extra_paid` (a list by plan month; `null` for months not logged yet).

## Reference: summary JSON

- `annual_income`: money in, or revenue for business.
- `annual_fixed`: bills, or fixed costs including payroll for business.
- `annual_variable`: variable spending.
- `annual_out`, `net`, `savings_rate`: totals for the year.
- `paychecks`: the number of pay or payroll dates.
- `monthly`: a list of 12 entries, each with `in`, `out` and `net`.
- `backup`: present when `--backup` saved a copy of the previous file.

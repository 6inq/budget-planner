# Budget Planner: a Claude Skill

Turn a short conversation with Claude into a professional Excel budget. Say **"Start Budgeting"**, answer a few multiple-choice questions, and Claude builds a polished, fully formula-driven `.xlsx` workbook with an annual dashboard, charts, and a month-by-month breakdown, for **personal** or **business** finances, in **light** or **dark** mode.

![Personal budget dashboard, dark mode](docs/images/personal-dark-Dashboard.png)

| Business, light mode | Personal, light mode |
|---|---|
| ![Business dashboard](docs/images/business-light-Dashboard.png) | ![Personal dashboard, light](docs/images/personal-light-Dashboard.png) |

## Features

- **Annual overview dashboard** as the first sheet: six KPI cards, a Money In vs Money Out chart by month with a Net trend line, a spending-by-category chart, and a 12-month breakdown table.
- **Personal or Business**:
  - Personal tracks paychecks, bills, expenses, savings goals, and debts.
  - Business tracks revenue, payroll, fixed and variable costs, reserve goals, and loans, and reports net profit and margin.
- **Your choice of chart**: pie, doughnut, bar, or column.
- **Light or dark mode** with four accent colors: violet, teal, rose, and gold.
- **Pay schedule aware**: weekly, bi-weekly, semi-monthly (any two of the 1st, the 15th, and the last day of the month), monthly, or irregular. Every payday for the year is listed on its own sheet, and months with an extra paycheck or payroll run are highlighted on the dashboard.
- **Live formulas everywhere**: change any number in Excel and every total, chart, and KPI updates. Nothing is a static value.
- **Fill it yourself or let Claude do it**: get a ready-to-type template with starter rows and dropdowns, or tell Claude your numbers in chat and it fills the workbook in for you.
- **Optional trackers**:
  - Savings or reserve goals: progress bars and the monthly amount needed to hit the target date.
  - Debts or loans: months to payoff, estimated interest, and payoff date.
- Works in **Microsoft Excel**, and also opens in Google Sheets, LibreOffice, and Numbers.

## How it works

Once the skill is installed, start a chat and say **Start Budgeting** (or "make me a budget", "business budget", and so on). Claude asks a few rounds of questions:

1. Personal or Business finances?
2. Dashboard chart: Pie, Doughnut, Bar, or Column?
3. Light Mode or Dark Mode?
4. Accent color?
5. How often do you get paid (or, for a business, how often do you run payroll)?
6. If semi-monthly: which two days? (1st, 15th, or last day of the month)
7. Budget period, currency, and optional trackers.
8. **Do you want to fill out the sheet, or me?**

What happens next depends on the last answer:

- **"I'll fill it out myself"**: you get the workbook with starter rows already in place, plus clear instructions on which sheet gets which numbers.
- **"You fill it out for me"**: Claude builds the workbook, asks for your income, bills, and expenses in plain language, fills them in, and reports your annual totals.

## Installation

### Claude.ai (web, desktop, or mobile app)

1. Download **[budget-planner.zip](https://github.com/6inq/budget-planner/releases/latest/download/budget-planner.zip)** from the latest release.
2. In Claude, open **Settings > Capabilities** and make sure code execution is turned on.
3. Under **Skills**, choose **Upload skill** and select the zip file.
4. Start a new chat and say **Start Budgeting**.

### Claude Code (plugin)

```
/plugin marketplace add 6inq/budget-planner
/plugin install budget-planner@6inq
```

### Claude Code (manual)

Clone the repo and copy the skill folder into your personal skills directory:

```bash
git clone https://github.com/6inq/budget-planner.git
cp -r budget-planner/skills/budget-planner ~/.claude/skills/
```

On Windows (PowerShell):

```powershell
git clone https://github.com/6inq/budget-planner.git
Copy-Item -Recurse budget-planner\skills\budget-planner "$HOME\.claude\skills\"
```

You can also just ask Claude Code: *"Install the skill from https://github.com/6inq/budget-planner"*.

### Requirements

- Python 3.10 or newer
- [openpyxl](https://pypi.org/project/openpyxl/). The skill installs it automatically if it's missing, or you can run `pip install openpyxl`.

## What's in the workbook

| Sheet | Personal | Business |
|---|---|---|
| Dashboard | Money In, Money Out, Net Savings, Savings Rate, Avg Monthly Spend, Paychecks | Revenue, Expenses, Net Profit, Profit Margin, Avg Monthly Burn, Payroll Runs |
| Income / Revenue | Your paycheck row (automatic) plus other income per month | Revenue streams per month |
| Paychecks / Payroll | Every payday in the year, with an editable take-home amount | Every payroll date, with an editable cost per run |
| Expenses | Bills (fixed) and Expenses (variable), with Category and Type dropdowns | Fixed and Variable costs, with payroll pulled in automatically |
| Savings Goals / Reserve Goals (optional) | Target, saved so far, progress, months left, needed per month | Same, for cash reserves and tax set-asides |
| Debt Tracker / Loans & Credit (optional) | Balance, APR, minimum payment, months to payoff, estimated interest | Same, for business cards and loans |

More screenshots:

| Paycheck calendar | Bills & expenses |
|---|---|
| ![Paycheck calendar](docs/images/personal-dark-Paychecks.png) | ![Expenses sheet](docs/images/personal-dark-Expenses.png) |

Sample workbooks you can open right away are in [`samples/`](samples).

## Using the generator without Claude

The skill is a thin layer over a plain Python script, so you can also run it directly:

```bash
python skills/budget-planner/scripts/build_budget.py --config my-budget.json --out "My Budget.xlsx"
```

Start from [`example_config.json`](skills/budget-planner/scripts/example_config.json) (personal) or [`example_business_config.json`](skills/budget-planner/scripts/example_business_config.json) (business). The script prints a JSON summary of the annual totals. If the config is invalid, it prints `ERROR: ...` and exits with code 2.

### Config reference

| Field | Values |
|---|---|
| `budget_type` | `personal` (default) or `business` |
| `year`, `start_month` | e.g. `2026`, `1`. The budget covers 12 months from `start_month`. |
| `mode` | `light` or `dark` |
| `accent` | `violet`, `teal`, `rose`, `gold` |
| `category_chart` | `pie`, `doughnut`, `bar`, `column` |
| `currency_symbol` | e.g. `$`, `€`, `£` |
| `pay.frequency` | `weekly`, `biweekly`, `semimonthly`, `monthly`, `none` |
| `pay.amount` | Take-home pay per paycheck (personal) or cost per payroll run (business) |
| `pay.first_payday` | Weekly or bi-weekly only: `"YYYY-MM-DD"` |
| `pay.semimonthly_days` | Exactly two of `"1"`, `"15"`, `"last"` |
| `pay.monthly_day` | `"1"`, `"15"`, `"last"`, or a day from 1 to 28 |
| `features` | `{"savings_goals": true, "debt_tracker": true}` |
| `income[]` | `name`, `category`, `amount`, `frequency` (`weekly`, `biweekly`, `semimonthly`, `monthly`, `quarterly`, `yearly`, `one_time`), optional `month`, or a `months` list of 12 numbers |
| `expenses[]` | Same as income, plus `type` (`Bill` or `Expense` for personal, `Fixed` or `Variable` for business) and an optional `due_day` |
| `savings_goals[]` | `name`, `target`, `saved`, `target_date` |
| `debts[]` | `name`, `balance`, `apr`, `min_payment` |

## Development

```bash
pip install openpyxl
python -m unittest discover -s tests -v
```

The test suite covers pay schedules (including leap years and 3-paycheck months), frequency expansion, config validation, workbook structure, the summary math, and the CLI. Tests run on every push via GitHub Actions.

### Project layout

```
skills/budget-planner/      The skill itself (what gets installed)
  SKILL.md                  Questionnaire flow and instructions for Claude
  scripts/build_budget.py   Workbook generator
  scripts/profiles.py       Personal and Business labels, categories, starter rows
  scripts/example_*.json    Example configs
tests/                      Unit tests
samples/                    Example workbooks
docs/images/                Screenshots
.claude-plugin/             Claude Code plugin and marketplace manifests
```

Contributions are welcome. Please open an issue or a pull request, and include tests for new behavior.

## Notes and limitations

- Weekly and bi-weekly bills are spread evenly across months (52/12 and 26/12 per month). Paychecks use exact calendar dates.
- Paydays that fall on a weekend are not moved to the nearest business day.
- This tool helps you organize your own numbers. It is not financial, tax, or legal advice.

## License

[MIT](LICENSE). Free to use, modify, and share.

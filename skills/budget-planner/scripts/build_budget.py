"""Budget Planner workbook generator (Personal or Business).

Usage:
    python build_budget.py --config budget-config.json [--out "Budget Planner 2026.xlsx"]

Builds a formula-driven .xlsx: an annual overview Dashboard, Income/Revenue,
a pay or payroll calendar, Expenses, and optional goal and debt trackers.
Prints a JSON summary on success; on bad input prints "ERROR: ..." to stderr
and exits with code 2.
"""

import argparse
import calendar
import json
import re
import shutil
import sys
import zipfile
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Font as DrawingFont,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
    RichTextProperties,
)
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
from profiles import PROFILES  # noqa: E402

# ---------------------------------------------------------------- constants

FONT_NAME = "Segoe UI"
CREDIT_TEXT = "Budget Planner - Claude Skill by 6inq"
CREDIT_URL = "https://github.com/6inq/budget-planner"
SETTINGS_CELL = "E2"  # on the hidden Lists sheet; lets read_workbook.py rebuild with the same choices
SETTINGS_FORMAT = 1

MODES = {
    "light": {
        "bg": "F5F6FA", "card": "FFFFFF", "text": "1F2937", "muted": "6B7280",
        "border": "E3E6EE", "input": "FFFFFF", "zebra": "F9FAFC", "calc": "EEF0F6",
    },
    "dark": {
        "bg": "141830", "card": "1E2342", "text": "F1F3F9", "muted": "9AA3C2",
        "border": "2E3560", "input": "252B50", "zebra": "222848", "calc": "1A1F3A",
    },
}
ACCENTS = {"violet": "7C5CFC", "teal": "14B8A6", "rose": "F43F5E", "gold": "F5A524"}
POSITIVE = "22C55E"
NEGATIVE = "EF4444"
BASE_PALETTE = [
    "22D3EE", "F472B6", "FBBF24", "34D399", "60A5FA", "F87171",
    "A78BFA", "FB923C", "4ADE80", "E879F9", "94A3B8", "2DD4BF",
]

CHARTS = ("pie", "doughnut", "bar", "column")
PAY_FREQUENCIES = ("weekly", "biweekly", "semimonthly", "monthly", "none")
LINE_FREQUENCIES = ("weekly", "biweekly", "semimonthly", "monthly", "quarterly", "yearly", "one_time")
SEMIMONTHLY_TOKENS = ("1", "15", "last")

# Sheet geometry (1-based Excel rows/columns).
HEADER_ROW = 5
FIRST_ROW = 6
INCOME_INPUT_ROWS = 15
EXPENSE_INPUT_ROWS = 60
PAYCHECK_EXTRA_ROWS = 6
TRACKER_ROWS = 10
PLAN_HORIZONS = {"car": 72, "house": 360}  # schedule months per loan type
PLAN_MONTHS = PLAN_HORIZONS["car"]
PLAN_INPUT_ROW = 7  # Payoff Plan: first input row
PLAN_GOAL_OFFSET = 4  # Payoff Plan: row of the extra-payment goal, counted from PLAN_INPUT_ROW
PLAN_HEADER_ROW = 21
PLAN_HELPER_COL = 14  # N: hidden full-length columns for the no-extra comparison
INCOME_MONTH_COL = 4    # D
EXPENSE_MONTH_COL = 6   # F
EXPENSE_TOTAL_COL = EXPENSE_MONTH_COL + 12  # R
PAID_COL = EXPENSE_TOTAL_COL + 2  # T: first "Paid" month column on Expenses
PAID_MARK = "Paid"
SHEET_FILL_COLS = PAID_COL + 13
SHEET_FILL_ROWS = 200
MONTH_TABLE_HEADER = 29
CATEGORY_COL = 11  # K


# ---------------------------------------------------------------- dates

def add_months(day: date, months: int) -> date:
    index = day.month - 1 + months
    return date(day.year + index // 12, index % 12 + 1, 1)


def window_months(config: dict) -> list[date]:
    start = date(config["year"], config["start_month"], 1)
    return [add_months(start, i) for i in range(12)]


def window_bounds(config: dict) -> tuple[date, date]:
    start = window_months(config)[0]
    return start, add_months(start, 12) - timedelta(days=1)


def month_index(config: dict, day: date) -> int:
    return (day.year - config["year"]) * 12 + day.month - config["start_month"]


def resolve_day(year: int, month: int, token) -> date:
    last = calendar.monthrange(year, month)[1]
    if str(token) == "last":
        return date(year, month, last)
    return date(year, month, min(int(token), last))


def first_friday(start: date) -> date:
    return start + timedelta(days=(4 - start.weekday()) % 7)


def period_label(config: dict) -> str:
    start, end = window_bounds(config)
    return f"{start:%b %Y} - {end:%b %Y}"


# ---------------------------------------------------------------- validation

def fail(message: str):
    raise ValueError(message)


def check_choice(value, allowed, field: str):
    if value not in allowed:
        fail(f"'{field}' must be one of {', '.join(allowed)} (got {value!r})")


def check_amount(value, field: str) -> float:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        fail(f"'{field}' must be a number (got {value!r})")
    if value < 0:
        fail(f"'{field}' cannot be negative (got {value})")
    return float(value)


def check_int_range(value, low: int, high: int, field: str) -> int:
    if isinstance(value, bool) or not str(value).strip().isdigit() or not low <= int(value) <= high:
        fail(f"'{field}' must be a whole number from {low} to {high} (got {value!r})")
    return int(value)


def parse_date(value, field: str) -> date:
    try:
        return date.fromisoformat(str(value))
    except ValueError:
        fail(f"'{field}' must be a date like 2026-01-31 (got {value!r})")


def validate_pay(pay: dict) -> dict:
    frequency = pay.get("frequency", "monthly")
    check_choice(frequency, PAY_FREQUENCIES, "pay.frequency")
    clean = {"frequency": frequency, "amount": check_amount(pay.get("amount", 0), "pay.amount")}
    if frequency in ("weekly", "biweekly") and pay.get("first_payday"):
        clean["first_payday"] = parse_date(pay["first_payday"], "pay.first_payday")
    if frequency == "semimonthly":
        clean["semimonthly_days"] = validate_semimonthly(pay.get("semimonthly_days"))
    if frequency == "monthly":
        clean["monthly_day"] = validate_monthly_day(pay.get("monthly_day", "1"))
    return clean


def validate_paychecks(entries, config: dict):
    """Optional explicit paycheck list (written by read_workbook.py) that replaces the generated schedule."""
    if entries is None:
        return None
    if not isinstance(entries, list):
        fail("'pay.paychecks' must be a list of {date, amount, note}")
    start, end = window_bounds(config)
    clean = []
    for entry in entries:
        day = parse_date(entry.get("date"), "pay.paychecks.date")
        if not start <= day <= end:
            fail(f"paycheck date {day} is outside the budget period {period_label(config)}")
        clean.append({"date": day, "amount": check_amount(entry.get("amount") or 0, "pay.paychecks.amount"),
                      "note": str(entry.get("note") or "")})
    return sorted(clean, key=lambda p: p["date"])


def validate_semimonthly(days) -> list[str]:
    if not isinstance(days, list) or len(days) != 2:
        fail("'pay.semimonthly_days' must list exactly 2 paydays from: 1, 15, last")
    tokens = [str(d).strip().lower() for d in days]
    for token in tokens:
        check_choice(token, SEMIMONTHLY_TOKENS, "pay.semimonthly_days")
    if tokens[0] == tokens[1]:
        fail("'pay.semimonthly_days' must be two different days")
    return sorted(tokens, key=SEMIMONTHLY_TOKENS.index)


def validate_monthly_day(day) -> str:
    token = str(day).strip().lower()
    if token == "last":
        return token
    return str(check_int_range(token, 1, 28, "pay.monthly_day"))


def validate_line(item: dict, field: str, types=None) -> dict:
    if not item.get("name"):
        fail(f"every entry in '{field}' needs a 'name'")
    clean = {"name": str(item["name"]), "category": str(item.get("category") or "Other")}
    if types:
        clean["type"] = item.get("type", types[1])
        check_choice(clean["type"], types, f"{field}.type")
        due = item.get("due_day")
        clean["due_day"] = None if due is None else check_int_range(due, 1, 31, f"{field}.due_day")
        clean["paid"] = validate_paid(item.get("paid"), clean["name"], field)
    if "months" in item:
        months = item["months"]
        if not isinstance(months, list) or len(months) != 12:
            fail(f"'{field}.months' for '{clean['name']}' must have exactly 12 numbers")
        clean["months"] = [check_amount(v, f"{field}.months") for v in months]
        return clean
    clean["amount"] = check_amount(item.get("amount", 0), f"{field}.amount")
    clean["frequency"] = item.get("frequency", "monthly")
    check_choice(clean["frequency"], LINE_FREQUENCIES, f"{field}.frequency")
    month = item.get("month")
    clean["month"] = None if month is None else check_int_range(month, 1, 12, f"{field}.month")
    return clean


def validate_paid(paid, name: str, field: str) -> list[bool]:
    if paid is None:
        return [False] * 12
    if not isinstance(paid, list) or len(paid) != 12 or not all(isinstance(v, bool) for v in paid):
        fail(f"'{field}.paid' for '{name}' must be a list of 12 true/false values")
    return list(paid)


def validate_goal(item: dict) -> dict:
    target_date = item.get("target_date")
    return {
        "name": str(item.get("name") or "Goal"),
        "target": check_amount(item.get("target", 0), "savings_goals.target"),
        "saved": check_amount(item.get("saved", 0), "savings_goals.saved"),
        "target_date": parse_date(target_date, "savings_goals.target_date") if target_date else None,
    }


def validate_debt(item: dict) -> dict:
    return {
        "name": str(item.get("name") or "Debt"),
        "balance": check_amount(item.get("balance", 0), "debts.balance"),
        "apr": check_amount(item.get("apr", 0), "debts.apr"),
        "min_payment": check_amount(item.get("min_payment", 0), "debts.min_payment"),
    }


# key, label, number format ("money" = currency); row = PLAN_INPUT_ROW + position
PLAN_INPUTS = [
    ("loan_name", "Loan name", None),
    ("balance", "Loan balance today", "money"),
    ("apr", "Interest rate (APR %)", '0.00"%"'),
    ("payment", "Regular monthly payment", "money"),
    ("asset_value", "Vehicle value today", "money"),
    ("value_drop", "Value drop per month (estimate)", "money"),
    ("take_home", "Take-home per paycheck", "money"),
    ("paychecks_per_month", "Paychecks per month", "0.00"),
    ("monthly_bills", "Monthly bills (not this loan or savings)", "money"),
    ("savings_percent", "Savings (% of take-home)", '0"%"'),
    ("buffer", "Monthly buffer for surprises", "money"),
    ("start", "Plan starts", "mmm d, yyyy"),
    ("target_gap", "Target: underwater by at most", "money"),
]
PLAN_AMOUNT_KEYS = [k for k, _, _ in PLAN_INPUTS if k not in ("loan_name", "start")]
# Wording that differs for a home loan; a home's value can rise, so its value change may be negative.
PLAN_TYPE_TEXT = {
    "car": {"title": "Car Loan Payoff Plan", "asset": "Vehicle Value"},
    "house": {"title": "Home Loan Payoff Plan", "asset": "Home Value",
              "payment": "Regular payment (principal + interest)", "asset_value": "Home value today",
              "value_drop": "Value change per month (negative if rising)"},
}


def plan_label(key: str, label: str, loan_type: str) -> str:
    return PLAN_TYPE_TEXT[loan_type].get(key, label)


def plan_row(key: str) -> int:
    return PLAN_INPUT_ROW + [k for k, _, _ in PLAN_INPUTS].index(key)


def validate_payoff_plan(plan: dict) -> dict:
    """Optional loan payoff plan. Every field may be left out; the sheet then shows blank inputs."""
    loan_type = plan.get("loan_type", "car")
    check_choice(loan_type, tuple(PLAN_HORIZONS), "payoff_plan.loan_type")
    clean = {"loan_type": loan_type, "loan_name": str(plan.get("loan_name") or "")}
    for key in PLAN_AMOUNT_KEYS:
        value = plan.get(key)
        if value is not None and key == "value_drop" and isinstance(value, (int, float)) and value < 0:
            clean[key] = float(value)  # a rising value
            continue
        clean[key] = None if value is None else check_amount(value, f"payoff_plan.{key}")
    clean["start"] = parse_date(plan["start"], "payoff_plan.start") if plan.get("start") else None
    extra = plan.get("extra_paid") or []
    months = PLAN_HORIZONS[loan_type]
    if not isinstance(extra, list) or len(extra) > months:
        fail(f"'payoff_plan.extra_paid' must be a list of at most {months} amounts")
    clean["extra_paid"] = [None if v is None else check_amount(v, "payoff_plan.extra_paid") for v in extra]
    return clean


def validate_config(raw: dict) -> dict:
    if "year" not in raw:
        fail("'year' is required")
    year = raw["year"]
    if not isinstance(year, int) or isinstance(year, bool) or not 2000 <= year <= 2100:
        fail(f"'year' must be a whole number from 2000 to 2100 (got {year!r})")
    budget_type = raw.get("budget_type", "personal")
    check_choice(budget_type, tuple(PROFILES), "budget_type")
    profile = PROFILES[budget_type]
    symbol = str(raw.get("currency_symbol", "$"))
    if not 1 <= len(symbol) <= 3 or '"' in symbol:
        fail("'currency_symbol' must be 1-3 characters without quotes")
    config = {
        "budget_type": budget_type,
        "profile": profile,
        "year": year,
        "start_month": check_int_range(raw.get("start_month", 1), 1, 12, "start_month"),
        "title": raw.get("title") or profile["default_title"].format(year=year),
        "owner": raw.get("owner") or "",
        "currency_symbol": symbol,
        "mode": raw.get("mode", "light"),
        "accent": raw.get("accent", "violet"),
        "category_chart": raw.get("category_chart", "doughnut"),
        "pay": validate_pay(raw.get("pay") or {}),
    }
    config["pay"]["paychecks"] = validate_paychecks((raw.get("pay") or {}).get("paychecks"), config)
    check_choice(config["mode"], tuple(MODES), "mode")
    check_choice(config["accent"], tuple(ACCENTS), "accent")
    check_choice(config["category_chart"], CHARTS, "category_chart")
    features = raw.get("features") or {}
    config["features"] = {"savings_goals": bool(features.get("savings_goals")),
                          "debt_tracker": bool(features.get("debt_tracker")),
                          "payoff_plan": bool(features.get("payoff_plan"))}
    config["income"] = [validate_line(i, "income") for i in raw.get("income") or []]
    config["expenses"] = [validate_line(e, "expenses", profile["types"]) for e in raw.get("expenses") or []]
    config["savings_goals"] = [validate_goal(g) for g in raw.get("savings_goals") or []]
    config["debts"] = [validate_debt(d) for d in raw.get("debts") or []]
    config["payoff_plan"] = validate_payoff_plan(raw.get("payoff_plan") or {})
    return config


# ---------------------------------------------------------------- schedule math

def generate_paydays(config: dict) -> list[date]:
    pay = config["pay"]
    start, end = window_bounds(config)
    if pay["frequency"] == "none":
        return []
    if pay["frequency"] in ("weekly", "biweekly"):
        step = timedelta(days=7 if pay["frequency"] == "weekly" else 14)
        return stepped_dates(pay.get("first_payday") or first_friday(start), step, start, end)
    tokens = pay["semimonthly_days"] if pay["frequency"] == "semimonthly" else [pay["monthly_day"]]
    return [resolve_day(m.year, m.month, t) for m in window_months(config) for t in tokens]


def paycheck_entries(config: dict) -> list[dict]:
    """Paydays with their amounts: the explicit list if given, else the generated schedule."""
    if config["pay"].get("paychecks") is not None:
        return config["pay"]["paychecks"]
    return [{"date": d, "amount": config["pay"]["amount"], "note": ""} for d in generate_paydays(config)]


def stepped_dates(anchor: date, step: timedelta, start: date, end: date) -> list[date]:
    # Align the cadence to the earliest date inside the window, whichever side the anchor is on.
    while anchor < start:
        anchor += step
    while anchor - step >= start:
        anchor -= step
    dates = []
    while anchor <= end:
        dates.append(anchor)
        anchor += step
    return dates


def expand_line(item: dict, config: dict) -> list[float]:
    if "months" in item:
        return list(item["months"])
    amount, frequency = item["amount"], item["frequency"]
    per_month = {
        "monthly": amount,
        "weekly": amount * 52 / 12,
        "biweekly": amount * 26 / 12,
        "semimonthly": amount * 2,
    }
    if frequency in per_month:
        return [per_month[frequency]] * 12
    first = first_window_index(config, item.get("month"))
    step = 3 if frequency == "quarterly" else 12
    return [amount if i >= first and (i - first) % step == 0 else 0.0 for i in range(12)]


def first_window_index(config: dict, calendar_month) -> int:
    if calendar_month is None:
        return 0
    return next(i for i, m in enumerate(window_months(config)) if m.month == calendar_month)


def category_list(config: dict) -> list[str]:
    return merge_unique(config["profile"]["categories"], [e["category"] for e in config["expenses"]])


def income_category_list(config: dict) -> list[str]:
    return merge_unique(config["profile"]["income_categories"], [i["category"] for i in config["income"]])


def merge_unique(base: list[str], extra: list[str]) -> list[str]:
    seen = {c.lower() for c in base}
    merged = list(base)
    for name in extra:
        if name.lower() not in seen:
            seen.add(name.lower())
            merged.append(name)
    return merged


def compute_summary(config: dict) -> dict:
    paydays = paycheck_entries(config)
    pay_by_month = [0.0] * 12
    for entry in paydays:
        pay_by_month[month_index(config, entry["date"])] += entry["amount"]
    fixed_type = config["profile"]["types"][0]
    money_in = list(pay_by_month) if config["profile"]["pay_is_income"] else [0.0] * 12
    fixed = [0.0] * 12 if config["profile"]["pay_is_income"] else list(pay_by_month)
    variable = [0.0] * 12
    for item in config["income"]:
        money_in = add_lists(money_in, rounded(expand_line(item, config)))
    for item in config["expenses"]:
        values = rounded(expand_line(item, config))
        if item["type"] == fixed_type:
            fixed = add_lists(fixed, values)
        else:
            variable = add_lists(variable, values)
    return summary_dict(config, paydays, money_in, fixed, variable)


def add_lists(a: list[float], b: list[float]) -> list[float]:
    return [x + y for x, y in zip(a, b)]


def rounded(values: list[float]) -> list[float]:
    return [round(v, 2) for v in values]


def summary_dict(config, paydays, money_in, fixed, variable) -> dict:
    total_in = round(sum(money_in), 2)
    total_out = round(sum(fixed) + sum(variable), 2)
    net = round(total_in - total_out, 2)
    monthly = [
        {"month": f"{m:%b %Y}", "in": round(money_in[i], 2), "out": round(fixed[i] + variable[i], 2),
         "net": round(money_in[i] - fixed[i] - variable[i], 2)}
        for i, m in enumerate(window_months(config))
    ]
    return {
        "budget_type": config["budget_type"],
        "period": period_label(config),
        "paychecks": len(paydays),
        "annual_income": total_in,
        "annual_fixed": round(sum(fixed), 2),
        "annual_variable": round(sum(variable), 2),
        "annual_out": total_out,
        "net": net,
        "savings_rate": round(net / total_in, 4) if total_in else None,
        "monthly": monthly,
    }


def frequency_text(pay: dict) -> str:
    labels = {"1": "1st", "15": "15th", "last": "Last day"}
    simple = {"weekly": "Weekly", "biweekly": "Bi-weekly", "none": "Irregular"}
    if pay["frequency"] in simple:
        return simple[pay["frequency"]]
    if pay["frequency"] == "semimonthly":
        first, second = (labels[d] for d in pay["semimonthly_days"])
        return f"Semi-monthly ({first} & {second})"
    return f"Monthly ({labels.get(pay['monthly_day']) or ordinal(int(pay['monthly_day']))})"


def ordinal(day: int) -> str:
    suffix = "th" if 10 <= day % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


def sheet_ref(name: str) -> str:
    return f"'{name}'"


# ---------------------------------------------------------------- styling

class Theme:
    def __init__(self, config: dict):
        self.colors = dict(MODES[config["mode"]])
        self.accent = ACCENTS[config["accent"]]
        self.palette = [self.accent] + BASE_PALETTE
        sym = config["currency_symbol"]
        self.money = f'"{sym}"#,##0.00;-"{sym}"#,##0.00;"{sym}"0.00'
        self.money_whole = f'"{sym}"#,##0;-"{sym}"#,##0;"{sym}"0'
        self.money_label = f'"{sym}"#,##0;;;'  # hides zero-value data labels
        self.accent_tint = blend(self.accent, self.colors["card"], 0.22)

    def __getitem__(self, key: str) -> str:
        return self.colors[key]

    def font(self, size=10, bold=False, color=None, italic=False) -> Font:
        return Font(name=FONT_NAME, size=size, bold=bold, italic=italic, color=color or self["text"])

    @staticmethod
    def fill(color: str) -> PatternFill:
        return PatternFill("solid", start_color=color, end_color=color)

    def box(self) -> Border:
        side = Side(style="thin", color=self["border"])
        return Border(left=side, right=side, top=side, bottom=side)


def blend(color: str, base: str, weight: float) -> str:
    mix = [round(int(color[i:i + 2], 16) * weight + int(base[i:i + 2], 16) * (1 - weight)) for i in (0, 2, 4)]
    return "".join(f"{v:02X}" for v in mix)


def paint_sheet(ws, theme: Theme):
    fill, font = theme.fill(theme["bg"]), theme.font()
    for row in ws.iter_rows(min_row=1, max_row=SHEET_FILL_ROWS, max_col=SHEET_FILL_COLS):
        for cell in row:
            cell.fill = fill
            cell.font = font
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = theme.accent
    ws.column_dimensions["A"].width = 2
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def style_cell(cell, theme: Theme, fill=None, font=None, fmt=None, align=None):
    cell.fill = theme.fill(fill or theme["card"])
    cell.font = font or theme.font()
    cell.border = theme.box()
    if fmt:
        cell.number_format = fmt
    if align:
        cell.alignment = Alignment(horizontal=align, vertical="center")


def write_banner(ws, theme: Theme, title: str, subtitle: str, last_col: int):
    end = get_column_letter(last_col)
    ws.merge_cells(f"B2:{end}2")
    ws.merge_cells(f"B3:{end}3")
    ws["B2"].value = title
    ws["B2"].font = theme.font(22, bold=True)
    ws["B2"].alignment = Alignment(vertical="center")
    ws["B3"].value = subtitle
    ws["B3"].font = theme.font(10, color=theme["muted"], italic=True)
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 18
    write_credit(ws, theme, last_col)


def write_credit(ws, theme: Theme, last_col: int):
    """Small, muted attribution on every sheet (row 4 is an empty spacer row) and in the print footer."""
    ws.merge_cells(f"B4:{get_column_letter(last_col)}4")
    cell = ws["B4"]
    cell.value = CREDIT_TEXT
    cell.hyperlink = CREDIT_URL
    cell.font = theme.font(8, color=theme["muted"], italic=True)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[4].height = 14
    ws.oddFooter.right.text = CREDIT_TEXT
    ws.oddFooter.right.size = 8
    ws.oddFooter.right.color = "9AA3C2"


def write_header(ws, theme: Theme, row: int, start_col: int, labels: list):
    for offset, label in enumerate(labels):
        cell = ws.cell(row=row, column=start_col + offset, value=label)
        style_cell(cell, theme, fill=theme.accent, font=theme.font(10, True, "FFFFFF"), align="center")
        if isinstance(label, date):
            cell.number_format = "mmm yyyy"
    ws.row_dimensions[row].height = 24


def style_total_row(ws, theme: Theme, row: int, first_col: int, last_col: int, height=22):
    top = Side(style="medium", color=theme.accent)
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = theme.fill(theme["card"])
        cell.font = theme.font(10, bold=True)
        cell.border = Border(top=top, bottom=Side(style="thin", color=theme["border"]))
    if height:
        ws.row_dimensions[row].height = height


def set_print_area(ws, last_col: int, total_row: int):
    # The painted background spans far past the table; print only the table.
    ws.print_area = f"A1:{get_column_letter(last_col + 1)}{total_row + 1}"


def input_fill(theme: Theme, row: int) -> str:
    return theme["zebra"] if (row - FIRST_ROW) % 2 else theme["input"]


def write_section_title(ws, theme: Theme, cell_ref: str, text: str):
    ws[cell_ref].value = text
    ws[cell_ref].font = theme.font(12, bold=True)


def write_card(ws, theme: Theme, col: str, top: int, card: tuple, value_size=18):
    """A 2-column KPI tile spanning 3 rows: label, big value, caption."""
    label, formula, fmt, caption = card
    start = ws[f"{col}1"].column
    end = get_column_letter(start + 1)
    gap = Side(style="thick", color=theme["bg"])
    accent_top = Side(style="thick", color=theme.accent)
    contents = [(label, theme.font(9, True, theme["muted"])), (formula, theme.font(value_size, True)),
                (caption, theme.font(8, color=theme["muted"]))]
    for offset, (value, font) in enumerate(contents):
        row = top + offset
        ws.merge_cells(f"{col}{row}:{end}{row}")
        for c in range(start, start + 2):
            ws.cell(row=row, column=c).fill = theme.fill(theme["card"])
            ws.cell(row=row, column=c).border = Border(
                left=gap, right=gap, top=accent_top if offset == 0 else None, bottom=gap if offset == 2 else None)
        cell = ws[f"{col}{row}"]
        cell.value, cell.font = value, font
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        if offset == 1:
            cell.number_format = fmt
    for offset, height in enumerate((22, value_size * 1.8, 18)):
        ws.row_dimensions[top + offset].height = height


# ---------------------------------------------------------------- chart styling

def char_props(color: str, size=900, bold=False) -> CharacterProperties:
    return CharacterProperties(sz=size, b=bold, solidFill=color, latin=DrawingFont(typeface=FONT_NAME))


def rich_text(color: str, size=900) -> RichText:
    props = char_props(color, size)
    return RichText(bodyPr=RichTextProperties(), p=[Paragraph(pPr=ParagraphProperties(defRPr=props), endParaRPr=props)])


def chart_title(text: str, color: str) -> Title:
    props = char_props(color, 1200, True)
    para = Paragraph(pPr=ParagraphProperties(defRPr=props), r=[RegularTextRun(rPr=props, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def style_chart(chart, theme: Theme, title: str):
    chart.title = chart_title(title, theme["text"])
    chart.graphical_properties = GraphicalProperties(solidFill=theme["card"], ln=LineProperties(solidFill=theme["border"]))
    chart.plot_area.graphicalProperties = GraphicalProperties(noFill=True)
    chart.roundedCorners = True
    if chart.legend is not None:
        chart.legend.position = "b"
        chart.legend.overlay = False
        chart.legend.txPr = rich_text(theme["text"])


def style_axes(chart, theme: Theme, value_format: str):
    for axis in (chart.x_axis, chart.y_axis):
        axis.delete = False
        axis.txPr = rich_text(theme["muted"])
        axis.spPr = GraphicalProperties(ln=LineProperties(solidFill=theme["border"]))
    chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=theme["border"]))
    chart.y_axis.number_format = value_format


def color_series(series, color: str, line_width=None):
    series.graphicalProperties.solidFill = color
    series.graphicalProperties.line.solidFill = color
    if line_width:
        series.graphicalProperties.line.width = line_width


def color_points(series, theme: Theme, count: int):
    for idx in range(count):
        point = DataPoint(idx=idx)
        point.graphicalProperties.solidFill = theme.palette[idx % len(theme.palette)]
        point.graphicalProperties.line.solidFill = theme["card"]
        series.dPt.append(point)


def data_labels(theme: Theme, percent: bool, fmt: str) -> DataLabelList:
    labels = DataLabelList()
    labels.showPercent = percent
    labels.showVal = not percent
    labels.showCatName = labels.showSerName = labels.showLegendKey = False
    labels.numFmt = fmt
    labels.txPr = rich_text(theme["text"], 900)
    return labels


# ---------------------------------------------------------------- sheets: inputs

def build_lists(wb, config: dict) -> dict:
    name = config["profile"]["sheets"]["lists"]
    ws = wb.create_sheet(name)
    columns = {
        "A": ("Expense Categories", category_list(config)),
        "B": ("Types", list(config["profile"]["types"])),
        "C": ("Income Categories", income_category_list(config)),
    }
    ranges = {}
    for letter, (title, values) in columns.items():
        ws[f"{letter}1"] = title
        for i, value in enumerate(values, start=2):
            ws[f"{letter}{i}"] = value
        ranges[letter] = f"={name}!${letter}$2:${letter}${len(values) + 1}"
    ws["E1"] = "Planner Settings"
    ws[SETTINGS_CELL] = json.dumps(export_settings(config))
    ws.sheet_state = "hidden"
    return {"categories": ranges["A"], "types": ranges["B"], "income": ranges["C"]}


def export_settings(config: dict) -> dict:
    """The questionnaire answers (no financial data), saved inside the workbook for later updates."""
    pay = {k: v for k, v in config["pay"].items() if k != "paychecks"}
    if "first_payday" in pay:
        pay["first_payday"] = pay["first_payday"].isoformat()
    keys = ("budget_type", "title", "owner", "year", "start_month", "currency_symbol",
            "mode", "accent", "category_chart", "features")
    settings = {"format": SETTINGS_FORMAT, **{k: config[k] for k in keys}, "pay": pay}
    if config["features"]["payoff_plan"]:
        settings["payoff_loan_type"] = config["payoff_plan"]["loan_type"]
    return settings


def add_dropdown(ws, formula: str, cell_range: str):
    validation = DataValidation(type="list", formula1=formula, allow_blank=True)
    validation.add(cell_range)
    ws.add_data_validation(validation)


def write_month_values(ws, theme: Theme, row: int, first_col: int, values):
    for i in range(12):
        value = values[i] if values else None
        cell = ws.cell(row=row, column=first_col + i, value=round(value, 2) if value else None)
        style_cell(cell, theme, fill=input_fill(theme, row), fmt=theme.money)


def write_row_total(ws, theme: Theme, row: int, first_col: int):
    span = f"{get_column_letter(first_col)}{row}:{get_column_letter(first_col + 11)}{row}"
    # Blank until the row has numbers, so empty input rows stay clean.
    cell = ws.cell(row=row, column=first_col + 12, value=f'=IF(COUNT({span})=0,"",SUM({span}))')
    style_cell(cell, theme, fill=theme["calc"], font=theme.font(10, True), fmt=theme.money)


def write_column_totals(ws, theme: Theme, row: int, first_col: int, last_col: int, first_row: int):
    for col in range(first_col, last_col + 1):
        letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col, value=f"=SUM({letter}{first_row}:{letter}{row - 1})")
        cell.number_format = theme.money


def write_formula_month_row(ws, theme: Theme, row: int, first_col: int, label: str, category: str, formula_for):
    """A locked-looking row whose 12 month cells come from formula_for(header_cell_ref)."""
    style_cell(ws.cell(row=row, column=2, value=label), theme, fill=theme["calc"], font=theme.font(10, True))
    style_cell(ws.cell(row=row, column=3, value=category), theme, fill=theme["calc"])
    for i in range(12):
        col = first_col + i
        formula = formula_for(f"{get_column_letter(col)}${HEADER_ROW}")
        style_cell(ws.cell(row=row, column=col, value=formula), theme, fill=theme["calc"], fmt=theme.money)
    write_row_total(ws, theme, row, first_col)


def write_line_row(ws, theme: Theme, row: int, item, config: dict, month_col: int):
    fill = input_fill(theme, row)
    values = [item["name"], item["category"]] if item else [None, None]
    for col, value in zip((2, 3), values):
        style_cell(ws.cell(row=row, column=col, value=value), theme, fill=fill)
    amounts = expand_line(item, config) if item and ("months" in item or "amount" in item) else None
    write_month_values(ws, theme, row, month_col, amounts)
    write_row_total(ws, theme, row, month_col)


def build_paychecks(wb, theme: Theme, config: dict) -> dict:
    profile = config["profile"]
    name = profile["sheets"]["pay"]
    ws = wb.create_sheet(name)
    paint_sheet(ws, theme)
    write_banner(ws, theme, profile["pay_title"], profile["pay_help"], 6)
    date_label = "Pay Date" if profile["pay_is_income"] else "Payroll Date"
    write_header(ws, theme, HEADER_ROW, 2, ["#", date_label, "Month", profile["pay_amount_header"], "Notes"])
    entries = paycheck_entries(config)
    rows = len(entries) + PAYCHECK_EXTRA_ROWS
    for offset in range(rows):
        entry = entries[offset] if offset < len(entries) else None
        write_paycheck_row(ws, theme, FIRST_ROW + offset, offset + 1, entry)
    last = FIRST_ROW + rows - 1
    total = last + 1
    ws.cell(row=total, column=2, value="TOTAL")
    ws.cell(row=total, column=5, value=f"=SUM(E{FIRST_ROW}:E{last})").number_format = theme.money
    style_total_row(ws, theme, total, 2, 6)
    set_print_area(ws, 6, total)
    for letter, width in zip("BCDEF", (6, 22, 14, 16, 36)):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = f"A{FIRST_ROW}"
    ref = sheet_ref(name)
    return {"month": f"{ref}!$D${FIRST_ROW}:$D${last}", "pay": f"{ref}!$E${FIRST_ROW}:$E${last}"}


def write_paycheck_row(ws, theme: Theme, row: int, number: int, entry):
    fill = input_fill(theme, row)
    day = entry["date"] if entry else None
    amount = entry["amount"] if entry else None
    style_cell(ws.cell(row=row, column=2, value=number if day else None), theme, fill=theme["calc"],
               font=theme.font(10, color=theme["muted"]), align="center")
    style_cell(ws.cell(row=row, column=3, value=day), theme, fill=fill, fmt="ddd, mmm d, yyyy")
    month = f'=IF(C{row}="","",DATE(YEAR(C{row}),MONTH(C{row}),1))'
    style_cell(ws.cell(row=row, column=4, value=month), theme, fill=theme["calc"], fmt="mmm yyyy")
    style_cell(ws.cell(row=row, column=5, value=amount or None), theme, fill=fill, fmt=theme.money)
    style_cell(ws.cell(row=row, column=6, value=(entry or {}).get("note") or None), theme, fill=fill)


def pay_formula(paychecks: dict):
    return lambda header: f"=SUMIFS({paychecks['pay']},{paychecks['month']},{header})"


def pay_row_label(config: dict) -> str:
    return f"{config['profile']['pay_row']} ({frequency_text(config['pay'])})"


def build_income(wb, theme: Theme, config: dict, lists: dict, paychecks: dict) -> dict:
    profile = config["profile"]
    ws = wb.create_sheet(profile["sheets"]["income"])
    paint_sheet(ws, theme)
    last_col = INCOME_MONTH_COL + 12
    write_banner(ws, theme, profile["income_title"], profile["income_help"], last_col)
    write_header(ws, theme, HEADER_ROW, 2, ["Source", "Category"] + window_months(config) + ["Annual Total"])
    row = FIRST_ROW
    if profile["pay_is_income"]:
        write_formula_month_row(ws, theme, row, INCOME_MONTH_COL, pay_row_label(config),
                                profile["pay_category"], pay_formula(paychecks))
        row += 1
    items = config["income"] or [{"name": n, "category": c} for n, c in profile["starter_income"]]
    count = max(INCOME_INPUT_ROWS, len(items))
    for offset in range(count):
        write_line_row(ws, theme, row + offset, items[offset] if offset < len(items) else None, config, INCOME_MONTH_COL)
    last = row + count - 1
    add_dropdown(ws, lists["income"], f"C{row}:C{last}")
    total = last + 1
    ws.cell(row=total, column=2, value=profile["income_total"])
    write_column_totals(ws, theme, total, INCOME_MONTH_COL, last_col, FIRST_ROW)
    style_total_row(ws, theme, total, 2, last_col)
    set_print_area(ws, last_col, total)
    set_widths(ws, {"B": 32, "C": 18}, INCOME_MONTH_COL, last_col)
    ws.freeze_panes = f"C{FIRST_ROW}"
    return {"total_row": total}


def build_expenses(wb, theme: Theme, config: dict, lists: dict, paychecks: dict) -> dict:
    profile = config["profile"]
    ws = wb.create_sheet(profile["sheets"]["expenses"])
    paint_sheet(ws, theme)
    last_col = EXPENSE_TOTAL_COL
    write_banner(ws, theme, profile["expenses_title"], profile["expenses_help"], last_col)
    write_header(ws, theme, HEADER_ROW, 2,
                 ["Name", "Category", "Type", "Due Day"] + window_months(config) + ["Annual Total"])
    row = FIRST_ROW
    if not profile["pay_is_income"]:
        write_payroll_expense_row(ws, theme, config, paychecks)
        row += 1
    items = config["expenses"] or starter_expense_items(profile)
    count = max(EXPENSE_INPUT_ROWS, len(items) + 10)
    for offset in range(count):
        write_expense_row(ws, theme, row + offset, items[offset] if offset < len(items) else None, config)
    last = row + count - 1
    add_dropdown(ws, lists["categories"], f"C{row}:C{last}")
    add_dropdown(ws, lists["types"], f"D{row}:D{last}")
    ws.conditional_formatting.add(f"D{FIRST_ROW}:D{last}", FormulaRule(
        formula=[f'$D{FIRST_ROW}="{profile["types"][0]}"'], font=Font(name=FONT_NAME, bold=True, color=theme.accent)))
    total = last + 1
    ws.cell(row=total, column=2, value="TOTAL")
    write_column_totals(ws, theme, total, EXPENSE_MONTH_COL, last_col, FIRST_ROW)
    style_total_row(ws, theme, total, 2, last_col)
    write_paid_block(ws, theme, config, items, row, last)
    write_paid_totals(ws, theme, total, last)
    set_print_area(ws, PAID_COL + 11, total + 2)
    set_widths(ws, {"B": 30, "C": 22, "D": 11, "E": 9}, EXPENSE_MONTH_COL, last_col)
    ws.column_dimensions[get_column_letter(last_col + 1)].width = 3
    for i in range(12):
        ws.column_dimensions[get_column_letter(PAID_COL + i)].width = 11
    ws.freeze_panes = f"C{FIRST_ROW}"
    return {"first": FIRST_ROW, "last": last}


def write_paid_block(ws, theme: Theme, config: dict, items: list, first: int, last: int):
    """One "Paid" dropdown per bill per month, right of the amounts. Marks never change any total."""
    write_header(ws, theme, HEADER_ROW, PAID_COL, window_months(config))
    for i in range(12):
        ws.cell(row=HEADER_ROW, column=PAID_COL + i).number_format = f'"{PAID_MARK}" mmm'
    for offset, row in enumerate(range(first, last + 1)):
        item = items[offset] if offset < len(items) else None
        paid = (item or {}).get("paid") or [False] * 12
        for i in range(12):
            cell = ws.cell(row=row, column=PAID_COL + i, value=PAID_MARK if paid[i] else None)
            style_cell(cell, theme, fill=input_fill(theme, row), align="center")
    add_dropdown(ws, f'"{PAID_MARK}"', paid_range(first, last))
    green = Font(name=FONT_NAME, bold=True, color=POSITIVE)
    tint = theme.fill(blend(POSITIVE, theme["card"], 0.18))
    test = f'{get_column_letter(PAID_COL)}{first}="{PAID_MARK}"'
    amounts = f"{get_column_letter(EXPENSE_MONTH_COL)}{first}:{get_column_letter(EXPENSE_MONTH_COL + 11)}{last}"
    for cell_range in (amounts, paid_range(first, last)):
        ws.conditional_formatting.add(cell_range, FormulaRule(formula=[test], font=green, fill=tint))


def paid_range(first: int, last: int) -> str:
    return f"{get_column_letter(PAID_COL)}{first}:{get_column_letter(PAID_COL + 11)}{last}"


def write_paid_totals(ws, theme: Theme, total: int, last: int):
    """PAID and LEFT TO PAY rows under TOTAL, per month and for the year."""
    paid_row, left_row = total + 1, total + 2
    for row, label, color in ((paid_row, "PAID", POSITIVE), (left_row, "LEFT TO PAY", None)):
        ws.cell(row=row, column=2, value=label).font = theme.font(10, True, color)
    for i in range(12):
        amount = get_column_letter(EXPENSE_MONTH_COL + i)
        mark = get_column_letter(PAID_COL + i)
        paid = (f'=SUMIFS({amount}{FIRST_ROW}:{amount}{last},'
                f'{mark}{FIRST_ROW}:{mark}{last},"{PAID_MARK}")')
        cells = ((paid_row, paid, POSITIVE), (left_row, f"={amount}{total}-{amount}{paid_row}", None))
        for row, formula, color in cells:
            cell = ws.cell(row=row, column=EXPENSE_MONTH_COL + i, value=formula)
            cell.number_format, cell.font = theme.money, theme.font(10, True, color)
    first_amount, last_amount = get_column_letter(EXPENSE_MONTH_COL), get_column_letter(EXPENSE_MONTH_COL + 11)
    for row, color in ((paid_row, POSITIVE), (left_row, None)):
        cell = ws.cell(row=row, column=EXPENSE_TOTAL_COL, value=f"=SUM({first_amount}{row}:{last_amount}{row})")
        cell.number_format, cell.font = theme.money, theme.font(10, True, color)


def write_payroll_expense_row(ws, theme: Theme, config: dict, paychecks: dict):
    profile = config["profile"]
    write_formula_month_row(ws, theme, FIRST_ROW, EXPENSE_MONTH_COL, pay_row_label(config),
                            profile["pay_category"], pay_formula(paychecks))
    style_cell(ws.cell(row=FIRST_ROW, column=4, value=profile["types"][0]), theme, fill=theme["calc"], align="center")
    style_cell(ws.cell(row=FIRST_ROW, column=5), theme, fill=theme["calc"])


def starter_expense_items(profile: dict) -> list[dict]:
    return [{"name": n, "category": c, "type": t, "due_day": d} for n, c, t, d in profile["starter_expenses"]]


def write_expense_row(ws, theme: Theme, row: int, item, config: dict):
    write_line_row(ws, theme, row, item, config, EXPENSE_MONTH_COL)
    fill = input_fill(theme, row)
    style_cell(ws.cell(row=row, column=4, value=item["type"] if item else None), theme, fill=fill, align="center")
    style_cell(ws.cell(row=row, column=5, value=item.get("due_day") if item else None), theme, fill=fill, align="center")


def set_widths(ws, fixed: dict, first_month_col: int, last_col: int):
    for letter, width in fixed.items():
        ws.column_dimensions[letter].width = width
    for col in range(first_month_col, last_col):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions[get_column_letter(last_col)].width = 15


# ---------------------------------------------------------------- sheets: trackers

def build_savings(wb, theme: Theme, config: dict) -> dict:
    profile = config["profile"]
    ws = wb.create_sheet(profile["sheets"]["goals"])
    paint_sheet(ws, theme)
    write_banner(ws, theme, profile["goals_title"], profile["goals_help"], 9)
    write_header(ws, theme, HEADER_ROW, 2, ["Goal", "Target", "Saved So Far", "Remaining", "Progress",
                                            "Target Date", "Months Left", "Needed / Month"])
    goals = config["savings_goals"] or [{"name": n} for n in profile["starter_goals"]]
    last = FIRST_ROW + max(TRACKER_ROWS, len(goals)) - 1
    for row in range(FIRST_ROW, last + 1):
        index = row - FIRST_ROW
        write_goal_row(ws, theme, row, goals[index] if index < len(goals) else {})
    add_data_bar(ws, f"F{FIRST_ROW}:F{last}", theme.accent, end_type="num", end_value=1)
    total = last + 1
    ws.cell(row=total, column=2, value="TOTAL")
    write_column_totals(ws, theme, total, 3, 5, FIRST_ROW)
    ws.cell(row=total, column=6, value=f"=IFERROR(D{total}/C{total},0)").number_format = "0.0%"
    style_total_row(ws, theme, total, 2, 9)
    set_print_area(ws, 9, total)
    for letter, width in zip("BCDEFGHI", (28, 15, 15, 15, 13, 15, 13, 16)):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = f"C{FIRST_ROW}"
    return {"total_row": total}


def write_goal_row(ws, theme: Theme, row: int, goal: dict):
    fill = input_fill(theme, row)
    inputs = {2: (goal.get("name"), None), 3: (goal.get("target"), theme.money),
              4: (goal.get("saved"), theme.money), 7: (goal.get("target_date"), "mmm d, yyyy")}
    for col, (value, fmt) in inputs.items():
        style_cell(ws.cell(row=row, column=col, value=value or None), theme, fill=fill, fmt=fmt)
    formulas = {
        5: (f'=IF(C{row}="","",MAX(0,C{row}-D{row}))', theme.money),
        6: (f'=IF(C{row}="","",IFERROR(MIN(1,D{row}/C{row}),0))', "0.0%"),
        8: (f'=IF(G{row}="","",IFERROR(MAX(0,DATEDIF(TODAY(),G{row},"m")),0))', "0"),
        9: (f'=IF(OR(E{row}="",H{row}=""),"",E{row}/MAX(1,H{row}))', theme.money),
    }
    for col, (formula, fmt) in formulas.items():
        style_cell(ws.cell(row=row, column=col, value=formula), theme, fill=theme["calc"], fmt=fmt)


def build_debts(wb, theme: Theme, config: dict) -> dict:
    profile = config["profile"]
    ws = wb.create_sheet(profile["sheets"]["debts"])
    paint_sheet(ws, theme)
    write_banner(ws, theme, profile["debts_title"], profile["debts_help"], 8)
    write_header(ws, theme, HEADER_ROW, 2, [profile["debt_header"], "Balance", "APR %", "Min Payment",
                                            "Months to Payoff", "Est. Interest", "Payoff Date"])
    debts = config["debts"] or [{"name": n} for n in profile["starter_debts"]]
    last = FIRST_ROW + max(TRACKER_ROWS, len(debts)) - 1
    for row in range(FIRST_ROW, last + 1):
        index = row - FIRST_ROW
        write_debt_row(ws, theme, row, debts[index] if index < len(debts) else {})
    total = last + 1
    ws.cell(row=total, column=2, value="TOTAL")
    for letter in ("C", "E"):
        ws[f"{letter}{total}"] = f"=SUM({letter}{FIRST_ROW}:{letter}{last})"
        ws[f"{letter}{total}"].number_format = theme.money
    style_total_row(ws, theme, total, 2, 8)
    set_print_area(ws, 8, total)
    for letter, width in zip("BCDEFGH", (28, 15, 10, 15, 16, 15, 15)):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = f"C{FIRST_ROW}"
    return {"total_row": total}


def write_debt_row(ws, theme: Theme, row: int, debt: dict):
    fill = input_fill(theme, row)
    inputs = {2: (debt.get("name"), None), 3: (debt.get("balance"), theme.money),
              4: (debt.get("apr"), '0.00"%"'), 5: (debt.get("min_payment"), theme.money)}
    for col, (value, fmt) in inputs.items():
        style_cell(ws.cell(row=row, column=col, value=value or None), theme, fill=fill, fmt=fmt)
    months = (f'=IF(OR(C{row}="",E{row}=""),"",IF(N(D{row})=0,ROUNDUP(C{row}/E{row},0),'
              f'IFERROR(ROUNDUP(NPER(D{row}/1200,-E{row},C{row}),0),"Never")))')
    formulas = {
        6: (months, "0"),
        7: (f'=IF(ISNUMBER(F{row}),MAX(0,F{row}*E{row}-C{row}),"")', theme.money),
        8: (f'=IF(ISNUMBER(F{row}),EDATE(TODAY(),F{row}),"")', "mmm yyyy"),
    }
    for col, (formula, fmt) in formulas.items():
        style_cell(ws.cell(row=row, column=col, value=formula), theme, fill=theme["calc"], fmt=fmt, align="right")


# ---------------------------------------------------------------- sheets: payoff plan

def build_payoff_plan(wb, theme: Theme, config: dict):
    """Extra-payment goal from take-home minus bills, savings and a buffer, then a month-by-month
    schedule showing when the loan drops below the car's or home's value and when it is paid off."""
    ws = wb.create_sheet(config["profile"]["sheets"]["payoff"])
    paint_sheet(ws, theme)
    loan_type = config["payoff_plan"]["loan_type"]
    text = PLAN_TYPE_TEXT[loan_type]
    months = PLAN_HORIZONS[loan_type]
    write_banner(ws, theme, text["title"],
                 "Change any number under Your Numbers; the goal and every month below update. "
                 "Log what you really paid in Extra Paid.", 12)
    write_plan_inputs(ws, theme, config)
    first, last = PLAN_HEADER_ROW + 1, PLAN_HEADER_ROW + months
    write_plan_results(ws, theme, first, last)
    write_header(ws, theme, PLAN_HEADER_ROW, 2, ["Month", "Start Balance", "Interest", "Regular Payment",
                                                 "Extra Goal", "Extra Paid", "End Balance", text["asset"],
                                                 "Underwater By", "Status", "Balance (no extra)"])
    extra = config["payoff_plan"]["extra_paid"]
    for n in range(months):
        write_plan_row(ws, theme, first + n, n, extra[n] if n < len(extra) else None)
    add_plan_rules(ws, theme, first, last)
    set_print_area(ws, 12, last)
    for letter, width in zip("BCDEFGHIJKL", (40, 15, 12, 16, 14, 14, 15, 15, 15, 14, 18)):
        ws.column_dimensions[letter].width = width
    for i, label in enumerate(("Month", "Balance (no extra)", text["asset"])):
        letter = get_column_letter(PLAN_HELPER_COL + i)
        ws.cell(row=PLAN_HEADER_ROW, column=PLAN_HELPER_COL + i, value=label)
        ws.column_dimensions[letter].hidden = True
    ws.freeze_panes = f"C{first}"


def plan_cells() -> dict:
    return {key: f"$C${plan_row(key)}" for key, _, _ in PLAN_INPUTS}


def write_plan_inputs(ws, theme: Theme, config: dict):
    plan, pay = config["payoff_plan"], config["pay"]
    per_month = {"weekly": 52 / 12, "biweekly": 26 / 12, "semimonthly": 2, "monthly": 1}
    defaults = {"paychecks_per_month": per_month.get(pay["frequency"])}
    write_section_title(ws, theme, f"B{PLAN_INPUT_ROW - 1}", "Your Numbers")
    for key, label, fmt in PLAN_INPUTS:
        row = plan_row(key)
        value = plan.get(key)
        if value is None:
            value = defaults.get(key)
        label = plan_label(key, label, plan["loan_type"])
        style_cell(ws.cell(row=row, column=2, value=label), theme, fill=theme["calc"])
        style_cell(ws.cell(row=row, column=3, value=value if value != "" else None), theme,
                   fill=theme["input"], fmt=theme.money if fmt == "money" else fmt, align="right")


def write_plan_results(ws, theme: Theme, first: int, last: int):
    c = plan_cells()
    months = f"$B${first}:$B${last}"

    n_col, o_col, p_col = (get_column_letter(PLAN_HELPER_COL + i) for i in range(3))
    helper_months = f"${n_col}${first}:${n_col}${last}"
    plain = f"${o_col}${first}:${o_col}${last}"
    plain_value = f"${p_col}${first}:${p_col}${last}"

    def first_month(condition: str, months=months) -> str:
        return (f'=IF(N({c["balance"]})=0,"",IFERROR(INDEX({months},'
                f'MATCH(TRUE,INDEX({condition},0),0)),"After {(last - first + 1) // 12} years"))')

    row = PLAN_INPUT_ROW
    target = c["target_gap"]
    target_month = first_month(f"$J${first}:$J${last}<={target}")  # drop its "=" to nest it below
    results = [
        ("Monthly take-home", f"=N({c['take_home']})*N({c['paychecks_per_month']})", theme.money),
        ("Bills + regular loan payment", f"=N({c['monthly_bills']})+N({c['payment']})", theme.money),
        ("Savings", f"=ROUND(G{row}*N({c['savings_percent']})/100,2)", theme.money),
        ("Buffer", f"=N({c['buffer']})", theme.money),
        ("EXTRA PAYMENT GOAL (per month)", f"=MAX(0,ROUND(G{row}-G{row + 1}-G{row + 2}-G{row + 3},2))", theme.money),
        ("Above water by", first_month(f"$J${first}:$J${last}<=0"), "mmm yyyy"),
        ("Above water by (regular payments only)",
         first_month(f"{plain}-{plain_value}<=0", helper_months), "mmm yyyy"),
        ("Paid off by", first_month(f"$H${first}:$H${last}<=0"), "mmm yyyy"),
        ("Paid off by (regular payments only)", first_month(f"{plain}<=0", helper_months), "mmm yyyy"),
        ("Target reached by", f'=IF(ISNUMBER({target}),{target_month[1:]},"")', "mmm yyyy"),
    ]
    write_section_title(ws, theme, f"E{PLAN_INPUT_ROW - 1}", "Your Plan")
    for i, (label, formula, fmt) in enumerate(results):
        r = row + i
        goal = i == PLAN_GOAL_OFFSET
        color = theme.accent if goal else None
        ws.merge_cells(f"E{r}:F{r}")
        style_cell(ws.cell(row=r, column=5, value=label), theme, fill=theme["calc"],
                   font=theme.font(10, bold=goal, color=color))
        style_cell(ws.cell(row=r, column=6), theme, fill=theme["calc"])
        style_cell(ws.cell(row=r, column=7, value=formula), theme, fill=theme["calc"], fmt=fmt, align="right",
                   font=theme.font(12 if goal else 10, bold=True, color=color))


def write_plan_row(ws, theme: Theme, row: int, n: int, extra):
    c = plan_cells()
    goal = f"$G${PLAN_INPUT_ROW + PLAN_GOAL_OFFSET}"
    n_col, o_col, p_col = (get_column_letter(PLAN_HELPER_COL + i) for i in range(3))
    month = f'IF({c["start"]}="","",EDATE({c["start"]},{n}))'
    value = f"MAX(0,N({c['asset_value']})-N({c['value_drop']})*{n + 1})"
    prev_plain = f"N({c['balance']})" if n == 0 else f"{o_col}{row - 1}"
    formulas = {
        2: (month, "mmm yyyy"),
        3: (f"N({c['balance']})" if n == 0 else f"H{row - 1}", theme.money),
        4: (f"ROUND(C{row}*N({c['apr']})/1200,2)", theme.money),
        5: (f"MIN(N({c['payment']}),C{row}+D{row})", theme.money),
        6: (f"MIN({goal},MAX(0,C{row}+D{row}-E{row}))", theme.money),
        8: (f'MAX(0,C{row}+D{row}-E{row}-IF(G{row}="",F{row},G{row}))', theme.money),
        9: (value, theme.money),
        10: (f"H{row}-I{row}", theme.money),
        11: (f'IF(N({c["balance"]})=0,"",IF(H{row}<=0,"Paid off",IF(J{row}<=0,"Above water","Underwater")))',
             None),
        12: (f"{o_col}{row}", theme.money),
    }
    # Only the payoff month and the months before it are shown.
    active = None if n == 0 else f"N(H{row - 1})>0"
    fill = theme["zebra"] if n % 2 else theme["card"]
    for col, (formula, fmt) in formulas.items():
        value_text = f"={formula}" if active is None else f'=IF({active},{formula},"")'
        style_cell(ws.cell(row=row, column=col, value=value_text), theme, fill=fill, fmt=fmt,
                   align="left" if col == 2 else "right")
    style_cell(ws.cell(row=row, column=7, value=extra), theme, fill=theme["input"], fmt=theme.money, align="right")
    helpers = (f"={month}", f"=MAX(0,ROUND({prev_plain}*(1+N({c['apr']})/1200),2)-N({c['payment']}))", f"={value}")
    for i, formula in enumerate(helpers):
        ws.cell(row=row, column=PLAN_HELPER_COL + i, value=formula)


def add_plan_rules(ws, theme: Theme, first: int, last: int):
    # Months after the payoff month are blank: hide their stripes and borders too.
    hidden_side = Side(style="thin", color=theme["bg"])
    ws.conditional_formatting.add(f"B{first}:L{last}", FormulaRule(
        formula=[f'$C{first}=""'], stopIfTrue=True, fill=theme.fill(theme["bg"]),
        border=Border(left=hidden_side, right=hidden_side, top=hidden_side, bottom=hidden_side)))
    # Highlight the first month the gap is within the target.
    target = plan_cells()["target_gap"]
    ws.conditional_formatting.add(f"B{first}:L{last}", FormulaRule(
        formula=[f"AND(ISNUMBER({target}),$J{first}<={target},OR(ROW()={first},N($J{first - 1})>{target}))"],
        fill=theme.fill(theme.accent_tint), font=Font(name=FONT_NAME, bold=True, color=theme["text"])))
    status = f"K{first}:K{last}"
    for text, color in (("Underwater", NEGATIVE), ("Above water", POSITIVE), ("Paid off", POSITIVE)):
        ws.conditional_formatting.add(status, FormulaRule(
            formula=[f'K{first}="{text}"'], font=Font(name=FONT_NAME, bold=True, color=color)))
    # Underwater By: a positive gap is bad here, so the usual sign colors are flipped.
    for operator, color in (("greaterThan", NEGATIVE), ("lessThanOrEqual", POSITIVE)):
        ws.conditional_formatting.add(f"J{first}:J{last}", CellIsRule(
            operator=operator, formula=["0"], font=Font(name=FONT_NAME, bold=True, color=color)))


# ---------------------------------------------------------------- dashboard

def build_dashboard(wb, theme: Theme, config: dict, refs: dict):
    profile = config["profile"]
    ws = wb.create_sheet(profile["sheets"]["dashboard"])
    paint_sheet(ws, theme)
    ws.sheet_view.zoomScale = 90
    subtitle = f"Annual Overview  |  {period_label(config)}  |  {profile['pay_prefix']} {frequency_text(config['pay'])}"
    if config["owner"]:
        subtitle += f"  |  {config['owner']}"
    write_banner(ws, theme, config["title"], subtitle, 14)
    for col in range(2, 15):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["J"].width = 3
    ws.column_dimensions["K"].width = 22
    ws.column_dimensions["L"].width = 18
    month_total = write_month_table(ws, theme, config, refs)
    cat_total = write_category_table(ws, theme, config, refs)
    write_kpis(ws, theme, profile, month_total)
    add_month_chart(ws, theme, profile)
    add_category_chart(ws, theme, config, cat_total)
    tracker_row = max(month_total, cat_total) + 3
    has_trackers = config["features"]["savings_goals"] or config["features"]["debt_tracker"]
    write_tracker_cards(ws, theme, config, refs, tracker_row)
    ws.print_area = f"A1:O{tracker_row + 4 if has_trackers else tracker_row}"
    ws.page_setup.fitToHeight = 1


def write_kpis(ws, theme: Theme, profile: dict, total: int):
    # Month table total row: C=count, D=in, E=fixed, F=variable, G=out, H=net.
    values = [
        (f"=D{total}", theme.money_whole), (f"=G{total}", theme.money_whole),
        (f"=H{total}", theme.money_whole), (f"=IFERROR(H{total}/D{total},0)", "0.0%"),
        (f"=G{total}/12", theme.money_whole), (f"=C{total}", "0"),
    ]
    for col, (label, caption), (formula, fmt) in zip("BDFHKM", profile["kpis"], values):
        write_card(ws, theme, col, 5, (label, formula, fmt, caption), value_size=20)
    for letter in ("F", "H"):
        add_sign_colors(ws, f"{letter}6", size=20)


def add_data_bar(ws, cell_range: str, color: str, end_type="max", end_value=None):
    # Cap bar length at half the cell so right-aligned numbers stay readable.
    ws.conditional_formatting.add(cell_range, DataBarRule(
        start_type="num", start_value=0, end_type=end_type, end_value=end_value,
        color=color, minLength=0, maxLength=30))


def add_sign_colors(ws, cell_range: str, size=None):
    for operator, color in (("lessThan", NEGATIVE), ("greaterThanOrEqual", POSITIVE)):
        ws.conditional_formatting.add(cell_range, CellIsRule(
            operator=operator, formula=["0"], font=Font(name=FONT_NAME, size=size, bold=True, color=color)))


def write_month_table(ws, theme: Theme, config: dict, refs: dict) -> int:
    header = MONTH_TABLE_HEADER
    write_section_title(ws, theme, f"B{header - 1}", "Monthly Breakdown")
    write_header(ws, theme, header, 2, config["profile"]["month_headers"])
    first, last = header + 1, header + 12
    for i, month in enumerate(window_months(config)):
        write_month_row(ws, theme, config, first + i, i, month, refs)
    total = last + 1
    ws.cell(row=total, column=2, value="TOTAL")
    for col in range(3, 9):
        letter = get_column_letter(col)
        ws.cell(row=total, column=col, value=f"=SUM({letter}{first}:{letter}{last})")
    ws.cell(row=total, column=9, value=f"=IFERROR(H{total}/D{total},0)")
    style_total_row(ws, theme, total, 2, 9, height=None)
    for col in range(3, 10):
        ws.cell(row=total, column=col).number_format = {3: "0", 9: "0.0%"}.get(col, theme.money)
        ws.cell(row=total, column=col).alignment = Alignment(horizontal="right")
    add_month_table_rules(ws, theme, first, last, total)
    return total


def write_month_row(ws, theme: Theme, config: dict, row: int, index: int, month: date, refs: dict):
    fixed_type, variable_type = config["profile"]["types"]
    income_col = get_column_letter(INCOME_MONTH_COL + index)
    exp_col = get_column_letter(EXPENSE_MONTH_COL + index)
    exp = refs["expenses_ref"]
    exp_range = f"{exp}!${exp_col}${refs['exp_first']}:${exp_col}${refs['exp_last']}"
    type_range = f"{exp}!$D${refs['exp_first']}:$D${refs['exp_last']}"
    values = [
        (month, "mmm yyyy"),
        (f"=COUNTIFS({refs['pay_month']},B{row})", "0"),
        (f"={refs['income_ref']}!{income_col}{refs['income_total']}", theme.money),
        (f'=SUMIFS({exp_range},{type_range},"{fixed_type}")', theme.money),
        (f'=SUMIFS({exp_range},{type_range},"{variable_type}")', theme.money),
        (f"=E{row}+F{row}", theme.money),
        (f"=D{row}-G{row}", theme.money),
        (f"=IFERROR(H{row}/D{row},0)", "0.0%"),
    ]
    fill = theme["zebra"] if index % 2 else theme["card"]
    for offset, (value, fmt) in enumerate(values):
        style_cell(ws.cell(row=row, column=2 + offset, value=value), theme, fill=fill, fmt=fmt,
                   align="left" if offset == 0 else "right")


def add_month_table_rules(ws, theme: Theme, first: int, last: int, total: int):
    add_sign_colors(ws, f"H{first}:H{total}")
    # Highlight 3-paycheck (bi-weekly) or 5-paycheck (weekly) months.
    ws.conditional_formatting.add(f"B{first}:C{last}", FormulaRule(
        formula=[f"$C{first}>MIN($C${first}:$C${last})"], fill=theme.fill(theme.accent_tint),
        font=Font(name=FONT_NAME, bold=True, color=theme["text"])))


def write_category_table(ws, theme: Theme, config: dict, refs: dict) -> int:
    profile = config["profile"]
    header = MONTH_TABLE_HEADER
    write_section_title(ws, theme, f"K{header - 1}", profile["category_title"])
    write_header(ws, theme, header, CATEGORY_COL, ["Category", "Annual", "Monthly Avg", profile["category_share"]])
    categories = category_list(config)
    first, last = header + 1, header + len(categories)
    total = last + 1
    exp, total_col = refs["expenses_ref"], get_column_letter(EXPENSE_TOTAL_COL)
    cat_range = f"{exp}!$C${refs['exp_first']}:$C${refs['exp_last']}"
    sum_range = f"{exp}!${total_col}${refs['exp_first']}:${total_col}${refs['exp_last']}"
    for i, name in enumerate(categories):
        row = first + i
        fill = theme["zebra"] if i % 2 else theme["card"]
        cells = [(name, None), (f"=SUMIF({cat_range},K{row},{sum_range})", theme.money),
                 (f"=L{row}/12", theme.money), (f"=IFERROR(L{row}/$L${total},0)", "0.0%")]
        for offset, (value, fmt) in enumerate(cells):
            style_cell(ws.cell(row=row, column=CATEGORY_COL + offset, value=value), theme, fill=fill, fmt=fmt,
                       align="left" if offset == 0 else "right")
    ws.cell(row=total, column=CATEGORY_COL, value="TOTAL")
    for letter in ("L", "M"):
        ws[f"{letter}{total}"] = f"=SUM({letter}{first}:{letter}{last})"
    ws[f"N{total}"] = f"=IF(L{total}>0,1,0)"
    style_total_row(ws, theme, total, CATEGORY_COL, 14, height=None)
    for letter, fmt in (("L", theme.money), ("M", theme.money), ("N", "0.0%")):
        ws[f"{letter}{total}"].number_format = fmt
    add_data_bar(ws, f"L{first}:L{last}", theme.accent)
    return total


def add_month_chart(ws, theme: Theme, profile: dict):
    first, last = MONTH_TABLE_HEADER + 1, MONTH_TABLE_HEADER + 12
    bars = BarChart()
    bars.type, bars.grouping, bars.gapWidth, bars.overlap = "col", "clustered", 60, -10
    for col, color in ((4, theme.accent), (7, theme.palette[1])):
        bars.add_data(Reference(ws, min_col=col, min_row=MONTH_TABLE_HEADER, max_row=last), titles_from_data=True)
        color_series(bars.series[-1], color)
    net = LineChart()
    net.add_data(Reference(ws, min_col=8, min_row=MONTH_TABLE_HEADER, max_row=last), titles_from_data=True)
    line = net.series[0]
    color_series(line, POSITIVE, line_width=28575)
    line.smooth = True
    line.marker.symbol = "circle"
    line.marker.size = 5
    line.marker.graphicalProperties = GraphicalProperties(solidFill=POSITIVE, ln=LineProperties(solidFill=POSITIVE))
    months = Reference(ws, min_col=2, min_row=first, max_row=last)
    bars.set_categories(months)
    net.set_categories(months)
    style_chart(bars, theme, profile["flow_chart_title"])
    style_axes(bars, theme, theme.money_whole)
    bars.x_axis.number_format = "mmm"
    bars += net
    bars.width, bars.height = 18.6, 9.2
    ws.add_chart(bars, "B9")


def add_category_chart(ws, theme: Theme, config: dict, cat_total: int):
    first, last = MONTH_TABLE_HEADER + 1, cat_total - 1
    kind = config["category_chart"]
    chart = make_category_chart(kind)
    chart.add_data(Reference(ws, min_col=12, min_row=MONTH_TABLE_HEADER, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=CATEGORY_COL, min_row=first, max_row=last))
    series = chart.series[0]
    color_points(series, theme, last - first + 1)
    style_chart(chart, theme, config["profile"]["category_title"])
    if kind in ("pie", "doughnut"):
        chart.legend.position = "r"
        series.dLbls = data_labels(theme, True, '[<0.03]"";0%')  # hide labels on tiny slices
    else:
        chart.legend = None
        chart.varyColors = True
        chart.gapWidth = 50
        style_axes(chart, theme, theme.money_whole)
        chart.x_axis.tickLblSkip = 1
        series.dLbls = data_labels(theme, False, theme.money_label)
    chart.width, chart.height = 15.4, 9.2
    ws.add_chart(chart, "I9")


def make_category_chart(kind: str):
    if kind == "pie":
        return PieChart()
    if kind == "doughnut":
        chart = DoughnutChart()
        chart.holeSize = 55
        return chart
    chart = BarChart()
    chart.type = "bar" if kind == "bar" else "col"
    if kind == "bar":
        chart.x_axis.scaling.orientation = "maxMin"  # first category on top
        chart.y_axis.crosses = "max"  # keeps the value axis at the bottom after flipping
    return chart


def write_tracker_cards(ws, theme: Theme, config: dict, refs: dict, row: int):
    profile = config["profile"]
    has_goals = config["features"]["savings_goals"]
    if has_goals:
        goals, total = sheet_ref(profile["sheets"]["goals"]), refs["goals_total"]
        write_section_title(ws, theme, f"B{row}", profile["goals_title"])
        cards = [("B", ("TOTAL TARGET", f"={goals}!C{total}", theme.money_whole, "across all goals")),
                 ("D", ("SAVED SO FAR", f"={goals}!D{total}", theme.money_whole, "toward your goals")),
                 ("F", ("PROGRESS", f"=IFERROR({goals}!D{total}/{goals}!C{total},0)", "0.0%", "of total target"))]
        for col, card in cards:
            write_card(ws, theme, col, row + 1, card)
    if config["features"]["debt_tracker"]:
        debts, total = sheet_ref(profile["sheets"]["debts"]), refs["debts_total"]
        first, second = ("H", "K") if has_goals else ("B", "D")
        write_section_title(ws, theme, f"{first}{row}", profile["debts_title"])
        write_card(ws, theme, first, row + 1,
                   (profile["debt_card"], f"={debts}!C{total}", theme.money_whole, "current balances"))
        write_card(ws, theme, second, row + 1,
                   ("MIN PAYMENTS", f"={debts}!E{total}", theme.money_whole, "per month"))


# ---------------------------------------------------------------- assembly

def build_workbook(config: dict) -> Workbook:
    theme = Theme(config)
    sheets = config["profile"]["sheets"]
    wb = Workbook()
    wb.remove(wb.active)
    lists = build_lists(wb, config)
    paychecks = build_paychecks(wb, theme, config)
    income = build_income(wb, theme, config, lists, paychecks)
    expenses = build_expenses(wb, theme, config, lists, paychecks)
    refs = {"pay_month": paychecks["month"], "income_total": income["total_row"],
            "income_ref": sheet_ref(sheets["income"]), "expenses_ref": sheet_ref(sheets["expenses"]),
            "exp_first": expenses["first"], "exp_last": expenses["last"]}
    if config["features"]["savings_goals"]:
        refs["goals_total"] = build_savings(wb, theme, config)["total_row"]
    if config["features"]["debt_tracker"]:
        refs["debts_total"] = build_debts(wb, theme, config)["total_row"]
    if config["features"]["payoff_plan"]:
        build_payoff_plan(wb, theme, config)
    build_dashboard(wb, theme, config, refs)
    order_sheets(wb, sheets)
    wb.properties.creator = CREDIT_TEXT
    wb.properties.title = config["title"]
    wb.calculation.fullCalcOnLoad = True
    return wb


def order_sheets(wb, sheets: dict):
    order = [sheets[k] for k in ("dashboard", "income", "pay", "expenses", "goals", "debts", "payoff", "lists")]
    for target, title in enumerate(t for t in order if t in wb.sheetnames):
        wb.move_sheet(title, offset=target - wb.sheetnames.index(title))
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == sheets["dashboard"]


def default_output_name(config: dict) -> str:
    prefix = config["profile"]["file_prefix"]
    if config["start_month"] == 1:
        return f"{prefix} {config['year']}.xlsx"
    start, end = window_bounds(config)
    return f"{prefix} {start:%b %Y}-{end:%b %Y}.xlsx"


def unlink_label_formats(path: Path):
    """openpyxl omits sourceLinked on data-label number formats, so Excel ignores them."""
    pattern = re.compile(rb'<numFmt formatCode="([^"]*)"\s*/>')
    with zipfile.ZipFile(path) as src:
        parts = {info: src.read(info.filename) for info in src.infolist()}
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for info, data in parts.items():
            if info.filename.startswith("xl/charts/chart"):
                data = pattern.sub(rb'<numFmt formatCode="\1" sourceLinked="0" />', data)
            dst.writestr(info, data)


def backup_path(path: Path) -> Path:
    return path.with_name(f"{path.stem} (backup){path.suffix}")


def build_from_config(raw: dict, out: str | None = None, backup: bool = False) -> dict:
    config = validate_config(raw)
    path = Path(out or default_output_name(config)).resolve()
    result = {"file": str(path)}
    if backup and path.exists():
        shutil.copy2(path, backup_path(path))
        result["backup"] = str(backup_path(path))
    build_workbook(config).save(path)
    unlink_label_formats(path)
    return {**result, **compute_summary(config)}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Generate a Budget Planner Excel workbook.")
    parser.add_argument("--config", required=True, help="Path to budget-config.json")
    parser.add_argument("--out", help="Output .xlsx path")
    parser.add_argument("--backup", action="store_true",
                        help="If the output file exists, copy it to '<name> (backup).xlsx' first")
    args = parser.parse_args(argv)
    try:
        # utf-8-sig accepts the BOM that Windows PowerShell adds when saving files.
        raw = json.loads(Path(args.config).read_text(encoding="utf-8-sig"))
        result = build_from_config(raw, args.out, args.backup)
    except FileNotFoundError:
        print(f"ERROR: config file not found: {args.config}", file=sys.stderr)
        return 2
    except json.JSONDecodeError as exc:
        print(f"ERROR: config is not valid JSON: {exc}", file=sys.stderr)
        return 2
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except PermissionError:
        print("ERROR: cannot write the workbook. Close it in Excel and try again.", file=sys.stderr)
        return 2
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""Summarize bank or credit card CSV exports into exact totals for the budget.

Usage:
    python import_transactions.py statement.csv [more.csv ...] [--start 2026-01] [--invert] [--day-first]

Reads only the date, description and amount columns (account numbers and other
columns are ignored), groups transactions by merchant, and prints JSON with
exact per-month totals and likely recurring bills. Sign convention in the
output: money_out is spending, money_in is deposits. Use --invert for card
exports that list purchases as positive numbers.
Everything happens locally; this script never opens a network connection.
"""

import argparse
import csv
import json
import re
import statistics
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

DATE_HEADERS = ("date", "transaction date", "trans date", "posted date", "post date", "posting date", "booking date")
DESC_HEADERS = ("description", "payee", "merchant", "name", "memo", "details", "transaction description",
                "original description")
AMOUNT_HEADERS = ("amount", "transaction amount", "amount (usd)")
DEBIT_HEADERS = ("debit", "debits", "withdrawal", "withdrawals", "money out", "debit amount", "paid out")
CREDIT_HEADERS = ("credit", "credits", "deposit", "deposits", "money in", "credit amount", "paid in")
TYPE_HEADERS = ("transaction type", "type", "credit/debit", "debit/credit", "cr/dr", "dr/cr")
OUTFLOW_TYPES = ("debit", "dr", "withdrawal", "purchase", "sale")
INFLOW_TYPES = ("credit", "cr", "deposit")
DATE_FORMATS_MONTH_FIRST = ("%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d %b %Y", "%b %d, %Y",
                            "%d-%b-%Y", "%d-%b-%y")
DATE_FORMATS_DAY_FIRST = ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d %b %Y", "%b %d, %Y",
                          "%d-%b-%Y", "%d-%b-%y")
NOISE_PREFIXES = re.compile(
    r"^(POS |DEBIT CARD PURCHASE |DEBIT PURCHASE |PURCHASE AUTHORIZED ON \S+ |CHECKCARD \S+ |CHECKCARD |"
    r"RECURRING PAYMENT |ACH DEBIT |ACH CREDIT |ACH |ONLINE PAYMENT |ELECTRONIC PAYMENT |PREAUTHORIZED DEBIT |"
    r"VISA PURCHASE |SQ \*|TST\* |PAYPAL \*)+")
HEADER_SCAN_ROWS = 15


# ---------------------------------------------------------------- parsing

def find_column(headers: list[str], candidates: tuple):
    lowered = [h.strip().lower() for h in headers]
    for name in candidates:
        if name in lowered:
            return lowered.index(name)
    for i, header in enumerate(lowered):
        if any(name in header for name in candidates):
            return i
    return None


def exact_column(headers: list[str], candidates: tuple):
    # Exact match only: "type" as a substring would also hit "Transaction Type Code" and similar.
    lowered = [h.strip().lower() for h in headers]
    return next((lowered.index(name) for name in candidates if name in lowered), None)


def locate_header(rows: list[list[str]]):
    """Banks often put account info above the real header; find the first row that looks like one."""
    for index, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        has_date = find_column(row, DATE_HEADERS) is not None
        has_amount = find_column(row, AMOUNT_HEADERS) is not None or find_column(row, DEBIT_HEADERS) is not None
        if has_date and has_amount:
            return index
    raise ValueError("could not find a header row with a date column and an amount (or debit/credit) column")


def parse_amount(text: str):
    value = (text or "").strip()
    if not value:
        return None
    negative = value.startswith("(") and value.endswith(")") or value.endswith("-") or value.upper().endswith("DR")
    cleaned = re.sub(r"[^0-9.\-]", "", value.replace("CR", "").replace("DR", ""))
    if cleaned in ("", "-", "."):
        return None
    number = float(cleaned.rstrip("-"))
    return -abs(number) if negative else number


def detect_day_first(date_texts: list[str]) -> bool:
    """dd/mm if any slash or dash date has a first field above 12."""
    for text in date_texts:
        parts = re.split(r"[/-]", text.strip())
        if len(parts) == 3 and parts[0].isdigit() and len(parts[0]) <= 2 and int(parts[0]) > 12:
            return True
    return False


def parse_date(text: str, day_first: bool):
    formats = DATE_FORMATS_DAY_FIRST if day_first else DATE_FORMATS_MONTH_FIRST
    cleaned = text.strip().split(" ")[0] if re.match(r"^\d", text.strip()) and ":" in text else text.strip()
    for fmt in formats:
        try:
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
    return None


def signed_amount(row: list[str], columns: dict):
    if columns["amount"] is not None:
        amount = parse_amount(cell(row, columns["amount"]))
        kind = cell(row, columns["type"]).strip().lower() if columns["type"] is not None else ""
        if amount is not None and kind in OUTFLOW_TYPES:
            return -abs(amount)
        if amount is not None and kind in INFLOW_TYPES:
            return abs(amount)
        return amount
    debit = parse_amount(cell(row, columns["debit"])) if columns["debit"] is not None else None
    credit = parse_amount(cell(row, columns["credit"])) if columns["credit"] is not None else None
    if debit is None and credit is None:
        return None
    return (credit or 0) - abs(debit or 0)


def cell(row: list[str], index: int) -> str:
    return row[index] if index < len(row) else ""


def normalize_merchant(description: str) -> str:
    text = NOISE_PREFIXES.sub("", description.upper().strip())
    text = re.sub(r"[#*]\s*\S*\d\S*", " ", text)     # store numbers like #1234 or *AB12C
    text = re.sub(r"\b\d{2}/\d{2}(/\d{2,4})?\b", " ", text)
    text = re.sub(r"\b[\dX]{3,}\b", " ", text)       # long digit runs, masked card numbers
    text = re.sub(r"[^A-Z0-9&' ]", " ", text)
    words = text.split()[:4]
    return " ".join(words).title() if words else "Unknown"


def read_csv(path: Path) -> list[list[str]]:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    return [row for row in csv.reader(text.splitlines(), dialect) if any(c.strip() for c in row)]


def load_transactions(path: Path, invert: bool, day_first) -> tuple[list[dict], int]:
    rows = read_csv(path)
    if not rows:
        raise ValueError(f"{path.name} is empty")
    header_index = locate_header(rows)
    headers, body = rows[header_index], rows[header_index + 1:]
    columns = {"date": find_column(headers, DATE_HEADERS), "desc": find_column(headers, DESC_HEADERS),
               "amount": find_column(headers, AMOUNT_HEADERS), "debit": find_column(headers, DEBIT_HEADERS),
               "credit": find_column(headers, CREDIT_HEADERS), "type": exact_column(headers, TYPE_HEADERS)}
    if day_first is None:
        day_first = detect_day_first([cell(r, columns["date"]) for r in body])
    transactions, skipped = [], 0
    for row in body:
        day = parse_date(cell(row, columns["date"]), day_first)
        amount = signed_amount(row, columns)
        if day is None or amount is None or amount == 0:
            skipped += 1
            continue
        description = cell(row, columns["desc"]) if columns["desc"] is not None else ""
        transactions.append({"date": day, "amount": -amount if invert else amount,
                             "description": description.strip(), "merchant": normalize_merchant(description)})
    return transactions, skipped


# ---------------------------------------------------------------- summarizing

def month_key(day: date) -> str:
    return f"{day.year:04d}-{day.month:02d}"


def window_keys(start: str) -> list[str]:
    match = re.fullmatch(r"(\d{4})-(\d{1,2})", start or "")
    if not match or not 1 <= int(match.group(2)) <= 12:
        raise ValueError("--start must look like 2026-01")
    year, month = int(match.group(1)), int(match.group(2))
    keys = []
    for offset in range(12):
        index = month - 1 + offset
        keys.append(f"{year + index // 12:04d}-{index % 12 + 1:02d}")
    return keys


def summarize_group(merchant: str, items: list[dict], all_months: list[str], window) -> dict:
    by_month = defaultdict(float)
    for item in items:
        by_month[month_key(item["date"])] += abs(item["amount"])
    monthly = {k: round(v, 2) for k, v in sorted(by_month.items())}
    values = list(monthly.values())
    group = {
        "merchant": merchant,
        "count": len(items),
        "total": round(sum(values), 2),
        "months": monthly,
        "average_per_month": round(sum(values) / max(1, len(all_months)), 2),
        "recurring": is_recurring(values, len(all_months)),
        "typical_day": int(statistics.median(i["date"].day for i in items)),
        "examples": sorted({i["description"] for i in items})[:2],
    }
    if window:
        group["window_months"] = [monthly.get(k, 0.0) for k in window]
    return group


def is_recurring(values: list[float], months_covered: int) -> bool:
    """Shows up in most covered months (at least 2) with a fairly steady amount."""
    needed = max(2, round(months_covered * 0.75))
    if len(values) < needed:
        return False
    mean = statistics.mean(values)
    spread = statistics.pstdev(values) / mean if mean else 1
    return spread <= 0.25


def summarize(transactions: list[dict], window) -> dict:
    if not transactions:
        raise ValueError("no transactions could be read from the file(s)")
    all_months = sorted({month_key(t["date"]) for t in transactions})
    result = {"date_range": [min(t["date"] for t in transactions).isoformat(),
                             max(t["date"] for t in transactions).isoformat()],
              "months_covered": all_months, "transactions": len(transactions)}
    for direction, keep in (("money_out", lambda a: a < 0), ("money_in", lambda a: a > 0)):
        groups = defaultdict(list)
        for t in transactions:
            if keep(t["amount"]):
                groups[t["merchant"]].append(t)
        summaries = [summarize_group(m, items, all_months, window) for m, items in groups.items()]
        result[direction] = sorted(summaries, key=lambda g: g["total"], reverse=True)
        result[f"{direction}_total"] = round(sum(g["total"] for g in summaries), 2)
    if window:
        result["window"] = window
    return result


def run(paths: list[str], start=None, invert=False, day_first=None) -> dict:
    window = window_keys(start) if start else None
    transactions, skipped, sources = [], 0, []
    for name in paths:
        path = Path(name)
        found, missed = load_transactions(path, invert, day_first)
        transactions += found
        skipped += missed
        sources.append({"file": path.name, "transactions": len(found), "skipped_rows": missed})
    return {"sources": sources, "skipped_rows": skipped, **summarize(transactions, window)}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Summarize bank or card CSV exports for the budget.")
    parser.add_argument("files", nargs="+", help="CSV exports")
    parser.add_argument("--start", help="Budget window start as YYYY-MM; adds 12-month arrays to each merchant")
    parser.add_argument("--invert", action="store_true", help="Flip signs (card exports with purchases positive)")
    parser.add_argument("--day-first", action="store_true", default=None, help="Dates are dd/mm/yyyy")
    args = parser.parse_args(argv)
    try:
        result = run(args.files, args.start, args.invert, args.day_first)
    except FileNotFoundError as exc:
        print(f"ERROR: file not found: {exc.filename}", file=sys.stderr)
        return 2
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())

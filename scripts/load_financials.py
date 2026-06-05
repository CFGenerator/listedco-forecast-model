#!/usr/bin/env python3
"""Load historical financial data from local files or AkShare.

Sign convention for template rows:
- expense rows in the Excel template are stored as negative values
- income or tax benefit rows are stored as positive values

Income tax rule:
- source `INCOME_TAX` > 0 means tax expense and should become a negative template value
- source `INCOME_TAX` < 0 means tax benefit / tax refund and should become a positive template value
- therefore map template tax row as `-INCOME_TAX`, not `-abs(INCOME_TAX)`
"""

from __future__ import annotations

import csv
import json
import sys
from datetime import datetime, timezone
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
WORKSPACE_DIR = SKILL_DIR.parent
LOCAL_PYDEPS = WORKSPACE_DIR / ".pydeps"

if LOCAL_PYDEPS.exists():
    sys.path.insert(0, str(LOCAL_PYDEPS))


def load_json(path: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        if isinstance(data.get("historical_financials"), list):
            return data["historical_financials"]
        return [data]
    raise ValueError("Unsupported JSON structure")


def load_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def normalize_ticker(ticker: str) -> str:
    raw = ticker.strip().upper()
    if raw.startswith(("SH", "SZ", "BJ")) and len(raw) == 8:
        return raw
    if raw.endswith((".SH", ".SZ", ".BJ")):
        code, market = raw.split(".")
        return f"{market}{code}"
    if len(raw) == 6:
        if raw.startswith(("6", "9")):
            return f"SH{raw}"
        if raw.startswith(("0", "3")):
            return f"SZ{raw}"
        if raw.startswith("8"):
            return f"BJ{raw}"
    raise ValueError(f"Unsupported ticker format: {ticker}")


def safe_number(value):
    if value in (None, "", "None"):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def value_from_row(row: dict, *keys: str):
    for key in keys:
        if key in row:
            value = safe_number(row.get(key))
            if value is not None:
                return value
    return None


def build_template_payload(df, normalized_ticker: str) -> dict:
    records = df.sort_values("REPORT_DATE").tail(5).copy()
    years = [int(str(value)[:4]) for value in records["REPORT_DATE"].tolist()]

    revenue = []
    gross_profit = []
    sales_marketing = []
    business_tax_surcharges = []
    general_admin = []
    research = []
    operating_profit = []
    net_interest = []
    profit_before_tax = []
    tax_expenses = []
    raw_records = []
    net_income_validation = []

    for _, row in records.iterrows():
        row_dict = row.to_dict()

        top_line = value_from_row(row_dict, "OPERATE_INCOME", "TOTAL_OPERATE_INCOME")
        operate_cost = value_from_row(row_dict, "OPERATE_COST")
        gross_profit_value = None if top_line is None or operate_cost is None else top_line - operate_cost

        interest_expense = value_from_row(row_dict, "FE_INTEREST_EXPENSE", "INTEREST_EXPENSE")
        interest_income = value_from_row(row_dict, "FE_INTEREST_INCOME", "INTEREST_INCOME")
        finance_expense = value_from_row(row_dict, "FINANCE_EXPENSE")

        if interest_expense is not None or interest_income is not None:
            net_interest_value = (interest_expense or 0.0) - (interest_income or 0.0)
        else:
            net_interest_value = finance_expense

        tax_value = value_from_row(row_dict, "INCOME_TAX")
        # Eastmoney/AkShare uses positive values for tax expense and negative values for tax benefit.
        # The template stores expenses as negative rows, so invert the sign instead of forcing abs().
        tax_expense_value = None if tax_value is None else -tax_value

        revenue.append(top_line)
        gross_profit.append(gross_profit_value)
        business_tax_surcharges.append(negate_if_present(value_from_row(row_dict, "OPERATE_TAX_ADD")))
        sales_marketing.append(negate_if_present(value_from_row(row_dict, "SALE_EXPENSE")))
        general_admin.append(negate_if_present(value_from_row(row_dict, "MANAGE_EXPENSE")))
        research.append(
            negate_if_present(value_from_row(row_dict, "RESEARCH_EXPENSE", "ME_RESEARCH_EXPENSE"))
        )
        operating_profit.append(value_from_row(row_dict, "OPERATE_PROFIT"))
        net_interest.append(negate_if_present(net_interest_value))
        total_profit_value = value_from_row(row_dict, "TOTAL_PROFIT")
        profit_before_tax.append(total_profit_value)
        tax_expenses.append(tax_expense_value)
        netprofit_value = value_from_row(row_dict, "NETPROFIT")
        computed_net_income_value = (
            None if total_profit_value is None or tax_expense_value is None else total_profit_value + tax_expense_value
        )
        validation_delta = (
            None
            if computed_net_income_value is None or netprofit_value is None
            else computed_net_income_value - netprofit_value
        )
        validation_passed = (
            None
            if validation_delta is None
            else abs(validation_delta) <= 1.0
        )
        net_income_validation.append(
            {
                "year": int(str(row_dict["REPORT_DATE"])[:4]),
                "template_net_income_rmbm": to_rmb_million_scalar(computed_net_income_value),
                "akshare_netprofit_rmbm": to_rmb_million_scalar(netprofit_value),
                "delta_rmbm": to_rmb_million_scalar(validation_delta),
                "passed": validation_passed,
            }
        )
        raw_records.append(
            {
                "year": int(str(row_dict["REPORT_DATE"])[:4]),
                "security_name": row_dict.get("SECURITY_NAME_ABBR"),
                "revenue": top_line,
                "operate_cost": operate_cost,
                "gross_profit": gross_profit_value,
                "business_tax_surcharges": value_from_row(row_dict, "OPERATE_TAX_ADD"),
                "sale_expense": value_from_row(row_dict, "SALE_EXPENSE"),
                "manage_expense": value_from_row(row_dict, "MANAGE_EXPENSE"),
                "research_expense": value_from_row(row_dict, "RESEARCH_EXPENSE", "ME_RESEARCH_EXPENSE"),
                "operate_profit": value_from_row(row_dict, "OPERATE_PROFIT"),
                "total_profit": total_profit_value,
                "income_tax": tax_value,
                "netprofit": netprofit_value,
                "computed_net_income": computed_net_income_value,
            }
        )

    company_name = raw_records[-1].get("security_name") if raw_records else None
    historical_financials = {
        "revenue": to_rmb_million(revenue),
        "gross_profit": to_rmb_million(gross_profit),
        "business_tax_surcharges": to_rmb_million(business_tax_surcharges),
        "sales_marketing_expenses": to_rmb_million(sales_marketing),
        "general_admin_expenses": to_rmb_million(general_admin),
        "research_expenses": to_rmb_million(research),
        "operating_profit": to_rmb_million(operating_profit),
        "net_interest_expenses": to_rmb_million(net_interest),
        "profit_before_tax": to_rmb_million(profit_before_tax),
        "tax_expenses": to_rmb_million(tax_expenses),
    }
    missing_metrics = [
        metric_name
        for metric_name, values in historical_financials.items()
        if all(value is None for value in values)
    ]
    validation_issues = [item for item in net_income_validation if item.get("passed") is False]
    return {
        "ticker": normalized_ticker,
        "company_name": company_name,
        "currency": "CNY",
        "template_unit": "RMBm",
        "historical_years": years,
        "historical_financials": historical_financials,
        "missing_metrics": missing_metrics,
        "net_income_validation": net_income_validation,
        "validation_issues": validation_issues,
        "raw_records": raw_records,
        "fetched_at": datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
    }


def negate_if_present(value):
    if value is None:
        return None
    return -abs(value)


def to_rmb_million(values: list[float | None]) -> list[float | None]:
    result = []
    for value in values:
        if value is None:
            result.append(None)
        else:
            result.append(round(value / 1_000_000, 2))
    return result


def to_rmb_million_scalar(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value / 1_000_000, 2)


def fetch_akshare_yearly(symbol: str) -> dict:
    try:
        import akshare as ak
    except ImportError as exc:
        raise RuntimeError(
            "AkShare is not available. Install dependencies into the workspace .pydeps directory first."
        ) from exc

    df = ak.stock_profit_sheet_by_yearly_em(symbol=symbol)
    if df.empty:
        raise RuntimeError(f"No AkShare yearly profit-sheet data returned for {symbol}.")
    return build_template_payload(df, symbol)


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: load_financials.py <input.json|input.csv|ticker>", file=sys.stderr)
        return 1

    source = sys.argv[1]
    path = Path(source)

    if path.exists():
        suffix = path.suffix.lower()
        if suffix == ".json":
            records = load_json(path)
        elif suffix == ".csv":
            records = load_csv(path)
        else:
            print("Only .json and .csv are supported for local files.", file=sys.stderr)
            return 1
        print(json.dumps({"historical_financials": records}, ensure_ascii=False, indent=2))
        return 0

    ticker = normalize_ticker(source)
    payload = fetch_akshare_yearly(ticker)
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

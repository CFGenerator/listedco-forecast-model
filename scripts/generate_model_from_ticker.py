#!/usr/bin/env python3
"""Generate a new Excel model file from a ticker without modifying the source template."""

from __future__ import annotations

import json
import re
import sys
from datetime import date
from pathlib import Path

from fill_template import build_blank_template, fill_historical_block, load_mapping
from load_financials import fetch_akshare_yearly, normalize_ticker


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
DEFAULT_MAPPING = SKILL_DIR / "assets" / "template-mapping.json"


def sanitize_filename_part(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\\\|?*]+', "-", value).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or "model"


def default_output_path(ticker: str, company_name: str | None) -> Path:
    base_name = sanitize_filename_part(company_name or "listedco")
    ticker_name = sanitize_filename_part(normalize_ticker(ticker))
    today = date.today().isoformat()
    return SKILL_DIR / "assets" / f"{base_name}-{ticker_name}-{today}.xlsx"


def generate_model(ticker: str, output_path: Path, mapping_path: Path) -> dict:
    normalized_ticker = normalize_ticker(ticker)
    payload = fetch_akshare_yearly(normalized_ticker)
    mapping = load_mapping(mapping_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = build_blank_template()
    worksheet = workbook[mapping["sheet"]]
    fill_historical_block(worksheet, mapping, payload)
    workbook.save(output_path)
    return payload


def main() -> int:
    if len(sys.argv) not in (2, 3):
        print(
            "Usage: generate_model_from_ticker.py <ticker> [output.xlsx]",
            file=sys.stderr,
        )
        return 1

    ticker = sys.argv[1]

    if len(sys.argv) >= 3:
        output_path = Path(sys.argv[2])
    else:
        preview_payload = fetch_akshare_yearly(normalize_ticker(ticker))
        output_path = default_output_path(ticker, preview_payload.get("company_name"))

    payload = generate_model(ticker, output_path, DEFAULT_MAPPING)
    print(
        json.dumps(
            {
                "ticker": payload["ticker"],
                "company_name": payload.get("company_name"),
                "historical_years": payload["historical_years"],
                "template_unit": payload.get("template_unit"),
                "missing_metrics": payload.get("missing_metrics", []),
                "net_income_validation_passed": not payload.get("validation_issues", []),
                "validation_issues": payload.get("validation_issues", []),
                "output_file": str(output_path),
                "template_file": None,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

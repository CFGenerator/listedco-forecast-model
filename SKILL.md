---
name: listedco-forecast-model
description: |
  Generate a 5-year P&L forecast Excel model from an A-share ticker. Use when the user
  asks to build a valuation/forecast model for a Chinese A-share listed company, requests
  a 5-year P&L Excel template from a ticker like "688041.SH" or "600519.SH", or wants
  Akshare data turned into a structured P&L workbook with a forward forecast. Triggers on
  "forecast model", "P&L model", "盈利预测", "估值模型", "Akshare 拉数据做模型", or any
  A-share 6-digit ticker with SH/SZ/BJ prefix.

  Do NOT use for: banks, insurers, brokers, or other financial institutions
  (sector-specific statements needed); non-A-share stocks (Akshare endpoint is A-share
  only); one-off quarterly P&L snapshots (annual 5Y only); balance sheet or cash flow
  models (P&L only); loading a pre-built Excel template (this skill builds the workbook
  from code, no .xlsx template bundled).
---

# Listedco 5Y P&L Forecast Model

## Inputs to collect
- **Ticker** (required). A-share format. The script auto-normalizes:
  - `688041.SH` / `600519.SH` — full `.SH` form
  - `SH688041` / `SZ000001` / `BJ830xxx` — leading market code
  - `688041` / `600519` — bare 6-digit code (heuristic: 6/9 → SH, 0/3 → SZ, 8 → BJ)
- **Output path** (optional). Defaults to `assets/<company_name>-<ticker>-<today>.xlsx` inside this skill directory.

## Setup
```bash
pip install akshare openpyxl
```
If Akshare is blocked, `scripts/load_financials.py` accepts a pre-fetched `.json` or `.csv` file as input — see the "Failure handling" section.

## Procedure

### 1. Confirm the company is in scope
This skill is for **non-financial** listed companies with a standard P&L structure. If the ticker is a bank, insurer, broker, or other financial institution, stop and tell the user the standard P&L template does not fit — they need a sector-specific approach.

### 2. Run the end-to-end generator
```bash
python scripts/generate_model_from_ticker.py 688041.SH output.xlsx
```
This:
1. Calls `ak.stock_profit_sheet_by_yearly_em(symbol=...)` (Akshare → Eastmoney annual P&L endpoint).
2. Sorts by `REPORT_DATE`, takes the latest 5 years, normalizes to the template row schema.
3. Validates `Net income = Profit before tax + Tax expenses` against source `NETPROFIT` for each year.
4. Builds the workbook directly in code (no `.xlsx` template bundled).
5. Saves to the given path. JSON summary on stdout.

Step-by-step for inspection / debugging:
```bash
python scripts/load_financials.py 688041.SH > financials.json
python scripts/fill_template.py assets/template-mapping.json financials.json output.xlsx
```

### 3. Validate the net income bridge BEFORE declaring success
The stdout JSON contains `net_income_validation_passed` (boolean) and `validation_issues` (per-year deltas). The check is: `template Net income (row 28) = Profit before tax (row 26) + Tax expenses (row 27)` must match source `NETPROFIT` within 1 RMBm. If a year fails, surface the per-year delta to the user — do not silently export as if the data is clean.

### 4. Read row-mapping before changing the template
`scripts/fill_template.py` constructs the workbook layout in code. If you need to add a row, change a formula, or shift a column, read `references/template_mapping.md` first. The constants you'll touch most:
- **Input rows** (blue cells, columns C:G): `{2, 5, 8, 10, 12, 14, 19, 22, 26, 27}`
- **Assumption rows** (light-blue cells, columns H:L): `{32, 33, 34, 35, 36, 37, 38, 39, 40, 41}`
- **Derived rows** (do not overwrite): `{3, 4, 6, 9, 11, 13, 15, 17, 20, 23, 25, 28, 29}`

The forecast `Operating profit` (row 19) formula intentionally includes row 8 (`Business tax and surcharges`) in addition to the obvious 5/10/12/14/16 — keep it that way to match the original template.

## Output contract
- A new `.xlsx` at the given path with one sheet `P&L`, 5 historical years filled into the blue input cells, 5 forecast years computed by formulas driven by the assumption block (rows 32:41).
- JSON on stdout from the runner: `ticker`, `company_name`, `historical_years`, `template_unit`, `missing_metrics`, `net_income_validation_passed`, `validation_issues`, `output_file`, `template_file: null`.
- In your final reply, always tell the user: the file path, the net income validation outcome, and any missing metrics.

## Failure handling
- **`akshare` not installed**: tell the user to `pip install akshare openpyxl`. If still blocked, fall back to a pre-fetched JSON/CSV: `python scripts/load_financials.py path/to/data.json > financials.json` then run `fill_template.py`.
- **Net income bridge mismatch > 1 RMBm in any year**: surface the per-year delta in your reply. Ask the user whether to (a) override source `NETPROFIT` with the bridge value, (b) keep both and add an `Adjustment` row, or (c) abort.
- **Ticker normalization failure**: only A-share 6-digit codes (SH/SZ/BJ) are supported. Confirm the format with the user.
- **Empty Akshare response**: confirm the ticker is a real A-share listed company. Some recently-listed companies have thin annual history and may return fewer than 5 years.
- **Tax sign confusion** (`INCOME_TAX` field): source positive = tax expense (template stores negative); source negative = tax refund (template stores positive). The script maps `template tax = -INCOME_TAX`. Do **not** use `abs()` — that erases genuine refunds. See `references/metric_definitions.md`.
- **Missing metrics in source data**: the script records them in `missing_metrics`. Leave the corresponding blue cells blank; do not invent values, and do not overwrite formula-driven rows to force the model to balance.

## File map
- `scripts/load_financials.py` — fetch & normalize 5Y P&L (Akshare / Eastmoney) with offline JSON/CSV fallback
- `scripts/fill_template.py` — build the workbook structure from code, fill historical inputs
- `scripts/generate_model_from_ticker.py` — end-to-end ticker-to-xlsx runner
- `references/template_mapping.md` — row-by-row template rules (read before editing layout)
- `references/metric_definitions.md` — field meanings and sign conventions
- `references/forecast_rules.md` — default forecast approach and guardrails
- `references/sector_templates.md` — sector-specific override notes
- `assets/template-mapping.json` — machine-readable metric → row mapping
- `assets/template-fill-sample.json` — example payload for manual fill testing

## Examples

**Input**: "用 688041.SH 给我做一份 5 年 P&L 预测模型"

**Output**: runs `python scripts/generate_model_from_ticker.py 688041.SH`, produces
`<skill-dir>/assets/海光信息-SH688041-<today>.xlsx`, reports:
- 5 historical years covered (e.g. 2021A-2025A)
- Net income validation outcome (pass/fail per year, with delta in RMBm)
- Any missing metrics (e.g. if Eastmoney returned NaN for `RESEARCH_EXPENSE` in a year, list it)
- File path the user can open

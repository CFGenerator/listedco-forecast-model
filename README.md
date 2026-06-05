# ListedCo 5Y P&L Model Skill

Generate a simple listed-company P&L model from a ticker, including historical 5-year P&L data and a forward 5-year P&L forecast, without relying on a bundled Excel master template.

The skill fetches the latest 5 fiscal years of annual P&L data, builds the workbook structure directly from code, validates the net income bridge against source `NETPROFIT`, generates a simple 5-year forward forecast using embedded assumption logic, and exports a new workbook.

## What It Does

* Accepts a listed-company ticker, such as `688041.SH`
* Fetches annual P&L data through AkShare / Eastmoney
* Builds the workbook layout, formulas, and formatting directly from code
* Generates a simple 5-year forward P&L forecast using the template's assumption block
* Preserves the template's formulas, formatting, and forecast logic
* Validates `Net income = Profit before tax + Tax expenses` against source `NETPROFIT`

## Ticker Formats

The script auto-normalizes the input to Eastmoney's `symbol` field. All three forms are accepted:

| Input         | Example       | Notes                                |
| ------------- | ------------- | ------------------------------------ |
| Full `.SH`    | `600519.SH`   | Preferred                            |
| Leading code  | `SH600519`    | Equivalent                           |
| Bare 6-digit  | `600519`      | Heuristic: 6/9 → SH, 0/3 → SZ, 8 → BJ |

## Usage

Install dependencies as needed:

```bash
pip install akshare openpyxl
```

Generate a model:

```bash
python scripts/generate_model_from_ticker.py 688041.SH output.xlsx
```

The script builds the workbook from code, fills the historical input cells, and writes a new Excel file.

## Output Contract

The runner prints a JSON summary to stdout. Key fields:

| Field                          | Meaning                                                                                |
| ------------------------------ | -------------------------------------------------------------------------------------- |
| `ticker`                       | Normalized ticker used in the call to Eastmoney                                        |
| `company_name`                 | Company name as returned by the source                                                 |
| `historical_years`             | Years covered, e.g. `[2021, 2022, 2023, 2024, 2025]`                                    |
| `template_unit`                | Reporting unit, typically `RMBm`                                                      |
| `missing_metrics`              | Source fields that returned empty/NaN; corresponding cells are left blank               |
| `net_income_validation_passed` | `true` only if all 5 years reconcile within 1 RMBm                                     |
| `validation_issues`            | Per-year deltas for any year that failed the bridge check                               |
| `output_file`                  | Absolute path to the generated `.xlsx`                                                 |
| `template_file`                | Always `null` — no `.xlsx` template is bundled                                        |

Always report three things to the user: the file path, the net income validation outcome, and any missing metrics.

## Forecast Assumptions

The forecast section uses the assumption block generated directly by the workbook-building code.

By default, the projection-period assumptions are set to remain consistent with `2025A`, including:

| Assumption                              | Excel row |
| --------------------------------------- | --------- |
| Revenue growth                          | 32        |
| Gross margin                            | 33        |
| Business tax and surcharges as % of sales | 34      |
| Sales & marketing expenses as % of sales  | 35      |
| General & administration expenses as % of sales | 36  |
| Research expenses as % of sales         | 37        |
| Other operating expenses as % of sales  | 38        |
| Net interest expenses as % of sales     | 39        |
| Other non-operating items as % of sales | 40        |
| Corporate income tax rate               | 41        |

Users should adjust these assumptions based on the company's actual business outlook, industry conditions, and their own forecast view.

## Main Files

* `SKILL.md`: Mavis agent contract — trigger description, procedure, and output contract (read by the agent, not by humans)
* `scripts/load_financials.py`: Fetch and normalize historical P&L data
* `scripts/fill_template.py`: Fill the workbook from prepared JSON
* `scripts/generate_model_from_ticker.py`: End-to-end ticker-to-model generation
* `references/template_mapping.md`: Human-readable template mapping notes
* `references/metric_definitions.md`: Field meanings and sign conventions
* `references/forecast_rules.md`: Default forecast approach and guardrails
* `references/sector_templates.md`: Sector-specific override notes
* `assets/template-mapping.json`: Machine-readable metric-to-row mapping
* `assets/template-fill-sample.json`: Example payload for manual fill testing

## Failure Handling

| Situation                              | What to do                                                                                |
| -------------------------------------- | ----------------------------------------------------------------------------------------- |
| `akshare` not installed                | `pip install akshare openpyxl`                                                            |
| Akshare blocked                        | Pre-fetch to JSON/CSV, then `python scripts/load_financials.py data.json > financials.json` |
| Net income bridge mismatch > 1 RMBm    | Surface per-year delta; ask user to override source `NETPROFIT`, add an `Adjustment` row, or abort |
| Ticker not recognized                  | Confirm the 6-digit A-share format with the user                                          |
| Source field is empty/NaN              | Cell is left blank; the metric is recorded in `missing_metrics` — do not invent values     |
| Tax sign confusion (`INCOME_TAX`)      | Map as `-INCOME_TAX`; never use `abs()` (it would erase genuine refunds)                  |

## Not Designed For

This skill is intentionally narrow. Do not use it for:

* Banks, insurers, brokers, or other financial institutions
* Non-A-share tickers (Eastmoney endpoint is A-share only)
* Quarterly P&L snapshots (annual 5Y only)
* Balance sheet or cash-flow models (P&L only)

## Notes

* Expense rows in the template use negative values.
* `INCOME_TAX` is mapped as `-INCOME_TAX` so tax refunds remain positive in the template.

## License

MIT.

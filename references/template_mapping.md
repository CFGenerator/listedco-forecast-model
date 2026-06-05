# Template Mapping

Use the code-generated workbook layout defined in `scripts/fill_template.py` as the canonical structure.

## Workbook structure

- Sheet name: `P&L`
- Historical actual years: columns `B:F`
- Forecast years: columns `G:K`
- Year headers:
  - `B1:F1` are `2021A:2025A`
  - `G1:K1` are `2026E:2030E`

## Historical input policy

Only fill the historical input cells in the actual period.
Preserve in generated output:
- Formula structure
- Assumption cells
- Forecast formulas
- Number formats
- Code-defined styles

## Historical row mapping

Map normalized historical metrics to these rows:

- `Revenue` -> row `2` -> cells `B2:F2`
- `Gross profit` -> row `5` -> cells `B5:F5`
- `Business tax and surcharges` -> row `8` -> cells `B8:F8`
- `Sales & marketing expenses` -> row `10` -> cells `B10:F10`
- `General & administration expenses` -> row `12` -> cells `B12:F12`
- `Research expenses` -> row `14` -> cells `B14:F14`
- `Operating profit` -> row `19` -> cells `B19:F19`
- `Net interest expenses` -> row `22` -> cells `B22:F22`
- `Profit before tax` -> row `26` -> cells `B26:F26`
- `Tax expenses` -> row `27` -> cells `B27:F27`

## Derived historical rows

Do not overwrite these rows in historical periods because the template already derives them:

- `COGS` -> row `4`
- `Growth rate` -> row `3`
- `Gross margin` -> row `6`
- `As % of sales` rows -> `9`, `11`, `13`, `15`, `17`, `23`, `25`
- `Other operating expenses` -> row `16`
- `Operating profit margin` -> row `20`
- `Other non-operating income / expenses` -> row `24`
- `Net income` -> row `28`
- `Net income margin` -> row `29`
- Assumption block -> rows `32:41`

## Forecast formula note

The forecast `Operating profit` row at `19` must include:

- `Gross profit`
- `Business tax and surcharges`
- `Sales & marketing expenses`
- `General & administration expenses`
- `Research expenses`
- `Other operating expenses`

In the current template, the forecast formula should therefore include row `8` in addition to rows `5`, `10`, `12`, `14`, and `16`.

## Metric keys

Use this normalized schema when filling the workbook:

```json
{
  "historical_years": [2021, 2022, 2023, 2024, 2025],
  "historical_financials": {
    "revenue": [0, 0, 0, 0, 0],
    "gross_profit": [0, 0, 0, 0, 0],
    "business_tax_surcharges": [0, 0, 0, 0, 0],
    "sales_marketing_expenses": [0, 0, 0, 0, 0],
    "general_admin_expenses": [0, 0, 0, 0, 0],
    "research_expenses": [0, 0, 0, 0, 0],
    "operating_profit": [0, 0, 0, 0, 0],
    "net_interest_expenses": [0, 0, 0, 0, 0],
    "profit_before_tax": [0, 0, 0, 0, 0],
    "tax_expenses": [0, 0, 0, 0, 0]
  }
}
```

## Missing-data rule

If a source field is unavailable:

- Leave the blue cell blank rather than inventing a value.
- Record the missing metric in the output summary.
- Do not overwrite formula-driven rows to force the model to balance.

## Tax sign rule

For the historical `Tax expenses` row at `27`:

- store tax expense as a negative template value
- store tax benefit or tax refund as a positive template value
- when the source field is `INCOME_TAX`, map it as `-INCOME_TAX`

Do not force the sign with `abs()`, because that would erase genuine tax refunds.

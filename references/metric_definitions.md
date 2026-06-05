# Metric Definitions

Use these definitions consistently in the scaffold.

## Core metrics

- `revenue`: Operating revenue used as the top line for forecasting.
- `gross_profit`: Revenue minus cost of goods sold.
- `gross_margin`: `gross_profit / revenue`.
- `operating_profit`: Profit after operating expenses and before non-operating items.
- `net_profit`: Net profit for the period. Do not assume this equals parent net profit.
- `parent_net_profit`: Net profit attributable to the parent company.
- `shares_outstanding`: Share count used for simple EPS estimation.
- `tax_expenses`: Template tax row. Store tax expense as a negative value and tax benefit / tax refund as a positive value.

## Field handling rules

- Keep `net_profit` and `parent_net_profit` separate.
- If only one of the two is available, retain the original label in the normalized data.
- Do not map financial-company metrics into this general template without explicit user approval.
- For `INCOME_TAX`, preserve the economics of the source sign:
  - positive source value = tax expense
  - negative source value = tax benefit or tax refund
  - template value must be `-INCOME_TAX`

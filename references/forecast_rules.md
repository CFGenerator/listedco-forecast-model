# Forecast Rules

## Objective

Produce a lightweight forecast that is auditable and easy to modify.

## Default approach

Start with a P&L-driven model:

1. Forecast revenue growth.
2. Forecast gross margin.
3. Forecast sales, admin, and R&D expense ratios.
4. Derive operating profit and net profit.
5. Estimate EPS if share count is available.

## Guardrails

- Prefer explicit user assumptions to inferred assumptions.
- Flag one-off gains and losses instead of rolling them into steady-state margins.
- Avoid averaging through obvious cyclical peaks and troughs without a note.
- Stop and explain the limitation for banks, insurers, brokers, and other financial companies.
- Keep a `base`, `bull`, and `bear` scenario whenever assumptions are uncertain.

## Escalation path

Extend only when requested:

- Add working capital and capex for a simple cash-flow layer.
- Add balance-sheet drivers for a partial three-statement model.
- Add valuation outputs such as PE or DCF only after the operating forecast is stable.

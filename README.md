# E-commerce Settlement Reconciliation & True-Profit Calculator

A Python tool that reconciles Amazon and Shopify settlement payouts against
the store's own fee schedule — order by order — and works out the store's
**true** profit after COGS, platform fees, refunds, ad spend and shipping.
Every payout that is missing, duplicated, overcharged or miscalculated is
itemised in an exceptions report with its dollar impact.

Built by **Waqar Ahmed**, spreadsheet & data specialist (7 years' experience:
data cleaning, reconciliation, B2B lead enrichment, Excel / Google Sheets
dashboards and automation).

## Who it is for

E-commerce sellers on Amazon, Shopify or both who export their orders and
settlement reports and want to know:

- what each SKU, channel and country actually earned after *every* cost;
- whether the platforms paid what the fee schedule says they should have;
- exactly which orders leaked money, and how much.

No marketplace account access is needed — the tool works entirely from
exported CSV files.

## What it produces

One Excel workbook with five sheets:

1. **Profit by SKU** — revenue, refunds, COGS, platform fees, allocated ad
   spend and shipping per product, down to true net profit per SKU.
2. **Profit by Channel & Country** — the same view split by channel × country.
3. **Settlement Reconciliation** — every order: expected payout (computed from
   the fee schedule) vs actual payout (from the settlement files), classified
   Matched / Underpaid / Overpaid / Missing / Duplicate.
4. **Exceptions Report** — every discrepancy itemised with type, order,
   expected vs actual amounts, money impact and a plain-English explanation.
5. **Summary** — one-page overview with totals, net leakage and method notes.

## Demo result (included sample data)

The sample data is 100% synthetic: a fictional store ("Northwind Home Goods"),
October 2026, 120 orders across Amazon and Shopify, with 14 realistic payout
errors deliberately planted in the settlement files. The tool found all 14.

| Measure | Value |
|---|---|
| Orders reconciled | 120 (189 units) |
| Gross revenue | $9,824.97 |
| True net profit (after all costs) | $313.68 |
| Expected payouts per fee schedule | $7,897.58 |
| Actually paid by the platforms | $7,628.69 |
| **Net leakage** | **$268.89** |
| Exceptions found | 14 of 14 planted |
| Gross value of all discrepancies | $589.25 |

Exceptions caught: 5 payouts missing from settlement, 4 fee overcharges above
the fee schedule, 2 duplicate payouts, 2 refunds deducted twice, 1 wrong
payout amount.

## Run it

Requirements: Python 3.10+ and `openpyxl`.

```bash
pip install -r requirements.txt
python3 demo.py
```

`demo.py` runs the engine and then the independent validator. Real console
output from the included sample data:

```text
===== python3 reconcile.py =====
workbook saved: output/Ecommerce-Profit-Settlement-Reconciliation.xlsx
revenue=9824.97 net_profit=313.68 exceptions=14 abs_impact=589.25 expected=7897.58 actual=7628.69 leakage=268.89

===== python3 validate.py =====
...
PASS  Planted exceptions == found exceptions  [planted=14 found=14]

32/32 checks passed
ALL CHECKS PASSED
```

Or run the steps individually:

```bash
python3 reconcile.py     # builds output/Ecommerce-Profit-Settlement-Reconciliation.xlsx
python3 validate.py      # independently re-verifies every figure (32 checks)
python3 generate_sample_data.py   # regenerates the synthetic sample data (seeded, deterministic)
```

## Repository layout

| File | Purpose |
|---|---|
| `reconcile.py` | The reconciliation engine. Reads the CSVs, computes expected payouts from the fee schedule, classifies every order, builds the workbook. |
| `validate.py` | Independent checker. Shares no code with the engine; recomputes everything from the raw CSVs and asserts the workbook ties out. |
| `generate_sample_data.py` | Creates the synthetic sample dataset, including the 14 planted payout errors (recorded in `sample_data/plants.json`). |
| `demo.py` | One-command demo: engine + validator. |
| `sample_data/` | Input CSVs (`orders.csv`, `settlement_amazon.csv`, `shopify_payouts.csv`, `ad_spend.csv`, `shipping_costs.csv`), the `fee_schedule.md` rate card, and the expected results (`expected_output.xlsx`, `expected_results.json`). |
| `output/` | Where a fresh run writes the workbook and `computed.json`. |

## Method notes (the details are the point)

- All money maths uses `Decimal`, computed at order level and rounded to
  cents, so totals tie out exactly across every sheet — no float drift.
- Ad spend is allocated to orders and SKUs by gross-revenue share within each
  channel using the largest-remainder method, exact to the cent.
- COGS is counted only on non-refunded orders (refunded goods are assumed
  returned to inventory); shipping is charged on all orders.
- Fee checks are made against the fee schedule in `sample_data/fee_schedule.md`;
  in a client engagement that is the client's own confirmed rate card.
- The validator is deliberately independent: if the engine and the validator
  disagree, the run fails loudly instead of shipping a wrong number.

## Scope and honesty notes

- Works from exported CSV/report files only; no account access is requested.
- Tax/VAT treatment and accounting judgements stay with the client's
  accountant — this tool reconciles payouts and profit, it does not file
  anything.
- Figures are only as complete as the exports provided; missing source files
  are reported as gaps, never guessed.
- All sample data in this repository is fictional and generated for
  demonstration. No real client data is included.

## About

Waqar Ahmed — spreadsheet & data specialist. Services: e-commerce and invoice
reconciliation, data cleaning, CRM cleanup, B2B lead lists and enrichment,
Excel / Google Sheets dashboards and automation. Written communication only.

- Contra: https://contra.com/mr_world_z7lsrx86
- LinkedIn: https://www.linkedin.com/in/waqar-ahmed-4347a631b/

## License

MIT — see `LICENSE`.

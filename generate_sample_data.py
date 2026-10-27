#!/usr/bin/env python3
"""Generate 100% fictional input data for the E-commerce Profit & Settlement
Reconciliation sample: Northwind Home Goods, October 2026, Amazon + Shopify.

Planted discrepancies (recorded in plants.json):
  - 5 orders missing from settlement (3 Amazon, 2 Shopify)
  - 4 fee overcharges vs the fee schedule (2 Amazon @18% referral, 2 Shopify @3.9%)
  - 2 duplicate payouts (1 per channel)
  - 2 refunded orders with the refund deducted twice (1 per channel)
  - 1 payout in a wrong amount (Shopify, $13.47 short)
Total planted exceptions: 14.
"""
import csv, json, random
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

BASE = Path(__file__).parent
IN = BASE / "sample_data"
IN.mkdir(parents=True, exist_ok=True)

CENT = Decimal("0.01")
def q(x):  # quantize to cents
    return Decimal(x).quantize(CENT, rounding=ROUND_HALF_UP)
def f2(d):  # Decimal -> string with 2dp for CSV
    return f"{q(d):.2f}"

rng = random.Random(20261025)

PRODUCTS = [
    # sku, name, unit price USD, unit COGS USD, weight
    ("NHG-001", "Ceramic Vase Set",      "49.99", "18.50", 12),
    ("NHG-002", "Bamboo Bath Tray",      "34.95", "12.25", 10),
    ("NHG-003", "Linen Throw Blanket",   "59.00", "22.00", 11),
    ("NHG-004", "Stoneware Dinner Set",  "89.99", "38.75",  9),
    ("NHG-005", "Scented Candle Trio",   "27.50",  "9.80", 13),
    ("NHG-006", "Wall Mirror Round",     "74.25", "30.40",  8),
    ("NHG-007", "Cotton Sheet Set",      "65.00", "26.10", 10),
    ("NHG-008", "Brass Table Lamp",      "92.00", "41.60",  7),
    ("NHG-009", "Storage Basket Set",    "39.99", "14.90", 11),
    ("NHG-010", "Ceramic Planter Duo",   "32.75", "11.35",  9),
]
PRICE = {p[0]: Decimal(p[2]) for p in PRODUCTS}

# ---------------------------------------------------------------- orders
orders = []
skus = [p[0] for p in PRODUCTS]
sku_w = [p[4] for p in PRODUCTS]
for i in range(120):
    oid = f"ORD-{1001 + i}"
    day = 1 + rng.randrange(31)
    channel = "Amazon" if rng.random() < 0.55 else "Shopify"
    country = rng.choices(["US", "CA", "UK", "DE"], weights=[55, 15, 18, 12])[0]
    sku = rng.choices(skus, weights=sku_w)[0]
    qty = rng.choices([1, 2, 3], weights=[55, 32, 13])[0]
    orders.append({
        "order_id": oid, "date": f"2026-10-{day:02d}", "channel": channel,
        "sku": sku, "country": country, "quantity": qty, "status": "Completed",
    })

def pick(pool, n, used):
    """Pick n orders from pool whose ids are not in `used`; mark them used."""
    cands = [o for o in pool if o["order_id"] not in used]
    chosen = rng.sample(cands, n)
    for o in chosen:
        used.add(o["order_id"])
    return chosen

used = set()
amz = [o for o in orders if o["channel"] == "Amazon"]
shp = [o for o in orders if o["channel"] == "Shopify"]

# refunds: 4 Amazon + 3 Shopify (kept disjoint from other plants)
refunded = pick(amz, 4, used) + pick(shp, 3, used)
for o in refunded:
    o["status"] = "Refunded"

# plant selection (all Completed orders, disjoint)
plants = []
missing      = pick(amz, 3, used) + pick(shp, 2, used)
overcharge   = pick(amz, 2, used) + pick(shp, 2, used)
duplicate    = pick(amz, 1, used) + pick(shp, 1, used)
wrong_amount = pick(shp, 1, used)
dbl_refund   = [o for o in refunded if o["channel"] == "Amazon"][:1] + \
               [o for o in refunded if o["channel"] == "Shopify"][:1]

for o in missing:      plants.append({"type": "Missing from settlement", "order_id": o["order_id"], "channel": o["channel"]})
for o in overcharge:   plants.append({"type": "Fee overcharge", "order_id": o["order_id"], "channel": o["channel"],
                                      "detail": "Amazon referral charged at 18% instead of 15%" if o["channel"] == "Amazon"
                                                else "Shopify processing charged at 3.9% instead of 2.9%"})
for o in duplicate:    plants.append({"type": "Duplicate payout", "order_id": o["order_id"], "channel": o["channel"]})
for o in wrong_amount: plants.append({"type": "Wrong payout amount", "order_id": o["order_id"], "channel": o["channel"],
                                      "detail": "Payout $13.47 less than the settlement arithmetic"})
for o in dbl_refund:   plants.append({"type": "Refund deducted twice", "order_id": o["order_id"], "channel": o["channel"]})

missing_ids = {o["order_id"] for o in missing}
overcharge_ids = {o["order_id"] for o in overcharge}
duplicate_ids = {o["order_id"] for o in duplicate}
wrong_ids = {o["order_id"] for o in wrong_amount}
dbl_ids = {o["order_id"] for o in dbl_refund}

def revenue(o):
    return q(PRICE[o["sku"]] * o["quantity"])

def expected_fees(o):
    rev = revenue(o)
    if o["channel"] == "Amazon":
        if o["status"] == "Refunded":   # 20% of the 15% referral kept as admin fee + fulfilment kept
            return q(rev * Decimal("0.03")) + Decimal("4.25")
        return q(rev * Decimal("0.15")) + Decimal("4.25")
    return q(rev * Decimal("0.029")) + Decimal("0.30")   # Shopify: processing never refunded

# ---------------------------------------------------------------- settlements
amz_rows, shp_rows = [], []
for o in orders:
    if o["order_id"] in missing_ids:
        continue
    rev = revenue(o)
    is_ref = o["status"] == "Refunded"
    refund_ded = (rev * 2) if o["order_id"] in dbl_ids else (rev if is_ref else Decimal("0"))
    sdate = f"2026-10-{min(31, int(o['date'][-2:]) + 4):02d}" if int(o['date'][-2:]) <= 27 else "2026-11-04"
    if o["channel"] == "Amazon":
        fee = expected_fees(o)
        referral = fee - Decimal("4.25"); fulfil = Decimal("4.25")
        if o["order_id"] in overcharge_ids:
            referral = q(rev * Decimal("0.18")) if not is_ref else q(rev * Decimal("0.06"))
            fee = referral + fulfil
        payout = q(rev - fee - refund_ded)
        row = {"order_id": o["order_id"], "settlement_date": sdate,
               "gross_sales": f2(rev), "referral_fee": f2(referral),
               "fulfillment_fee": f2(fulfil), "refund_deduction": f2(refund_ded),
               "payout_amount": f2(payout)}
        amz_rows.append(row)
        if o["order_id"] in duplicate_ids:
            amz_rows.append(dict(row))
    else:
        fee = expected_fees(o)
        if o["order_id"] in overcharge_ids:
            fee = q(rev * Decimal("0.039")) + Decimal("0.30")
        payout = q(rev - fee - refund_ded)
        if o["order_id"] in wrong_ids:
            payout = q(payout - Decimal("13.47"))
        week = (int(o["date"][-2:]) - 1) // 7 + 1
        row = {"order_id": o["order_id"], "payout_batch": f"PO-2026-W{week}",
               "payout_date": sdate, "gross_sales": f2(rev),
               "processing_fee": f2(fee), "refund_deduction": f2(refund_ded),
               "payout_amount": f2(payout)}
        shp_rows.append(row)
        if o["order_id"] in duplicate_ids:
            shp_rows.append(dict(row))

# ---------------------------------------------------------------- ad spend
ad_rows = []
for day in range(1, 32):
    ad_rows.append({"date": f"2026-10-{day:02d}", "channel": "Amazon",
                    "ad_spend": f2(Decimal(str(round(rng.uniform(36, 74), 2))))})
    ad_rows.append({"date": f"2026-10-{day:02d}", "channel": "Shopify",
                    "ad_spend": f2(Decimal(str(round(rng.uniform(20, 49), 2))))})

# ---------------------------------------------------------------- shipping
BASE_SHIP = {"US": Decimal("6.95"), "CA": Decimal("9.50"), "UK": Decimal("11.25"), "DE": Decimal("10.75")}
ship_rows = []
for o in orders:
    cost = BASE_SHIP[o["country"]] + Decimal("1.50") * (o["quantity"] - 1) \
           + Decimal(str(round(rng.uniform(-0.60, 0.60), 2)))
    ship_rows.append({"order_id": o["order_id"], "shipping_cost": f2(cost)})

# ---------------------------------------------------------------- write files
def write_csv(name, rows, fields):
    with open(IN / name, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=fields)
        w.writeheader(); w.writerows(rows)

write_csv("orders.csv", [
    {"order_id": o["order_id"], "date": o["date"], "channel": o["channel"], "sku": o["sku"],
     "country": o["country"], "quantity": o["quantity"],
     "unit_price": f2(PRICE[o["sku"]]), "currency": "USD", "order_status": o["status"]}
    for o in orders],
    ["order_id", "date", "channel", "sku", "country", "quantity", "unit_price", "currency", "order_status"])
write_csv("settlement_amazon.csv", amz_rows,
          ["order_id", "settlement_date", "gross_sales", "referral_fee", "fulfillment_fee", "refund_deduction", "payout_amount"])
write_csv("shopify_payouts.csv", shp_rows,
          ["order_id", "payout_batch", "payout_date", "gross_sales", "processing_fee", "refund_deduction", "payout_amount"])
write_csv("ad_spend.csv", ad_rows, ["date", "channel", "ad_spend"])
write_csv("shipping_costs.csv", ship_rows, ["order_id", "shipping_cost"])

(IN / "fee_schedule.md").write_text("""# Northwind Home Goods — Platform Fee Schedule (the rules reconciliation checks against)

All amounts USD. Fees are charged per order.

## Amazon
- Referral fee: **15% of gross item revenue**.
- Fulfilment fee: **$4.25 flat per order** (charged on the sale; NOT returned if the order is later refunded).
- Refunded order: customer is refunded 100% of revenue. Amazon returns the referral fee
  but keeps **20% of the referral fee (= 3% of revenue) as a refund administration fee**.
  Expected net payout for a refunded order = -(3% of revenue) - $4.25.

## Shopify
- Payment processing fee: **2.9% of gross revenue + $0.30 per order**.
- Processing fees are **never returned** on refunds.
- Refunded order: 100% of revenue is deducted back.
  Expected net payout for a refunded order = -(2.9% of revenue + $0.30).

## Clean settlement arithmetic (every clean row must satisfy this)
`payout_amount = gross_sales - platform_fees - refund_deduction`
where refund_deduction = full revenue for refunded orders, otherwise 0.
Each order must appear in settlement **exactly once**, with fees exactly per the schedule above.
""")

(IN / "plants.json").write_text(json.dumps(
    {"store": "Northwind Home Goods", "period": "October 2026",
     "planted_exceptions": plants, "planted_count": len(plants)}, indent=2))

print(f"orders={len(orders)} amazon_settlement_rows={len(amz_rows)} shopify_rows={len(shp_rows)} "
      f"refunded={len(refunded)} planted={len(plants)}")
for p in plants:
    print(" PLANT:", p["type"], p["order_id"], p["channel"])

# Northwind Home Goods — Platform Fee Schedule (the rules reconciliation checks against)

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

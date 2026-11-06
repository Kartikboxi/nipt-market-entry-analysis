-- Q3. How much disclosed capital has gone into acquisitions and funding over time?
-- INR amount is used where reported; otherwise USD is converted at 95 INR/USD (1 USD m = Rs 9.5 cr).
-- A running total (window SUM) shows the cumulative build-up.

WITH money AS (
    SELECT
        event_date,
        company,
        event_type,
        counterparty,
        COALESCE(amount_inr_cr, amount_usd_m * 9.5) AS amount_cr
    FROM industry_events
    WHERE event_type IN ('Acquisition', 'Funding')
      AND (amount_inr_cr IS NOT NULL OR amount_usd_m IS NOT NULL)
)
SELECT
    event_date,
    company,
    event_type,
    counterparty,
    ROUND(amount_cr, 1)                                              AS amount_cr,
    ROUND(SUM(amount_cr) OVER (ORDER BY event_date
                               ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW), 1) AS cumulative_cr
FROM money
ORDER BY event_date;

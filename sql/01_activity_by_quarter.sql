-- Q1. How is deal activity trending? Events per calendar quarter, split by type.
-- Conditional aggregation (SUM of CASE) turns event types into columns.

SELECT
    strftime('%Y', event_date) || '-Q' || ((CAST(strftime('%m', event_date) AS INTEGER) + 2) / 3) AS quarter,
    SUM(CASE WHEN event_type = 'Acquisition' THEN 1 ELSE 0 END)                        AS acquisitions,
    SUM(CASE WHEN event_type = 'Funding' THEN 1 ELSE 0 END)                            AS funding_rounds,
    SUM(CASE WHEN event_type IN ('Joint venture', 'Partnership') THEN 1 ELSE 0 END)    AS jvs_partnerships,
    SUM(CASE WHEN event_type IN ('Product launch', 'Capability', 'Results') THEN 1 ELSE 0 END) AS other,
    COUNT(*)                                                                           AS total_events
FROM industry_events
GROUP BY quarter
ORDER BY quarter;

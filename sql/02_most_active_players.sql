-- Q2. Who is consolidating the market? Inorganic moves per company, with disclosed value.

SELECT
    e.company,
    c.segment,
    COUNT(*)                                                                AS inorganic_moves,
    SUM(CASE WHEN e.event_type = 'Acquisition' THEN 1 ELSE 0 END)           AS acquisitions,
    SUM(CASE WHEN e.event_type IN ('Joint venture', 'Partnership') THEN 1 ELSE 0 END) AS jvs_partnerships,
    COALESCE(SUM(CASE WHEN e.event_type = 'Acquisition' THEN e.amount_inr_cr END), 0) AS disclosed_deal_value_cr,
    MAX(e.event_date)                                                       AS latest_move,
    DENSE_RANK() OVER (ORDER BY COUNT(*) DESC)                              AS activity_rank
FROM industry_events e
JOIN companies c ON c.company = e.company
WHERE e.event_type IN ('Acquisition', 'Joint venture', 'Partnership')
GROUP BY e.company, c.segment
ORDER BY inorganic_moves DESC, disclosed_deal_value_cr DESC;

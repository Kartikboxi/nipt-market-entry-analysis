-- Q5. Are genomics specialists behaving differently from pathology chains?
-- Compares the mix of moves by segment and the share each segment holds.

SELECT
    c.segment,
    COUNT(*)                                                         AS events,
    ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 1)               AS share_of_all_events_pct,
    SUM(CASE WHEN e.event_type = 'Acquisition' THEN 1 ELSE 0 END)    AS acquisitions,
    SUM(CASE WHEN e.event_type = 'Funding' THEN 1 ELSE 0 END)        AS funding_rounds,
    SUM(CASE WHEN e.event_type = 'Product launch' OR e.event_type = 'Capability' THEN 1 ELSE 0 END) AS capability_moves,
    COUNT(DISTINCT e.company)                                        AS active_companies
FROM industry_events e
JOIN companies c ON c.company = e.company
GROUP BY c.segment
ORDER BY events DESC;

-- Q4. What happened in the 90 days before the report date? Feeds the monthly management update.
-- Report date is fixed at 2026-10-31 so the output is reproducible; change it each month.

WITH params AS (SELECT DATE('2026-10-31') AS report_date)
SELECT
    e.event_date,
    e.company,
    c.segment,
    e.event_type,
    e.counterparty,
    e.geography,
    e.summary
FROM industry_events e
JOIN companies c ON c.company = e.company
CROSS JOIN params p
WHERE e.event_date >  DATE(p.report_date, '-90 days')
  AND e.event_date <= p.report_date
ORDER BY e.event_date DESC;

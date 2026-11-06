-- Competitor and industry event tracker for Indian diagnostics and genomics.

DROP TABLE IF EXISTS companies;
DROP TABLE IF EXISTS industry_events;

CREATE TABLE companies (
    company  TEXT PRIMARY KEY,
    segment  TEXT NOT NULL,          -- 'Pathology chain' or 'Genomics specialist'
    listed   TEXT NOT NULL           -- 'Y' / 'N'
);

CREATE TABLE industry_events (
    event_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    event_date      TEXT NOT NULL,   -- ISO date, YYYY-MM-DD
    company         TEXT NOT NULL REFERENCES companies(company),
    event_type      TEXT NOT NULL,   -- Acquisition, Funding, Joint venture, Partnership, Product launch, Capability, Results
    counterparty    TEXT,
    amount_inr_cr   REAL,            -- NULL when not disclosed
    amount_usd_m    REAL,
    geography       TEXT,
    summary         TEXT,
    source_url      TEXT
);

CREATE INDEX idx_events_date ON industry_events(event_date);
CREATE INDEX idx_events_company ON industry_events(company);

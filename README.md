# NIPT in India: Market Sizing, Entry Strategy & Competitor Tracker

A market-entry assessment of India's **non-invasive prenatal testing (NIPT)** market, built in **Excel** and **SQL**. NIPT is a blood test that screens a pregnancy for chromosomal conditions such as Down syndrome.

The question it answers: **how big is the NIPT market, how fast will it grow, and how should a diagnostics company without its own genomics lab enter it: build a lab, partner with a reference lab, or acquire one?**

It also includes a competitor and deal tracker for Indian diagnostics and genomics, which feeds an auto-generated monthly update for management.

## What's inside

| Part | What it does |
|---|---|
| `data/` | Sourced inputs: births, delivery mix, NIPT prices from 7 listings, published market estimates, and 20 industry events (2025-26) |
| `outputs/NIPT_Market_Entry_Model.xlsx` | Excel model: scenarios, market sizing, 5-year forecast, investment case, entry scorecard, event tracker |
| `sql/` | SQLite schema and 5 tracker queries (conditional aggregation, window functions, CTEs, date filters) |
| `scripts/run_sql.py` | Loads the tracker into SQLite and runs the queries |
| `scripts/build_model.py` | Builds the Excel model with live formulas |
| `scripts/generate_reports.py` | Writes the monthly management update and a 6-slide deck from the SQL results and the model |
| `outputs/monthly_update_2026-10.md` | October 2026 industry update |
| `outputs/NIPT_Entry_Management_Deck.pptx` | Management summary deck |

### The Excel model

- **Scenarios**: Low / Base / High assumptions with a drop-down. Every sheet reads the active case through `INDEX/MATCH`.
- **Market_Sizing**: a bottom-up funnel (births, then facility births, then private-facility births, then families who can afford the test, then uptake), TAM / SAM, and a cross-check that turns each published estimate into the share of births it implies are tested
- **Forecast**: 2026-2031 tests, price per test, market value and a new entrant's share ramp
- **Investment_Case**: build-own-lab vs partner-with-reference-lab cash flows, with `NPV`, `IRR`, break-even volume and the crossover volume where building starts to beat partnering
- **Entry_Options**: weighted scorecard for Build, Partner and Acquire
- **Events**: the deal tracker with `COUNTIFS` / `SUMIFS` summaries

## Key findings

**1. Today's market is about Rs 940 cr (base case), and only 2.7% of births are tested.**

| Estimate | Market (Rs cr) | Implied share of births tested |
|---|---|---|
| This model: Low | 419 | 1.4% |
| This model: Base | 938 | 2.7% |
| This model: High | 1,874 | 5.0% |
| MarketsandMarkets (2024) | 1,800 | 5.1% |
| DPI Research (2028 forecast) | 589 | 1.7% |
| MedGenome estimate (via The Ken) | 400 | 1.1% |

Published estimates differ by 4-5x. Turning each into implied penetration shows they line up with the Low and High ends of the bottom-up range, so the disagreement is really about how many pregnancies are tested.

**2. Volume grows about 12% a year but value only about 8%**, because realised prices fall about 4% a year. The base case reaches Rs 1,374 cr by 2031. All the growth comes from higher uptake, since births are flat to slightly falling.

**3. Building a lab pays only if volume passes about 17,000 tests a year.**

| Scenario | NPV: build own lab (Rs cr) | NPV: partner (Rs cr) |
|---|---|---|
| Low | -24.5 | -4.5 |
| Base | 12.2 | 8.6 |
| High | 137.3 | 70.2 |

Building has the higher NPV in the Base and High cases, but it loses the most if adoption is slow. In the base case an entrant passes the crossover volume in 2029.

**4. Recommendation: enter through a reference-lab partner now** (about Rs 1 cr at risk, live in weeks). Build or buy a lab only once run-rate volume passes about 17,000 tests a year. This keeps most of the upside while capping the downside.

**5. Competitors are buying capability rather than building it.** The tracker logs 13 moves in 2026 so far vs 7 in all of 2025. Neuberg has done 4 JVs and partnerships, Metropolis bought an oncology and genomics lab (Core Diagnostics, Rs 247 cr), and MedGenome has made 3 bolt-on acquisitions since November 2025.

## How to run

```bash
pip install -r requirements.txt
python scripts/run_sql.py           # tracker database + query results
python scripts/build_model.py       # Excel model (LibreOffice optional, used to store calculated values)
python scripts/generate_reports.py  # monthly update + deck (needs the two steps above)
```

## Assumptions and limitations

- **Sourced:** 2024 births (Civil Registration System), institutional delivery share (NFHS-6), private share of facility births (NFHS-5), NIPT prices (7 provider listings, Nov 2026), USD/INR of about 95.
- **My assumptions (in yellow, editable):** affordability share, uptake and its growth, price decline, entrant share, and every cost in the investment case (Rs 15 cr lab capex, Rs 6,000 variable cost per test, Rs 9,500 reference-lab price). These are illustrative and would need vendor quotes and expert calls before any real decision.
- USD estimates are converted at today's rate, which slightly overstates older USD figures in rupee terms.
- The tracker covers 20 publicly reported events. It is a sample of notable moves, not a full census of deals.
- NPVs cover 2026-31 only, with no terminal value, so they understate the long-run value of both options equally.

## Sources

Births: [CRS 2024 report coverage](https://www.outlookindia.com/national/india-registered-over-25-crore-births-in-2024-what-the-civil-registration-system-report-reveals) · Institutional births: [NFHS-6 coverage](https://www.dtnext.in/news/national/institutional-deliveries-in-india-up-from-886-in-2019-2020-to-906-in-2023-24-survey) · Private share: [CEDA Ashoka on NFHS-5](https://ceda.ashoka.edu.in/?p=5590) · Market estimates: [MarketsandMarkets](https://www.marketsandmarkets.com/Market-Reports/geography/non-invasive-prenatal-testing-market/India), [DPI Research](https://pdf.marketpublishers.com/dpi/india-non-invasive-prenatal-testing-nipt-market-n-forecast-2022-2028.pdf), [The Ken](https://the-ken.com/story/nipt-is-expensive-but-india-cant-afford-to-miss-out/) · Prices and events: links in `data/`

*Self-directed learning project. Not medical or investment advice.*

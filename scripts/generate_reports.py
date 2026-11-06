"""Generate the monthly industry update (Markdown) and the management deck (PowerPoint).

Reads:  outputs/tracker.db                     (run scripts/run_sql.py first)
        outputs/NIPT_Market_Entry_Model.xlsx   (run scripts/build_model.py first)
Writes: outputs/monthly_update_<YYYY-MM>.md
        outputs/NIPT_Entry_Management_Deck.pptx

Usage: python scripts/generate_reports.py [--report-date 2026-10-31]
"""
import argparse
import shutil
import sqlite3
import subprocess
import tempfile
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs"
DB = OUT / "tracker.db"
MODEL = OUT / "NIPT_Market_Entry_Model.xlsx"

NAVY = RGBColor(0x1F, 0x4E, 0x78)
GREY = RGBColor(0x59, 0x59, 0x59)


# ------------------------------------------------------------------ data
def query(conn, sql, params=()):
    cur = conn.execute(sql, params)
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def tracker_facts(report_date):
    conn = sqlite3.connect(DB)
    rd = report_date.isoformat()
    f = {}
    f["last90"] = query(conn, """
        SELECT e.event_date, e.company, c.segment, e.event_type, e.counterparty, e.geography, e.summary
        FROM industry_events e JOIN companies c ON c.company = e.company
        WHERE e.event_date > DATE(?, '-90 days') AND e.event_date <= ?
        ORDER BY e.event_date DESC""", (rd, rd))
    f["last365"] = query(conn, """
        SELECT event_type, COUNT(*) AS n FROM industry_events
        WHERE event_date > DATE(?, '-365 days') AND event_date <= ?
        GROUP BY event_type ORDER BY n DESC""", (rd, rd))
    f["by_year"] = query(conn, """
        SELECT strftime('%Y', event_date) AS yr,
               SUM(event_type = 'Acquisition') AS acquisitions,
               SUM(event_type IN ('Joint venture', 'Partnership')) AS jvs,
               COUNT(*) AS total
        FROM industry_events WHERE event_date <= ? GROUP BY yr ORDER BY yr""", (rd,))
    f["scoreboard"] = query(conn, """
        SELECT e.company, c.segment,
               SUM(e.event_type = 'Acquisition') AS acquisitions,
               SUM(e.event_type IN ('Joint venture', 'Partnership')) AS jvs,
               COALESCE(SUM(CASE WHEN e.event_type = 'Acquisition' THEN e.amount_inr_cr END), 0) AS value_cr,
               MAX(e.event_date) AS latest
        FROM industry_events e JOIN companies c ON c.company = e.company
        WHERE e.event_type IN ('Acquisition', 'Joint venture', 'Partnership') AND e.event_date <= ?
        GROUP BY e.company ORDER BY acquisitions + jvs DESC, value_cr DESC""", (rd,))
    f["capital"] = query(conn, """
        SELECT ROUND(SUM(COALESCE(amount_inr_cr, amount_usd_m * 9.5)), 0) AS total_cr
        FROM industry_events
        WHERE event_type IN ('Acquisition', 'Funding') AND event_date <= ?""", (rd,))[0]["total_cr"]
    conn.close()
    return f


def model_values(path, scenario=None):
    """Read key outputs from the model. If a scenario is given, switch it and recalculate with LibreOffice."""
    if scenario:
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            return None
        tmp = Path(tempfile.mkdtemp())
        wb = load_workbook(path)
        wb["Scenarios"]["C3"] = scenario
        src = tmp / "scenario.xlsx"
        wb.save(src)
        (tmp / "out").mkdir()
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(tmp / "out"), str(src)],
                       check=True, capture_output=True)
        path = tmp / "out" / "scenario.xlsx"
    wb = load_workbook(path, data_only=True)
    ms, fc, ic = wb["Market_Sizing"], wb["Forecast"], wb["Investment_Case"]
    return {
        "scenario": wb["Scenarios"]["C3"].value,
        "births": ms["C5"].value, "tests": ms["C12"].value, "penetration": ms["C13"].value,
        "price": ms["C14"].value, "market_cr": ms["C15"].value, "market_usd": ms["C16"].value,
        "sam_cr": ms["C22"].value,
        "years": [fc[f"{c}4"].value for c in "CDEFGH"],
        "market_series": [fc[f"{c}13"].value for c in "CDEFGH"],
        "tests_series": [fc[f"{c}11"].value for c in "CDEFGH"],
        "value_cagr": fc["C19"].value, "volume_cagr": fc["C20"].value,
        "entrant_tests_2031": fc["H16"].value, "entrant_rev_2031": fc["H17"].value,
        "npv_build": ic["C46"].value, "npv_partner": ic["D46"].value,
        "crossover": ic["C52"].value, "crossover_year": ic["C53"].value, "recommendation": ic["C54"].value,
        "published": [(ms[f"A{r}"].value, ms[f"B{r}"].value, ms[f"C{r}"].value, ms[f"E{r}"].value) for r in (31, 32, 33)],
    }


def plural(n, word, plural_word=None):
    return f"{n} {word if n == 1 else (plural_word or word + 's')}"


def lakh(n):
    return f"{n / 1e5:,.1f} lakh"


# --------------------------------------------------------------- markdown
def write_update(report_date, t, m):
    month = report_date.strftime("%B %Y")
    yrs = {r["yr"]: r for r in t["by_year"]}
    this_year, last_year = str(report_date.year), str(report_date.year - 1)
    top = t["scoreboard"][0]
    lines = [
        f"# Diagnostics & Genomics Industry Update: {month}",
        "",
        f"*Prepared for senior management. Data to {report_date:%d %b %Y}. Sources linked in `data/industry_events.csv`.*",
        "",
        "## Headlines",
        "",
        f"- **{len(t['last90'])} notable moves in the last 90 days**, led by {top['company']} "
        f"({plural(top['acquisitions'], 'acquisition')} and {plural(top['jvs'], 'JV or partnership', 'JVs or partnerships')} "
        f"since Jan-2025).",
        f"- **Deal flow is picking up:** the tracker logs {yrs.get(this_year, {}).get('total', 0)} events in {this_year} "
        f"so far vs {yrs.get(last_year, {}).get('total', 0)} in all of {last_year} (tracked sample, not a full census); "
        f"{yrs.get(this_year, {}).get('acquisitions', 0)} acquisitions this year.",
        f"- **About Rs {t['capital']:,.0f} cr of disclosed capital** has gone into tracked acquisitions and funding rounds since Jan-2025.",
        f"- **NIPT market (base case):** about Rs {m['market_cr']:,.0f} cr in 2026 ({lakh(m['tests'])} tests, "
        f"{m['penetration']:.1%} of births), growing about {m['value_cagr']:.0%} a year in value and "
        f"{m['volume_cagr']:.0%} in volume to 2031.",
        "",
        "## Moves in the last 90 days",
        "",
        "| Date | Company | Type | Counterparty | Where | What happened |",
        "|---|---|---|---|---|---|",
    ]
    for e in t["last90"]:
        lines.append(f"| {e['event_date']} | {e['company']} | {e['event_type']} | {e['counterparty'] or ''} | "
                     f"{e['geography']} | {e['summary']} |")
    lines += [
        "",
        "## Consolidation scoreboard (Jan-2025 to date)",
        "",
        "| Company | Segment | Acquisitions | JVs / partnerships | Disclosed deal value (Rs cr) | Latest move |",
        "|---|---|---|---|---|---|",
    ]
    for s in t["scoreboard"]:
        val = f"{s['value_cr']:,.0f}" if s["value_cr"] else "not disclosed"
        lines.append(f"| {s['company']} | {s['segment']} | {s['acquisitions']} | {s['jvs']} | {val} | {s['latest']} |")
    lines += [
        "",
        "## NIPT market snapshot (base case)",
        "",
        "| Metric | Value |",
        "|---|---|",
        f"| Births, 2026E | {m['births'] / 1e7:,.2f} crore |",
        f"| NIPT tests, 2026E | {lakh(m['tests'])} ({m['penetration']:.1%} of births) |",
        f"| Realised price per test | Rs {m['price']:,.0f} |",
        f"| Market value, 2026E | Rs {m['market_cr']:,.0f} cr (USD {m['market_usd']:,.0f} m) |",
        f"| Market value, 2031E | Rs {m['market_series'][-1]:,.0f} cr |",
        f"| Serviceable market if every addressable pregnancy were tested | Rs {m['sam_cr']:,.0f} cr |",
        "",
        "## What to watch next month",
        "",
        "- Whether pathology chains keep adding genomics capability through JVs and partnerships instead of building it in-house.",
        "- Follow-on bolt-on acquisitions by genomics specialists, and any price moves in prenatal screening as volumes grow.",
        "- Funding rounds for oncology genomics start-ups, which compete for the same lab capacity and talent.",
        "",
    ]
    path = OUT / f"monthly_update_{report_date:%Y-%m}.md"
    path.write_text("\n".join(lines))
    return path


# --------------------------------------------------------------- deck
def add_title(slide, text, sub=None):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.9))
    p = box.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = NAVY
    if sub:
        p2 = box.text_frame.add_paragraph()
        p2.text = sub
        p2.font.size = Pt(14)
        p2.font.color.rgb = GREY


def add_bullets(slide, items, left, top, width, height, size=16):
    tf = slide.shapes.add_textbox(left, top, width, height).text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.size = Pt(size)
        p.space_after = Pt(8)


def add_table(slide, rows, left, top, width, col_widths=None, size=12):
    shape = slide.shapes.add_table(len(rows), len(rows[0]), left, top, width, Inches(0.4) * len(rows))
    table = shape.table
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(val)
            para = cell.text_frame.paragraphs[0]
            para.font.size = Pt(size)
            if r == 0:
                para.font.bold = True
                para.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w


def write_deck(report_date, t, base, scen):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    # 1. Title
    s = prs.slides.add_slide(blank)
    box = s.shapes.add_textbox(Inches(0.8), Inches(2.4), Inches(11.5), Inches(2.5)).text_frame
    box.paragraphs[0].text = "Entering India's NIPT market"
    box.paragraphs[0].font.size = Pt(40)
    box.paragraphs[0].font.bold = True
    box.paragraphs[0].font.color.rgb = NAVY
    for txt, size in [("Market size, growth, entry economics and competitive landscape", 20),
                      (f"Analysis as of {report_date:%d %b %Y}  |  Kartik Kumar Boxi", 14)]:
        p = box.add_paragraph()
        p.text = txt
        p.font.size = Pt(size)
        p.font.color.rgb = GREY

    # 2. Market size
    s = prs.slides.add_slide(blank)
    add_title(s, f"NIPT is a ~Rs {base['market_cr']:,.0f} cr market today, with only {base['penetration']:.1%} of births tested",
              "Bottom-up sizing for 2026 against published estimates")
    rows = [["Estimate", "Market (Rs cr)", "Implied share of births tested"]]
    for name in ("Low", "Base", "High"):
        if scen.get(name):
            rows.append([f"This model: {name} case", f"{scen[name]['market_cr']:,.0f}", f"{scen[name]['penetration']:.1%}"])
    for pub, yr, cr, pct in base["published"]:
        label = f"{pub} ({yr})" if yr not in (None, "n/a") else pub
        rows.append([label, f"{cr:,.0f}", f"{pct:.1%}"])
    add_table(s, rows, Inches(0.6), Inches(1.6), Inches(7.4), [Inches(3.6), Inches(1.8), Inches(2.0)], size=13)
    add_bullets(s, [
        f"Funnel: {base['births'] / 1e7:.2f} cr births, 30% of facility births are private, of which 55% can afford a "
        f"~Rs {base['price']:,.0f} test, and 15% of those are tested today.",
        "Published estimates differ by 4-5x. They line up with the Low and High ends of the bottom-up range, "
        "so the gap comes down to the assumed penetration.",
        f"Headroom is large: testing every addressable pregnancy would be a Rs {base['sam_cr']:,.0f} cr market.",
    ], Inches(8.3), Inches(1.6), Inches(4.6), Inches(5), size=15)

    # 3. Outlook chart
    s = prs.slides.add_slide(blank)
    add_title(s, f"Volumes grow ~{base['volume_cagr']:.0%} a year, value ~{base['value_cagr']:.0%} as prices keep falling",
              "Market value by scenario, Rs crore")
    cd = CategoryChartData()
    cd.categories = [str(y) for y in base["years"]]
    for name in ("Low", "Base", "High"):
        if scen.get(name):
            cd.add_series(f"{name} case", [round(v) for v in scen[name]["market_series"]])
    chart = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.6), Inches(1.5), Inches(8), Inches(5.5), cd).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.tick_labels.font.size = Pt(12)
    chart.category_axis.tick_labels.font.size = Pt(12)
    add_bullets(s, [
        f"Base case: Rs {base['market_series'][0]:,.0f} cr (2026) to Rs {base['market_series'][-1]:,.0f} cr (2031).",
        "Drivers: uptake among private-care pregnancies rises 3 points a year as obstetricians adopt NIPT as standard screening.",
        "Drag: realised prices fall about 4% a year as sequencing costs drop and competition grows.",
        "Births drift down slowly, so all growth comes from penetration, not population.",
    ], Inches(8.9), Inches(1.6), Inches(4.1), Inches(5.5), size=14)

    # 4. Entry economics
    s = prs.slides.add_slide(blank)
    add_title(s, "Building a lab pays only if volumes reach ~17,000 tests a year",
              "NPV 2026-31 at a 14% hurdle rate, Rs crore (no terminal value)")
    rows = [["Scenario", "Build own lab", "Partner (send-out)", "Model says"]]
    for name in ("Low", "Base", "High"):
        if scen.get(name):
            v = scen[name]
            rows.append([name, f"{v['npv_build']:,.1f}", f"{v['npv_partner']:,.1f}", v["recommendation"]])
    add_table(s, rows, Inches(0.6), Inches(1.6), Inches(12.1),
              [Inches(1.4), Inches(1.8), Inches(2.0), Inches(6.9)], size=13)
    add_bullets(s, [
        f"Crossover: owning the lab beats partnering above ~{base['crossover']:,.0f} tests a year; "
        f"in the base case the entrant gets there in {base['crossover_year']}.",
        "Build needs ~Rs 15 cr upfront and loses the most in the Low case; partnering needs ~Rs 1 cr and can launch in weeks.",
        "Cost inputs are illustrative and need vendor quotes before any decision.",
    ], Inches(0.6), Inches(4.2), Inches(12), Inches(3), size=15)

    # 5. Competitive activity
    s = prs.slides.add_slide(blank)
    add_title(s, "Pathology chains are buying and partnering their way into specialty testing",
              "Tracked moves, Jan-2025 to date")
    rows = [["Company", "Segment", "Acquisitions", "JVs / partnerships", "Latest"]]
    for r in t["scoreboard"]:
        rows.append([r["company"], r["segment"], r["acquisitions"], r["jvs"], r["latest"]])
    add_table(s, rows, Inches(0.6), Inches(1.6), Inches(7.6),
              [Inches(2.4), Inches(1.9), Inches(1.1), Inches(1.2), Inches(1.0)], size=12)
    recent = [f"{e['event_date']}  {e['company']}: {e['summary']}" for e in t["last90"][:4]]
    add_bullets(s, ["Last 90 days:"] + recent, Inches(8.5), Inches(1.6), Inches(4.5), Inches(5.5), size=13)

    # 6. Recommendation
    s = prs.slides.add_slide(blank)
    add_title(s, "Recommendation: enter through a partner now, keep the option to build or buy")
    add_bullets(s, [
        "1. Launch NIPT through a reference-lab partner: about Rs 1 cr at risk and live within weeks.",
        f"2. Use the partner phase to test demand. Build an in-house lab, or buy a small genetics lab (the top option "
        f"on the qualitative scorecard), only once run-rate volume passes ~{base['crossover']:,.0f} tests a year.",
        "3. Before committing capex, validate the three assumptions that drive the answer: share gain, realised price "
        "and cost per test.",
        "4. Track monthly: tests per referring obstetrician, realised price, turnaround time, and competitor launches "
        "and deals (see the monthly update).",
        "Why not build now? In the base case building has the higher NPV, but it is the only option that loses heavily "
        "if adoption is slow. Partnering first caps that downside at little cost.",
    ], Inches(0.6), Inches(1.5), Inches(12), Inches(5.5), size=17)

    path = OUT / "NIPT_Entry_Management_Deck.pptx"
    prs.save(path)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--report-date", default="2026-10-31")
    rd = date.fromisoformat(ap.parse_args().report_date)
    facts = tracker_facts(rd)
    base = model_values(MODEL)
    scen = {name: model_values(MODEL, name) for name in ("Low", "Base", "High")}
    print("wrote", write_update(rd, facts, base))
    print("wrote", write_deck(rd, facts, base, scen))


if __name__ == "__main__":
    main()

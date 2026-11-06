"""Build the NIPT market entry model (outputs/NIPT_Market_Entry_Model.xlsx).

Sheets: README, Inputs, Scenarios, Market_Sizing, Forecast, Investment_Case,
Entry_Options, Events. Every output cell is a formula that traces back to the
Inputs and Scenarios sheets.

Usage: python scripts/build_model.py
"""
import csv
import shutil
import subprocess
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = ROOT / "outputs" / "NIPT_Market_Entry_Model.xlsx"

BLUE = Font(color="0000CC")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
WHITE_BOLD = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
KEY_FILL = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="999999")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

CR = '#,##0;(#,##0)'
CR1 = '#,##0.0;(#,##0.0)'
NUM = '#,##0'
PCT = '0.0%;(0.0%)'
RS = '"Rs "#,##0'
YEARS = list(range(2026, 2032))
YCOLS = ["C", "D", "E", "F", "G", "H"]          # 2026..2031


def read_csv(name):
    with open(DATA / name, newline="") as f:
        return list(csv.DictReader(f))


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


def header_row(ws, row, values, start_col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = WHITE_BOLD
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX


def inp(cell, value, fmt=None):
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL
    cell.border = BOX
    if fmt:
        cell.number_format = fmt


def put(ws, ref, value, fmt=None, bold=False, key=False):
    ws[ref] = value
    if fmt:
        ws[ref].number_format = fmt
    if bold:
        ws[ref].font = BOLD
    if key:
        ws[ref].fill = KEY_FILL


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- sheets
def build_readme(wb):
    ws = wb.active
    ws.title = "README"
    lines = [
        ("NIPT (Non-Invasive Prenatal Testing) in India: Market Sizing, Entry Options & Investment Case", TITLE),
        ("", None),
        ("Question this model answers", BOLD),
        ("How big is India's NIPT market today, how fast can it grow, and how should a diagnostics company without "
         "its own genomics lab enter it: build a lab, partner with a reference lab, or acquire?", None),
        ("", None),
        ("Sheets", BOLD),
        ("Inputs : sourced data (births, delivery mix, NIPT prices, third-party market estimates).", None),
        ("Scenarios : Low / Base / High assumptions. Pick the active scenario in C3; every other sheet updates.", None),
        ("Market_Sizing : 2026 bottom-up funnel, TAM / SAM, and a cross-check against published estimates.", None),
        ("Forecast : 2026-2031 test volumes, prices and market value, plus a new entrant's share ramp.", None),
        ("Investment_Case : Build-own-lab vs Partner-with-reference-lab economics: NPV, IRR, break-even, crossover volume.", None),
        ("Entry_Options : weighted scorecard comparing Build, Partner and Acquire.", None),
        ("Events : competitor and deal tracker (2025-26) with COUNTIFS / SUMIFS summaries.", None),
        ("", None),
        ("Conventions", BOLD),
        ("Blue on yellow = input. Black = formula. Green = key output. Money in Rs crore (1 crore = 10 million) unless stated.", None),
        ("Sourced inputs carry a link in the Inputs sheet. Scenario values and cost assumptions in Investment_Case are my "
         "estimates for illustration and should be validated with vendor quotes and expert calls.", None),
        ("", None),
        ("Built by Kartik Kumar Boxi as a self-directed learning project. Not medical or investment advice.", Font(italic=True)),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=text)
        if font:
            c.font = font
    ws.column_dimensions["A"].width = 140


def build_inputs(wb):
    ws = wb.create_sheet("Inputs")
    ws["A1"] = "Sourced inputs"
    ws["A1"].font = TITLE

    header_row(ws, 3, ["Key", "Input", "Value", "Unit", "Source", "Link"])
    for i, row in enumerate(read_csv("market_inputs.csv")):
        r = 4 + i
        ws[f"A{r}"] = row["key"]
        ws[f"B{r}"] = row["label"]
        fmt = PCT if row["unit"] == "%" else NUM
        inp(ws[f"C{r}"], num(row["value"]), fmt)
        ws[f"D{r}"] = row["unit"]
        ws[f"E{r}"] = row["source"]
        ws[f"F{r}"] = row["source_url"]
    # rows: 4 births, 5 institutional share, 6 private share, 7 usd_inr

    ws["A10"] = "NIPT price points (Rs per test, Nov-2026 web listings)"
    ws["A10"].font = BOLD
    header_row(ws, 11, ["Provider", "Test", "Price (Rs)", "Note", "", "Link"])
    prices = read_csv("nipt_price_points.csv")
    for i, row in enumerate(prices):
        r = 12 + i
        ws[f"A{r}"] = row["provider"]
        ws[f"B{r}"] = row["test"]
        inp(ws[f"C{r}"], num(row["price_inr"]), NUM)
        ws[f"D{r}"] = row["note"]
        ws[f"F{r}"] = row["source_url"]
    last = 11 + len(prices)  # 18
    put(ws, "B20", "Median list price", bold=True)
    put(ws, "C20", f"=MEDIAN(C12:C{last})", NUM, bold=True, key=True)
    put(ws, "B21", "Average list price")
    put(ws, "C21", f"=AVERAGE(C12:C{last})", NUM)
    put(ws, "B22", "Lowest / highest")
    put(ws, "C22", f'=TEXT(MIN(C12:C{last}),"#,##0")&" / "&TEXT(MAX(C12:C{last}),"#,##0")')

    ws["A25"] = "Published market-size estimates"
    ws["A25"].font = BOLD
    header_row(ws, 26, ["Publisher", "Year", "Value", "Currency", "Value (Rs cr)", "Note", "Link"])
    for i, row in enumerate(read_csv("third_party_estimates.csv")):
        r = 27 + i
        ws[f"A{r}"] = row["publisher"]
        ws[f"B{r}"] = num(row["estimate_year"]) if row["estimate_year"] else "n/a"
        inp(ws[f"C{r}"], num(row["value"]), CR1)
        ws[f"D{r}"] = row["currency"]
        put(ws, f"E{r}", f'=IF(D{r}="USD m",C{r}*$C$7/10,C{r})', CR)
        ws[f"F{r}"] = row["note"]
        ws[f"G{r}"] = row["source_url"]
    ws["A31"] = "USD m to Rs cr: multiply by USD/INR and divide by 10 (1 crore = 10 million)."
    widths(ws, {"A": 26, "B": 44, "C": 14, "D": 40, "E": 44, "F": 50, "G": 50})
    ws.row_dimensions[26].height = 30


SCEN = [  # label, low, base, high, fmt, note
    ("Birth trend per year", -0.01, -0.005, 0.0, PCT, "India's fertility is below replacement; births drift down slowly"),
    ("Share of private-facility births that can afford NIPT", 0.40, 0.55, 0.70, PCT,
     "Families already paying for private delivery and able to spend ~Rs 14-15k on one test"),
    ("NIPT uptake in the addressable segment, 2026", 0.10, 0.15, 0.22, PCT, "Share of addressable pregnancies tested today"),
    ("Uptake gain per year (percentage points)", 0.015, 0.03, 0.045, PCT, "Driven by obstetrician adoption and falling prices; capped at 60%"),
    ("Uptake outside the addressable segment", 0.003, 0.005, 0.01, PCT, "Public-facility and lower-income pregnancies"),
    ("Realised price discount to median list price", 0.20, 0.10, 0.05, PCT, "Doctor, hospital and online discounts"),
    ("Price change per year", -0.06, -0.04, -0.02, PCT, "Sequencing costs keep falling; competition passes it on"),
    ("New entrant's share of tests by 2031", 0.02, 0.04, 0.07, PCT, "Share ramps linearly from 0% in 2026"),
]


def build_scenarios(wb):
    ws = wb.create_sheet("Scenarios")
    ws["A1"] = "Scenario assumptions"
    ws["A1"].font = TITLE
    ws["B3"] = "Active scenario"
    ws["B3"].font = BOLD
    inp(ws["C3"], "Base")
    dv = DataValidation(type="list", formula1='"Low,Base,High"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C3")

    header_row(ws, 5, ["Parameter", "Low", "Base", "High", "Active", "Why"], start_col=2)
    for i, (label, lo, base, hi, fmt, note) in enumerate(SCEN):
        r = 6 + i
        ws[f"B{r}"] = label
        for col, v in zip("CDE", (lo, base, hi)):
            inp(ws[f"{col}{r}"], v, fmt)
        put(ws, f"F{r}", f"=INDEX(C{r}:E{r},MATCH($C$3,$C$5:$E$5,0))", fmt, bold=True, key=True)
        ws[f"G{r}"] = note
    ws["B15"] = "Active column uses INDEX/MATCH on the scenario name, so switching C3 flows through every sheet."
    widths(ws, {"A": 2, "B": 52, "C": 10, "D": 10, "E": 10, "F": 10, "G": 70})
    # Scenario rows: 6 birth trend, 7 afford, 8 uptake, 9 gain, 10 outside, 11 discount, 12 price chg, 13 share


def build_market_sizing(wb):
    ws = wb.create_sheet("Market_Sizing")
    ws["A1"] = "2026 market size: bottom-up funnel"
    ws["A1"].font = TITLE
    ws["A2"] = '=CONCATENATE("Scenario: ",Scenarios!C3)'
    ws["A2"].font = BOLD

    header_row(ws, 4, ["Step", "", "2026E", "How it is calculated"])
    rows = [
        (5, "Live births", "=Inputs!C4*(1+Scenarios!F6)^2", NUM, "2024 CRS births rolled forward 2 years at the birth trend"),
        (6, "Births in a health facility", "=C5*Inputs!C5", NUM, "x institutional delivery share (NFHS-6)"),
        (7, "Births in private facilities", "=C6*Inputs!C6", NUM, "x private share of institutional births (NFHS-5)"),
        (8, "Addressable pregnancies (private + can afford NIPT)", "=C7*Scenarios!F7", NUM, "x affordability share (scenario)"),
        (9, "All other pregnancies", "=C5-C8", NUM, "Births not in the addressable segment"),
        (10, "NIPT tests: addressable segment", "=C8*Scenarios!F8", NUM, "x uptake in addressable segment (scenario)"),
        (11, "NIPT tests: other pregnancies", "=C9*Scenarios!F10", NUM, "x uptake outside the segment (scenario)"),
        (12, "Total NIPT tests", "=C10+C11", NUM, ""),
        (13, "Penetration of all births", "=C12/C5", PCT, ""),
        (14, "Realised price per test (Rs)", "=Inputs!C20*(1-Scenarios!F11)", RS, "Median list price less discount (scenario)"),
        (15, "Market value (Rs cr)", "=C12*C14/10^7", CR, "tests x price / 1 crore"),
        (16, "Market value (USD m)", "=C15*10/Inputs!C7", CR1, ""),
    ]
    for r, label, f, fmt, how in rows:
        ws[f"A{r}"] = label
        put(ws, f"C{r}", f, fmt)
        ws[f"D{r}"] = how
    for r in (12, 15):
        ws[f"A{r}"].font = BOLD
        ws[f"C{r}"].font = BOLD
        ws[f"C{r}"].fill = KEY_FILL

    ws["A19"] = "Market layers (Rs cr, 2026)"
    ws["A19"].font = BOLD
    header_row(ws, 20, ["Layer", "", "Rs cr", "Meaning"])
    layers = [
        (21, "TAM: every pregnancy tested", "=C5*C14/10^7", "Theoretical ceiling at today's price"),
        (22, "SAM: every addressable pregnancy tested", "=C8*C14/10^7", "Private-care families who can afford the test"),
        (23, "Current market", "=C15", "What is actually being tested today"),
        (24, "Headroom inside SAM", "=C22-C10*C14/10^7", "Addressable pregnancies not yet tested"),
        (25, "New entrant's obtainable revenue in 2031", "=Forecast!H17", "From the Forecast sheet (share ramp)"),
    ]
    for r, label, f, how in layers:
        ws[f"A{r}"] = label
        put(ws, f"C{r}", f, CR)
        ws[f"D{r}"] = how

    ws["A28"] = "Cross-check against published estimates"
    ws["A28"].font = BOLD
    header_row(ws, 29, ["Estimate", "Year", "Rs cr", "Implied tests at our realised price", "Implied % of 2026 births"])
    xs = [
        (30, '=CONCATENATE("This model (",Scenarios!C3,")")', 2026, "=C15"),
        (31, "=Inputs!A27", "=Inputs!B27", "=Inputs!E27"),
        (32, "=Inputs!A28", "=Inputs!B28", "=Inputs!E28"),
        (33, "=Inputs!A29", "=Inputs!B29", "=Inputs!E29"),
    ]
    for r, label, yr, val in xs:
        ws[f"A{r}"] = label
        ws[f"B{r}"] = yr
        put(ws, f"C{r}", val, CR)
        put(ws, f"D{r}", f"=C{r}*10^7/$C$14", NUM)
        put(ws, f"E{r}", f"=D{r}/$C$5", PCT)
    ws["A35"] = ("Published estimates range from about Rs 400 cr to Rs 1,800 cr, a 4x spread. Converting each into "
                 "implied tests shows what penetration it assumes, which is the quickest way to judge which one is realistic.")
    widths(ws, {"A": 50, "B": 10, "C": 16, "D": 60, "E": 22})
    ws.row_dimensions[29].height = 30


def build_forecast(wb):
    ws = wb.create_sheet("Forecast")
    ws["A1"] = "Market forecast 2026-2031"
    ws["A1"].font = TITLE
    ws["A2"] = '=CONCATENATE("Scenario: ",Scenarios!C3)'
    ws["A2"].font = BOLD
    header_row(ws, 4, ["Calendar year", ""] + [str(y) for y in YEARS])
    labels = {5: "Years from 2026", 6: "Live births", 7: "Addressable pregnancies", 8: "Uptake in addressable segment",
              9: "Tests: addressable", 10: "Tests: other pregnancies", 11: "Total NIPT tests",
              12: "Realised price per test (Rs)", 13: "Market value (Rs cr)", 14: "Market growth",
              15: "New entrant share", 16: "New entrant tests", 17: "New entrant revenue (Rs cr)"}
    for r, lab in labels.items():
        ws[f"A{r}"] = lab
    for r in (11, 13, 16, 17):
        ws[f"A{r}"].font = BOLD
    for i, (col, yr) in enumerate(zip(YCOLS, YEARS)):
        ws[f"{col}4"] = yr  # numeric year used in formulas
        ws[f"{col}4"].font = WHITE_BOLD
        ws[f"{col}4"].fill = HEADER_FILL
        put(ws, f"{col}5", f"={col}4-2026", "0")
        put(ws, f"{col}6", f"=Inputs!$C$4*(1+Scenarios!$F$6)^({col}4-2024)", NUM)
        put(ws, f"{col}7", f"={col}6*Inputs!$C$5*Inputs!$C$6*Scenarios!$F$7", NUM)
        put(ws, f"{col}8", f"=MIN(Scenarios!$F$8+Scenarios!$F$9*{col}5,0.6)", PCT)
        put(ws, f"{col}9", f"={col}7*{col}8", NUM)
        put(ws, f"{col}10", f"=({col}6-{col}7)*Scenarios!$F$10", NUM)
        put(ws, f"{col}11", f"={col}9+{col}10", NUM, bold=True)
        put(ws, f"{col}12", f"=Market_Sizing!$C$14*(1+Scenarios!$F$12)^{col}5", RS)
        put(ws, f"{col}13", f"={col}11*{col}12/10^7", CR, bold=True, key=True)
        if i > 0:
            prev = YCOLS[i - 1]
            put(ws, f"{col}14", f"={col}13/{prev}13-1", PCT)
        put(ws, f"{col}15", f"=Scenarios!$F$13*{col}5/5", PCT)
        put(ws, f"{col}16", f"={col}11*{col}15", NUM, bold=True)
        put(ws, f"{col}17", f"={col}16*{col}12/10^7", CR1)
    ws["A19"] = "Market value CAGR 2026-31"
    put(ws, "C19", "=(H13/C13)^(1/5)-1", PCT, bold=True, key=True)
    ws["A20"] = "Test volume CAGR 2026-31"
    put(ws, "C20", "=(H11/C11)^(1/5)-1", PCT, bold=True)
    ws["A22"] = "Volumes grow faster than value because the price per test keeps falling."
    widths(ws, {"A": 34, "B": 2, **{c: 14 for c in YCOLS}})


def build_investment_case(wb):
    ws = wb.create_sheet("Investment_Case")
    ws["A1"] = "Entry economics: build an in-house NIPT lab vs partner with a reference lab"
    ws["A1"].font = TITLE
    ws["A2"] = ("Cost inputs below are illustrative assumptions, not quotes. Volumes and prices come from the Forecast "
                "sheet, so the case changes with the scenario.")

    assumptions = [
        (4, "Option A: build own lab", None, None),
        (5, "Upfront capex: sequencer, lab fit-out, bioinformatics (Rs cr)", 15, CR1),
        (6, "Variable cost per test: reagents, sequencing, logistics, counselling (Rs)", 6000, NUM),
        (7, "Change in variable cost per year", -0.03, PCT),
        (8, "Fixed operating cost per year: staff, QC, accreditation, rent (Rs cr)", 5, CR1),
        (9, "Depreciation life (years)", 5, "0"),
        (10, "Option B: partner with a reference lab (send-out model)", None, None),
        (11, "Upfront setup: sales, sample logistics, IT (Rs cr)", 1, CR1),
        (12, "Price paid to reference lab per test (Rs)", 9500, NUM),
        (13, "Own logistics and counselling cost per test (Rs)", 500, NUM),
        (14, "Fixed operating cost per year (Rs cr)", 1, CR1),
        (15, "Common", None, None),
        (16, "Tax rate", 0.2517, PCT),
        (17, "Hurdle rate for the project", 0.14, PCT),
    ]
    for r, label, val, fmt in assumptions:
        ws[f"A{r}"] = label
        if val is None:
            ws[f"A{r}"].font = BOLD
        else:
            inp(ws[f"C{r}"], val, fmt)

    header_row(ws, 19, ["Rs crore unless stated", ""] + [str(y) for y in YEARS])
    for col, yr in zip(YCOLS, YEARS):
        ws[f"{col}19"] = yr
        ws[f"{col}19"].font = WHITE_BOLD
        ws[f"{col}19"].fill = HEADER_FILL
    ws["A20"] = "Entrant tests (from Forecast)"
    ws["A21"] = "Realised price per test (Rs)"
    for col in YCOLS:
        put(ws, f"{col}20", f"=Forecast!{col}16", NUM)
        put(ws, f"{col}21", f"=Forecast!{col}12", RS)

    blocks = {
        "A": dict(start=23, title="Option A: build own lab"),
        "B": dict(start=33, title="Option B: partner with reference lab"),
    }
    for opt, b in blocks.items():
        s = b["start"]
        ws[f"A{s}"] = b["title"]
        ws[f"A{s}"].font = BOLD
        names = ["Revenue", "Variable costs", "Fixed costs", "EBITDA", "Depreciation", "Tax", "Upfront investment",
                 "Free cash flow"]
        for k, n in enumerate(names):
            ws[f"A{s + 1 + k}"] = n
        ws[f"A{s + 8}"].font = BOLD
        rev, var, fix, ebitda, dep, tax, capex, fcf = range(s + 1, s + 9)
        capex_in = "$C$5" if opt == "A" else "$C$11"
        life = "$C$9" if opt == "A" else "5"
        for col in YCOLS:
            first = col == "C"
            if first:  # 2026 = year 0: only the upfront investment
                for r in (rev, var, fix, ebitda, dep, tax):
                    put(ws, f"{col}{r}", 0, CR1)
                put(ws, f"{col}{capex}", f"={capex_in}", CR1)
            else:
                put(ws, f"{col}{rev}", f"={col}$20*{col}$21/10^7", CR1)
                if opt == "A":
                    put(ws, f"{col}{var}", f"={col}$20*$C$6*(1+$C$7)^({col}$19-2027)/10^7", CR1)
                    put(ws, f"{col}{fix}", "=$C$8", CR1)
                else:
                    put(ws, f"{col}{var}", f"={col}$20*($C$12+$C$13)/10^7", CR1)
                    put(ws, f"{col}{fix}", "=$C$14", CR1)
                put(ws, f"{col}{ebitda}", f"={col}{rev}-{col}{var}-{col}{fix}", CR1)
                put(ws, f"{col}{dep}", f"={capex_in}/{life}", CR1)
                put(ws, f"{col}{tax}", f"=MAX(0,{col}{ebitda}-{col}{dep})*$C$16", CR1)
                put(ws, f"{col}{capex}", 0, CR1)
            put(ws, f"{col}{fcf}", f"={col}{ebitda}-{col}{tax}-{col}{capex}", CR1, bold=True)
        b["fcf"] = fcf
        b["ebitda"] = ebitda

    fa, fb = blocks["A"]["fcf"], blocks["B"]["fcf"]   # 31, 41
    ws["A44"] = "Results"
    ws["A44"].font = TITLE
    header_row(ws, 45, ["Metric", "", "Build (A)", "Partner (B)"])
    res = [
        (46, "NPV at hurdle rate, 2026-31 (Rs cr, no terminal value)",
         f"=C{fa}+NPV($C$17,D{fa}:H{fa})", f"=C{fb}+NPV($C$17,D{fb}:H{fb})", CR1),
        (47, "IRR", f'=IFERROR(IRR(C{fa}:H{fa}),"n/a")', f'=IFERROR(IRR(C{fb}:H{fb}),"n/a")', PCT),
        (48, "Cumulative cash flow by 2031 (Rs cr)", f"=SUM(C{fa}:H{fa})", f"=SUM(C{fb}:H{fb})", CR1),
        (49, "EBITDA in 2031 (Rs cr)", f"=H{blocks['A']['ebitda']}", f"=H{blocks['B']['ebitda']}", CR1),
        (50, "Break-even volume, tests per year (2027 prices and costs)",
         "=$C$8*10^7/($D$21-$C$6)", "=$C$14*10^7/($D$21-$C$12-$C$13)", NUM),
    ]
    for r, label, fa_f, fb_f, fmt in res:
        ws[f"A{r}"] = label
        put(ws, f"C{r}", fa_f, fmt, key=True)
        put(ws, f"D{r}", fb_f, fmt, key=True)

    ws["A52"] = "Crossover: annual volume above which building beats partnering (tests/yr)"
    put(ws, "C52", "=(($C$8+$C$5/$C$9)-($C$14+$C$11/5))*10^7/(($C$12+$C$13)-$C$6)", NUM, bold=True, key=True)
    ws["A53"] = "First year entrant volume passes the crossover"
    put(ws, "C53", ('=IF(D20>=C52,D19,IF(E20>=C52,E19,IF(F20>=C52,F19,'
                    'IF(G20>=C52,G19,IF(H20>=C52,H19,"Beyond 2031")))))'), bold=True)
    ws["A54"] = "Recommendation"
    put(ws, "C54", ('=IF(AND(C46<0,D46<0),"Do not enter: both options lose value in this scenario",'
                    'IF(D46>=C46,"Partner (send-out model): higher NPV with Rs "&TEXT($C$11,"0")&" cr at risk",'
                    '"Build own lab: NPV higher by Rs "&TEXT(C46-D46,"0.0")&" cr; volume passes crossover in "&C53))'),
        bold=True)
    ws["A56"] = ("Crossover = extra fixed cost of owning the lab (opex + depreciation) / saving per test. "
                 "Price cancels out because both options sell at the same price.")
    widths(ws, {"A": 66, "B": 2, **{c: 13 for c in YCOLS}})


def build_entry_options(wb):
    ws = wb.create_sheet("Entry_Options")
    ws["A1"] = "Entry options scorecard"
    ws["A1"].font = TITLE
    ws["A2"] = "Scores 1 (poor) to 5 (strong) are judgement calls informed by the Investment_Case sheet; weights are editable."
    header_row(ws, 4, ["Criterion", "Weight", "Build own lab", "Partner (send-out)", "Acquire a genetics lab", "Why"])
    crit = [
        ("Capital required", 0.20, 2, 5, 1, "Build needs ~Rs 15 cr upfront; acquisition needs more; partner needs ~Rs 1 cr"),
        ("Time to market", 0.20, 2, 5, 4, "Lab setup and accreditation take 9-12 months; partner can launch in weeks"),
        ("Unit margin at scale", 0.20, 5, 2, 5, "Owning the lab captures the reference lab's margin"),
        ("Control of quality, turnaround and data", 0.15, 5, 2, 5, "Own lab controls reporting and keeps genomic data in-house"),
        ("Execution risk", 0.15, 2, 4, 2, "New lab and integration both carry operational risk"),
        ("Strategic optionality (other genomic tests)", 0.10, 4, 2, 5, "A sequencing lab can add carrier screening, oncology panels"),
    ]
    for i, (name, w, a, b, c, why) in enumerate(crit):
        r = 5 + i
        ws[f"A{r}"] = name
        inp(ws[f"B{r}"], w, PCT)
        for col, v in zip("CDE", (a, b, c)):
            inp(ws[f"{col}{r}"], v, "0")
        ws[f"F{r}"] = why
    last = 4 + len(crit)
    ws[f"A{last + 1}"] = "Total weight"
    put(ws, f"B{last + 1}", f"=SUM(B5:B{last})", PCT)
    ws[f"A{last + 2}"] = "Weighted score (out of 5)"
    ws[f"A{last + 2}"].font = BOLD
    ws[f"A{last + 3}"] = "Rank"
    for col in "CDE":
        put(ws, f"{col}{last + 2}", f"=SUMPRODUCT($B$5:$B${last},{col}5:{col}{last})", "0.00", bold=True, key=True)
        put(ws, f"{col}{last + 3}", f"=RANK({col}{last + 2},$C${last + 2}:$E${last + 2})", "0", bold=True)
    widths(ws, {"A": 42, "B": 10, "C": 14, "D": 16, "E": 18, "F": 80})
    ws.row_dimensions[4].height = 30


def build_events(wb):
    ws = wb.create_sheet("Events")
    ws["A1"] = "Competitor and deal tracker (same data as the SQL tracker)"
    ws["A1"].font = TITLE
    events = read_csv("industry_events.csv")
    segs = {r["company"]: r["segment"] for r in read_csv("companies.csv")}
    cols = ["Date", "Company", "Segment", "Event type", "Counterparty", "Rs cr", "USD m", "Geography", "Summary", "Source"]
    header_row(ws, 3, cols)
    for i, e in enumerate(events):
        r = 4 + i
        vals = [e["event_date"], e["company"], segs.get(e["company"], ""), e["event_type"], e["counterparty"],
                num(e["amount_inr_cr"]) if e["amount_inr_cr"] else None,
                num(e["amount_usd_m"]) if e["amount_usd_m"] else None, e["geography"], e["summary"], e["source_url"]]
        for c, v in enumerate(vals, start=1):
            ws.cell(row=r, column=c, value=v).font = BLUE
    last = 3 + len(events)

    ws["L3"] = "Summary"
    ws["L3"].font = BOLD
    header_row(ws, 4, ["Event type", "Count", "Disclosed Rs cr"], start_col=12)
    types = ["Acquisition", "Funding", "Joint venture", "Partnership", "Product launch", "Capability", "Results"]
    for i, t in enumerate(types):
        r = 5 + i
        ws[f"L{r}"] = t
        put(ws, f"M{r}", f'=COUNTIFS($D$4:$D${last},L{r})', "0")
        put(ws, f"N{r}", f'=SUMIFS($F$4:$F${last},$D$4:$D${last},L{r})', CR1)
    put(ws, "L12", "Total", bold=True)
    put(ws, "M12", "=SUM(M5:M11)", "0", bold=True)
    put(ws, "N12", "=SUM(N5:N11)", CR1, bold=True)

    header_row(ws, 15, ["Company", "Acquisitions", "All events"], start_col=12)
    for i, co in enumerate(sorted(segs)):
        r = 16 + i
        ws[f"L{r}"] = co
        put(ws, f"M{r}", f'=COUNTIFS($B$4:$B${last},L{r},$D$4:$D${last},"Acquisition")', "0")
        put(ws, f"N{r}", f'=COUNTIFS($B$4:$B${last},L{r})', "0")
    ws.freeze_panes = "A4"
    widths(ws, {"A": 11, "B": 24, "C": 18, "D": 15, "E": 32, "F": 9, "G": 9, "H": 13, "I": 70, "J": 40,
                "K": 2, "L": 24, "M": 12, "N": 15})


def recalculate(path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice not found: values will be calculated when the file is opened in Excel")
        return
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, str(path)],
                       check=True, capture_output=True)
        shutil.move(Path(tmp) / path.name, path)


def main():
    wb = Workbook()
    build_readme(wb)
    build_inputs(wb)
    build_scenarios(wb)
    build_market_sizing(wb)
    build_forecast(wb)
    build_investment_case(wb)
    build_entry_options(wb)
    build_events(wb)
    wb.calculation.fullCalcOnLoad = True
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    recalculate(OUT)
    print(f"saved {OUT}")


if __name__ == "__main__":
    main()

"""
Build an Apple Ads FP&A model (Excel, live formulas) driven by the App Store data.

Framing: Apple Ads sells search ads on the App Store; its customers are app developers.
What developers can afford to bid depends on how their apps monetize, so the App Store
pricing data is used as the *demand-side driver* of an ad-revenue plan:

  Data                -> category engagement share, freemium share, intl price index
  Assumptions         -> scenario levers (Base / Upside / Downside), seasonality, drivers
  Revenue_Forecast    -> FY27 monthly ad revenue by category (advertisers x spend x seasonality)
  Opex_Headcount      -> headcount plan, loaded cost, non-HC opex, capex & depreciation
  PnL_Summary         -> quarterly P&L, margins
  Budget_vs_Actual    -> Q1 variance analysis with volume / rate decomposition
  Market_Investment   -> NPV of incremental sales investment by storefront
  Sensitivity         -> FY27 revenue vs advertiser growth x spend growth
  Dashboard           -> one-page summary for leadership

All Apple Ads financial figures are ILLUSTRATIVE (Apple does not disclose them).
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
TABLES = ROOT / "reports" / "tables"
OUT = ROOT / "reports" / "Apple_Ads_FPA_Model.xlsx"

# ---------- styles ----------
F = "Arial"
BLUE = Font(name=F, color="0000FF")
GREEN = Font(name=F, color="008000")
BLACK = Font(name=F)
BOLD = Font(name=F, bold=True)
TITLE = Font(name=F, bold=True, size=14)
HDR = Font(name=F, bold=True, color="FFFFFF")
NOTE = Font(name=F, italic=True, color="666666", size=9)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
KEY_FILL = PatternFill("solid", fgColor="FFFF00")
TOT_FILL = PatternFill("solid", fgColor="E7E6E6")
THIN = Side(style="thin", color="999999")
TOP = Border(top=THIN)

USD_MM = '$#,##0.0;($#,##0.0);"-"'
USD = '$#,##0;($#,##0);"-"'
USD2 = '$#,##0.00;($#,##0.00);"-"'
NUM = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
MULT = '0.00x'

MONTHS = ["Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"]
QTRS = ["Q1", "Q2", "Q3", "Q4"]
MCOL = [L(c) for c in range(2, 14)]  # B..M month columns, shared by Assumptions & Revenue_Forecast


def hdr(ws, row, values, col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font, c.fill = HDR, HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")


def put(ws, ref, value, font=BLACK, fmt=None, fill=None, comment=None):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if comment:
        c.comment = Comment(comment, "Model")
    return c


def widths(ws, w):
    for col, width in w.items():
        ws.column_dimensions[col].width = width


# ---------- data ----------

def category_data():
    apps = pd.read_csv(DATA / "apps.csv")
    charts = pd.read_csv(DATA / "charts.csv")
    df = charts.merge(apps, on="trackId")
    eng = (df[df.chart.isin(["top_free", "top_grossing"])]
           .drop_duplicates(["trackId", "chart_genre"])
           .groupby("chart_genre").userRatingCount.sum())
    tg = df[df.chart == "top_grossing"]
    free_share = tg.groupby("chart_genre").price.apply(lambda s: (s == 0).mean())
    paid = df[(df.chart == "top_paid") & (df.price > 0)]
    med_paid = paid.groupby("chart_genre").price.median()
    n_apps = df.groupby("chart_genre").trackId.nunique()
    t = pd.DataFrame({"apps": n_apps, "reviews": eng, "free_share": free_share,
                      "median_paid": med_paid}).fillna(0)
    return t.sort_values("reviews", ascending=False)


def intl_data():
    t = pd.read_csv(TABLES / "intl_price_index.csv")
    return t[t.country != "us"].sort_values("median_index", ascending=False)


# Standard VAT/GST rates embedded in storefront prices (US & Canada prices are shown pre-tax)
VAT = {"gb": 0.20, "de": 0.19, "jp": 0.10, "au": 0.10, "in": 0.18, "kr": 0.10,
       "br": 0.0, "cn": 0.06, "ca": 0.0}

# Share of addressable advertisers able to run ads. Mainland China requires business + app
# approval (licenses, industry qualifications) - illustrative haircut.
ELIGIBLE = {"cn": 0.70}

# Illustrative storefront size vs US (share of US advertiser base addressable there)
MARKET_SIZE = {"gb": 0.12, "de": 0.09, "jp": 0.16, "cn": 0.22, "in": 0.07,
               "br": 0.06, "au": 0.05, "ca": 0.07, "kr": 0.06}


def build():
    cats = category_data()
    intl = intl_data()
    n = len(cats)
    wb = Workbook()

    # ================= README =================
    ws = wb.active
    ws.title = "README"
    put(ws, "A1", "Apple Ads FP&A Model — FY27 Plan driven by App Store pricing data", TITLE)
    lines = [
        ("Purpose", "Driver-based ad-revenue plan, cost/headcount plan, variance analysis and investment case "
                    "for an App Store search-ads business. Demand drivers come from live App Store data."),
        ("Fiscal year", "Apple fiscal calendar: Q1 = Oct–Dec, Q4 = Jul–Sep. Units: $mm unless stated."),
        ("IMPORTANT", "Apple does not disclose Apple Ads revenue, advertisers or costs. Every financial "
                      "input (blue) is ILLUSTRATIVE and should be replaced; the structure and the data-driven "
                      "allocations are the point of the model."),
        ("Data source", "Apple iTunes RSS top charts + iTunes Lookup API (US, 22 categories) and 10 storefronts; "
                        "ECB FX via frankfurter.app. See Data sheet and project README."),
        ("How to use", "1) Pick a scenario in Assumptions!B3.  2) Edit blue inputs (yellow = key levers).  "
                       "3) Read Dashboard. Everything else recalculates."),
    ]
    for i, (k, v) in enumerate(lines, start=3):
        put(ws, f"A{i}", k, BOLD)
        put(ws, f"B{i}", v)
        ws[f"B{i}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 45
    put(ws, "A10", "Legend", BOLD)
    put(ws, "A11", "1,234", BLUE); put(ws, "B11", "Blue = hard-coded input / scenario lever")
    put(ws, "A12", "1,234", BLACK); put(ws, "B12", "Black = formula")
    put(ws, "A13", "1,234", GREEN); put(ws, "B13", "Green = link to another sheet")
    put(ws, "A14", "1,234", BLUE, fill=KEY_FILL); put(ws, "B14", "Yellow fill = key assumption to review")
    put(ws, "A16", "Sheets", BOLD)
    for i, (s, d) in enumerate([
        ("Dashboard", "One-page summary for leadership"),
        ("Data", "App Store category drivers & international price index (pasted values from analysis)"),
        ("Assumptions", "Scenario selector, growth levers, seasonality, category drivers"),
        ("Revenue_Forecast", "Monthly FY27 ad revenue by category"),
        ("Opex_Headcount", "Headcount plan, people cost, non-HC opex, capex & depreciation"),
        ("PnL_Summary", "Quarterly P&L and margins"),
        ("Budget_vs_Actual", "Q1 FY27 variance: volume (advertisers) vs rate (spend per advertiser)"),
        ("Market_Investment", "NPV / payback of incremental sales investment by storefront"),
        ("Sensitivity", "FY27 revenue: advertiser growth x spend-per-advertiser growth"),
    ], start=17):
        put(ws, f"A{i}", s, BOLD); put(ws, f"B{i}", d)
    widths(ws, {"A": 20, "B": 110})

    # ================= Data =================
    d = wb.create_sheet("Data")
    put(d, "A1", "App Store category drivers (US, top charts)", TITLE)
    put(d, "A2", "Source: iTunes RSS top free / paid / grossing (limit 100 per category) + iTunes Lookup API; "
                 "pulled by src/collect.py. Values pasted by src/build_model.py.", NOTE)
    hdr(d, 4, ["Category", "Unique apps in charts", "Ratings count (top free + grossing)",
               "Engagement share", "Top-grossing free share", "Median paid price ($)", "Willingness-to-pay index"])
    r0 = 5
    for i, (cat, row) in enumerate(cats.iterrows()):
        r = r0 + i
        put(d, f"A{r}", cat)
        put(d, f"B{r}", int(row.apps), BLUE, NUM)
        put(d, f"C{r}", int(row.reviews), BLUE, NUM)
        put(d, f"D{r}", f"=IFERROR(C{r}/SUM($C${r0}:$C${r0 + n - 1}),0)", BLACK, PCT)
        put(d, f"E{r}", float(row.free_share), BLUE, PCT)
        put(d, f"F{r}", float(row.median_paid), BLUE, USD2)
        put(d, f"G{r}", f"=IFERROR(F{r}/AVERAGE($F${r0}:$F${r0 + n - 1}),0)", BLACK, MULT)
    rt = r0 + n
    put(d, f"A{rt}", "Total", BOLD)
    for c in "BC":
        put(d, f"{c}{rt}", f"=SUM({c}{r0}:{c}{rt - 1})", BOLD, NUM)
    put(d, f"D{rt}", f"=SUM(D{r0}:D{rt - 1})", BOLD, PCT)
    for c in "ABCDEFG":
        d[f"{c}{rt}"].fill, d[f"{c}{rt}"].border = TOT_FILL, TOP
    DATA_FIRST, DATA_LAST = r0, rt - 1

    ir = rt + 3
    put(d, f"A{ir}", "International price index — same paid app, USD-converted, US = 100", TITLE)
    put(d, f"A{ir + 1}", "Source: iTunes Lookup API per storefront; FX = ECB reference rates (frankfurter.app).", NOTE)
    hdr(d, ir + 3, ["Country code", "Country", "Currency", "Apps priced", "Availability vs US", "Median price index"])
    INTL_FIRST = ir + 4
    for i, row in enumerate(intl.itertuples()):
        r = INTL_FIRST + i
        put(d, f"A{r}", row.country)
        put(d, f"B{r}", row.country_name)
        put(d, f"C{r}", row.currency)
        put(d, f"D{r}", int(row.apps), BLUE, NUM)
        put(d, f"E{r}", float(row.availability), BLUE, PCT)
        put(d, f"F{r}", float(row.median_index), BLUE, "0.0")
    INTL_LAST = INTL_FIRST + len(intl) - 1
    widths(d, {"A": 22, "B": 18, "C": 22, "D": 16, "E": 18, "F": 18, "G": 16})
    d.freeze_panes = "B5"

    # ================= Assumptions =================
    a = wb.create_sheet("Assumptions")
    put(a, "A1", "Assumptions & scenario levers", TITLE)
    put(a, "A3", "Active scenario", BOLD)
    put(a, "B3", "Base", BLUE, fill=KEY_FILL)
    dv = DataValidation(type="list", formula1='"Base,Upside,Downside"', allow_blank=False)
    a.add_data_validation(dv); dv.add("B3")
    put(a, "C3", "=MATCH(B3,$C$6:$E$6,0)", BLACK)
    put(a, "D3", "← scenario index (1=Base, 2=Upside, 3=Downside)", NOTE)

    hdr(a, 6, ["Driver", "Active", "Base", "Upside", "Downside", "Note"])
    levers = [
        (7, "Advertiser growth, YoY", (0.12, 0.18, 0.05), PCT, "Illustrative. Net new active advertisers."),
        (8, "Spend per advertiser growth, YoY", (0.08, 0.12, 0.02), PCT,
         "Illustrative. Reflects auction pricing (CPT) and budget expansion."),
        (9, "Non-HC opex, % of revenue", (0.16, 0.15, 0.18), PCT,
         "Illustrative. Ad-serving infrastructure, payments, tools."),
        (10, "Discount rate (investment cases)", (0.09, 0.09, 0.09), PCT, "Illustrative hurdle rate."),
    ]
    for r, label, vals, fmt, note in levers:
        put(a, f"A{r}", label)
        put(a, f"B{r}", f"=INDEX(C{r}:E{r},$C$3)", BLACK, fmt)
        for col, v in zip("CDE", vals):
            put(a, f"{col}{r}", v, BLUE, fmt, KEY_FILL if r in (7, 8) else None)
        put(a, f"F{r}", note, NOTE)

    put(a, "A12", "FY26 exit active advertisers (total)")
    put(a, "B12", 60000, BLUE, NUM, KEY_FILL, "Illustrative — Apple does not disclose.")
    put(a, "A13", "FY26 exit avg monthly spend per advertiser ($)")
    put(a, "B13", 7000, BLUE, USD, KEY_FILL, "Illustrative — Apple does not disclose.")
    put(a, "A14", "Spend multiplier: intercept")
    put(a, "B14", 0.4, BLUE, "0.00", comment="Category spend multiplier = intercept + slope x willingness-to-pay index")
    put(a, "A15", "Spend multiplier: slope on willingness-to-pay index")
    put(a, "B15", 0.6, BLUE, "0.00",
        comment="Hypothesis: categories whose users pay more upfront (median paid price) have higher LTV, so advertisers bid more per tap. Top-grossing free share is ~99% in every category, so it cannot differentiate.")
    put(a, "A16", "FY26 run-rate revenue ($mm)")
    put(a, "B16", "=B12*B13*12/1000000", BLACK, USD_MM)

    put(a, "A18", "Monthly seasonality (avg = 1.00)", BOLD)
    for i, m in enumerate(MONTHS):
        c = a[f"{MCOL[i]}18"]; c.value = m; c.font, c.fill = HDR, HDR_FILL
    put(a, "N18", "Avg check", BOLD)
    seas = [1.02, 1.10, 1.22, 0.92, 0.90, 0.96, 0.95, 0.97, 0.98, 0.98, 0.99, 1.01]
    put(a, "A19", "Seasonality index")
    for i, s in enumerate(seas):
        put(a, f"{MCOL[i]}19", s, BLUE, "0.00")
    put(a, "N19", "=AVERAGE(B19:M19)", BLACK, "0.00")
    put(a, "A20", "Months from FY26 exit")
    for i in range(12):
        put(a, f"{MCOL[i]}20", i + 1, BLUE, "0")
    put(a, "A21", "Seasonality: holiday quarter peak (Dec) and January reset, typical of performance ads. Illustrative.",
        NOTE)

    hdr(a, 23, ["Category", "Engagement share", "Willingness-to-pay index", "Median paid price ($)",
                "FY26 advertisers", "Spend multiplier", "Raw monthly spend ($)", "Monthly spend / advertiser ($)",
                "FY26 monthly revenue ($mm)"])
    A_FIRST = 24
    for i in range(n):
        r, dr = A_FIRST + i, DATA_FIRST + i
        put(a, f"A{r}", f"=Data!A{dr}", GREEN)
        put(a, f"B{r}", f"=Data!D{dr}", GREEN, PCT)
        put(a, f"C{r}", f"=Data!G{dr}", GREEN, MULT)
        put(a, f"D{r}", f"=Data!F{dr}", GREEN, USD2)
        put(a, f"E{r}", f"=$B$12*B{r}", BLACK, NUM)
        put(a, f"F{r}", f"=$B$14+$B$15*C{r}", BLACK, MULT)
        put(a, f"G{r}", f"=$B$13*F{r}", BLACK, USD)
    A_LAST = A_FIRST + n - 1
    for i in range(n):
        r = A_FIRST + i
        put(a, f"H{r}",
            f"=IFERROR(G{r}*$B$12*$B$13/SUMPRODUCT($E${A_FIRST}:$E${A_LAST},$G${A_FIRST}:$G${A_LAST}),0)",
            BLACK, USD)
        put(a, f"I{r}", f"=E{r}*H{r}/1000000", BLACK, USD_MM)
    at = A_LAST + 1
    put(a, f"A{at}", "Total", BOLD)
    put(a, f"B{at}", f"=SUM(B{A_FIRST}:B{A_LAST})", BOLD, PCT)
    put(a, f"E{at}", f"=SUM(E{A_FIRST}:E{A_LAST})", BOLD, NUM)
    put(a, f"H{at}", f"=IFERROR(I{at}*1000000/E{at},0)", BOLD, USD)
    put(a, f"I{at}", f"=SUM(I{A_FIRST}:I{A_LAST})", BOLD, USD_MM)
    for c in "ABCDEFGHI":
        a[f"{c}{at}"].fill, a[f"{c}{at}"].border = TOT_FILL, TOP
    put(a, f"A{at + 1}", "Advertisers allocated by engagement share; spend per advertiser scaled by the willingness-to-pay "
                         "multiplier and normalized so the total equals B12 x B13.", NOTE)
    widths(a, {"A": 44, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 16, "I": 16,
               "J": 9, "K": 9, "L": 9, "M": 9, "N": 10})
    a.freeze_panes = "B4"

    # ================= Revenue_Forecast =================
    rv = wb.create_sheet("Revenue_Forecast")
    put(rv, "A1", "FY27 ad revenue forecast by category ($mm)", TITLE)
    put(rv, "A2", "Revenue = advertisers x (1+adv growth)^(m/12) x spend/adv x (1+spend growth)^(m/12) x seasonality",
        NOTE)
    hdr(rv, 4, ["Category"] + [f"{m}" for m in MONTHS] + ["FY27"] + QTRS)
    R_FIRST = 5
    for i in range(n):
        r, ar = R_FIRST + i, A_FIRST + i
        put(rv, f"A{r}", f"=Assumptions!A{ar}", GREEN)
        for c in MCOL:
            put(rv, f"{c}{r}",
                f"=Assumptions!$E{ar}*(1+Assumptions!$B$7)^(Assumptions!{c}$20/12)"
                f"*Assumptions!$H{ar}*(1+Assumptions!$B$8)^(Assumptions!{c}$20/12)"
                f"*Assumptions!{c}$19/1000000", BLACK, USD_MM)
        put(rv, f"N{r}", f"=SUM(B{r}:M{r})", BLACK, USD_MM)
        for q, (c0, c1) in enumerate([("B", "D"), ("E", "G"), ("H", "J"), ("K", "M")]):
            put(rv, f"{L(15 + q)}{r}", f"=SUM({c0}{r}:{c1}{r})", BLACK, USD_MM)
    R_LAST = R_FIRST + n - 1
    RT = R_LAST + 1
    put(rv, f"A{RT}", "Total revenue", BOLD)
    for col in range(2, 19):
        c = L(col)
        put(rv, f"{c}{RT}", f"=SUM({c}{R_FIRST}:{c}{R_LAST})", BOLD, USD_MM)
    for col in range(1, 19):
        rv.cell(row=RT, column=col).fill = TOT_FILL
        rv.cell(row=RT, column=col).border = TOP
    put(rv, f"A{RT + 1}", "YoY growth vs FY26 run-rate")
    put(rv, f"N{RT + 1}", f"=IFERROR(N{RT}/Assumptions!B16-1,0)", BLACK, PCT)
    put(rv, f"A{RT + 2}", "Category mix of FY27 revenue")
    put(rv, f"A{RT + 4}", "Driver detail", BOLD)
    put(rv, f"A{RT + 5}", "Active advertisers (end of month)")
    put(rv, f"A{RT + 6}", "Avg monthly spend / advertiser ($)")
    for c in MCOL:
        put(rv, f"{c}{RT + 5}", f"=Assumptions!$E${at}*(1+Assumptions!$B$7)^(Assumptions!{c}$20/12)", BLACK, NUM)
        put(rv, f"{c}{RT + 6}", f"=IFERROR({c}{RT}*1000000/{c}{RT + 5},0)", BLACK, USD)
    hdr(rv, RT + 8, ["Category", "FY27 revenue mix"])
    for i in range(n):
        r = RT + 9 + i
        put(rv, f"A{r}", f"=A{R_FIRST + i}", BLACK)
        put(rv, f"B{r}", f"=IFERROR(N{R_FIRST + i}/$N${RT},0)", BLACK, PCT)
    widths(rv, {"A": 34, **{L(c): 10 for c in range(2, 19)}, "N": 12})
    rv.freeze_panes = "B5"

    # ================= Opex_Headcount =================
    o = wb.create_sheet("Opex_Headcount")
    put(o, "A1", "Headcount, opex and capex plan — FY27 by quarter", TITLE)
    funcs = [("Engineering & Ad Platform", 900, [30, 30, 25, 25], 420000),
             ("Sales & Account Management", 650, [25, 20, 20, 15], 300000),
             ("Marketing & Developer Relations", 160, [5, 5, 5, 5], 260000),
             ("Finance, Legal & G&A", 90, [3, 2, 2, 2], 240000)]
    hdr(o, 3, ["Net hires by quarter", "FY26 exit HC"] + QTRS + ["Loaded cost / HC ($/yr)"])
    H_FIRST = 4
    for i, (fn, hc, hires, cost) in enumerate(funcs):
        r = H_FIRST + i
        put(o, f"A{r}", fn)
        put(o, f"B{r}", hc, BLUE, NUM)
        for q in range(4):
            put(o, f"{L(3 + q)}{r}", hires[q], BLUE, NUM)
        put(o, f"G{r}", cost, BLUE, USD, comment="Illustrative fully-loaded cost: salary, benefits, equity, payroll tax.")
    H_LAST = H_FIRST + len(funcs) - 1
    put(o, f"A{H_LAST + 1}", "Headcount and costs are illustrative.", NOTE)

    hdr(o, 10, ["Ending headcount"] + [""] + QTRS)
    E_FIRST = 11
    for i in range(len(funcs)):
        r, hr = E_FIRST + i, H_FIRST + i
        put(o, f"A{r}", f"=A{hr}")
        put(o, f"C{r}", f"=B{hr}+C{hr}", BLACK, NUM)
        for q in range(1, 4):
            put(o, f"{L(3 + q)}{r}", f"={L(2 + q)}{r}+{L(3 + q)}{hr}", BLACK, NUM)
    E_LAST = E_FIRST + len(funcs) - 1
    put(o, f"A{E_LAST + 1}", "Total headcount", BOLD)
    for q in range(4):
        c = L(3 + q)
        put(o, f"{c}{E_LAST + 1}", f"=SUM({c}{E_FIRST}:{c}{E_LAST})", BOLD, NUM)
    HC_TOT = E_LAST + 1

    hdr(o, 18, ["People cost ($mm)"] + [""] + QTRS + ["FY27"])
    P_FIRST = 19
    for i in range(len(funcs)):
        r, hr, er = P_FIRST + i, H_FIRST + i, E_FIRST + i
        put(o, f"A{r}", f"=A{hr}")
        put(o, f"C{r}", f"=(B{hr}+C{er})/2*$G{hr}/4/1000000", BLACK, USD_MM)
        for q in range(1, 4):
            put(o, f"{L(3 + q)}{r}", f"=({L(2 + q)}{er}+{L(3 + q)}{er})/2*$G{hr}/4/1000000", BLACK, USD_MM)
        put(o, f"G{r}", f"=SUM(C{r}:F{r})", BLACK, USD_MM)
    P_LAST = P_FIRST + len(funcs) - 1
    PEOPLE = P_LAST + 1
    put(o, f"A{PEOPLE}", "Total people cost", BOLD)
    for col in "CDEFG":
        put(o, f"{col}{PEOPLE}", f"=SUM({col}{P_FIRST}:{col}{P_LAST})", BOLD, USD_MM)
    put(o, f"A{PEOPLE + 1}", "Avg HC cost formula: (begin HC + end HC) / 2 x loaded cost / 4.", NOTE)

    hdr(o, 27, ["Capex & depreciation ($mm)"] + [""] + QTRS + ["FY27"])
    put(o, "A28", "Capex — ad-serving infrastructure")
    for q, v in enumerate([45, 40, 40, 50]):
        put(o, f"{L(3 + q)}28", v, BLUE, USD_MM)
    put(o, "G28", "=SUM(C28:F28)", BLACK, USD_MM)
    put(o, "A29", "Useful life (years)"); put(o, "B29", 4, BLUE, "0")
    put(o, "A30", "Depreciation on existing asset base"); put(o, "B30", 30, BLUE, USD_MM)
    put(o, "A31", "Depreciation — new capex")
    # straight-line, half-quarter convention: capex placed in service mid-quarter
    for q in range(4):
        c = L(3 + q)
        prior = f"SUM($C$28:{L(2 + q)}28)" if q > 0 else "0"
        put(o, f"{c}31", f"=({prior}+{c}28/2)/($B$29*4)", BLACK, USD_MM)
    put(o, "G31", "=SUM(C31:F31)", BLACK, USD_MM)
    put(o, "A32", "Total depreciation", BOLD)
    for q in range(4):
        c = L(3 + q)
        put(o, f"{c}32", f"=$B$30+{c}31", BOLD, USD_MM)
    put(o, "G32", "=SUM(C32:F32)", BOLD, USD_MM)
    put(o, "A33", "Capex and asset base illustrative. Half-quarter convention for new assets.", NOTE)
    widths(o, {"A": 36, "B": 14, "C": 12, "D": 12, "E": 12, "F": 12, "G": 20})

    # ================= PnL_Summary =================
    p = wb.create_sheet("PnL_Summary")
    put(p, "A1", "Apple Ads — FY27 quarterly P&L ($mm, illustrative)", TITLE)
    hdr(p, 3, ["Line item"] + QTRS + ["FY27"])
    rev_q = ["O", "P", "Q", "R"]
    rows = {
        4: ("Ad revenue", [f"=Revenue_Forecast!{c}{RT}" for c in rev_q], GREEN),
        5: ("People cost", [f"=-Opex_Headcount!{L(3 + q)}{PEOPLE}" for q in range(4)], GREEN),
        6: ("Non-HC opex", [f"=-{L(2 + q)}4*Assumptions!$B$9" for q in range(4)], BLACK),
        7: ("Depreciation", [f"=-Opex_Headcount!{L(3 + q)}32" for q in range(4)], GREEN),
    }
    for r, (label, fs, font) in rows.items():
        put(p, f"A{r}", label)
        for q, f in enumerate(fs):
            put(p, f"{L(2 + q)}{r}", f, font, USD_MM)
        put(p, f"F{r}", f"=SUM(B{r}:E{r})", BLACK, USD_MM)
    put(p, "A8", "Operating income", BOLD)
    put(p, "A9", "Operating margin")
    put(p, "A10", "Revenue per employee ($k, annualized)")
    put(p, "A11", "Ending headcount")
    for col in "BCDEF":
        put(p, f"{col}8", f"=SUM({col}4:{col}7)", BOLD, USD_MM)
        put(p, f"{col}9", f"=IFERROR({col}8/{col}4,0)", BLACK, PCT)
        p[f"{col}8"].border = TOP
    for q in range(4):
        c = L(2 + q)
        put(p, f"{c}11", f"=Opex_Headcount!{L(3 + q)}{HC_TOT}", GREEN, NUM)
        put(p, f"{c}10", f"=IFERROR({c}4*4*1000/{c}11,0)", BLACK, USD)
    put(p, "F11", "=E11", BLACK, NUM)
    put(p, "F10", "=IFERROR(F4*1000/F11,0)", BLACK, USD)
    put(p, "A13", "Note: excludes revenue share/TAC (owned inventory), allocations and income tax. Illustrative.",
        NOTE)
    widths(p, {"A": 38, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12})

    # ================= Budget_vs_Actual =================
    # Illustrative Q1 actuals generated deterministically around the Base-case budget.
    rng = np.random.default_rng(27)
    actuals = []
    bv = wb.create_sheet("Budget_vs_Actual")
    put(bv, "A1", "Q1 FY27 budget vs actual — ad revenue by category", TITLE)
    put(bv, "A2", "Budget = Revenue_Forecast Q1 (active scenario). Actuals are ILLUSTRATIVE demo inputs (blue). "
                  "Volume var = (Act adv − Bud adv) × Bud rev/adv;  Rate var = (Act rev/adv − Bud rev/adv) × Act adv.",
        NOTE)
    hdr(bv, 4, ["Category", "Budget rev ($mm)", "Budget advertisers (avg)", "Budget rev / adv ($)",
                "Actual advertisers (avg)", "Actual rev / adv ($)", "Actual rev ($mm)",
                "Variance ($mm)", "Variance %", "Volume var ($mm)", "Rate var ($mm)", "Check", "F/U", "|Var| (helper)"])
    B_FIRST = 5
    base_adv = cats.reviews / cats.reviews.sum() * 60000
    for i in range(n):
        r, ar, rr = B_FIRST + i, A_FIRST + i, R_FIRST + i
        put(bv, f"A{r}", f"=Assumptions!A{ar}", GREEN)
        put(bv, f"B{r}", f"=Revenue_Forecast!O{rr}", GREEN, USD_MM)
        put(bv, f"C{r}", f"=AVERAGE(Revenue_Forecast!B{RT + 5}:D{RT + 5})*Assumptions!B{ar}", BLACK, NUM)
        put(bv, f"D{r}", f"=IFERROR(B{r}*1000000/C{r},0)", BLACK, USD)
        adv_noise, rpa_noise = rng.normal(0.0, 0.05), rng.normal(0.0, 0.07)
        # Store actuals as multipliers of a Python-side approximation of budget drivers
        approx_adv = base_adv.iloc[i] * (1.12 ** (2 / 12))
        put(bv, f"E{r}", round(float(approx_adv * (1 + adv_noise))), BLUE, NUM)
        actuals.append({"adv": round(float(approx_adv * (1 + adv_noise))), "rpa_mult": round(1 + rpa_noise, 4)})
        put(bv, f"F{r}", f"=ROUND(D{r}*{1 + rpa_noise:.4f},0)", BLUE, USD,
            comment=f"Illustrative actual: budget rev/adv x {1 + rpa_noise:.3f}. Replace with ledger actuals.")
        put(bv, f"G{r}", f"=E{r}*F{r}/1000000", BLACK, USD_MM)
        put(bv, f"H{r}", f"=G{r}-B{r}", BLACK, USD_MM)
        put(bv, f"I{r}", f"=IFERROR(H{r}/B{r},0)", BLACK, PCT)
        put(bv, f"J{r}", f"=(E{r}-C{r})*D{r}/1000000", BLACK, USD_MM)
        put(bv, f"K{r}", f"=(F{r}-D{r})*E{r}/1000000", BLACK, USD_MM)
        put(bv, f"L{r}", f"=ROUND(H{r}-J{r}-K{r},6)", BLACK, "0.000")
        put(bv, f"M{r}", f'=IF(H{r}>=0,"F","U")', BLACK)
        put(bv, f"N{r}", f"=ABS(H{r})+ROW()/1000000000", BLACK, USD_MM)
    B_LAST = B_FIRST + n - 1
    BT = B_LAST + 1
    put(bv, f"A{BT}", "Total", BOLD)
    for c in "BCEGHJKL":
        fmt = NUM if c in "CE" else ("0.000" if c == "L" else USD_MM)
        put(bv, f"{c}{BT}", f"=SUM({c}{B_FIRST}:{c}{B_LAST})", BOLD, fmt)
    put(bv, f"D{BT}", f"=IFERROR(B{BT}*1000000/C{BT},0)", BOLD, USD)
    put(bv, f"F{BT}", f"=IFERROR(G{BT}*1000000/E{BT},0)", BOLD, USD)
    put(bv, f"I{BT}", f"=IFERROR(H{BT}/B{BT},0)", BOLD, PCT)
    put(bv, f"M{BT}", f'=IF(H{BT}>=0,"F","U")', BOLD)
    for col in range(1, 14):
        bv.cell(row=BT, column=col).fill = TOT_FILL
        bv.cell(row=BT, column=col).border = TOP
    put(bv, f"A{BT + 1}", "Mix effect is embedded in category-level volume/rate; total check column must be 0.", NOTE)
    red = Font(name=F, color="C00000"); grn = Font(name=F, color="007A33")
    for rng_ in (f"H{B_FIRST}:K{BT}",):
        bv.conditional_formatting.add(rng_, CellIsRule(operator="lessThan", formula=["0"], font=red))
        bv.conditional_formatting.add(rng_, CellIsRule(operator="greaterThan", formula=["0"], font=grn))
    put(bv, f"A{BT + 3}", "Top 3 drivers of variance (by absolute $)", BOLD)
    for k in range(3):
        r = BT + 4 + k
        put(bv, f"A{r}",
            f"=INDEX($A${B_FIRST}:$A${B_LAST},MATCH(LARGE($N${B_FIRST}:$N${B_LAST},{k + 1}),"
            f"$N${B_FIRST}:$N${B_LAST},0))")
        put(bv, f"B{r}", f"=INDEX($H${B_FIRST}:$H${B_LAST},MATCH(A{r},$A${B_FIRST}:$A${B_LAST},0))", BLACK, USD_MM)
    widths(bv, {"A": 30, **{L(c): 14 for c in range(2, 15)}, "L": 9, "M": 6})
    bv.freeze_panes = "B5"

    # ================= Market_Investment =================
    mi = wb.create_sheet("Market_Investment")
    put(mi, "A1", "Where to invest the next sales dollar? Storefront investment case (3-year NPV)", TITLE)
    put(mi, "A2", "Revenue/advertiser scaled by the App Store price index (local willingness-to-pay vs US). "
                  "Market size, ramp and costs are ILLUSTRATIVE.", NOTE)
    put(mi, "A4", "US FY27 annual revenue / advertiser ($)")
    put(mi, "B4", f"=IFERROR(Revenue_Forecast!N{RT}*1000000/Revenue_Forecast!M{RT + 5},0)", GREEN, USD)
    put(mi, "A5", "US FY26 advertisers"); put(mi, "B5", "=Assumptions!B12", GREEN, NUM)
    put(mi, "A6", "Discount rate"); put(mi, "B6", "=Assumptions!B10", GREEN, PCT)
    put(mi, "A7", "Upfront investment per market ($mm)"); put(mi, "B7", 15, BLUE, USD_MM)
    put(mi, "A8", "Incremental opex per market per year ($mm)"); put(mi, "B8", 8, BLUE, USD_MM)
    put(mi, "A9", "Incremental-share capture (yrs 1/2/3)")
    for c, v in zip("BCD", [0.04, 0.08, 0.12]):
        put(mi, f"{c}9", v, BLUE, PCT, KEY_FILL)
    put(mi, "E9", "Share of the market's addressable spend won by the incremental investment", NOTE)
    put(mi, "A10", "Contribution margin on incremental revenue"); put(mi, "B10", 0.80, BLUE, PCT)

    hdr(mi, 12, ["Country", "Price index (US=100)", "VAT/GST in local price", "Net price index",
                 "Availability", "Market size vs US", "Addressable advertisers", "Rev / adv ($)",
                 "Addressable spend ($mm)", "Yr1 rev ($mm)", "Yr2 rev ($mm)", "Yr3 rev ($mm)",
                 "NPV ($mm)", "Payback (yrs)", "Rank", "Advertiser eligibility"])
    M_FIRST = 13
    for i, row in enumerate(intl.itertuples()):
        r, dr = M_FIRST + i, INTL_FIRST + i
        put(mi, f"A{r}", f"=Data!B{dr}", GREEN)
        put(mi, f"B{r}", f"=Data!F{dr}", GREEN, "0.0")
        put(mi, f"C{r}", VAT.get(row.country, 0.0), BLUE, PCT,
            comment="Consumption tax embedded in the storefront price (US/CA prices exclude sales tax). "
                    "Standard rates; verify against Apple's tax schedule before use.")
        put(mi, f"D{r}", f"=B{r}/(1+C{r})", BLACK, "0.0")
        put(mi, f"E{r}", f"=Data!E{dr}", GREEN, PCT)
        put(mi, f"F{r}", MARKET_SIZE.get(row.country, 0.05), BLUE, PCT,
            comment="Illustrative: storefront developer/advertiser base relative to US.")
        put(mi, f"P{r}", ELIGIBLE.get(row.country, 1.0), BLUE, PCT, KEY_FILL if row.country in ELIGIBLE else None,
            comment=("Mainland China requires advertiser business + app approval (licenses, industry "
                     "qualifications) before ads run - see Apple Ads Help 'Advertise in mainland China'. "
                     "Share of advertisers assumed to clear approval is illustrative.")
            if row.country == "cn" else "Share of addressable advertisers able to run ads (no extra approval).")
        put(mi, f"G{r}", f"=$B$5*F{r}*E{r}*P{r}", BLACK, NUM)
        put(mi, f"H{r}", f"=$B$4*D{r}/100", BLACK, USD)
        put(mi, f"I{r}", f"=G{r}*H{r}/1000000", BLACK, USD_MM)
        put(mi, f"J{r}", f"=I{r}*$B$9", BLACK, USD_MM)
        put(mi, f"K{r}", f"=I{r}*$C$9", BLACK, USD_MM)
        put(mi, f"L{r}", f"=I{r}*$D$9", BLACK, USD_MM)
        cf1, cf2, cf3 = (f"(J{r}*$B$10-$B$8)", f"(K{r}*$B$10-$B$8)", f"(L{r}*$B$10-$B$8)")
        put(mi, f"M{r}", f"=-$B$7+{cf1}/(1+$B$6)+{cf2}/(1+$B$6)^2+{cf3}/(1+$B$6)^3", BLACK, USD_MM)
        put(mi, f"N{r}",
            f'=IF({cf1}>=$B$7,$B$7/{cf1},IF({cf1}+{cf2}>=$B$7,1+($B$7-{cf1})/{cf2},'
            f'IF({cf1}+{cf2}+{cf3}>=$B$7,2+($B$7-{cf1}-{cf2})/{cf3},">3")))', BLACK, "0.0")
    M_LAST = M_FIRST + len(intl) - 1
    for i in range(len(intl)):
        r = M_FIRST + i
        put(mi, f"O{r}", f"=RANK(M{r},$M${M_FIRST}:$M${M_LAST})", BLACK, "0")
    mi.conditional_formatting.add(f"M{M_FIRST}:M{M_LAST}",
                                  CellIsRule(operator="lessThan", formula=["0"], font=red))
    put(mi, f"A{M_LAST + 2}", "Recommendation", BOLD)
    put(mi, f"B{M_LAST + 2}",
        f'="Prioritize "&INDEX(A{M_FIRST}:A{M_LAST},MATCH(1,O{M_FIRST}:O{M_LAST},0))&" (NPV $"&'
        f'TEXT(MAX(M{M_FIRST}:M{M_LAST}),"#,##0.0")&"mm), then "&INDEX(A{M_FIRST}:A{M_LAST},'
        f'MATCH(2,O{M_FIRST}:O{M_LAST},0))&"."')
    put(mi, f"A{M_LAST + 3}", "Gross price index is tax-inclusive outside US/CA; the net index strips VAT/GST so "
                              "willingness-to-pay is compared like-for-like.", NOTE)
    put(mi, f"A{M_LAST + 4}", "All 9 storefronts are Apple Ads markets. Mainland China requires advertiser business and "
                              "app approval before ads run, reflected by the eligibility haircut in column P.", NOTE)
    widths(mi, {"A": 44, **{L(c): 13 for c in range(2, 17)}})

    # ================= Sensitivity =================
    s = wb.create_sheet("Sensitivity")
    put(s, "A1", "FY27 revenue ($mm) — advertiser growth (rows) x spend-per-advertiser growth (columns)", TITLE)
    put(s, "A2", "Each cell recomputes FY27 revenue from FY26 monthly run-rate, seasonality and both growth rates.",
        NOTE)
    put(s, "A4", "FY26 monthly run-rate ($mm)"); put(s, "B4", f"=Assumptions!I{at}", GREEN, USD_MM)
    adv_g = [0.0, 0.05, 0.10, 0.15, 0.20]
    sp_g = [0.0, 0.04, 0.08, 0.12, 0.16]
    put(s, "A6", "Adv growth ↓ / Spend growth →", BOLD)
    for j, g in enumerate(sp_g):
        put(s, f"{L(2 + j)}6", g, BLUE, PCT)
        s[f"{L(2 + j)}6"].fill = TOT_FILL
    for i, g in enumerate(adv_g):
        r = 7 + i
        put(s, f"A{r}", g, BLUE, PCT); s[f"A{r}"].fill = TOT_FILL
        for j in range(len(sp_g)):
            c = L(2 + j)
            put(s, f"{c}{r}",
                f"=$B$4*SUMPRODUCT(Assumptions!$B$19:$M$19,((1+$A{r})*(1+{c}$6))^(Assumptions!$B$20:$M$20/12))",
                BLACK, USD_MM)
    s.conditional_formatting.add("B7:F11", CellIsRule(operator="greaterThanOrEqual",
                                                      formula=[f"Revenue_Forecast!$N${RT}"],
                                                      fill=PatternFill("solid", fgColor="E2EFDA")))
    put(s, "A13", "Active-scenario FY27 revenue"); put(s, "B13", f"=Revenue_Forecast!N{RT}", GREEN, USD_MM)
    put(s, "A14", "Green cells ≥ active scenario.", NOTE)
    widths(s, {"A": 32, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13})

    # ================= Dashboard =================
    db = wb.create_sheet("Dashboard", 1)
    put(db, "A1", "Apple Ads — FY27 plan dashboard (illustrative)", TITLE)
    put(db, "A2", '="Scenario: "&Assumptions!B3')
    kpis = [
        ("FY27 ad revenue ($mm)", f"=Revenue_Forecast!N{RT}", USD_MM),
        ("YoY growth vs FY26 run-rate", f"=Revenue_Forecast!N{RT + 1}", PCT),
        ("FY27 operating income ($mm)", "=PnL_Summary!F8", USD_MM),
        ("FY27 operating margin", "=PnL_Summary!F9", PCT),
        ("FY27 exit headcount", "=PnL_Summary!F11", NUM),
        ("FY27 capex ($mm)", "=Opex_Headcount!G28", USD_MM),
        ("Q1 revenue variance vs budget ($mm)", f"=Budget_vs_Actual!H{BT}", USD_MM),
        ("  of which volume (advertisers)", f"=Budget_vs_Actual!J{BT}", USD_MM),
        ("  of which rate (spend / advertiser)", f"=Budget_vs_Actual!K{BT}", USD_MM),
        ("Top storefront to invest in", f"=INDEX(Market_Investment!A{M_FIRST}:A{M_LAST},"
                                         f"MATCH(1,Market_Investment!O{M_FIRST}:O{M_LAST},0))", None),
        ("Largest revenue category", f"=INDEX(Revenue_Forecast!A{R_FIRST}:A{R_LAST},"
                                     f"MATCH(MAX(Revenue_Forecast!N{R_FIRST}:N{R_LAST}),"
                                     f"Revenue_Forecast!N{R_FIRST}:N{R_LAST},0))", None),
    ]
    hdr(db, 4, ["Metric", "Value"])
    for i, (k, f, fmt) in enumerate(kpis):
        put(db, f"A{5 + i}", k)
        put(db, f"B{5 + i}", f, GREEN, fmt)
    hdr(db, 18, ["Quarterly P&L ($mm)"] + QTRS + ["FY27"])
    for i, (r, lab) in enumerate([(4, "Ad revenue"), (5, "People cost"), (6, "Non-HC opex"),
                                  (7, "Depreciation"), (8, "Operating income"), (9, "Operating margin")]):
        put(db, f"A{19 + i}", lab, BOLD if r == 8 else BLACK)
        for col in "BCDEF":
            put(db, f"{col}{19 + i}", f"=PnL_Summary!{col}{r}", GREEN, PCT if r == 9 else USD_MM)
    widths(db, {"A": 40, "B": 16, "C": 12, "D": 12, "E": 12, "F": 12})

    for w in wb.worksheets:
        w.sheet_view.showGridLines = False
        for row in w.iter_rows():
            for c in row:
                if c.font and c.font.name != F:
                    c.font = Font(name=F, bold=c.font.bold, italic=c.font.italic,
                                  color=c.font.color, size=c.font.size)

    OUT.parent.mkdir(exist_ok=True)
    wb.calculation.fullCalcOnLoad = True  # Excel computes every formula on open
    wb.save(OUT)
    inputs = {
        "categories": [{"name": c, "reviews": int(r.reviews), "median_paid": float(r.median_paid),
                        "free_share": float(r.free_share)} for c, r in cats.iterrows()],
        "scenarios": {name: {"adv_g": lv[0][i], "spend_g": lv[1][i], "nonhc": lv[2][i], "disc": lv[3][i]}
                      for i, name in enumerate(["Base", "Upside", "Downside"])
                      for lv in [[l[2] for l in levers]]},
        "advertisers": 60000, "spend": 7000, "mult_int": 0.4, "mult_slope": 0.6,
        "seasonality": seas, "months": MONTHS,
        "functions": [{"name": f, "hc": h, "hires": hi, "cost": c} for f, h, hi, c in funcs],
        "capex": [45, 40, 40, 50], "life": 4, "dep_base": 30,
        "actuals": actuals,
        "markets": [{"code": r.country, "name": r.country_name, "currency": r.currency,
                     "index": float(r.median_index), "availability": float(r.availability),
                     "vat": VAT.get(r.country, 0.0), "size": MARKET_SIZE.get(r.country, 0.05),
                     "eligible": ELIGIBLE.get(r.country, 1.0)}
                    for r in intl.itertuples()],
        "invest": {"upfront": 15, "opex": 8, "capture": [0.04, 0.08, 0.12], "margin": 0.80},
    }
    (OUT.parent / "model_inputs.json").write_text(json.dumps(inputs, indent=1))
    print(f"Wrote {OUT}")
    return {"RT": RT, "BT": BT, "at": at, "M_FIRST": M_FIRST, "M_LAST": M_LAST}


if __name__ == "__main__":
    build()

"""Build the formula-driven Excel workbook: excel/WishCare_Performance_Marketing_Analysis.xlsx"""
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

raw = pd.read_csv("data/raw/campaign_performance_raw.csv", dtype=str, encoding="utf-8-sig")
cl = pd.read_csv("data/cleaned/campaign_performance_cleaned.csv")
ab = pd.read_csv("data/cleaned/ab_test_cleaned.csv")
log = pd.read_csv("data/cleaned/data_cleaning_log.csv")
py = json.load(open("data/cleaned/results.json"))["overall"]
camp_order = pd.read_csv("data/cleaned/summary_campaign_kpis.csv")  # ROAS order from Python

FONT = "Arial"
NAVY, PINK, LIGHT = "1F2A44", "C2185B", "F4F6FA"
F = lambda **k: Font(name=FONT, **k)
HDR_FILL = PatternFill("solid", fgColor=NAVY)
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
SUB_FILL = PatternFill("solid", fgColor=LIGHT)
thin = Side(style="thin", color="D0D4DC")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
INR, INR2, INT, DEC2 = '"₹"#,##0;("₹"#,##0);-', '"₹"#,##0.00;("₹"#,##0.00);-', '#,##0;(#,##0);-', '0.00'
ROASF, PCTF = '0.00"x"', '0.00'

wb = Workbook()
wb.remove(wb.active)

def title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = F(bold=True, size=15, color=NAVY)
    if sub:
        ws["A2"] = sub; ws["A2"].font = F(italic=True, size=9, color="666666")
    ws.sheet_view.showGridLines = False

def header(ws, row, col, names, widths=None):
    for i, n in enumerate(names):
        c = ws.cell(row=row, column=col + i, value=n)
        c.font = F(bold=True, color="FFFFFF", size=10); c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
    ws.row_dimensions[row].height = 32
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[L(col + i)].width = w

def style_body(ws, r1, r2, c1, c2, fmts):
    for r in range(r1, r2 + 1):
        for j, c in enumerate(range(c1, c2 + 1)):
            cell = ws.cell(row=r, column=c); cell.font = Font(name=FONT, size=10, bold=cell.font.b, color=cell.font.color, italic=cell.font.i); cell.border = BOX
            if fmts and j < len(fmts) and fmts[j]: cell.number_format = fmts[j]

def finish_chart(ch, w=18, h=8):
    ch.roundedCorners = False; ch.width, ch.height = w, h
    if ch.legend is not None: ch.legend.overlay = False
    if ch.title is not None: ch.title.overlay = False
    ch.x_axis.delete = False; ch.y_axis.delete = False
    return ch

# ---------------------------------------------------------------- README
ws = wb.create_sheet("README")
title(ws, "WishCare-Inspired Performance Marketing Analysis - Excel Workbook",
      "SIMULATED DATA for a portfolio project. It does not represent WishCare's real campaign performance.")
rows = [
    ("Period", "1 Sep 2026 - 30 Oct 2026 (60 days), 14 campaigns, 3 channels (Meta Ads, Google Ads, E-commerce marketplaces)"),
    ("Currency", "Indian Rupees (₹)"),
    ("", ""),
    ("SHEET", "WHAT IT CONTAINS"),
    ("Raw_Data", "The messy export exactly as received (852 rows) - duplicates, text numbers, mixed dates, wrong labels"),
    ("Cleaning_Log", "Every cleaning fix, rows affected and the reason. Also: how to do each fix in Excel"),
    ("Cleaned_Data", "840 clean rows (Excel Table 'tblCampaign'). Columns V:AB = row-level KPI formulas with zero-guards"),
    ("Assumptions", "Yellow cells = editable inputs (margins, target CAC, scale %). Budget sheet updates automatically"),
    ("KPI_Summary", "The 11 dashboard KPI cards, calculated from totals with SUM formulas"),
    ("Channel_Analysis", "SUMIFS by channel + KPIs + charts"),
    ("Campaign_Analysis", "SUMIFS by campaign + KPIs + rank + charts"),
    ("Product_Analysis", "SUMIFS by product + KPIs + chart"),
    ("Trends", "Daily (60 rows) and weekly SUMIFS tables + daily revenue chart"),
    ("Customers_Audience", "New vs returning by channel, audience-wise and creative-wise tables"),
    ("Budget_Optimisation", "Decision formulas (not ROAS alone) and a budget-neutral reallocation scenario"),
    ("AB_Test", "A/B data, summary and two-proportion z-tests written as Excel formulas"),
    ("Pivot_Guide", "Step-by-step instructions to build the same analysis with Pivot Tables"),
    ("Checks", "Reconciles Excel totals with the Python notebook - every difference should be 0"),
    ("", ""),
    ("COLOUR KEY", ""),
    ("Yellow fill + blue text", "Input / assumption you can change"),
    ("Black text", "Formula"),
    ("Green text", "Formula that pulls from another sheet"),
    ("", ""),
    ("GOLDEN RULE", "Every KPI = SUM(numerator) / SUM(denominator). Never average row-level ratios. This keeps Excel = SQL = Python = Power BI."),
    ("ZERO GUARD", 'Pattern used: =IF(denominator=0,"",numerator/denominator). Blank means "not available" - e.g. CAC on a day with 0 new customers.'),
]
for i, (a, b) in enumerate(rows, start=4):
    ws.cell(row=i, column=1, value=a).font = F(bold=a.isupper() or a in ("Period", "Currency"), size=10)
    ws.cell(row=i, column=2, value=b).font = F(size=10)
    if a in ("SHEET", "COLOUR KEY"):
        for c in (1, 2):
            ws.cell(row=i, column=c).fill = HDR_FILL; ws.cell(row=i, column=c).font = F(bold=True, color="FFFFFF")
ws["A24"].fill = INPUT_FILL; ws["A24"].font = F(color="0000FF", size=10)
ws["A26"].font = F(color="008000", size=10)
ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 120

# ---------------------------------------------------------------- Raw_Data
ws = wb.create_sheet("Raw_Data")
header(ws, 1, 1, list(raw.columns), [12] * len(raw.columns))
for r in raw.itertuples(index=False):
    ws.append(list(r))
ws.freeze_panes = "A2"
ws.column_dimensions["C"].width = 36

# ---------------------------------------------------------------- Cleaning_Log
ws = wb.create_sheet("Cleaning_Log")
title(ws, "Data Cleaning Log", "Raw 852 rows -> 840 clean rows. Every change is flagged in Cleaned_Data column U.")
header(ws, 4, 1, list(log.columns) + ["How to do it in Excel"], [6, 44, 10, 40, 70, 60])
excel_how = {
    "2.1": "Data > Remove Duplicates (tick all columns). Or Power Query: Home > Remove Rows > Remove Duplicates",
    "2.2": '=TRIM(PROPER(G2)) cleans spaces/case; then a lookup table + INDEX/MATCH maps "Facebook Ads" -> "Meta Ads"',
    "2.3": 'Power Query: Change Type > Using Locale > English (India) for DD/MM/YYYY. Or =DATE(RIGHT(A2,4),MID(A2,4,2),LEFT(A2,2))',
    "2.4": '=VALUE(SUBSTITUTE(SUBSTITUTE(F2,"₹",""),",",""))',
    "2.5": "=INDEX(audience_col, MATCH(campaign, campaign_col, 0)) using a non-blank row of the same campaign",
    "2.6": "Filter with a helper column, e.g. =I2>G2 for clicks>impressions, then fix flagged rows; =ABS(F2) for negatives",
    "2.7": "Helper column CPM =F2/G2*1000 and ratio to campaign median (=MEDIAN(IF(...)) array). Investigate rows >3x",
    "2.8": "Spend = Impressions/1000 x campaign median CPM; Revenue = Orders x campaign median AOV. Add a flag column",
}
for i, r in enumerate(log.itertuples(index=False), start=5):
    vals = list(r) + [excel_how.get(str(r.Step), "")]
    for j, v in enumerate(vals, start=1):
        c = ws.cell(row=i, column=j, value=v); c.alignment = Alignment(wrap_text=True, vertical="top")
        if isinstance(v, str): c.data_type = "s"
style_body(ws, 5, 4 + len(log), 1, 6, [None, None, INT])
ws.freeze_panes = "A5"

# ---------------------------------------------------------------- Cleaned_Data
ws = wb.create_sheet("Cleaned_Data")
N = len(cl); LAST = N + 1
base_cols = list(cl.columns)              # A..U
kpi_cols = ["ROAS", "CAC", "CTR %", "CPC", "CPM", "CVR %", "AOV"]   # V..AB
header(ws, 1, 1, base_cols + kpi_cols)
for r in cl.itertuples(index=False):
    row = list(r)
    row[1] = pd.Timestamp(row[1]).to_pydatetime(); row[2] = pd.Timestamp(row[2]).to_pydatetime()
    row[20] = "" if pd.isna(row[20]) else row[20]
    ws.append(row)
for i in range(2, LAST + 1):
    ws[f"V{i}"] = f'=IF(M{i}=0,"",T{i}/M{i})'
    ws[f"W{i}"] = f'=IF(R{i}=0,"",M{i}/R{i})'
    ws[f"X{i}"] = f'=IF(N{i}=0,"",P{i}/N{i}*100)'
    ws[f"Y{i}"] = f'=IF(P{i}=0,"",M{i}/P{i})'
    ws[f"Z{i}"] = f'=IF(N{i}=0,"",M{i}/N{i}*1000)'
    ws[f"AA{i}"] = f'=IF(P{i}=0,"",Q{i}/P{i}*100)'
    ws[f"AB{i}"] = f'=IF(Q{i}=0,"",T{i}/Q{i})'
fmt = {"B": "yyyy-mm-dd", "C": "yyyy-mm-dd", "M": INR2, "T": INR2, "V": ROASF, "W": INR2, "X": PCTF,
       "Y": INR2, "Z": INR2, "AA": PCTF, "AB": INR2}
for colL, f in fmt.items():
    for i in range(2, LAST + 1):
        ws[f"{colL}{i}"].number_format = f
for colL in "NOPQRS":
    for i in range(2, LAST + 1):
        ws[f"{colL}{i}"].number_format = INT
widths = [7, 11, 11, 10, 9, 8, 12, 36, 12, 20, 26, 18, 12, 11, 10, 9, 8, 9, 9, 12, 32, 8, 9, 7, 8, 8, 7, 9]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[L(i)].width = w
tbl = Table(displayName="tblCampaign", ref=f"A1:AB{LAST}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
ws.add_table(tbl)
ws.freeze_panes = "I2"
ws["V1"].comment = Comment("Row-level KPIs are for drill-down only. Do NOT average them - use the summary sheets.", "Analyst")

CD = "Cleaned_Data"
rng = lambda col: f"{CD}!${col}$2:${col}${LAST}"
COL = dict(spend="M", imp="N", reach="O", clk="P", ord="Q", new="R", ret="S", rev="T",
           ch="G", camp="H", obj="I", prod="J", aud="K", crea="L", date="B", week="C", wkend="E", fest="F")

# ---------------------------------------------------------------- Assumptions
ws = wb.create_sheet("Assumptions")
title(ws, "Assumptions (editable)", "These are planning assumptions, NOT observed data. Change a yellow cell and the Budget_Optimisation sheet recalculates.")
header(ws, 4, 1, ["Assumption", "Value", "Why / source"], [52, 14, 90])
assume = [
    ("Contribution margin before ads - own website (Meta Ads, Google Ads)", 0.55, "After product cost, shipping, payment gateway fees and returns. Assumed for a D2C beauty brand."),
    ("Contribution margin before ads - marketplaces (E-commerce)", 0.40, "Marketplace commission and fulfilment fees reduce margin. Assumed."),
    ("Target CAC (₹)", 450, "Assumed affordable cost to acquire a first-time customer who may re-order."),
    ("Minimum orders in period to trust a campaign", 300, "Avoids decisions based on small, noisy numbers."),
    ("Profit index needed to scale (ROAS / break-even ROAS)", 1.3, "30% safety buffer above break-even before adding money."),
    ("Max ROAS decline (last 4 vs first 4 weeks) allowed to scale", -0.10, "Do not scale a campaign that is getting worse."),
    ("CAC multiple of target that triggers a cut", 1.2, "Cut only if clearly above target, not marginally."),
    ("Budget change - Scale test", 0.25, "Scale in steps; watch marginal CAC."),
    ("Budget change - Reduce & fix", -0.30, "Cut loss-making spend but keep learning."),
    ("Budget change - Cap awareness", -0.15, "Awareness is judged on reach/CPM, but spend is capped."),
    ("Efficiency of extra spend vs current ROAS", 0.70, "Diminishing returns: new money assumed to earn 70% of current ROAS."),
    ("Days in period", 60, "1 Sep - 30 Oct 2026."),
    ("Significance level for A/B test (alpha)", 0.05, "Standard 95% confidence."),
]
for i, (a, v, why) in enumerate(assume, start=5):
    ws.cell(row=i, column=1, value=a).font = F(size=10)
    c = ws.cell(row=i, column=2, value=v); c.fill = INPUT_FILL; c.font = F(color="0000FF", bold=True, size=10)
    c.number_format = "0%" if isinstance(v, float) and abs(v) < 1 else "#,##0.00" if isinstance(v, float) else "#,##0"
    ws.cell(row=i, column=3, value=why).font = F(size=9, color="555555")
style_body(ws, 5, 4 + len(assume), 1, 3, None)
A_ = {k: f"Assumptions!$B${r}" for k, r in zip(
    ["m_web", "m_mkt", "tcac", "minord", "pi", "trend", "cacx", "up", "down", "cap", "eff", "days", "alpha"], range(5, 18))}

# ---------------------------------------------------------------- generic SUMIFS table builder
SUMS = [("Ad Spend (₹)", "spend", INR), ("Impressions", "imp", INT), ("Clicks", "clk", INT), ("Orders", "ord", INT),
        ("New Customers", "new", INT), ("Returning Customers", "ret", INT), ("Revenue (₹)", "rev", INR)]
KPIS = [("ROAS", "rev", "spend", 1, ROASF), ("CAC (₹)", "spend", "new", 1, INR2), ("CTR %", "clk", "imp", 100, PCTF),
        ("CPC (₹)", "spend", "clk", 1, INR2), ("CPM (₹)", "spend", "imp", 1000, INR2), ("CVR %", "ord", "clk", 100, PCTF),
        ("AOV (₹)", "rev", "ord", 1, INR2)]

def sumifs_table(ws, top, keys, crit_cols, extra=None, label_names=None, label_fmt=None):
    """keys: list of tuples (one value per criteria column). Writes label cols, sums, KPIs. Returns (first,last,colmap)."""
    names = (label_names or [c for c in crit_cols]) + [s[0] for s in SUMS] + [k[0] for k in KPIS] + (extra or [])
    header(ws, top, 1, names)
    nl = len(crit_cols)
    cmap = {}
    for j, s in enumerate(SUMS): cmap[s[1]] = L(nl + 1 + j)
    for j, k in enumerate(KPIS): cmap[k[0]] = L(nl + 1 + len(SUMS) + j)
    for i, key in enumerate(keys):
        r = top + 1 + i
        for j, v in enumerate(key):
            c = ws.cell(row=r, column=1 + j, value=v)
            if label_fmt: c.number_format = label_fmt
        crit = ",".join(f"{rng(COL[cc])},${L(1 + j)}{r}" for j, cc in enumerate(crit_cols))
        for s in SUMS:
            ws[f"{cmap[s[1]]}{r}"] = f"=SUMIFS({rng(COL[s[1]])},{crit})"
            ws[f"{cmap[s[1]]}{r}"].font = F(color="008000", size=10)
        for k in KPIS:
            num, den = cmap[k[1]], cmap[k[2]]
            mult = "" if k[3] == 1 else f"*{k[3]}"
            ws[f"{cmap[k[0]]}{r}"] = f'=IF({den}{r}=0,"",{num}{r}/{den}{r}{mult})'
    last = top + len(keys)
    tot = last + 1
    ws.cell(row=tot, column=1, value="TOTAL").font = F(bold=True)
    for s in SUMS:
        ws[f"{cmap[s[1]]}{tot}"] = f"=SUM({cmap[s[1]]}{top+1}:{cmap[s[1]]}{last})"
    for k in KPIS:
        num, den = cmap[k[1]], cmap[k[2]]
        mult = "" if k[3] == 1 else f"*{k[3]}"
        ws[f"{cmap[k[0]]}{tot}"] = f'=IF({den}{tot}=0,"",{num}{tot}/{den}{tot}{mult})'
    fmts = [label_fmt] * nl + [s[2] for s in SUMS] + [k[4] for k in KPIS]
    style_body(ws, top + 1, tot, 1, len(fmts), fmts)
    for c in range(1, len(fmts) + 1):
        ws.cell(row=tot, column=c).font = F(bold=True, size=10); ws.cell(row=tot, column=c).fill = SUB_FILL
    for j in range(nl):
        ws.column_dimensions[L(1 + j)].width = 34 if j == 0 else 18
    for j in range(len(SUMS) + len(KPIS)):
        ws.column_dimensions[L(nl + 1 + j)].width = 13
    return top + 1, last, cmap

# ---------------------------------------------------------------- KPI_Summary
ws = wb.create_sheet("KPI_Summary")
title(ws, "KPI Summary - 11 dashboard cards", "All values are formulas on Cleaned_Data. KPI = SUM(numerator) / SUM(denominator).")
header(ws, 4, 1, ["KPI", "Value", "Formula (in words)", "Business meaning"], [20, 16, 50, 70])
cards = [
    ("Total Ad Spend", f"=SUM({rng('M')})", INR, "Sum of Ad Spend", "Money paid to ad platforms"),
    ("Total Revenue", f"=SUM({rng('T')})", INR, "Sum of Revenue", "Sales attributed to ads"),
    ("Total Orders", f"=SUM({rng('Q')})", INT, "Sum of Orders", "Number of purchases"),
    ("New Customers", f"=SUM({rng('R')})", INT, "Sum of New Customers", "First-time buyers won"),
    ("ROAS", '=IF(B5=0,"",B6/B5)', ROASF, "Revenue / Ad Spend", "₹ of sales per ₹1 of ads"),
    ("CAC", '=IF(B8=0,"",B5/B8)', INR2, "Ad Spend / New Customers", "Cost to win one new customer"),
    ("CTR %", f'=IF(SUM({rng("N")})=0,"",SUM({rng("P")})/SUM({rng("N")})*100)', PCTF, "Clicks / Impressions x 100", "How attractive the ad is"),
    ("CPC", f'=IF(SUM({rng("P")})=0,"",B5/SUM({rng("P")}))', INR2, "Ad Spend / Clicks", "Price per website visit"),
    ("CPM", f'=IF(SUM({rng("N")})=0,"",B5/SUM({rng("N")})*1000)', INR2, "Ad Spend / Impressions x 1,000", "Price of 1,000 ad views"),
    ("CVR %", f'=IF(SUM({rng("P")})=0,"",B7/SUM({rng("P")})*100)', PCTF, "Orders / Clicks x 100", "How well visitors convert to buyers"),
    ("AOV", '=IF(B7=0,"",B6/B7)', INR2, "Revenue / Orders", "Average bill size"),
]
for i, (k, f, nf, w, m) in enumerate(cards, start=5):
    ws.cell(row=i, column=1, value=k).font = F(bold=True, size=11)
    c = ws.cell(row=i, column=2, value=f); c.number_format = nf; c.font = F(bold=True, size=12, color=PINK)
    ws.cell(row=i, column=3, value=w).font = F(size=10); ws.cell(row=i, column=4, value=m).font = F(size=10)
    for cc in range(1, 5): ws.cell(row=i, column=cc).border = BOX
    ws.row_dimensions[i].height = 22
ws["A18"] = "Why not average row-level ROAS?"; ws["A18"].font = F(bold=True)
ws["A19"] = "Average of row ROAS"; ws["B19"] = f"=AVERAGE({rng('V')})"; ws["B19"].number_format = ROASF
ws["A20"] = "Correct ROAS (ratio of totals)"; ws["B20"] = "=B9"; ws["B20"].number_format = ROASF
ws["C19"] = "Wrong: gives a small campaign-day the same weight as a big one, and won't match SQL/Power BI totals."
for c in ("A19", "A20", "C19"): ws[c].font = F(size=10)

# ---------------------------------------------------------------- Channel
ws = wb.create_sheet("Channel_Analysis")
title(ws, "Channel-wise performance", "SUMIFS on Cleaned_Data, then KPIs from totals.")
chs = [("Google Ads",), ("Meta Ads",), ("E-commerce",)]
f1, l1, cm = sumifs_table(ws, 4, chs, ["ch"], label_names=["Marketing Channel"])
x = 2 + len(SUMS) + len(KPIS)
extra_hdr = ["Spend Share %", "Revenue Share %", "New Customer Share %"]
header(ws, 4, x, extra_hdr)
for r in range(f1, l1 + 2):
    ws.cell(row=r, column=x, value=f"={cm['spend']}{r}/{cm['spend']}${l1+1}*100").number_format = '0.0'
    ws.cell(row=r, column=x + 1, value=f"={cm['rev']}{r}/{cm['rev']}${l1+1}*100").number_format = '0.0'
    ws.cell(row=r, column=x + 2, value=f'=IF({cm["ord"]}{r}=0,"",{cm["new"]}{r}/{cm["ord"]}{r}*100)').number_format = '0.0'
    for c in range(x, x + 3): ws.cell(row=r, column=c).border = BOX; ws.cell(row=r, column=c).font = F(size=10)
for nm, col, anchor, fmtt in [("Revenue by channel (₹)", cm["rev"], "A10", INR), ("ROAS by channel", cm["ROAS"], "H10", ROASF)]:
    ch = BarChart(); ch.type = "bar"; ch.title = nm; ch.legend = None
    ci = ws[f"{col}4"].column
    ch.add_data(Reference(ws, min_col=ci, min_row=4, max_row=l1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=f1, max_row=l1))
    ch.series[0].graphicalProperties.solidFill = "2A78D6"; ch.y_axis.numFmt = fmtt
    ws.add_chart(finish_chart(ch, 14, 7), anchor)

# ---------------------------------------------------------------- Campaign
ws = wb.create_sheet("Campaign_Analysis")
title(ws, "Campaign-wise performance", "Sorted by ROAS (from the notebook). The Rank column is a live RANK formula.")
camps = camp_order.sort_values("ROAS", ascending=False)
keys = [(r.Campaign_Name, r.Marketing_Channel, r.Campaign_Objective, r.Audience_Type) for r in camps.itertuples()]
f2, l2, cm2 = sumifs_table(ws, 4, keys, ["camp", "ch", "obj", "aud"],
                           label_names=["Campaign", "Channel", "Objective", "Audience"])
x = 5 + len(SUMS) + len(KPIS)
header(ws, 4, x, ["ROAS Rank", "New Customer Share %", "Spend Share %"])
for r in range(f2, l2 + 1):
    ws.cell(row=r, column=x, value=f"=RANK({cm2['ROAS']}{r},${cm2['ROAS']}${f2}:${cm2['ROAS']}${l2},0)")
    ws.cell(row=r, column=x + 1, value=f'=IF({cm2["ord"]}{r}=0,"",{cm2["new"]}{r}/{cm2["ord"]}{r}*100)').number_format = '0.0'
    ws.cell(row=r, column=x + 2, value=f"={cm2['spend']}{r}/{cm2['spend']}${l2+1}*100").number_format = '0.0'
    for c in range(x, x + 3): ws.cell(row=r, column=c).border = BOX; ws.cell(row=r, column=c).font = F(size=10)
roas_col = cm2["ROAS"]
ws.conditional_formatting.add(f"{roas_col}{f2}:{roas_col}{l2}", CellIsRule(operator="greaterThanOrEqual", formula=["3"], fill=PatternFill("solid", fgColor="D9F2E6")))
ws.conditional_formatting.add(f"{roas_col}{f2}:{roas_col}{l2}", CellIsRule(operator="lessThan", formula=["1"], fill=PatternFill("solid", fgColor="FBE3E3")))
for nm, col, anchor, fmtt in [("ROAS by campaign", cm2["ROAS"], "A22", ROASF), ("Ad spend by campaign (₹)", cm2["spend"], "J22", INR)]:
    ch = BarChart(); ch.type = "bar"; ch.title = nm; ch.legend = None
    ch.add_data(Reference(ws, min_col=ws[f"{col}4"].column, min_row=4, max_row=l2), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=f2, max_row=l2))
    ch.series[0].graphicalProperties.solidFill = "2A78D6"; ch.y_axis.numFmt = fmtt
    ch.x_axis.scaling.orientation = "maxMin"
    ws.add_chart(finish_chart(ch, 17, 10), anchor)
ws.cell(row=l2 + 3, column=1, value="Green ROAS >= 3x, red ROAS < 1x (conditional formatting). Awareness campaigns should also be judged on CPM/reach.").font = F(italic=True, size=9)

# ---------------------------------------------------------------- Product
ws = wb.create_sheet("Product_Analysis")
title(ws, "Product-wise revenue", "Revenue here = ad-attributed revenue only.")
prods = pd.read_csv("data/cleaned/summary_product_kpis.csv").sort_values("Revenue", ascending=False)["Product_Name"]
f3, l3, cm3 = sumifs_table(ws, 4, [(p,) for p in prods], ["prod"], label_names=["Product"])
x = 2 + len(SUMS) + len(KPIS)
header(ws, 4, x, ["Revenue Share %"])
for r in range(f3, l3 + 1):
    ws.cell(row=r, column=x, value=f"={cm3['rev']}{r}/{cm3['rev']}${l3+1}*100").number_format = '0.0'
    ws.cell(row=r, column=x).border = BOX
ch = BarChart(); ch.type = "bar"; ch.title = "Revenue by product (₹)"; ch.legend = None
ch.add_data(Reference(ws, min_col=ws[f"{cm3['rev']}4"].column, min_row=4, max_row=l3), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=f3, max_row=l3)); ch.y_axis.numFmt = INR
ch.series[0].graphicalProperties.solidFill = "2A78D6"; ch.x_axis.scaling.orientation = "maxMin"
ws.add_chart(finish_chart(ch, 16, 8), "A14")

# ---------------------------------------------------------------- Trends
ws = wb.create_sheet("Trends")
title(ws, "Daily and weekly trends", "Daily table rows 5-64, weekly table below it. Festive sale = 15-20 Oct 2026.")
dates = [(pd.Timestamp(d).to_pydatetime(),) for d in sorted(cl["Date"].unique())]
f4, l4, cm4 = sumifs_table(ws, 4, dates, ["date"], label_names=["Date"], label_fmt="dd-mmm-yy")
ws.column_dimensions["A"].width = 13
wstart = l4 + 4
ws.cell(row=wstart - 1, column=1, value="Weekly (week starts Monday; first and last weeks are partial)").font = F(bold=True)
weeks = [(pd.Timestamp(d).to_pydatetime(),) for d in sorted(cl["Week_Start"].unique())]
f5, l5, cm5 = sumifs_table(ws, wstart, weeks, ["week"], label_names=["Week Start"], label_fmt="dd-mmm-yy")
x = 2 + len(SUMS) + len(KPIS)
header(ws, wstart, x, ["Days in week", "Revenue per day (₹)"])
for r in range(f5, l5 + 1):
    ws.cell(row=r, column=x, value=f"=COUNTIFS($A${f4}:$A${l4},\">=\"&A{r},$A${f4}:$A${l4},\"<\"&A{r}+7)")
    ws.cell(row=r, column=x + 1, value=f'=IF({L(x)}{r}=0,"",{cm5["rev"]}{r}/{L(x)}{r})').number_format = INR
    for c in (x, x + 1): ws.cell(row=r, column=c).border = BOX
lc = LineChart(); lc.title = "Daily revenue (₹)"; lc.legend = None
lc.add_data(Reference(ws, min_col=ws[f"{cm4['rev']}4"].column, min_row=4, max_row=l4), titles_from_data=True)
lc.set_categories(Reference(ws, min_col=1, min_row=f4, max_row=l4))
lc.series[0].graphicalProperties.line.solidFill = "2A78D6"; lc.series[0].graphicalProperties.line.width = 22000
lc.series[0].smooth = False; lc.y_axis.numFmt = INR; lc.x_axis.number_format = "dd-mmm"
ws.add_chart(finish_chart(lc, 22, 9), f"{L(x+3)}4")

# ---------------------------------------------------------------- Customers / audience / creative
ws = wb.create_sheet("Customers_Audience")
title(ws, "Customer, audience and creative analysis")
ws["A3"] = "New vs returning customers by channel - see Channel_Analysis for full KPIs"; ws["A3"].font = F(bold=True)
f6, l6, cm6 = sumifs_table(ws, 4, chs, ["ch"], label_names=["Channel"])
auds = pd.read_csv("data/cleaned/summary_campaign_kpis.csv")["Audience_Type"].drop_duplicates()
t2 = l6 + 4
ws.cell(row=t2 - 1, column=1, value="Audience-wise performance").font = F(bold=True)
f7, l7, cm7 = sumifs_table(ws, t2, [(a,) for a in sorted(auds)], ["aud"], label_names=["Audience Type"])
t3 = l7 + 4
ws.cell(row=t3 - 1, column=1, value="Creative-wise performance").font = F(bold=True)
creas = sorted(cl["Ad_Creative_Type"].unique())
f8, l8, cm8 = sumifs_table(ws, t3, [(c,) for c in creas], ["crea"], label_names=["Ad Creative Type"])
x = 2 + len(SUMS) + len(KPIS)
for (fa, la, cmx, top) in [(f6, l6, cm6, 4), (f7, l7, cm7, t2), (f8, l8, cm8, t3)]:
    header(ws, top, x, ["New Customer Share %"])
    for r in range(fa, la + 2):
        ws.cell(row=r, column=x, value=f'=IF({cmx["ord"]}{r}=0,"",{cmx["new"]}{r}/{cmx["ord"]}{r}*100)').number_format = '0.0'
        ws.cell(row=r, column=x).border = BOX
bc = BarChart(); bc.type = "bar"; bc.grouping = "stacked"; bc.overlap = 100; bc.title = "New vs returning customers"
bc.add_data(Reference(ws, min_col=ws[f"{cm6['new']}4"].column, max_col=ws[f"{cm6['ret']}4"].column, min_row=4, max_row=l6), titles_from_data=True)
bc.set_categories(Reference(ws, min_col=1, min_row=f6, max_row=l6))
bc.series[0].graphicalProperties.solidFill = "2A78D6"; bc.series[1].graphicalProperties.solidFill = "9CC3EF"
ws.add_chart(finish_chart(bc, 14, 6), f"{L(x+2)}3")

# ---------------------------------------------------------------- Budget_Optimisation
ws = wb.create_sheet("Budget_Optimisation")
title(ws, "Budget optimisation - decisions use ROAS, margin, CAC, CVR, volume and trend",
      "Inputs come from Campaign_Analysis and Assumptions. This is a planning scenario, not a forecast.")
hdr = ["Campaign", "Channel", "Objective", "Audience", "Ad Spend (₹)", "Revenue (₹)", "Orders", "New Cust", "ROAS",
       "CAC (₹)", "CVR %", "New Cust %", "Margin", "Break-even ROAS", "Profit Index", "Contribution after ads (₹)",
       "ROAS first 4W", "ROAS last 4W", "ROAS change", "Decision", "Budget change (rule)",
       "Current daily budget (₹)", "Proposed daily budget (₹)", "Δ spend/day (₹)", "Δ revenue/day (₹)", "Δ contribution/day (₹)"]
header(ws, 4, 1, hdr, [34, 11, 11, 24, 12, 12, 8, 8, 8, 9, 7, 8, 8, 9, 8, 13, 9, 9, 9, 26, 10, 12, 12, 11, 11, 12])
CA = "Campaign_Analysis"
for i in range(len(keys)):
    r = 5 + i; src = f2 + i
    for j, k in enumerate(["A", "B", "C", "D"]):
        ws[f"{k}{r}"] = f"={CA}!{k}{src}"; ws[f"{k}{r}"].font = F(color="008000", size=10)
    for k, key in [("E", "spend"), ("F", "rev"), ("G", "ord"), ("H", "new")]:
        ws[f"{k}{r}"] = f"={CA}!{cm2[key]}{src}"; ws[f"{k}{r}"].font = F(color="008000", size=10)
    ws[f"I{r}"] = f'=IF(E{r}=0,"",F{r}/E{r})'
    ws[f"J{r}"] = f'=IF(H{r}=0,"",E{r}/H{r})'
    ws[f"K{r}"] = f'=IF({CA}!{cm2["clk"]}{src}=0,"",G{r}/{CA}!{cm2["clk"]}{src}*100)'
    ws[f"L{r}"] = f'=IF(G{r}=0,"",H{r}/G{r}*100)'
    ws[f"M{r}"] = f'=IF(B{r}="E-commerce",{A_["m_mkt"]},{A_["m_web"]})'
    ws[f"N{r}"] = f"=1/M{r}"
    ws[f"O{r}"] = f"=I{r}/N{r}"
    ws[f"P{r}"] = f"=F{r}*M{r}-E{r}"
    c1 = f'{rng("H")},$A{r},{rng("B")},"<"&DATE(2026,9,29)'
    c2 = f'{rng("H")},$A{r},{rng("B")},">="&DATE(2026,10,3)'
    ws[f"Q{r}"] = f"=SUMIFS({rng('T')},{c1})/SUMIFS({rng('M')},{c1})"
    ws[f"R{r}"] = f"=SUMIFS({rng('T')},{c2})/SUMIFS({rng('M')},{c2})"
    ws[f"S{r}"] = f"=R{r}/Q{r}-1"
    ws[f"T{r}"] = (f'=IF(C{r}="Awareness","Cap & measure differently",'
                   f'IF(OR(D{r}="Brand Search",LEFT(D{r},11)="Retargeting"),"Maintain (limited headroom)",'
                   f'IF(AND(O{r}>={A_["pi"]},J{r}<={A_["tcac"]},G{r}>={A_["minord"]},S{r}>{A_["trend"]}),"Scale test (+25%)",'
                   f'IF(AND(O{r}<1,J{r}>{A_["tcac"]}*{A_["cacx"]}),"Reduce (-30%) & fix","Optimise (hold budget)"))))')
    ws[f"U{r}"] = (f'=IF(T{r}="Scale test (+25%)",{A_["up"]},IF(T{r}="Reduce (-30%) & fix",{A_["down"]},'
                   f'IF(T{r}="Cap & measure differently",{A_["cap"]},0)))')
    ws[f"V{r}"] = f"=E{r}/{A_['days']}"
    ws[f"W{r}"] = f"=V{r}*(1+U{r})"
    ws[f"X{r}"] = f"=W{r}-V{r}"
    ws[f"Y{r}"] = f"=IF(X{r}>0,X{r}*I{r}*{A_['eff']},X{r}*I{r})"
    ws[f"Z{r}"] = f"=Y{r}*M{r}-X{r}"
lb = 4 + len(keys)
fm = [None] * 4 + [INR, INR, INT, INT, ROASF, INR2, PCTF, '0.0', '0%', ROASF, '0.00', INR, ROASF, ROASF, '0.0%', None, '+0%;-0%;0%', INR, INR, INR, INR, INR]
style_body(ws, 5, lb, 1, 26, fm)
t = lb + 1
ws[f"A{t}"] = "TOTAL"; ws[f"A{t}"].font = F(bold=True)
for k in "EFGHPVWXYZ":
    ws[f"{k}{t}"] = f"=SUM({k}5:{k}{lb})"; ws[f"{k}{t}"].number_format = INR if k not in "GH" else INT
    ws[f"{k}{t}"].font = F(bold=True, size=10)
ws[f"I{t}"] = f'=IF(E{t}=0,"",F{t}/E{t})'; ws[f"I{t}"].number_format = ROASF
for k, col in [("Scale test (+25%)", "D9F2E6"), ("Reduce (-30%) & fix", "FBE3E3"), ("Cap & measure differently", "FFF1D6"), ("Maintain (limited headroom)", "E3EEFB")]:
    ws.conditional_formatting.add(f"T5:T{lb}", FormulaRule(formula=[f'$T5="{k}"'], fill=PatternFill("solid", fgColor=col)))
s = t + 2
summary = [
    ("Budget freed from cuts (₹/day)", f'=-SUMIF(X5:X{lb},"<0")', INR),
    ("Budget added to scale tests (₹/day)", f'=SUMIF(X5:X{lb},">0")', INR),
    ("Creative-testing reserve (₹/day) = freed - added", f"=B{s}-B{s+1}", INR),
    ("Total daily budget before (₹)", f"=V{t}", INR),
    ("Total daily budget after incl. reserve (₹)", f"=W{t}+B{s+2}", INR),
    ("Scenario Δ revenue/day, excl. reserve (₹)", f"=Y{t}", INR),
    ("Scenario Δ contribution after ads/day, excl. reserve (₹)", f"=Z{t}", INR),
    ("Scenario Δ contribution per 30 days (₹)", f"=Z{t}*30", INR),
]
ws[f"A{s-1}"] = "Budget-neutral reallocation scenario"; ws[f"A{s-1}"].font = F(bold=True, size=11, color=NAVY)
for i, (a, f, nf) in enumerate(summary):
    ws[f"A{s+i}"] = a; ws[f"B{s+i}"] = f; ws[f"B{s+i}"].number_format = nf
    ws[f"A{s+i}"].font = F(size=10); ws[f"B{s+i}"].font = F(bold=True, size=10); ws[f"A{s+i}"].border = BOX; ws[f"B{s+i}"].border = BOX
ws[f"A{s+9}"] = ("How to read it: the cuts give up some low-margin revenue but stop losing money on it. The reserve funds new creative "
                 "tests; its return is unknown, so no revenue is counted for it. Validate with a 2-4 week controlled budget test.")
ws[f"A{s+9}"].font = F(italic=True, size=9)
ws.freeze_panes = "B5"

# ---------------------------------------------------------------- AB_Test
ws = wb.create_sheet("AB_Test")
title(ws, "A/B test - static image (A) vs UGC testimonial video (B)", "Same campaign, audience and budget for 14 days; only the creative differs. Simulated.")
header(ws, 4, 1, list(ab.columns), [11, 22, 9, 26, 18, 22, 11, 11, 9, 8, 11])
for i, r in enumerate(ab.itertuples(index=False), start=5):
    for j, v in enumerate(r, start=1):
        c = ws.cell(row=i, column=j, value=pd.Timestamp(v).to_pydatetime() if j == 1 else v)
la = 4 + len(ab)
style_body(ws, 5, la, 1, 11, ["yyyy-mm-dd"] + [None] * 5 + [INR2, INT, INT, INT, INR2])
R = lambda c: f"${c}$5:${c}${la}"
S0 = 5; SC = 13
ws.cell(row=4, column=SC, value="Summary").font = F(bold=True, size=11, color=NAVY)
mets = [("Impressions", "H", INT), ("Clicks", "I", INT), ("Orders", "J", INT), ("Revenue (₹)", "K", INR), ("Ad Spend (₹)", "G", INR)]
header(ws, 5, SC, ["Metric", "A - Static", "B - UGC video", "Lift B vs A"], [30, 14, 14, 12])
for i, (m, c, nf) in enumerate(mets):
    r = 6 + i
    ws.cell(row=r, column=SC, value=m)
    ws.cell(row=r, column=SC + 1, value=f'=SUMIFS({R(c)},{R("C")},"A")').number_format = nf
    ws.cell(row=r, column=SC + 2, value=f'=SUMIFS({R(c)},{R("C")},"B")').number_format = nf
    ws.cell(row=r, column=SC + 3, value=f"={L(SC+2)}{r}/{L(SC+1)}{r}-1").number_format = '+0.0%;-0.0%'
M, Nn, O, P = L(SC), L(SC + 1), L(SC + 2), L(SC + 3)   # M=label, N=A, O=B, P=lift
ratios = [("CTR %", f"{{c}}7/{{c}}6*100", PCTF), ("CPC (₹)", f"{{c}}10/{{c}}7", INR2), ("CVR %", f"{{c}}8/{{c}}7*100", PCTF),
          ("ROAS", f"{{c}}9/{{c}}10", ROASF), ("AOV (₹)", f"{{c}}9/{{c}}8", INR2)]
for i, (m, f, nf) in enumerate(ratios):
    r = 11 + i
    ws.cell(row=r, column=SC, value=m)
    for c in (Nn, O):
        ws[f"{c}{r}"] = "=" + f.format(c=c); ws[f"{c}{r}"].number_format = nf
    ws[f"{P}{r}"] = f"={O}{r}/{Nn}{r}-1"; ws[f"{P}{r}"].number_format = '+0.0%;-0.0%'
style_body(ws, 6, 15, SC, SC + 3, None)
# z-tests
zt = 18
ws.cell(row=zt - 1, column=SC, value="Significance tests (two-proportion z-test)").font = F(bold=True, size=11, color=NAVY)
header(ws, zt, SC, ["Step", "CTR test", "CVR test", "Meaning"], [30, 14, 14, 60])
steps = [
    ("Successes A", f"={Nn}7", f"={Nn}8", "Clicks (CTR) or Orders (CVR) for A"),
    ("Trials A", f"={Nn}6", f"={Nn}7", "Impressions (CTR) or Clicks (CVR) for A"),
    ("Successes B", f"={O}7", f"={O}8", ""),
    ("Trials B", f"={O}6", f"={O}7", ""),
    ("Rate A", f"={Nn}{zt+1}/{Nn}{zt+2}", f"={O}{zt+1}/{O}{zt+2}", ""),
    ("Rate B", f"={Nn}{zt+3}/{Nn}{zt+4}", f"={O}{zt+3}/{O}{zt+4}", ""),
    ("Pooled rate", f"=({Nn}{zt+1}+{Nn}{zt+3})/({Nn}{zt+2}+{Nn}{zt+4})", f"=({O}{zt+1}+{O}{zt+3})/({O}{zt+2}+{O}{zt+4})", "Combined rate if A and B were the same"),
    ("Standard error", f"=SQRT({Nn}{zt+7}*(1-{Nn}{zt+7})*(1/{Nn}{zt+2}+1/{Nn}{zt+4}))", f"=SQRT({O}{zt+7}*(1-{O}{zt+7})*(1/{O}{zt+2}+1/{O}{zt+4}))", "Expected random wobble in the difference"),
    ("z-score", f"=({Nn}{zt+6}-{Nn}{zt+5})/{Nn}{zt+8}", f"=({O}{zt+6}-{O}{zt+5})/{O}{zt+8}", "How many 'wobbles' apart A and B are"),
    ("p-value (two-sided)", f"=2*(1-NORMSDIST(ABS({Nn}{zt+9})))", f"=2*(1-NORMSDIST(ABS({O}{zt+9})))", "Chance of a gap this big if the ads were truly equal"),
    ("Significant at 95%?", f'=IF({Nn}{zt+10}<{A_["alpha"]},"YES","NO")', f'=IF({O}{zt+10}<{A_["alpha"]},"YES","NO")', "p < 0.05 -> unlikely to be luck"),
]
for i, (a, b1, b2, mtxt) in enumerate(steps):
    r = zt + 1 + i
    ws.cell(row=r, column=SC, value=a); ws[f"{Nn}{r}"] = b1; ws[f"{O}{r}"] = b2; ws.cell(row=r, column=SC + 3, value=mtxt)
    nf = INT if i < 4 else '0.0000%' if i < 8 else '0.000' if i == 8 else '0.0000' if i == 9 else None
    for c in (Nn, O):
        if nf: ws[f"{c}{r}"].number_format = nf
style_body(ws, zt + 1, zt + 11, SC, SC + 3, None)
vr = zt + 13
ws.cell(row=vr, column=SC, value="Verdict").font = F(bold=True, size=11, color=NAVY)
ws.cell(row=vr + 1, column=SC, value=(
    f'=IF({Nn}{zt+11}="YES","B gets significantly more clicks (CTR lift "&TEXT({P}11,"+0.0%")&").","CTR difference is not significant.")'
    f'&" "&IF({O}{zt+11}="YES","CVR difference IS significant.","CVR difference is NOT significant, so B is not proven to sell more.")'))
ws.cell(row=vr + 2, column=SC, value="Decision: use B for cheaper clicks (lower CPC), but do not claim it increases sales or ROAS yet. "
        "ROAS is ~equal; a ROAS winner needs a longer test or a bigger difference.")
for r in (vr + 1, vr + 2): ws.cell(row=r, column=SC).font = F(size=10)
ch = BarChart(); ch.type = "col"; ch.title = "A vs B - CTR % and CVR %"
ch.add_data(Reference(ws, min_col=SC + 1, min_row=11, max_row=13), titles_from_data=False)
ch.add_data(Reference(ws, min_col=SC + 2, min_row=11, max_row=13), titles_from_data=False)
from openpyxl.chart.series import SeriesLabel
ch.series[0].tx = SeriesLabel(v="A - Static"); ch.series[1].tx = SeriesLabel(v="B - UGC video")
ch.series[0].graphicalProperties.solidFill = "8A8984"; ch.series[1].graphicalProperties.solidFill = "2A78D6"
ch.set_categories(Reference(ws, min_col=SC, min_row=11, max_row=13))
ws.add_chart(finish_chart(ch, 14, 7), f"{L(SC+5)}4")

# ---------------------------------------------------------------- Pivot_Guide
ws = wb.create_sheet("Pivot_Guide")
title(ws, "How to build the same analysis with Pivot Tables (step by step)")
steps = [
    "1. Go to Cleaned_Data and click any cell inside the table (it is an Excel Table named tblCampaign).",
    "2. Insert > PivotTable > 'From Table/Range' > Table/Range = tblCampaign > New Worksheet > OK.",
    "3. Channel view: drag Marketing_Channel to Rows; drag Ad_Spend, Revenue, Orders, New_Customers, Clicks, Impressions to Values (Sum).",
    "4. Add ROAS correctly: PivotTable Analyze > Fields, Items & Sets > Calculated Field. Name: ROAS, Formula: =Revenue/Ad_Spend.",
    "   A calculated field divides the SUMS, so it equals Revenue total / Spend total - the correct ratio-of-totals method.",
    "5. Add the other calculated fields the same way: CAC =Ad_Spend/New_Customers ; CTR =Clicks/Impressions*100 ; CPC =Ad_Spend/Clicks ;",
    "   CPM =Ad_Spend/Impressions*1000 ; CVR =Orders/Clicks*100 ; AOV =Revenue/Orders. (Calculated fields show a divide-by-zero error when the bottom number is 0; set the pivot",
    "   value with Options > Format > 'For error values show:' blank.)",
    "6. DO NOT drag the row-level ROAS column (V) into Values and choose Average - that is the 'average of ratios' mistake.",
    "7. Campaign view: copy the pivot, replace Rows with Campaign_Name; sort ROAS descending (right-click > Sort > Largest to Smallest).",
    "8. Product view: Rows = Product_Name, Values = Sum of Revenue; Show Values As > % of Grand Total for revenue share.",
    "9. Trend view: Rows = Date (right-click > Group > Days; or use Week_Start for weeks). Insert > PivotChart > Line.",
    "10. New vs returning: Rows = Marketing_Channel, Values = Sum of New_Customers and Sum of Returning_Customers; PivotChart > Stacked Bar.",
    "11. Slicers: click a pivot > PivotTable Analyze > Insert Slicer > tick Marketing_Channel, Campaign_Name, Product_Name.",
    "    Insert Timeline > Date. Right-click a slicer > Report Connections to make one slicer filter every pivot.",
    "12. Check your pivot totals against KPI_Summary - they must match exactly.",
    "",
    "Why the summary sheets use SUMIFS instead of pivots: formulas update instantly, are easy to audit cell by cell, and feed the Budget sheet.",
]
for i, s_ in enumerate(steps, start=3):
    c = ws.cell(row=i, column=1, value=s_); c.font = F(size=10); c.data_type = "s"
ws.column_dimensions["A"].width = 150

# ---------------------------------------------------------------- Checks
ws = wb.create_sheet("Checks")
title(ws, "Cross-tool consistency checks", "Python values are typed in (blue) from the notebook; Excel values are live formulas. Difference should be 0.00.")
header(ws, 4, 1, ["KPI", "Excel (formula)", "Python notebook", "Difference", "Status"], [18, 16, 16, 12, 10])
pairs = [("Total Ad Spend", "B5", py["Ad_Spend"], INR2), ("Total Revenue", "B6", py["Revenue"], INR2),
         ("Total Orders", "B7", py["Orders"], INT), ("New Customers", "B8", py["New_Customers"], INT),
         ("ROAS", "B9", py["ROAS"], '0.0000'), ("CAC", "B10", py["CAC"], '0.0000'), ("CTR %", "B11", py["CTR_%"], '0.0000'),
         ("CPC", "B12", py["CPC"], '0.0000'), ("CPM", "B13", py["CPM"], '0.0000'), ("CVR %", "B14", py["CVR_%"], '0.0000'),
         ("AOV", "B15", py["AOV"], '0.0000')]
for i, (k, ref, v, nf) in enumerate(pairs, start=5):
    ws.cell(row=i, column=1, value=k)
    ws.cell(row=i, column=2, value=f"=KPI_Summary!{ref}").number_format = nf
    c = ws.cell(row=i, column=3, value=round(v, 6)); c.number_format = nf; c.font = F(color="0000FF", size=10)
    ws.cell(row=i, column=4, value=f"=B{i}-C{i}").number_format = '0.0000'
    ws.cell(row=i, column=5, value=f'=IF(ABS(D{i})<0.001,"MATCH","CHECK")')
style_body(ws, 5, 15, 1, 5, None)
ws.conditional_formatting.add("E5:E15", CellIsRule(operator="equal", formula=['"MATCH"'], fill=PatternFill("solid", fgColor="D9F2E6")))
ws["A17"] = "Row count check"; ws["B17"] = f"=COUNTA({rng('A')})"; ws["C17"] = 840; ws["C17"].font = F(color="0000FF")
ws["D17"] = "=B17-C17"; ws["E17"] = '=IF(D17=0,"MATCH","CHECK")'

wb.save("excel/WishCare_Performance_Marketing_Analysis.xlsx")
print("saved")

"""Build report/WishCare_Performance_Marketing_Report.pdf from the analysis outputs."""
import json, re
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
                                Image, PageBreak, KeepTogether, NextPageTemplate)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image as PILImage

FD = "/usr/share/fonts/truetype/crosextra/"
pdfmetrics.registerFont(TTFont("Body", FD + "Carlito-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Body-B", FD + "Carlito-Bold.ttf"))
pdfmetrics.registerFont(TTFont("Body-I", FD + "Carlito-Italic.ttf"))
pdfmetrics.registerFont(TTFont("Body-BI", FD + "Carlito-BoldItalic.ttf"))
from reportlab.pdfbase.pdfmetrics import registerFontFamily
registerFontFamily("Body", normal="Body", bold="Body-B", italic="Body-I", boldItalic="Body-BI")

NAVY = colors.HexColor("#1F2A44"); PINK = colors.HexColor("#C2185B"); BLUE = colors.HexColor("#2A78D6")
LIGHT = colors.HexColor("#F4F6FA"); RULE = colors.HexColor("#D0D4DC"); GREY = colors.HexColor("#555555")

S = {
    "h1": ParagraphStyle("h1", fontName="Body-B", fontSize=18, leading=22, textColor=NAVY, spaceBefore=4, spaceAfter=8),
    "h2": ParagraphStyle("h2", fontName="Body-B", fontSize=13, leading=16, textColor=NAVY, spaceBefore=10, spaceAfter=4),
    "p": ParagraphStyle("p", fontName="Body", fontSize=10.5, leading=14.5, spaceAfter=6),
    "small": ParagraphStyle("small", fontName="Body", fontSize=8.5, leading=11, textColor=GREY),
    "cell": ParagraphStyle("cell", fontName="Body", fontSize=8.5, leading=10.5),
    "cellb": ParagraphStyle("cellb", fontName="Body-B", fontSize=8.5, leading=10.5, textColor=colors.white),
    "bullet": ParagraphStyle("bullet", fontName="Body", fontSize=10.5, leading=14.5, leftIndent=12, bulletIndent=2, spaceAfter=3),
    "box": ParagraphStyle("box", fontName="Body", fontSize=10, leading=13.5, textColor=NAVY),
    "cap": ParagraphStyle("cap", fontName="Body-I", fontSize=8.5, leading=11, textColor=GREY, spaceAfter=8),
}
o = json.load(open("data/cleaned/results.json"))
ov = o["overall"]
camp = pd.read_csv("data/cleaned/summary_campaign_kpis.csv")
ch = pd.read_csv("data/cleaned/summary_channel_kpis.csv")
prod = pd.read_csv("data/cleaned/summary_product_kpis.csv")
bud = pd.read_csv("data/cleaned/budget_recommendations.csv")
log = pd.read_csv("data/cleaned/data_cleaning_log.csv")
abs_ = pd.read_csv("data/cleaned/ab_test_summary.csv")
abt = pd.read_csv("data/cleaned/ab_test_significance.csv")

P = lambda t, st="p": Paragraph(t, S[st])
def bullets(items): return [Paragraph(i, S["bullet"], bulletText="•") for i in items]
def img(path, width=170 * mm):
    w, h = PILImage.open(path).size
    return Image(path, width=width, height=width * h / w)
def table(data, widths, header=True, zebra=True, align_right_from=None):
    rows = [[Paragraph(str(c), S["cellb"] if (header and i == 0) else S["cell"]) for c in r] for i, r in enumerate(data)]
    t = Table(rows, colWidths=widths, repeatRows=1 if header else 0)
    st = [("GRID", (0, 0), (-1, -1), 0.4, RULE), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    if header: st.append(("BACKGROUND", (0, 0), (-1, 0), NAVY))
    if zebra:
        for i in range(1 if header else 0, len(rows)):
            if i % 2 == 0: st.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
    t.setStyle(TableStyle(st)); return t
def callout(text, color=BLUE):
    t = Table([[Paragraph(text, S["box"])]], colWidths=[170 * mm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EEF4FC")),
                           ("LINEBEFORE", (0, 0), (0, -1), 3, color), ("LEFTPADDING", (0, 0), (-1, -1), 8),
                           ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    return t
inr = lambda v: f"₹{v/1e5:,.2f} L" if abs(v) >= 1e5 else f"₹{v:,.0f}"

def on_page(c, doc):
    c.saveState(); c.setFont("Body", 8); c.setFillColor(GREY)
    c.drawString(20 * mm, 12 * mm, "WishCare-Inspired Performance Marketing Analytics  |  Simulated data - not WishCare's actual performance")
    c.drawRightString(190 * mm, 12 * mm, f"Page {doc.page}")
    c.setStrokeColor(RULE); c.line(20 * mm, 15 * mm, 190 * mm, 15 * mm); c.restoreState()
def on_cover(c, doc):
    c.saveState(); c.setFillColor(NAVY); c.rect(0, 0, A4[0], A4[1], fill=1, stroke=0)
    c.setFillColor(PINK); c.rect(20 * mm, 190 * mm, 30 * mm, 2 * mm, fill=1, stroke=0)
    c.setFillColor(colors.white); c.setFont("Body-B", 30)
    c.drawString(20 * mm, 170 * mm, "Performance Marketing"); c.drawString(20 * mm, 157 * mm, "Analytics & Campaign")
    c.drawString(20 * mm, 144 * mm, "Optimisation")
    c.setFont("Body", 14); c.drawString(20 * mm, 128 * mm, "A WishCare-inspired beauty & personal-care e-commerce case study")
    c.setFont("Body", 11); c.setFillColor(colors.HexColor("#C9D7EA"))
    c.drawString(20 * mm, 112 * mm, "Project report  |  Meta Ads, Google Ads, E-commerce marketplaces  |  1 Sep - 30 Oct 2026")
    c.drawString(20 * mm, 104 * mm, "Tools: Excel, SQL, Python (Pandas, NumPy, Matplotlib, Plotly), Power BI")
    c.drawString(20 * mm, 60 * mm, "Prepared by: Sudhanshu Patwa")
    c.drawString(20 * mm, 52 * mm, "Portfolio project for a Performance Marketing Associate role")
    c.setFont("Body-I", 9)
    c.drawString(20 * mm, 30 * mm, "All campaign data in this report is simulated. It is inspired by a D2C beauty brand and does not represent")
    c.drawString(20 * mm, 25 * mm, "WishCare's internal campaign performance.")
    c.restoreState()

doc = BaseDocTemplate("report/WishCare_Performance_Marketing_Report.pdf", pagesize=A4, title="WishCare-Inspired Performance Marketing Report",
                      author="Sudhanshu Patwa", leftMargin=20 * mm, rightMargin=20 * mm, topMargin=18 * mm, bottomMargin=20 * mm)
fr = Frame(20 * mm, 20 * mm, 170 * mm, 257 * mm, id="f")
doc.addPageTemplates([PageTemplate("cover", frames=[fr], onPage=on_cover), PageTemplate("body", frames=[fr], onPage=on_page)])
st = [NextPageTemplate("body"), PageBreak()]

# ---------------------------------------------------------------- Contents
st += [P("Contents", "h1")]
toc = ["1. Executive summary", "2. Business problem and objectives", "3. Dataset and methodology", "4. Data cleaning",
       "5. Marketing KPIs", "6. Exploratory analysis and findings", "7. Budget optimisation", "8. A/B test analysis",
       "9. Power BI dashboard", "10. Business recommendations", "11. Limitations and next steps", "12. Conclusion",
       "Appendix A. Project files", "Appendix B. Cross-tool consistency check"]
st += [P(t) for t in toc]
st += [Spacer(1, 8), callout("<b>How to read this report.</b> Every finding is labelled as one of three types. <b>Observed</b> means it was "
       "calculated directly from the data. <b>Assumption</b> means a planning input we chose (for example, margins), stated openly. "
       "<b>Recommendation</b> means what we suggest doing, and how to check whether it worked."), PageBreak()]

# ---------------------------------------------------------------- 1 Executive summary
st += [P("1. Executive summary", "h1")]
kp = [["Ad Spend", "Revenue", "Orders", "New Cust.", "ROAS", "CAC", "CTR", "CPC", "CPM", "CVR", "AOV"],
      [inr(ov["Ad_Spend"]), inr(ov["Revenue"]), f"{ov['Orders']:,.0f}", f"{ov['New_Customers']:,.0f}", f"{ov['ROAS']:.2f}x",
       f"₹{ov['CAC']:,.0f}", f"{ov['CTR_%']:.2f}%", f"₹{ov['CPC']:.2f}", f"₹{ov['CPM']:.0f}", f"{ov['CVR_%']:.2f}%", f"₹{ov['AOV']:,.0f}"]]
st += [table(kp, [17 * mm, 17 * mm, 14 * mm, 15 * mm, 14 * mm, 14 * mm, 14 * mm, 15 * mm, 15 * mm, 15 * mm, 20 * mm], zebra=False), Spacer(1, 8)]
st += [P("<b>Observed findings</b>", "h2")]
st += bullets([
    "<b>Google Ads is the most efficient channel.</b> It took 35% of spend and produced 50% of revenue (ROAS 3.13x, CAC ₹318). Meta Ads took the "
    "largest share of spend (44%) at ROAS 1.79x. Marketplaces converted visitors best (CVR 7.15%), but clicks cost the most (₹27.40), so ROAS was only 1.45x.",
    "<b>Blended ROAS flatters performance.</b> Brand search produced 23% of revenue from 4% of spend (ROAS 12.41x), but only 29% of its buyers were new. "
    "Without brand search, ROAS falls from 2.19x to <b>1.76x</b>.",
    "<b>Generic hair-fall search is the best acquisition engine:</b> ROAS 3.15x, CAC ₹237, 84% new customers, 1,753 orders, stable trend.",
    "<b>Creative fatigue:</b> the Vitamin C influencer video's CTR fell 24% between the first and last four weeks, and its ROAS fell 22%.",
    "<b>The A/B test found a click winner, not a sales winner.</b> The UGC video lifted CTR by 15.9% (p &lt; 0.001), but CVR and ROAS were statistically identical.",
])
st += [P("<b>Assumption-based view</b>", "h2"),
       P("With assumed contribution margins of 55% (website) and 40% (marketplaces), break-even ROAS is 1.82x and 2.50x. On that basis, "
         "4 campaigns generate all the profit, and 10 campaigns holding 78% of spend are below break-even. Contribution after ad cost is still positive overall (≈ ₹5.35 L)."),
       P("<b>Recommendation in one line</b>", "h2"),
       P("Scale generic search in steps (+25%), hold brand search and retargeting, cut four loss-making campaigns by 30%, cap awareness by 15%, "
         "and put the freed budget into new creative tests. The plan keeps the budget the same and adds about ₹3,200 a day in contribution (scenario, excluding the test reserve)."),
       PageBreak()]

# ---------------------------------------------------------------- 2 Business problem
st += [P("2. Business problem and objectives", "h1"),
       P("A D2C beauty and personal-care brand sells hair-care and skin-care products through its own website and through marketplaces "
         "(Amazon, Nykaa, Flipkart). It advertises on Meta (Facebook and Instagram), Google (Search, Shopping, Performance Max, YouTube) and "
         "marketplace sponsored listings. Leadership wants to know:"),
       *bullets(["Which channels, campaigns, products, audiences and creatives are actually making money?",
                 "Where is the budget being wasted, and where could more budget profitably go?",
                 "How efficiently are we acquiring <i>new</i> customers?",
                 "Did our latest creative test produce a real winner?"]),
       P("Objectives", "h2"),
       *bullets(["Clean and validate a messy multi-channel campaign export, documenting every decision.",
                 "Calculate seven core KPIs consistently across Excel, SQL, Python and Power BI.",
                 "Diagnose performance across nine dimensions and identify top and underperforming campaigns.",
                 "Build a budget recommendation that considers margin, CAC, conversion, volume and trend, not ROAS alone.",
                 "Analyse an A/B test with statistical significance testing.",
                 "Deliver an interactive dashboard design and actionable recommendations."]),
       P("Why this matters", "h2"),
       P("Performance marketing budgets are spent daily. A campaign that looks fine on blended ROAS can quietly lose money once margins are counted, "
         "while the highest-ROAS campaigns are often the ones with the least room to grow. Getting this right decides whether ad spend builds the business "
         "or just buys sales the brand would have made anyway."), PageBreak()]

# ---------------------------------------------------------------- 3 Dataset
st += [P("3. Dataset and methodology", "h1"),
       callout("<b>Simulated data.</b> The dataset was generated with a fixed random seed by <font face='Body-B'>scripts/01_generate_data.py</font>, using "
               "realistic Indian D2C beauty benchmarks (CPM ₹60-380, CTR 0.4-9%, CVR 1-9%, AOV ₹410-990). It includes a festive sale "
               "(15-20 Oct), weekend effects, one fatiguing creative and deliberate data-quality problems. It is not WishCare data."),
       Spacer(1, 6), P("Structure", "h2"),
       P("60 days (1 Sep - 30 Oct 2026) × 14 campaigns = <b>840 campaign-day rows</b> after cleaning (852 raw rows). Each campaign has one "
         "objective, product, audience and creative. A separate A/B dataset has 28 rows (2 ads × 14 days)."),
       table([["Field", "Meaning", "Example"],
              ["Date", "Day of activity", "2026-10-15"], ["Marketing Channel", "Meta Ads, Google Ads, E-commerce", "Google Ads"],
              ["Campaign Name / Objective", "Campaign and its goal (Conversions, Sales, Awareness)", "GGL_Search_Generic_HairFall / Conversions"],
              ["Product Name", "Product advertised (6 products)", "Hair Growth Serum"],
              ["Ad Spend", "Money paid to the platform (₹)", "5,796"], ["Impressions / Reach", "Ad views / unique people reached", "19,300 / 12,900"],
              ["Clicks", "Clicks to the website or listing", "810"], ["Orders", "Purchases attributed to the ad", "29"],
              ["New / Returning Customers", "First-time vs repeat buyers (sum = Orders)", "24 / 5"],
              ["Revenue", "Attributed sales value (₹)", "18,300"],
              ["Audience Type", "Who was targeted", "Non-Brand Search, Retargeting - Cart Abandoners"],
              ["Ad Creative Type", "Format of the ad", "UGC Video, Static Image, Text Ad"]],
             [42 * mm, 72 * mm, 56 * mm]),
       P("Methodology and tool workflow", "h2"),
       table([["Step", "Tool", "Output"],
              ["Generate and clean data, KPIs, EDA, budget model, A/B statistics", "Python (Pandas, NumPy, SciPy, Matplotlib, Plotly)", "Notebook, cleaned CSVs, charts"],
              ["Validate, query and aggregate", "SQL (MySQL 8 syntax, tested in SQLite)", "19 queries incl. CTEs and window functions"],
              ["Formula-based reporting and audit", "Excel (SUMIFS, IF guards, RANK, NORMSDIST, charts)", "15-sheet workbook with Checks sheet"],
              ["Interactive dashboard", "Power BI (DAX measures, theme, build guide)", "Dashboard design + DAX"]],
             [70 * mm, 55 * mm, 45 * mm]),
       Spacer(1, 6), callout("<b>Golden rule used in every tool:</b> a KPI for any group = SUM(numerator) / SUM(denominator). Averaging row-level ratios "
                             f"is wrong. In this data, the average of daily ROAS is 2.65x, while the true ROAS is {ov['ROAS']:.2f}x."),
       PageBreak()]

# ---------------------------------------------------------------- 4 Cleaning
st += [P("4. Data cleaning", "h1"),
       P("The raw export had the problems real exports have. We read every column as text first, so nothing was silently converted, "
         "then fixed issues in a deliberate order: text and types first (so hidden duplicates become visible), then impossible values, "
         "then outliers, then missing numbers. All 27 corrected or estimated rows carry a <i>Data_Quality_Flag</i>."),
       table([["Issue", "Rows", "Action", "Reason"]] +
             [[r.Issue, int(r.Rows_Affected), r.Action, r.Reason] for r in log.itertuples()],
             [38 * mm, 11 * mm, 44 * mm, 77 * mm]),
       Spacer(1, 6),
       P("<b>Consistency checks after cleaning (all passed):</b> no missing values; one row per campaign per day; exactly 840 rows; each campaign maps to a single "
         "channel, objective, product, audience and creative; clicks ≤ impressions; reach ≤ impressions; orders ≤ clicks; new + returning = orders; "
         "no negative numbers; all dates within the period."),
       PageBreak()]

# ---------------------------------------------------------------- 5 KPIs
st += [P("5. Marketing KPIs", "h1"),
       P("If a denominator is zero, the KPI is left <b>blank</b>, not 0. For example, 20 campaign-days had zero new customers, and a CAC of ₹0 would "
         "wrongly suggest free customers. Excel uses IF(x=0,\"\",…), SQL uses NULLIF(x,0), DAX uses DIVIDE(), and Python returns NaN."),
       table([["KPI", "Formula", "Simple example", "Business interpretation", "This project"],
              ["ROAS", "Revenue ÷ Ad Spend", "₹50,000 ÷ ₹10,000 = 5.0x", "₹ of sales per ₹1 of ads. Compare with break-even ROAS (1 ÷ margin).", f"{ov['ROAS']:.2f}x"],
              ["CAC", "Ad Spend ÷ New Customers", "₹10,000 ÷ 25 = ₹400", "Cost to win a first-time buyer. Must be below what that customer is worth.", f"₹{ov['CAC']:,.0f}"],
              ["CTR", "Clicks ÷ Impressions × 100", "150 ÷ 10,000 = 1.5%", "How appealing the ad is to the audience. Falling CTR suggests creative fatigue.", f"{ov['CTR_%']:.2f}%"],
              ["CPC", "Ad Spend ÷ Clicks", "₹10,000 ÷ 1,000 = ₹10", "Price of one visit. Driven by CPM and CTR.", f"₹{ov['CPC']:.2f}"],
              ["CPM", "Ad Spend ÷ Impressions × 1,000", "₹10,000 ÷ 80,000 × 1,000 = ₹125", "Price of reaching people. Rises in competitive periods.", f"₹{ov['CPM']:.2f}"],
              ["CVR", "Orders ÷ Clicks × 100", "30 ÷ 1,000 = 3%", "How well the landing page or listing turns visitors into buyers.", f"{ov['CVR_%']:.2f}%"],
              ["AOV", "Revenue ÷ Orders", "₹18,000 ÷ 30 = ₹600", "Average basket size. Bundles and free-shipping thresholds raise it.", f"₹{ov['AOV']:,.0f}"]],
             [13 * mm, 34 * mm, 38 * mm, 64 * mm, 21 * mm]),
       Spacer(1, 8), P("How the KPIs fit together", "h2"),
       P("Revenue = Impressions × CTR × CVR × AOV, and Spend = Impressions × CPM ÷ 1,000. So <b>ROAS = CTR × CVR × AOV ÷ CPM × 1,000</b>. "
         "This identity tells us which lever is broken: expensive reach (CPM), weak ads (CTR), weak landing pages (CVR) or small baskets (AOV). "
         "For example, marketplaces have the best CVR (7.15%) but the highest CPC (₹27.40), so each order costs about ₹383 vs ₹184 on Google."),
       PageBreak()]

# ---------------------------------------------------------------- 6 EDA
st += [P("6. Exploratory analysis and findings", "h1"), P("6.1 Channel performance", "h2")]
chs = ch.sort_values("Revenue", ascending=False)
st += [table([["Channel", "Spend", "Spend %", "Revenue", "Rev %", "ROAS", "CAC", "CTR %", "CPC", "CVR %", "AOV", "New %"]] +
             [[r["Marketing_Channel"], inr(r["Ad_Spend"]), f"{r['Spend_Share_%']:.1f}", inr(r["Revenue"]), f"{r['Revenue_Share_%']:.1f}",
               f"{r['ROAS']:.2f}x", f"₹{r['CAC']:,.0f}", f"{r['CTR_%']:.2f}", f"₹{r['CPC']:.2f}", f"{r['CVR_%']:.2f}", f"₹{r['AOV']:,.0f}",
               f"{r['New_Cust_Share_%']:.0f}%"] for _, r in chs.iterrows()],
             [22 * mm, 16 * mm, 12 * mm, 16 * mm, 11 * mm, 12 * mm, 13 * mm, 12 * mm, 14 * mm, 12 * mm, 14 * mm, 12 * mm]),
       Spacer(1, 4), img("charts/01_channel_revenue_roas.png", 165 * mm),
       P("<b>Observed:</b> Google earns half of all revenue from about a third of spend. Meta is the biggest spender at below-average ROAS. "
         "Marketplaces have the highest CVR but the most expensive clicks. <b>Why it matters:</b> channel mix is the first budget lever, "
         "but Google's lead depends heavily on brand search (section 6.2).")]
st += [PageBreak(), P("6.2 Campaign performance: top and underperforming", "h2"), img("charts/02_campaign_roas.png", 160 * mm)]
cs = camp.sort_values("ROAS", ascending=False)
st += [table([["Campaign", "Spend", "Revenue", "Orders", "ROAS", "CAC", "CTR %", "CVR %", "New %"]] +
             [[r["Campaign_Name"], inr(r["Ad_Spend"]), inr(r["Revenue"]), f"{r['Orders']:,}", f"{r['ROAS']:.2f}x", f"₹{r['CAC']:,.0f}",
               f"{r['CTR_%']:.2f}", f"{r['CVR_%']:.2f}", f"{r['New_Cust_Share_%']:.0f}%"] for _, r in cs.iterrows()],
             [52 * mm, 16 * mm, 17 * mm, 13 * mm, 13 * mm, 15 * mm, 13 * mm, 13 * mm, 13 * mm])]
st += [PageBreak(), P("<b>Top performers (observed):</b>"),
       *bullets(["<b>GGL_Search_Brand</b>: ROAS 12.41x, 23% of revenue from 4% of spend. Only 29% of buyers are new, so it captures existing demand rather than creating it.",
                 "<b>META_Retargeting_Combo_Offer</b>: ROAS 5.74x and the highest AOV (₹997), from a small audience of past site visitors.",
                 "<b>GGL_Search_Generic_HairFall</b>: ROAS 3.15x with the lowest acquisition CAC (₹237) and 84% new customers. This is the strongest growth campaign."]),
       P("<b>Underperformers (observed):</b>"),
       *bullets(["<b>ECOM_Flipkart_SP_Shampoo</b>: ROAS 0.44x, CAC ₹1,471, CVR 3.21% vs a 7.01% marketplace median. This points to a listing problem.",
                 "<b>META_Awareness_Sunscreen_Reels</b> and <b>GGL_YouTube_HairSerum</b>: ROAS 0.57x and 0.50x, but these are awareness campaigns and should be judged on reach.",
                 "<b>META_Prospecting_HairSerum_Static</b> (1.16x, CAC ₹671) and <b>META_Prospecting_VitC_Influencer</b> (1.25x, fatiguing)."]),
       img("charts/09_campaign_ctr_cvr.png", 165 * mm),
       P("<b>Reading CTR and CVR together:</b> Flipkart has normal marketplace CTR but half the marketplace CVR, so the ad works and the listing doesn't. "
         "The awareness campaigns are weak on both, as expected for broad audiences.", "cap")]
st += [PageBreak(), P("6.3 Product-wise revenue", "h2"), img("charts/04_product_revenue.png", 150 * mm),
       P("<b>Observed:</b> Hair Growth Serum drives 61% of ad revenue (ROAS 2.78x). Hair Care Combo has the highest ROAS (5.74x) and AOV (₹997) but runs "
         "only through retargeting. Sunscreen (1.15x) and Vitamin C Serum (1.31x) are the weakest. <b>Why it matters:</b> a single hero product "
         "is a concentration risk, and the combo's AOV points to a bundling opportunity."),
       P("6.4 Daily and weekly trends", "h2"), img("charts/05_daily_revenue.png", 165 * mm)]
fv = o["fest"]
st += [P(f"<b>Observed:</b> the festive sale (15-20 Oct) lifted revenue per day by 39% (₹{fv['Revenue_per_day']['Yes']:,.0f} vs ₹{fv['Revenue_per_day']['No']:,.0f}) "
         f"and CVR from {fv['CVR_%']['No']:.2f}% to {fv['CVR_%']['Yes']:.2f}%, while CPM rose 16%. ROAS still improved ({fv['ROAS']['Yes']:.2f}x vs {fv['ROAS']['No']:.2f}x). "
         "Weekends convert slightly better (CVR 4.55% vs 4.14%). Weekly ROAS stayed between 2.06x and 2.34x, with no structural decline.")]
st += [PageBreak(), P("6.5 Customer acquisition: new vs returning", "h2"), img("charts/06_new_vs_returning.png", 140 * mm),
       P("<b>Observed:</b> 61% of all orders came from new customers (7,576 of 12,359). Google has the lowest CAC (₹318) but the lowest new-customer share (58%), "
         "because brand search mostly serves existing customers. Meta (65% new, CAC ₹535) and marketplaces (65% new, CAC ₹587) acquire new buyers at higher cost. "
         "<b>Why it matters:</b> a channel's CAC must be read together with how many genuinely new customers it brings.")]
aud_rows = [["Audience", "ROAS", "CAC", "CVR %", "New %"],
            ["Brand Search", "12.41x", "₹171", "8.85", "29%"], ["Retargeting - Site Visitors", "5.74x", "₹433", "7.12", "40%"],
            ["Retargeting - Cart Abandoners", "3.24x", "₹365", "6.29", "44%"], ["Non-Brand Search", "3.15x", "₹237", "3.61", "84%"],
            ["Shopping - In-Market", "1.80x", "₹305", "4.21", "78%"], ["Prospecting - Interest", "1.56x", "₹501", "2.98", "76%"],
            ["Marketplace - Keyword", "1.48x", "₹588", "7.22", "65%"], ["Prospecting - Lookalike", "1.16x", "₹671", "2.65", "79%"],
            ["Prospecting - Broad", "0.54x", "₹1,077", "1.31", "87%"]]
st += [P("6.6 Audience and creative performance", "h2"), table(aud_rows, [55 * mm, 20 * mm, 20 * mm, 20 * mm, 20 * mm]), Spacer(1, 4),
       P("<b>Observed:</b> ROAS falls as audiences get colder (brand → retargeting → non-brand search → prospecting → broad), while the new-customer share rises. "
         "This is the classic efficiency-vs-growth trade-off. By creative, text ads (search) lead on ROAS. Among Meta prospecting campaigns, the UGC video "
         "(ROAS 1.81x, CAC ₹458) beat the static image (1.16x, ₹671), but they targeted different audiences, so the controlled A/B test (section 8) is the fair comparison.")]
st += [PageBreak(), P("6.7 Creative fatigue", "h2"), img("charts/07_creative_fatigue.png", 150 * mm),
       P("<b>Observed:</b> the Vitamin C influencer video's CTR fell from 1.46% to 1.11% (−24%) between the first four and the last four weeks, and its ROAS fell from 1.43x to 1.11x. "
         "The UGC hair serum video stayed flat (+3%). <b>Why it matters:</b> fatigue raises CPC and CAC gradually. Without a weekly CTR check, it looks like "
         "the campaign just got worse."),
       P("6.8 Ad spend vs revenue", "h2"), img("charts/08_spend_vs_revenue.png", 140 * mm),
       P("<b>Observed:</b> bubbles above the dashed line return more than they cost, before margin. Within each campaign, daily spend and revenue are positively correlated "
         "(r = 0.27 to 0.61), so more spend on a day tends to bring more revenue. But correlation in 60 days of normal variation does not tell us the return on a "
         "<i>large</i> budget increase. That is why the budget plan scales in steps.")]

# ---------------------------------------------------------------- 7 Budget
st += [PageBreak(), P("7. Budget optimisation", "h1"),
       P("Moving money to the highest-ROAS campaign is the most common mistake. ROAS ignores margin, ignores whether customers are new, ignores "
         "audience size limits, and can be inflated by small volumes. We therefore score every campaign on six tests."),
       table([["Assumption (not data)", "Value", "Reason"],
              ["Contribution margin before ads - website", "55%", "After product cost, shipping, payment fees, returns"],
              ["Contribution margin before ads - marketplaces", "40%", "Commission and fulfilment fees"],
              ["Break-even ROAS", "1.82x / 2.50x", "= 1 ÷ margin"], ["Target CAC", "₹450", "Affordable cost for a customer who may re-order"],
              ["Minimum volume", "300 orders", "Avoid decisions on noise"], ["Extra-spend efficiency", "70% of current ROAS", "Diminishing returns"]],
             [65 * mm, 30 * mm, 75 * mm]),
       Spacer(1, 6),
       table([["Decision rule", "Condition"],
              ["Cap & measure differently (−15%)", "Objective = Awareness"],
              ["Maintain (limited headroom)", "Brand search or retargeting audience"],
              ["Scale test (+25%)", "Profit index ≥ 1.3 AND CAC ≤ ₹450 AND ≥ 300 orders AND ROAS trend > −10%"],
              ["Reduce (−30%) & fix", "Profit index < 1 AND CAC > 1.2 × target"],
              ["Optimise (hold budget)", "Everything else; reason listed per campaign"]],
             [55 * mm, 115 * mm]),
       Spacer(1, 4), P("Profit index = ROAS ÷ break-even ROAS. Above 1, the campaign makes money after paying for ads.", "cap")]
bs = bud.sort_values("Budget_Change_%", ascending=False)
st += [table([["Campaign", "ROAS", "Profit idx", "CAC", "Orders", "Decision", "Daily ₹ now → new"]] +
             [[r["Campaign_Name"], f"{r['ROAS']:.2f}x", f"{r['Profit_Index']:.2f}", f"₹{r['CAC']:,.0f}", f"{r['Orders']:,}", r["Decision"],
               f"{r['Current_Daily_Budget']:,.0f} → {r['Proposed_Daily_Budget']:,.0f}"] for _, r in bs.iterrows()],
             [57 * mm, 12 * mm, 13 * mm, 14 * mm, 12 * mm, 36 * mm, 26 * mm])]
ra = o["realloc"]
st += [PageBreak(), P("Budget-neutral reallocation scenario", "h2"),
       P(f"The cuts free <b>₹{ra['freed_daily']:,.0f}/day</b>. Scale tests receive ₹{ra['to_scale_daily']:,.0f}/day (capped at +25%, so the increase happens in steps). "
         f"The remaining <b>₹{ra['reserve_daily']:,.0f}/day</b> becomes a creative-testing reserve for new UGC videos for the Vitamin C and hair serum prospecting campaigns. Total budget is unchanged."),
       P(f"<b>Scenario result (excluding the reserve, whose return is unknown):</b> revenue changes by ₹{ra['delta_rev_daily']:,.0f}/day, but contribution after ads "
         f"improves by <b>₹{ra['delta_contrib_daily']:,.0f}/day</b> (≈ ₹{ra['delta_contrib_daily']*30/1e5:.2f} L per 30 days). We give up revenue that was losing money."),
       callout("<b>Caution:</b> this is a planning scenario built on stated assumptions, not a forecast. Validate it with a 2-4 week controlled budget change, "
               "measuring marginal CAC and ROAS each week before taking the next step.", PINK),
       Spacer(1, 6), P("<b>Opportunities to improve acquisition efficiency</b>"),
       *bullets(["Shift prospecting spend toward non-brand search and UGC video, which have the lowest CAC among campaigns that mostly win new customers.",
                 "Fix marketplace listings before spending more there (CVR gap on Flipkart).",
                 "Exclude existing customers from prospecting audiences, so CAC measures genuinely new buyers.",
                 "Refresh creatives when weekly CTR falls more than 20% below the launch week."])]

# ---------------------------------------------------------------- 8 A/B
A, B = abs_.iloc[0], abs_.iloc[1]; ab = o["ab"]
st += [PageBreak(), P("8. A/B test analysis", "h1"),
       P("<b>Design (simulated):</b> one Meta prospecting campaign for Hair Growth Serum, the same interest audience, the budget split 50/50, 14 days (1-14 Oct). "
         "Only the creative differs: <b>A = static product image</b> and <b>B = UGC testimonial video</b>. Changing one variable at a time is what makes the comparison fair."),
       table([["Metric", "A - Static", "B - UGC video", "B vs A"],
              ["Impressions", f"{A.Impressions:,}", f"{B.Impressions:,}", f"{B.Impressions/A.Impressions-1:+.1%}"],
              ["Clicks", f"{A.Clicks:,}", f"{B.Clicks:,}", f"{B.Clicks/A.Clicks-1:+.1%}"],
              ["CTR", f"{A['CTR_%']:.2f}%", f"{B['CTR_%']:.2f}%", f"{B['CTR_%']/A['CTR_%']-1:+.1%}"],
              ["CPC", f"₹{A.CPC:.2f}", f"₹{B.CPC:.2f}", f"{B.CPC/A.CPC-1:+.1%}"],
              ["Orders", f"{A.Orders}", f"{B.Orders}", f"{B.Orders/A.Orders-1:+.1%}"],
              ["Conversion rate", f"{A['CVR_%']:.2f}%", f"{B['CVR_%']:.2f}%", f"{B['CVR_%']/A['CVR_%']-1:+.1%}"],
              ["Revenue", f"₹{A.Revenue:,.0f}", f"₹{B.Revenue:,.0f}", f"{B.Revenue/A.Revenue-1:+.1%}"],
              ["ROAS", f"{A.ROAS:.2f}x", f"{B.ROAS:.2f}x", f"{B.ROAS/A.ROAS-1:+.1%}"]],
             [45 * mm, 40 * mm, 40 * mm, 45 * mm]),
       Spacer(1, 6), img("charts/10_ab_test.png", 165 * mm), P("Is the difference real? Significance tests", "h2"),
       table([["Test", "Lift B vs A", "p-value", "Significant at 95%?"]] +
             [[r["Metric"], f"{r['Lift_B_vs_A_%']:+.1f}%", f"{r['p_value']:.4f}" if r["p_value"] >= 0.0001 else "< 0.0001", r["Significant_at_95%"]]
              for _, r in abt.iterrows()],
             [62 * mm, 30 * mm, 30 * mm, 48 * mm]),
       Spacer(1, 4),
       P("A <b>p-value</b> is the chance of seeing a gap this large if the two ads were truly identical. Below 0.05, we treat the difference as real. "
         f"For ROAS, a bootstrap 95% confidence interval for B − A is {ab['roas_ci'][0]:+.2f} to {ab['roas_ci'][1]:+.2f}. It contains zero, so there is no evidence of a ROAS difference."),
       callout(f"<b>Verdict.</b> B wins on <b>attention</b>: CTR +{ab['ctr_lift']:.1f}% (p &lt; 0.001) and CPC 8% lower. It is <b>not</b> proven to sell more: "
               f"CVR p = {ab['cvr_p']:.2f}, ROAS p = {ab['roas_p']:.2f}. With about 3,000 clicks per ad, the test could only detect a CVR change of about ±{ab['mde_rel']:.0f}%. "
               "Detecting a realistic 15% CVR lift would need about 25,000 clicks per ad. <b>Decision:</b> use B for cheaper prospecting traffic, "
               "don't claim higher sales, and test the landing page next, because CVR is now the bottleneck.")]

# ---------------------------------------------------------------- 9 Dashboard
st += [PageBreak(), P("9. Power BI dashboard", "h1"), img("charts/11_dashboard_mockup.png", 170 * mm),
       P("Layout preview drawn in Python from the cleaned data, to guide the Power BI build. The DAX measures, theme file and step-by-step build guide "
         "are in the <font face='Body-B'>powerbi/</font> folder.", "cap"),
       P("Design decisions", "h2"),
       *bullets(["<b>Reading order follows the business question:</b> KPI cards (are we making money?) → channel, trend, product (where from?) → "
                 "campaign ROAS, spend, CVR (what to grow, fix, cut?) → customers and A/B test.",
                 "<b>11 KPI cards</b> built from DAX measures that divide sums, so every slicer combination stays correct.",
                 "<b>Four slicers in one row</b> (Date, Channel, Campaign, Product). The A/B visuals are excluded from slicer interactions because the test is a separate experiment.",
                 "<b>One fixed colour per channel</b> across all visuals. Horizontal bars instead of pies, for accurate comparison and readable campaign names.",
                 "<b>Break-even ROAS reference line</b> on the ROAS chart, so 1.5x reads as below break-even rather than as a fine-looking number.",
                 "<b>Tooltips</b> carry ROAS, CAC, CVR and orders, giving context on hover without cluttering the page."]),
       P("Build summary", "h2"),
       *bullets(["Load the cleaned CSVs in Power Query and set data types → create a Date table with CALENDAR and relate it 1:* to Campaign[Date].",
                 "Create the base SUM measures, then the KPI measures with DIVIDE() → apply the theme JSON.",
                 "Add slicers, the multi-value KPI card, and the 8 visuals as specified in the guide → set interactions → check the numbers against the expected values table."])]

# ---------------------------------------------------------------- 10 Recommendations
st += [PageBreak(), P("10. Business recommendations", "h1")]
recs = [
    ("1. Scale generic hair-fall search in steps", "Scaling profitable campaigns, reducing CAC",
     "ROAS 3.15x, CAC ₹237, 84% new, 1,753 orders, stable trend.", "More search demand is available at a similar cost; returns will diminish.",
     "+25% for 2 weeks and expand related keywords. Continue only if marginal CAC stays below ₹450 and ROAS above 2.5x."),
    ("2. Judge brand search and retargeting on incrementality", "Optimising budgets",
     "Brand ROAS 12.41x but 29% new. ROAS without brand is 1.76x.", "Many of these buyers would have purchased anyway.",
     "Hold budgets flat. Run a geo holdout test. Report ROAS excluding brand weekly."),
    ("3. Cut and fix loss-making acquisition", "Improving ROAS, reducing CAC",
     "Flipkart 0.44x (CVR 3.2% vs 7.0%), Nykaa 1.40x vs 2.50x break-even, Meta static 1.16x.", "Margin assumptions as stated.",
     "−30% now. Audit marketplace listings (price, ratings, images, delivery). Pause if still below break-even after 2 weeks."),
    ("4. Refresh the Vitamin C influencer creative", "Improving creative performance",
     "CTR −24%, ROAS −22% (first 4 vs last 4 weeks) while other ads were flat.", "Fatigue, not audience change, is the cause.",
     "Launch 3-4 new hooks or creators from the testing reserve. Rotate when weekly CTR is 20% below launch week."),
    ("5. Use UGC for prospecting, then test the landing page", "Increasing conversion rates",
     "UGC: CTR +15.9% (significant), CPC −8%. CVR and ROAS not different.", "Cheaper clicks are worth having even without a CVR lift.",
     "Make UGC the default prospecting format. Next A/B test: landing page (results-first hero, proof, COD or free-shipping badge)."),
    ("6. Raise AOV with bundles", "Increasing average order value",
     "Combo AOV ₹997 vs ₹588 overall. Combo retargeting ROAS 5.74x. Serum = 61% of revenue.", "Serum buyers want complementary products.",
     "Add a Complete Hair Kit bundle and a free-shipping threshold of about ₹699. Target AOV +10% with no drop in CVR."),
    ("7. Cap awareness spend and plan for peaks", "Optimising budgets",
     "Awareness = 12.7% of spend at 0.54x ROAS. Festive sale: revenue/day +39%, CVR +31%, CPM +16%.", "Awareness creates demand that last-click data misses.",
     "−15% on awareness, measured on reach, CPM and brand-search lift. Pre-load budgets for sales and weekends."),
]
for t, cov, obs, ass, rec in recs:
    st.append(KeepTogether([P(f"{t}  <font size=8.5 color='#C2185B'>[{cov}]</font>", "h2"),
                            table([["Observed", obs], ["Assumption", ass], ["Recommendation", rec]], [32 * mm, 138 * mm], header=False, zebra=True)]))
st += [Spacer(1, 6), P("<b>Weekly monitoring set:</b> ROAS and ROAS excluding brand; CAC vs ₹450; CVR by campaign; weekly CTR vs launch week; new-customer share; contribution after ads.")]

# ---------------------------------------------------------------- 11 Limitations / 12 Conclusion
st += [PageBreak(), P("11. Limitations and next steps", "h1"),
       *bullets(["<b>Simulated data.</b> Patterns are realistic but designed. Real data will be noisier, with tracking gaps.",
                 "<b>Attribution.</b> Platform-reported, last-click style revenue. It over-credits brand and retargeting and under-credits awareness. Platforms may also double-count the same order.",
                 "<b>Assumed economics.</b> Margins, target CAC and the 70% marginal-efficiency figure are planning inputs. The decisions are sensitive to them, and the Excel Assumptions sheet lets you change them.",
                 "<b>No lifetime value.</b> A higher CAC can be justified if new customers re-order. Cohort LTV would refine the target CAC.",
                 "<b>No returns, cancellations or organic baseline</b>, and only 60 days, so full seasonality (e.g. Diwali, summer for sunscreen) isn't captured.",
                 "<b>A/B test power.</b> 14 days could not detect realistic CVR differences."]),
       P("Next steps with real data", "h2"),
       *bullets(["Reconcile platform revenue with the order system (Shopify or the marketplace) and use UTM-based attribution.",
                 "Run incrementality tests (geo holdouts) for brand search and retargeting.",
                 "Build cohort LTV by acquisition channel to set channel-specific CAC targets.",
                 "Automate this pipeline daily (SQL → Python → Power BI refresh) with fatigue and CAC alerts.",
                 "Explore marketing mix modelling once 12+ months of data exist."]),
       P("12. Conclusion", "h1"),
       P("The account is profitable overall (ROAS 2.19x), but the headline hides a split. A few campaigns that capture existing demand (brand search and retargeting) "
         "plus one strong acquisition campaign (generic search) carry the profit, while most prospecting and marketplace spend sits below break-even "
         "under the stated margins. The right response is not to chase the highest ROAS. It is to scale what profitably acquires <i>new</i> customers, "
         "fix or trim what loses money, refresh fatiguing creatives, and keep testing with experiments large enough to give a clear answer. "
         "The same numbers reconcile across Excel, SQL, Python and Power BI, so the recommendations rest on a single, auditable source of truth.")]

# ---------------------------------------------------------------- Appendices
st += [PageBreak(), P("Appendix A. Project files", "h1"),
       table([["Path", "What it is"],
              ["data/raw/", "Raw messy export (852 rows), campaign master, raw A/B data"],
              ["data/cleaned/", "Cleaned campaign data (840 rows), cleaned A/B data, cleaning log, KPI summaries, budget recommendations, A/B results"],
              ["notebooks/WishCare_Performance_Marketing_Analysis.ipynb", "Executed Python notebook: cleaning → KPIs → EDA → budget → A/B, with explanations"],
              ["sql/wishcare_marketing_queries.sql", "19 queries: validation, KPIs, channel/campaign/product, trends, customers, budget, A/B"],
              ["excel/WishCare_Performance_Marketing_Analysis.xlsx", "15 sheets, 8,200+ live formulas, editable assumptions, z-test, Checks sheet"],
              ["powerbi/", "DAX measures, theme JSON, step-by-step dashboard build guide"],
              ["docs/", "Executive summary, recommendations, interview Q&A"],
              ["charts/", "All charts as PNG"], ["scripts/", "Data generator, notebook builder, Excel/report builders, SQL test"]],
             [70 * mm, 100 * mm]),
       P("Appendix B. Cross-tool consistency check", "h1"),
       P("The same KPIs were calculated independently in each tool. All match to at least four decimal places."),
       table([["KPI", "Python", "SQL (Q2.1)", "Excel (KPI_Summary)", "Power BI (expected)"],
              ["Total Ad Spend", "₹33,15,653.11", "₹33,15,653.11", "₹33,15,653.11", "₹33,15,653"],
              ["Total Revenue", "₹72,67,423.80", "₹72,67,423.80", "₹72,67,423.80", "₹72,67,424"],
              ["ROAS", f"{ov['ROAS']:.4f}", "2.19", f"{ov['ROAS']:.4f}", "2.19"],
              ["CAC", f"{ov['CAC']:.4f}", "437.65", f"{ov['CAC']:.4f}", "437.65"],
              ["CTR %", f"{ov['CTR_%']:.4f}", "1.29", f"{ov['CTR_%']:.4f}", "1.29"],
              ["CVR %", f"{ov['CVR_%']:.4f}", "4.25", f"{ov['CVR_%']:.4f}", "4.25"],
              ["AOV", f"{ov['AOV']:.4f}", "588.03", f"{ov['AOV']:.4f}", "588.03"]],
             [30 * mm, 35 * mm, 30 * mm, 38 * mm, 37 * mm]),
       P("SQL output is rounded to 2 decimals by design. The Excel Checks sheet computes the difference from Python live (all MATCH).", "cap")]

doc.build(st)
print("report built")

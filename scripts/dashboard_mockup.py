"""Static preview of the Power BI dashboard layout, drawn with matplotlib from the cleaned data.
It is a design mock-up to guide the Power BI build, not a Power BI export."""
import json
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

d = pd.read_csv("data/cleaned/campaign_performance_cleaned.csv", parse_dates=["Date"])
camp = pd.read_csv("data/cleaned/summary_campaign_kpis.csv")
ch = pd.read_csv("data/cleaned/summary_channel_kpis.csv")
prod = pd.read_csv("data/cleaned/summary_product_kpis.csv")
daily = pd.read_csv("data/cleaned/summary_daily_kpis.csv", parse_dates=["Date"])
ab = pd.read_csv("data/cleaned/ab_test_summary.csv")
o = json.load(open("data/cleaned/results.json"))["overall"]
CH = {"Meta Ads": "#2a78d6", "Google Ads": "#eb6834", "E-commerce": "#1baf7a"}
NAVY, BG, GREY = "#1f2a44", "#f4f6fa", "#8a8984"
plt.rcParams.update({"font.size": 8, "axes.spines.top": False, "axes.spines.right": False})

fig = plt.figure(figsize=(16, 10.2), facecolor=BG)
fig.text(0.012, 0.972, "WishCare-Inspired Performance Marketing Dashboard", fontsize=17, weight="bold", color=NAVY)
fig.text(0.012, 0.950, "Simulated data  |  1 Sep - 30 Oct 2026  |  Meta Ads, Google Ads, E-commerce", fontsize=9, color="#555")
# slicers
for i, s in enumerate(["Date range", "Marketing Channel", "Campaign", "Product"]):
    x = 0.60 + i * 0.099
    fig.patches.append(FancyBboxPatch((x, 0.947), 0.092, 0.03, boxstyle="round,pad=0.002,rounding_size=0.004",
                                      transform=fig.transFigure, fc="white", ec="#c9d0db"))
    fig.text(x + 0.006, 0.958, s + "  ▾", fontsize=8.5, color=NAVY)

cards = [("Ad Spend", f"₹{o['Ad_Spend']/1e5:.2f} L"), ("Revenue", f"₹{o['Revenue']/1e5:.2f} L"),
         ("Orders", f"{o['Orders']:,.0f}"), ("New Customers", f"{o['New_Customers']:,.0f}"),
         ("ROAS", f"{o['ROAS']:.2f}x"), ("CAC", f"₹{o['CAC']:,.0f}"), ("CTR", f"{o['CTR_%']:.2f}%"),
         ("CPC", f"₹{o['CPC']:.2f}"), ("CPM", f"₹{o['CPM']:.0f}"), ("CVR", f"{o['CVR_%']:.2f}%"), ("AOV", f"₹{o['AOV']:,.0f}")]
w = 0.0835
for i, (k, v) in enumerate(cards):
    x = 0.012 + i * (w + 0.006)
    fig.patches.append(FancyBboxPatch((x, 0.855), w, 0.075, boxstyle="round,pad=0.002,rounding_size=0.006",
                                      transform=fig.transFigure, fc="white", ec="#e3e6ec"))
    fig.text(x + 0.008, 0.906, k, fontsize=8.5, color="#555")
    fig.text(x + 0.008, 0.868, v, fontsize=14, weight="bold", color=NAVY)

def panel(rect, ttl):
    ax = fig.add_axes(rect); ax.set_facecolor("white"); ax.set_title(ttl, loc="left", fontsize=10, weight="bold", color=NAVY)
    ax.grid(axis="x", alpha=.25); return ax

# Row 1
ax = panel([0.075, 0.56, 0.17, 0.25], "Revenue by channel (₹ L)")
c = ch.sort_values("Revenue")
ax.barh(c.Marketing_Channel, c.Revenue / 1e5, color=[CH[x] for x in c.Marketing_Channel], height=.55)
for i, v in enumerate(c.Revenue / 1e5): ax.text(v, i, f" {v:.1f}", va="center")
ax = panel([0.30, 0.56, 0.38, 0.25], "Daily revenue trend (₹ L)")
ax.plot(daily.Date, daily.Revenue / 1e5, color="#2a78d6", lw=1.8)
ax.axvspan(pd.Timestamp("2026-10-15"), pd.Timestamp("2026-10-20"), color="#eda100", alpha=.15)
import matplotlib.dates as mdates
ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))
ax.grid(axis="y", alpha=.25)
ax = panel([0.76, 0.56, 0.22, 0.25], "Product-wise revenue (₹ L)")
p = prod.sort_values("Revenue")
ax.barh(p.Product_Name, p.Revenue / 1e5, color="#2a78d6", height=.55)
# Row 2
ax = panel([0.215, 0.08, 0.14, 0.40], "ROAS by campaign")
cc = camp.sort_values("ROAS")
short = cc.Campaign_Name.str.replace("META_", "M·").str.replace("GGL_", "G·").str.replace("ECOM_", "E·")
ax.barh(short, cc.ROAS, color=[CH[x] for x in cc.Marketing_Channel], height=.6); ax.tick_params(axis="y", labelsize=7)
ax.axvline(1.82, color=GREY, ls="--", lw=1); ax.text(1.9, -0.9, "break-even 1.82x (website)", fontsize=6.5, color=GREY)
ax = panel([0.39, 0.08, 0.13, 0.40], "Ad spend by campaign (₹ L)")
cs = camp.set_index("Campaign_Name").loc[cc.Campaign_Name]
ax.barh(short, cs.Ad_Spend / 1e5, color=[CH[x] for x in cs.Marketing_Channel], height=.6); ax.set_yticklabels([])
ax = panel([0.55, 0.08, 0.12, 0.40], "Campaign CVR %")
ax.barh(short, cs["CVR_%"], color=[CH[x] for x in cs.Marketing_Channel], height=.6); ax.set_yticklabels([])
ax = panel([0.73, 0.30, 0.25, 0.18], "New vs returning customers")
cn = ch.sort_values("Orders")
ax.barh(cn.Marketing_Channel, cn.New_Customers, color="#2a78d6", height=.55, label="New")
ax.barh(cn.Marketing_Channel, cn.Returning_Customers, left=cn.New_Customers + 20, color="#9cc3ef", height=.55, label="Returning")
ax.legend(frameon=False, fontsize=7, loc="lower right")
ax = panel([0.73, 0.06, 0.25, 0.17], "A/B test: A static vs B UGC video")
import numpy as np
m = ["CTR_%", "CVR_%", "ROAS"]; x = np.arange(3)
ax.bar(x - .18, ab.loc[0, m], .34, color=GREY, label="A"); ax.bar(x + .18, ab.loc[1, m], .34, color="#2a78d6", label="B")
ax.set_xticks(x, ["CTR %", "CVR %", "ROAS"]); ax.legend(frameon=False, fontsize=7); ax.grid(axis="y", alpha=.25)
fig.text(0.012, 0.47, "Legend", weight="bold", color=NAVY)
for i, (k, v) in enumerate(CH.items()):
    fig.patches.append(plt.Rectangle((0.012, 0.44 - i * 0.025), 0.01, 0.014, transform=fig.transFigure, color=v))
    fig.text(0.026, 0.442 - i * 0.025, k, fontsize=8)
fig.text(0.012, 0.012, "Layout mock-up drawn in Python to preview the Power BI design. Build steps: powerbi/Power_BI_Dashboard_Build_Guide.md",
         fontsize=8, color="#777", style="italic")
plt.savefig("charts/11_dashboard_mockup.png", dpi=150, facecolor=BG)
print("ok")

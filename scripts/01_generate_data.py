"""
01_generate_data.py
Creates SIMULATED campaign data for a WishCare-inspired beauty & personal-care brand.
These numbers are synthetic. They do NOT represent WishCare's real campaign performance.

Outputs
  data/raw/campaign_performance_raw.csv   (deliberately contains data-quality problems)
  data/raw/ab_test_raw.csv                (two ad variations, 14 days)
"""
import numpy as np
import pandas as pd
from pathlib import Path

rng = np.random.default_rng(42)          # fixed seed -> same data every run
ROOT = Path(__file__).resolve().parents[1]

dates = pd.date_range("2026-09-01", "2026-10-30", freq="D")    # 60 days
FESTIVE = (dates >= "2026-10-15") & (dates <= "2026-10-20")    # simulated festive sale

# ---------------------------------------------------------------------------
# Campaign master: one row per campaign (ad set). Base daily behaviour.
# spend = average daily spend (INR), cpm = cost per 1,000 impressions,
# ctr / cvr as fractions, aov = average order value, new = share of orders
# from first-time customers, freq = impressions per person reached.
# ---------------------------------------------------------------------------
C = [
 # channel,   campaign,                          objective,      product,               audience,                   creative,                 spend, cpm,  ctr,    cvr,   aov, new,  freq
 ("Meta Ads", "META_Prospecting_HairSerum_UGC",   "Conversions",  "Hair Growth Serum",   "Prospecting - Interest",   "UGC Video",              6500, 145, 0.0135, 0.030, 640, 0.78, 1.35),
 ("Meta Ads", "META_Prospecting_HairSerum_Static","Conversions",  "Hair Growth Serum",   "Prospecting - Lookalike",  "Static Image",           4200, 130, 0.0095, 0.026, 620, 0.80, 1.30),
 ("Meta Ads", "META_Retargeting_Shampoo_Carousel","Conversions",  "Rice Water Shampoo",  "Retargeting - Cart Abandoners","Carousel",           2600, 210, 0.0210, 0.058, 520, 0.42, 2.60),
 ("Meta Ads", "META_Awareness_Sunscreen_Reels",   "Awareness",    "Sunscreen SPF 50",    "Prospecting - Broad",      "Short Video (Reels)",    3800,  75, 0.0070, 0.012, 450, 0.85, 1.15),
 ("Meta Ads", "META_Prospecting_VitC_Influencer", "Conversions",  "Vitamin C Face Serum","Prospecting - Interest",   "Influencer Video",       5200, 150, 0.0160, 0.027, 560, 0.76, 1.40),
 ("Meta Ads", "META_Retargeting_Combo_Offer",     "Conversions",  "Hair Care Combo",     "Retargeting - Site Visitors","Static Image",         1500, 230, 0.0190, 0.070, 990, 0.35, 2.80),
 ("Google Ads","GGL_Search_Brand",                "Conversions",  "Hair Growth Serum",   "Brand Search",             "Text Ad",                2200, 380, 0.0900, 0.085, 610, 0.30, 1.60),
 ("Google Ads","GGL_Search_Generic_HairFall",     "Conversions",  "Hair Growth Serum",   "Non-Brand Search",         "Text Ad",                5600, 300, 0.0420, 0.034, 630, 0.82, 1.50),
 ("Google Ads","GGL_Shopping_Sunscreen",          "Sales",        "Sunscreen SPF 50",    "Shopping - In-Market",     "Product Listing Ad",     3400, 160, 0.0160, 0.038, 430, 0.74, 1.45),
 ("Google Ads","GGL_PMax_OnionHairOil",           "Sales",        "Onion Hair Oil",      "Performance Max - Mixed",  "Dynamic / Mixed",        4800, 120, 0.0110, 0.032, 480, 0.70, 1.40),
 ("Google Ads","GGL_YouTube_HairSerum",           "Awareness",    "Hair Growth Serum",   "Prospecting - Broad",      "Long Video",             3000,  60, 0.0040, 0.010, 620, 0.88, 1.20),
 ("E-commerce","ECOM_Amazon_SP_HairSerum",        "Sales",        "Hair Growth Serum",   "Marketplace - Keyword",    "Sponsored Product",      4400, 220, 0.0085, 0.085, 600, 0.68, 1.70),
 ("E-commerce","ECOM_Nykaa_Sponsored_VitC",       "Sales",        "Vitamin C Face Serum","Marketplace - Category",   "Sponsored Product",      3600, 200, 0.0075, 0.070, 540, 0.66, 1.60),
 ("E-commerce","ECOM_Flipkart_SP_Shampoo",        "Sales",        "Rice Water Shampoo",  "Marketplace - Keyword",    "Sponsored Product",      2800, 190, 0.0065, 0.030, 410, 0.64, 1.60),
]
cols = ["channel","campaign","objective","product","audience","creative",
        "spend","cpm","ctr","cvr","aov","new","freq"]
master = pd.DataFrame(C, columns=cols)

rows = []
for _, c in master.iterrows():
    for i, d in enumerate(dates):
        weekend = d.dayofweek >= 5
        fest = FESTIVE[i]
        spend = c.spend * rng.normal(1, 0.12) * (1.30 if fest else 1.0)
        cpm = c.cpm * rng.normal(1, 0.08) * (1.15 if fest else 1.0)   # auctions get pricier in sales
        impressions = int(spend / cpm * 1000)
        ctr = c.ctr * rng.normal(1, 0.10)
        # Creative fatigue: the influencer video loses ~40% of its CTR over 60 days
        if c.campaign == "META_Prospecting_VitC_Influencer":
            ctr *= 1 - 0.40 * i / (len(dates) - 1)
        clicks = int(rng.binomial(impressions, min(max(ctr, 0.0005), 0.5)))
        cvr = c.cvr * rng.normal(1, 0.12) * (1.10 if weekend else 1.0) * (1.25 if fest else 1.0)
        orders = int(rng.binomial(clicks, min(max(cvr, 0.001), 0.6)))
        new = int(rng.binomial(orders, c.new))
        revenue = round(orders * c.aov * rng.lognormal(0, 0.08) * (0.95 if fest else 1.0), 2)  # sale discount
        reach = int(impressions / (c.freq * rng.normal(1, 0.05)))
        rows.append([d.strftime("%Y-%m-%d"), c["channel"], c["campaign"], c["objective"], c["product"],
                     round(spend, 2), impressions, reach, clicks, orders, new, orders - new,
                     revenue, c["audience"], c["creative"]])

raw = pd.DataFrame(rows, columns=[
    "Date","Marketing_Channel","Campaign_Name","Campaign_Objective","Product_Name",
    "Ad_Spend","Impressions","Reach","Clicks","Orders","New_Customers",
    "Returning_Customers","Revenue","Audience_Type","Ad_Creative_Type"])
raw = raw.astype(object)   # allow text to be injected into numeric columns

# ---------------------------------------------------------------------------
# Inject realistic data-quality problems (the kind you see in real exports)
# ---------------------------------------------------------------------------
n = len(raw)
pick = lambda k: rng.choice(n, size=k, replace=False)

# 1. Inconsistent channel labels
for idx, label in zip(pick(14), ["meta ads","Facebook Ads","META ADS","google","Google ads ",
                                  " Google Ads","Ecommerce","e-commerce","Marketplace","meta",
                                  "Facebook","GOOGLE ADS","E commerce","Meta Ads "]):
    ch = raw.at[idx, "Marketing_Channel"]
    # only use a label that belongs to the row's true channel
    family = {"Meta Ads": ["meta ads","Facebook Ads","META ADS","meta","Facebook","Meta Ads "],
              "Google Ads": ["google","Google ads "," Google Ads","GOOGLE ADS"],
              "E-commerce": ["Ecommerce","e-commerce","Marketplace","E commerce"]}[ch]
    raw.at[idx, "Marketing_Channel"] = family[rng.integers(len(family))]

# 2. Campaign names with stray spaces / lower case
for idx in pick(8):
    raw.at[idx, "Campaign_Name"] = "  " + str(raw.at[idx, "Campaign_Name"]).lower() + " "

# 3. Mixed date formats (all unambiguous)
for idx in pick(20):
    d = pd.Timestamp(raw.at[idx, "Date"])
    raw.at[idx, "Date"] = d.strftime("%d/%m/%Y") if idx % 2 else d.strftime("%b %d, %Y")

# 4. Ad spend exported as text with currency symbol and thousands separator
for idx in pick(15):
    raw.at[idx, "Ad_Spend"] = f"₹{float(raw.at[idx, 'Ad_Spend']):,.2f}"

# 5. Missing values
for idx in pick(10): raw.at[idx, "Audience_Type"] = np.nan      # recoverable from campaign
for idx in pick(6):  raw.at[idx, "Ad_Creative_Type"] = np.nan   # recoverable from campaign
for idx in pick(8):  raw.at[idx, "Revenue"] = np.nan            # orders known, revenue lost
for idx in pick(4):  raw.at[idx, "Ad_Spend"] = np.nan           # not recoverable

# 6. Negative ad spend (sign entry error)
for idx in pick(3):
    v = raw.at[idx, "Ad_Spend"]
    if isinstance(v, float) and not np.isnan(v):
        raw.at[idx, "Ad_Spend"] = -v

# 7. Clicks and Impressions swapped
for idx in pick(3):
    raw.at[idx, "Clicks"], raw.at[idx, "Impressions"] = raw.at[idx, "Impressions"], raw.at[idx, "Clicks"]

# 8. Reach greater than impressions
for idx in pick(3):
    raw.at[idx, "Reach"] = int(int(raw.at[idx, "Impressions"]) * 1.4)

# 9. New + Returning does not add up to Orders
for idx in pick(4):
    raw.at[idx, "Returning_Customers"] = int(raw.at[idx, "Returning_Customers"]) + int(rng.integers(3, 9))

# 10. Decimal-shift typo: spend entered 10x too high (impressions unaffected)
for idx in pick(2):
    v = raw.at[idx, "Ad_Spend"]
    if isinstance(v, float) and not np.isnan(v) and v > 0:
        raw.at[idx, "Ad_Spend"] = round(v * 10, 2)

# 11. Exact duplicate rows (double export)
dups = raw.iloc[pick(12)]
raw = pd.concat([raw, dups], ignore_index=True)
raw = raw.sample(frac=1, random_state=7).reset_index(drop=True)   # shuffle like a messy export

(ROOT / "data/raw").mkdir(parents=True, exist_ok=True)
raw.to_csv(ROOT / "data/raw/campaign_performance_raw.csv", index=False, encoding="utf-8-sig")
master.to_csv(ROOT / "data/raw/campaign_master.csv", index=False)
print("raw rows:", len(raw))

# ---------------------------------------------------------------------------
# A/B test: same campaign, audience and budget; only the creative differs.
# Variation A = static product image, Variation B = UGC testimonial video.
# ---------------------------------------------------------------------------
ab_dates = pd.date_range("2026-10-01", periods=14, freq="D")
spec = {"A": dict(name="A - Static Product Image", cpm=140, ctr=0.0120, cvr=0.028, aov=620),
        "B": dict(name="B - UGC Testimonial Video", cpm=150, ctr=0.0138, cvr=0.029, aov=630)}
ab = []
for d in ab_dates:
    for v, s in spec.items():
        spend = 2500 * rng.normal(1, 0.06)
        imp = int(spend / (s["cpm"] * rng.normal(1, 0.05)) * 1000)
        clk = int(rng.binomial(imp, s["ctr"]))
        od = int(rng.binomial(clk, s["cvr"]))
        rev = round(od * s["aov"] * rng.lognormal(0, 0.08), 2)
        ab.append([d.strftime("%Y-%m-%d"), "META_ABTest_HairSerum", v, s["name"],
                   "Hair Growth Serum", "Prospecting - Interest", round(spend, 2), imp, clk, od, rev])
ab = pd.DataFrame(ab, columns=["Date","Campaign_Name","Variation","Variation_Description",
                               "Product_Name","Audience_Type","Ad_Spend","Impressions",
                               "Clicks","Orders","Revenue"])
ab.to_csv(ROOT / "data/raw/ab_test_raw.csv", index=False)
print("ab rows:", len(ab))

# Business Recommendations

*Simulated data. Each recommendation separates **Observed** (what the data shows), **Assumption** (what we
believe but did not measure) and **Recommendation** (what to do, and how to know if it worked).*

Assumptions used throughout: contribution margin before ads 55% (website) / 40% (marketplaces) → break-even ROAS
1.82x / 2.50x; target CAC ₹450.

---

### 1. Scale the generic hair-fall search campaign in steps (scaling profitable campaigns, reducing CAC)
- **Observed:** GGL_Search_Generic_HairFall has ROAS 3.15x, CAC ₹237 (the lowest of any acquisition campaign), 84% new
  customers, 1,753 orders, and a stable trend (+1% ROAS from the first four weeks to the last four).
- **Assumption:** there is more search demand available at a similar cost. More spend will usually cost more per customer (diminishing returns).
- **Recommendation:** raise the budget +25% for 2 weeks. Expand to related keywords ("hair fall serum for women", "hair regrowth serum").
  **Success metric:** marginal CAC stays under ₹450 and ROAS stays above 2.5x. Add the next +25% only if that holds.

### 2. Judge brand search and retargeting on incrementality, not ROAS (optimising budgets)
- **Observed:** brand search ROAS 12.41x but only 29% new customers. Retargeting ROAS 3.24x–5.74x with 40–44% new
  customers. Blended ROAS falls from 2.19x to 1.76x without brand search.
- **Assumption:** many of these buyers would have purchased anyway (they were already searching for the brand or had items in their cart).
- **Recommendation:** keep budgets flat. Run a **geo or holdout test** (pause brand ads in a few cities for 2 weeks) to
  measure true incremental sales. Report "ROAS excluding brand" as a standing KPI.

### 3. Fix or cut the loss-making acquisition campaigns (reducing CAC, improving ROAS)
- **Observed:** Flipkart shampoo (ROAS 0.44x, CAC ₹1,471, CVR 3.21% vs 7.01% marketplace median), Nykaa VitC (1.40x vs
  2.50x break-even), Meta static hair serum (1.16x, CAC ₹671).
- **Recommendation:** cut each by 30% now. For Flipkart, check the product page: price vs competitors, ratings,
  images and delivery promise, because a low CVR on a marketplace usually points to the listing, not the ad.
  Move Meta static spend toward the UGC format (see #5). Review again after 2 weeks: if ROAS is still below break-even, pause.

### 4. Refresh the Vitamin C influencer creative now (improving creative performance)
- **Observed:** CTR fell from 1.46% to 1.11% (−24%) and ROAS from 1.43x to 1.11x (−22%) between the first four and the last four weeks.
  The other Meta prospecting campaigns stayed stable, so the decline is specific to this creative.
- **Recommendation:** launch 3–4 new hooks or creators. Set a fatigue alert: if weekly CTR is more than 20% below the
  launch-week average, rotate the creative. Fund this from the ₹4,514/day creative-testing reserve.

### 5. Use UGC video for prospecting, but keep testing for sales impact (increasing conversion rates)
- **Observed:** in the controlled A/B test, the UGC video lifted CTR by 15.9% (p < 0.001) and cut CPC by 8%. CVR (2.94% vs 2.90%)
  and ROAS (1.64x vs 1.65x) were not different. The test could only have detected a CVR change of about ±42%.
- **Recommendation:** make UGC the default top-of-funnel format, because cheaper clicks are a real gain. Do **not** claim
  it increases sales. Next test: the **landing page** (CVR is now the bottleneck). Ideas: show the before/after
  result at the top, a "hair fall in 4 weeks" claim with evidence, and a COD or free-shipping badge.

### 6. Raise AOV with bundles (increasing average order value)
- **Observed:** Hair Care Combo has AOV ₹997 vs ₹588 overall, and its retargeting campaign has the second-highest ROAS (5.74x).
  Hair Growth Serum drives 61% of revenue.
- **Assumption:** buyers of the serum are likely to buy a matching shampoo or oil.
- **Recommendation:** add a "Complete Hair Kit" bundle and a free-shipping threshold just above the current AOV
  (e.g. ₹699) on serum landing pages. **Success metric:** AOV +10% with no drop in CVR.

### 7. Cap awareness spend and measure it properly; plan for festive peaks (optimising budgets)
- **Observed:** the two awareness campaigns use 12.7% of spend at ROAS 0.54x. During the festive sale, revenue per day rose
  39%, CVR rose from 4.11% to 5.37%, and CPM rose 16%. Weekends convert slightly better (CVR 4.55% vs 4.14%).
- **Assumption:** awareness ads create future demand (brand searches, later purchases) that last-click data doesn't capture.
- **Recommendation:** trim awareness spend by 15% and track it on CPM, reach, video views and changes in brand-search volume. Pre-load budgets
  for sale periods and weekends, because conversion improves enough to more than offset higher CPMs.

---

### How to monitor (weekly)
ROAS and ROAS excluding brand · CAC vs ₹450 · CVR by campaign · weekly CTR vs launch week (fatigue) ·
new-customer share · contribution after ads.

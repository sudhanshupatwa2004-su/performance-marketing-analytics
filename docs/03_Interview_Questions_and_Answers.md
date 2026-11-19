# Interview Questions & Answers: Performance Marketing Associate

Answers use this project's (simulated) numbers. Practise saying them out loud in your own words.
Pattern for every answer: **what I did → why → what I found → what I'd do next.**

---

## A. About the project

**1. Walk me through your project in 60 seconds.**
I built an end-to-end performance marketing analysis for a WishCare-inspired beauty brand using 60 days of simulated
data covering 14 campaigns across Meta, Google and marketplaces. I cleaned a messy export in Python (duplicates,
text currency values, mixed date formats, impossible values) and logged every fix. I calculated seven KPIs the same way in Excel,
SQL, Python and Power BI, and analysed performance by channel, campaign, product, audience and creative. Then I built a budget
plan that uses margin, CAC, conversion rate, volume and trend instead of ROAS alone. The key insights: Google search is the
most efficient channel; blended ROAS of 2.19x falls to 1.76x without brand search; one influencer creative is fatiguing;
and the A/B test found a click winner but not a sales winner.

**2. Why did you use simulated data?**
Real campaign data is confidential. Simulating it let me build realistic patterns (channel benchmarks, a festive sale,
creative fatigue) and realistic data-quality problems, so I could practise the whole workflow. I say clearly that the
numbers aren't WishCare's.

**3. What was the hardest part?**
Making the budget recommendation honest. The highest-ROAS campaign, brand search, is the one I recommended *not* scaling,
because only 29% of its buyers are new and many would probably have bought anyway. Explaining why "best ROAS" doesn't mean "best place
for more money" took the most thought.

## B. KPIs

**4. Define ROAS, CAC, CTR, CPC, CPM, CVR and AOV.**
ROAS = revenue ÷ ad spend (₹ of sales per ₹1 of ads). CAC = ad spend ÷ new customers. CTR = clicks ÷ impressions × 100.
CPC = spend ÷ clicks. CPM = spend ÷ impressions × 1,000. CVR = orders ÷ clicks × 100. AOV = revenue ÷ orders.
Overall in my data: 2.19x, ₹438, 1.29%, ₹11.39, ₹147, 4.25%, ₹588.

**5. How do these KPIs connect?**
Revenue = Impressions × CTR × CVR × AOV. So ROAS = (CTR × CVR × AOV) ÷ CPM × 1,000. That tells me which lever to pull:
a low ROAS can come from expensive impressions (CPM), weak ads (CTR), a weak landing page (CVR) or small baskets (AOV).

**6. Is a ROAS of 2x good?**
It depends on margin. Break-even ROAS = 1 ÷ contribution margin. With my assumed 55% margin on website sales, break-even
is 1.82x, so 2x is barely profitable. On marketplaces with an assumed 40% margin, break-even is 2.5x, so 2x loses money.

**7. Why calculate KPIs from totals and not average them?**
Averaging daily ROAS gives a ₹500 day the same weight as a ₹9,000 day. In my data, the average of row ROAS is 2.65x but
the true ROAS is 2.19x. I used SUM ÷ SUM everywhere, so Excel, SQL, Python and Power BI all match.

**8. How did you handle divide-by-zero?**
If the denominator is 0, the KPI is left blank: `IF(x=0,"",…)` in Excel, `NULLIF(x,0)` in SQL, `DIVIDE()` in DAX,
and a NaN-returning function in Python. Twenty campaign-days had zero new customers, so CAC is blank there, not 0
(0 would wrongly look like free customers).

**9. CAC vs CPA, what's the difference?**
CPA is cost per acquisition of any conversion, here an order (spend ÷ orders). CAC is cost per *new* customer.
CAC is higher, because returning buyers aren't counted in the denominator. I used CAC because the goal was customer acquisition.

## C. Data cleaning

**10. What problems did you find and how did you fix them?**
852 raw rows became 840. I removed 12 exact duplicates, mapped 13 channel spellings to 3 standard names, fixed campaign-name case and spaces,
parsed three date formats, converted "₹4,512.00" text to numbers, filled missing audience/creative from the same campaign,
fixed negative spend, swapped clicks/impressions where clicks were higher, capped reach at impressions, fixed
customers ≠ orders, corrected two decimal-shift spend errors, and estimated 4 missing spend and 8 missing revenue values (all flagged).

**11. Why did you correct the spend outliers but keep the revenue outliers?**
The spend rows had CPM about 13× the campaign median while impressions were normal, so the platform didn't deliver
13× more ads. That pattern means an extra digit was typed. The revenue outliers were real demand: 3 of the 12 were in the festive sale and the rest were normal
weekend variation. Removing real peaks would hide real demand.

**12. Why impute missing spend instead of dropping the row?**
Dropping the row also drops its revenue and orders, which distorts ROAS. CPM is stable within a campaign (it varies by less than 10%),
so impressions × median CPM is a tight estimate. I flagged those rows so the values can be replaced with invoice figures later.

## D. Analysis & findings

**13. Which channel performed best?**
Google Ads: 35% of spend, 50% of revenue, ROAS 3.13x, CAC ₹318. But a large part of that is brand search, so I also
looked at it without brand.

**14. E-commerce had the highest CVR but the lowest ROAS. How?**
Marketplace shoppers are ready to buy (CVR 7.15%), but clicks are expensive (CPC ₹27.40 vs ₹8–12 elsewhere).
Cost per order = CPC ÷ CVR ≈ ₹383, versus about ₹184 on Google. That makes it the most expensive way to get an order, and marketplace fees cut the margin further.

**15. How did you spot creative fatigue?**
I compared weekly CTR per campaign. The Vitamin C influencer video fell from 1.46% to 1.11% (−24%) between the first
and last four weeks, while the other Meta prospecting ads were flat. ROAS fell 22% with it.

**16. What did the festive sale show?**
Revenue per day rose 39% and CVR rose from 4.11% to 5.37%, while CPM rose 16% because auctions get more competitive. ROAS
still improved (2.30x vs 2.18x), so it's worth increasing budgets before sale periods.

## E. Budget optimisation

**17. How did you decide where to move budget?**
I scored each campaign on six things: ROAS vs margin-based break-even (a profit index), CAC vs the ₹450 target, CVR vs the
channel median, order volume (at least 300 orders), trend (first vs last four weeks), and campaign role (awareness, brand, retargeting).
Only campaigns that passed every test were scaled.

**18. Why not move everything to brand search?**
Its audience is people already searching for the brand. That's a fixed pool, and 71% of its buyers are returning customers.
Extra money would mostly buy sales that would happen anyway. I'd run a holdout test to measure true incrementality.

**19. Would you turn off the awareness campaigns because ROAS is 0.5x?**
No. They're built for reach, not immediate sales, and last-click attribution undercounts them. I'd cap them (−15%) and
judge them on CPM, reach, video views and changes in brand-search volume, not on ROAS.

**20. How confident are you in the reallocation impact?**
Not very, and I say that. It's a scenario with an explicit assumption that extra spend earns 70% of current ROAS. The real
answer comes from scaling in steps and measuring marginal CAC each week.

## F. A/B testing

**21. Which ad won?**
B (the UGC video) won on clicks: CTR 1.40% vs 1.21%, +15.9%, p < 0.001, with cheaper CPC. On conversion it did not:
CVR 2.90% vs 2.94% (p = 0.92) and ROAS 1.65x vs 1.64x (p = 0.96). So B gives cheaper traffic but is not proven to sell more.

**22. What is a p-value?**
The probability of seeing a difference at least this big if the two ads were actually equal. Below 0.05, we call it
statistically significant: it's unlikely to be luck.

**23. Why can't you call a winner on CVR?**
With about 3,000 clicks per ad, the test could only reliably detect a CVR change of roughly ±42%. The observed gap was 1.5%,
which is practically zero. To detect a 15% lift in CVR, we'd need roughly 25,000 clicks per variation (about 8x more data).

**24. What makes an A/B test valid?**
Change only one thing (here, the creative), split the audience randomly, use the same budget and dates, decide the sample size and success metric
in advance, and don't stop early just because one ad looks ahead.

## G. Tools & role

**25. Why use four tools?**
Excel for quick audits and stakeholder-friendly formulas, SQL for querying data where it lives, Python for cleaning,
statistics and repeatability, and Power BI for an interactive dashboard managers can filter themselves. I checked that
all four produce identical KPIs (there's a Checks sheet in the Excel file).

**26. How would you explain this dashboard to a non-technical manager?**
"Top row: are we making money from ads? Middle: where it's coming from. Bottom: which campaigns to grow, fix or cut,
and what our latest ad test showed. Use the slicers to look at any channel, campaign or product."

**27. What would you do in your first week as a Performance Marketing Associate?**
Audit tracking (pixel and conversion API, UTMs), build a daily KPI report like this one, check for creative fatigue, list
campaigns below break-even, and propose one well-designed test.

**28. What are the limitations of your analysis?**
Simulated data; last-click attribution; assumed margins and target CAC; no lifetime value, returns or organic baseline;
60 days only. With real data, I'd add cohort LTV, incrementality tests and marketing mix modelling.

**29. Meta CPMs suddenly rise 30%. What do you check?**
Seasonality or a sale period (auction competition), audience size or frequency (an audience that's too narrow or saturated),
creative quality ranking, and placement mix. Then check whether CTR and CVR held. If ROAS held, the increase may be acceptable.

**30. Which SQL concepts did you use?**
Aggregation with GROUP BY, NULLIF for zero-guards, CTEs (WITH), window functions (RANK for top and bottom campaigns,
LAG for week-over-week change), conditional aggregation (SUM(CASE WHEN …)) for the fatigue check, and CASE for
first-pass budget decisions.

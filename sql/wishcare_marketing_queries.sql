/* =====================================================================
   WishCare-Inspired Performance Marketing Analytics - SQL Query File
   ---------------------------------------------------------------------
   Data: SIMULATED. Not WishCare's real campaign data.
   Written for MySQL 8.0. Every query also runs in SQLite (tested),
   except the LOAD DATA command in Section 0, which is MySQL-only.

   KEY RULE (same as Excel, Python and Power BI):
     KPI = SUM(numerator) / SUM(denominator)   -> ratio of totals
     NEVER AVG(row_level_ratio)
   NULLIF(x, 0) turns a 0 denominator into NULL, so the division
   returns NULL ("not available") instead of an error or infinity.
   ===================================================================== */


/* ---------------------------------------------------------------------
   SECTION 0. Create the tables and load the cleaned CSVs
   --------------------------------------------------------------------- */
CREATE DATABASE IF NOT EXISTS wishcare_marketing;
USE wishcare_marketing;

DROP TABLE IF EXISTS campaign_performance;
CREATE TABLE campaign_performance (
    Row_ID               INT PRIMARY KEY,
    Date                 DATE          NOT NULL,
    Week_Start           DATE          NOT NULL,   -- Monday of the week
    Day_Name             VARCHAR(10),
    Is_Weekend           VARCHAR(10),
    Is_Festive_Sale      VARCHAR(3),
    Marketing_Channel    VARCHAR(20)   NOT NULL,
    Campaign_Name        VARCHAR(60)   NOT NULL,
    Campaign_Objective   VARCHAR(20),
    Product_Name         VARCHAR(40),
    Audience_Type        VARCHAR(40),
    Ad_Creative_Type     VARCHAR(30),
    Ad_Spend             DECIMAL(12,2) NOT NULL,
    Impressions          INT           NOT NULL,
    Reach                INT,
    Clicks               INT           NOT NULL,
    Orders               INT           NOT NULL,
    New_Customers        INT           NOT NULL,
    Returning_Customers  INT           NOT NULL,
    Revenue              DECIMAL(12,2) NOT NULL,
    Data_Quality_Flag    VARCHAR(120)
);

DROP TABLE IF EXISTS ab_test;
CREATE TABLE ab_test (
    Date                  DATE,
    Campaign_Name         VARCHAR(60),
    Variation             CHAR(1),
    Variation_Description VARCHAR(60),
    Product_Name          VARCHAR(40),
    Audience_Type         VARCHAR(40),
    Ad_Spend              DECIMAL(12,2),
    Impressions           INT,
    Clicks                INT,
    Orders                INT,
    Revenue               DECIMAL(12,2)
);

-- MySQL only. Change the path to where you saved the CSV.
-- (Or use MySQL Workbench: right-click table > Table Data Import Wizard.)
-- LOAD DATA LOCAL INFILE 'C:/wishcare/data/cleaned/campaign_performance_cleaned.csv'
-- INTO TABLE campaign_performance
-- FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"'
-- LINES TERMINATED BY '\n' IGNORE 1 ROWS;
-- LOAD DATA LOCAL INFILE 'C:/wishcare/data/cleaned/ab_test_cleaned.csv'
-- INTO TABLE ab_test FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"'
-- LINES TERMINATED BY '\n' IGNORE 1 ROWS;


/* ---------------------------------------------------------------------
   SECTION 1. Data validation - run these first. Every check should
   return 0 (or the expected number) on the cleaned table.
   --------------------------------------------------------------------- */

-- Q1.1 Row count and date range (expect 840 rows, 2026-09-01 to 2026-10-30)
SELECT COUNT(*) AS total_rows, MIN(Date) AS first_date, MAX(Date) AS last_date,
       COUNT(DISTINCT Campaign_Name) AS campaigns, COUNT(DISTINCT Date) AS days
FROM campaign_performance;

-- Q1.2 Duplicate campaign-days (expect no rows)
SELECT Date, Campaign_Name, COUNT(*) AS copies
FROM campaign_performance
GROUP BY Date, Campaign_Name
HAVING COUNT(*) > 1;

-- Q1.3 Logical errors (expect every column = 0)
SELECT
  SUM(CASE WHEN Clicks > Impressions THEN 1 ELSE 0 END)                    AS clicks_gt_impr,
  SUM(CASE WHEN Reach  > Impressions THEN 1 ELSE 0 END)                    AS reach_gt_impr,
  SUM(CASE WHEN Orders > Clicks      THEN 1 ELSE 0 END)                    AS orders_gt_clicks,
  SUM(CASE WHEN New_Customers + Returning_Customers <> Orders THEN 1 ELSE 0 END) AS customer_mismatch,
  SUM(CASE WHEN Ad_Spend < 0 OR Revenue < 0 THEN 1 ELSE 0 END)             AS negative_values,
  SUM(CASE WHEN Ad_Spend IS NULL OR Revenue IS NULL THEN 1 ELSE 0 END)     AS missing_values
FROM campaign_performance;

-- Q1.4 Rows corrected during cleaning (for transparency)
SELECT Data_Quality_Flag, COUNT(*) AS rows_flagged
FROM campaign_performance
WHERE Data_Quality_Flag IS NOT NULL AND Data_Quality_Flag <> ''
GROUP BY Data_Quality_Flag
ORDER BY rows_flagged DESC;


/* ---------------------------------------------------------------------
   SECTION 2. Overall KPIs (these feed the dashboard KPI cards)
   --------------------------------------------------------------------- */

-- Q2.1 Headline KPIs for the full period
SELECT
  ROUND(SUM(Ad_Spend), 2)                                        AS total_ad_spend,
  ROUND(SUM(Revenue), 2)                                         AS total_revenue,
  SUM(Orders)                                                    AS total_orders,
  SUM(New_Customers)                                             AS new_customers,
  ROUND(SUM(Revenue)  / NULLIF(SUM(Ad_Spend), 0), 2)             AS roas,
  ROUND(SUM(Ad_Spend) / NULLIF(SUM(New_Customers), 0), 2)        AS cac,
  ROUND(SUM(Clicks)   * 100.0 / NULLIF(SUM(Impressions), 0), 2)  AS ctr_pct,
  ROUND(SUM(Ad_Spend) / NULLIF(SUM(Clicks), 0), 2)               AS cpc,
  ROUND(SUM(Ad_Spend) * 1000.0 / NULLIF(SUM(Impressions), 0), 2) AS cpm,
  ROUND(SUM(Orders)   * 100.0 / NULLIF(SUM(Clicks), 0), 2)       AS cvr_pct,
  ROUND(SUM(Revenue)  / NULLIF(SUM(Orders), 0), 2)               AS aov
FROM campaign_performance;
-- Note: "* 100.0" (with a decimal) forces decimal maths. In some databases
-- integer / integer drops the decimals (e.g. 3/4 = 0).


/* ---------------------------------------------------------------------
   SECTION 3. Channel, campaign and product performance
   --------------------------------------------------------------------- */

-- Q3.1 Channel-wise performance with share of spend and revenue
WITH ch AS (
  SELECT Marketing_Channel,
         SUM(Ad_Spend) AS spend, SUM(Revenue) AS revenue, SUM(Impressions) AS impressions,
         SUM(Clicks) AS clicks, SUM(Orders) AS orders,
         SUM(New_Customers) AS new_cust, SUM(Returning_Customers) AS ret_cust
  FROM campaign_performance
  GROUP BY Marketing_Channel
)
SELECT Marketing_Channel,
       ROUND(spend, 2) AS ad_spend, ROUND(revenue, 2) AS revenue, orders, new_cust,
       ROUND(spend   * 100.0 / (SELECT SUM(Ad_Spend) FROM campaign_performance), 1) AS spend_share_pct,
       ROUND(revenue * 100.0 / (SELECT SUM(Revenue)  FROM campaign_performance), 1) AS revenue_share_pct,
       ROUND(revenue / NULLIF(spend, 0), 2)               AS roas,
       ROUND(spend   / NULLIF(new_cust, 0), 2)            AS cac,
       ROUND(clicks  * 100.0 / NULLIF(impressions, 0), 2) AS ctr_pct,
       ROUND(spend   / NULLIF(clicks, 0), 2)              AS cpc,
       ROUND(orders  * 100.0 / NULLIF(clicks, 0), 2)      AS cvr_pct,
       ROUND(revenue / NULLIF(orders, 0), 2)              AS aov,
       ROUND(new_cust * 100.0 / NULLIF(orders, 0), 1)     AS new_customer_share_pct
FROM ch
ORDER BY revenue DESC;

-- Q3.2 Campaign-wise performance, ranked by ROAS
SELECT Campaign_Name, Marketing_Channel, Campaign_Objective,
       ROUND(SUM(Ad_Spend), 2)                                       AS ad_spend,
       ROUND(SUM(Revenue), 2)                                        AS revenue,
       SUM(Orders)                                                   AS orders,
       SUM(New_Customers)                                            AS new_customers,
       ROUND(SUM(Revenue)  / NULLIF(SUM(Ad_Spend), 0), 2)            AS roas,
       ROUND(SUM(Ad_Spend) / NULLIF(SUM(New_Customers), 0), 2)       AS cac,
       ROUND(SUM(Clicks)   * 100.0 / NULLIF(SUM(Impressions), 0), 2) AS ctr_pct,
       ROUND(SUM(Ad_Spend) / NULLIF(SUM(Clicks), 0), 2)              AS cpc,
       ROUND(SUM(Orders)   * 100.0 / NULLIF(SUM(Clicks), 0), 2)      AS cvr_pct,
       ROUND(SUM(Revenue)  / NULLIF(SUM(Orders), 0), 2)              AS aov
FROM campaign_performance
GROUP BY Campaign_Name, Marketing_Channel, Campaign_Objective
ORDER BY roas DESC;

-- Q3.3 Top 3 and bottom 3 campaigns by ROAS using a window function
WITH c AS (
  SELECT Campaign_Name,
         SUM(Revenue) / NULLIF(SUM(Ad_Spend), 0) AS roas,
         SUM(Ad_Spend) / NULLIF(SUM(New_Customers), 0) AS cac
  FROM campaign_performance
  GROUP BY Campaign_Name
), r AS (
  SELECT Campaign_Name, roas, cac,
         RANK() OVER (ORDER BY roas DESC) AS rank_best,
         RANK() OVER (ORDER BY roas ASC)  AS rank_worst
  FROM c
)
SELECT CASE WHEN rank_best <= 3 THEN 'Top 3' ELSE 'Bottom 3' END AS bucket,
       Campaign_Name, ROUND(roas, 2) AS roas, ROUND(cac, 2) AS cac
FROM r
WHERE rank_best <= 3 OR rank_worst <= 3
ORDER BY roas DESC;

-- Q3.4 Product-wise revenue and share
SELECT Product_Name,
       ROUND(SUM(Revenue), 2) AS revenue,
       ROUND(SUM(Revenue) * 100.0 / (SELECT SUM(Revenue) FROM campaign_performance), 1) AS revenue_share_pct,
       SUM(Orders) AS orders,
       ROUND(SUM(Revenue) / NULLIF(SUM(Ad_Spend), 0), 2) AS roas,
       ROUND(SUM(Revenue) / NULLIF(SUM(Orders), 0), 2)   AS aov
FROM campaign_performance
GROUP BY Product_Name
ORDER BY revenue DESC;


/* ---------------------------------------------------------------------
   SECTION 4. Trends
   --------------------------------------------------------------------- */

-- Q4.1 Daily trend
SELECT Date,
       ROUND(SUM(Ad_Spend), 2) AS ad_spend,
       ROUND(SUM(Revenue), 2)  AS revenue,
       SUM(Orders)             AS orders,
       ROUND(SUM(Revenue) / NULLIF(SUM(Ad_Spend), 0), 2) AS roas
FROM campaign_performance
GROUP BY Date
ORDER BY Date;

-- Q4.2 Weekly trend with week-over-week revenue change (LAG window function)
WITH w AS (
  SELECT Week_Start, SUM(Ad_Spend) AS spend, SUM(Revenue) AS revenue,
         SUM(Orders) AS orders, SUM(New_Customers) AS new_cust,
         COUNT(DISTINCT Date) AS days_in_week
  FROM campaign_performance
  GROUP BY Week_Start
)
SELECT Week_Start, days_in_week,
       ROUND(spend, 2) AS ad_spend, ROUND(revenue, 2) AS revenue, orders,
       ROUND(revenue / NULLIF(spend, 0), 2)  AS roas,
       ROUND(spend / NULLIF(new_cust, 0), 2) AS cac,
       ROUND((revenue - LAG(revenue) OVER (ORDER BY Week_Start)) * 100.0
             / NULLIF(LAG(revenue) OVER (ORDER BY Week_Start), 0), 1) AS revenue_wow_pct
FROM w
ORDER BY Week_Start;
-- days_in_week matters: the first week (from Tue 1 Sep) and last week
-- (to Fri 30 Oct) are partial, so their totals look low. Compare averages per day.

-- Q4.3 Festive sale vs normal days (per-day averages so they are comparable)
SELECT Is_Festive_Sale,
       COUNT(DISTINCT Date)                                     AS days,
       ROUND(SUM(Revenue)  / COUNT(DISTINCT Date), 2)           AS revenue_per_day,
       ROUND(SUM(Ad_Spend) / COUNT(DISTINCT Date), 2)           AS spend_per_day,
       ROUND(SUM(Revenue)  / NULLIF(SUM(Ad_Spend), 0), 2)       AS roas,
       ROUND(SUM(Orders) * 100.0 / NULLIF(SUM(Clicks), 0), 2)   AS cvr_pct,
       ROUND(SUM(Ad_Spend) * 1000.0 / NULLIF(SUM(Impressions), 0), 2) AS cpm
FROM campaign_performance
GROUP BY Is_Festive_Sale;

-- Q4.4 Weekday vs weekend
SELECT Is_Weekend,
       ROUND(SUM(Revenue) / NULLIF(SUM(Ad_Spend), 0), 2)        AS roas,
       ROUND(SUM(Orders) * 100.0 / NULLIF(SUM(Clicks), 0), 2)   AS cvr_pct
FROM campaign_performance
GROUP BY Is_Weekend;


/* ---------------------------------------------------------------------
   SECTION 5. Customers, audiences and creatives
   --------------------------------------------------------------------- */

-- Q5.1 New vs returning customers by channel
SELECT Marketing_Channel,
       SUM(New_Customers) AS new_customers,
       SUM(Returning_Customers) AS returning_customers,
       ROUND(SUM(New_Customers) * 100.0 / NULLIF(SUM(Orders), 0), 1) AS new_share_pct,
       ROUND(SUM(Ad_Spend) / NULLIF(SUM(New_Customers), 0), 2)       AS cac
FROM campaign_performance
GROUP BY Marketing_Channel
ORDER BY cac;

-- Q5.2 Audience-wise performance
SELECT Audience_Type,
       ROUND(SUM(Ad_Spend), 2) AS ad_spend,
       ROUND(SUM(Revenue) / NULLIF(SUM(Ad_Spend), 0), 2)             AS roas,
       ROUND(SUM(Ad_Spend) / NULLIF(SUM(New_Customers), 0), 2)       AS cac,
       ROUND(SUM(Orders) * 100.0 / NULLIF(SUM(Clicks), 0), 2)        AS cvr_pct,
       ROUND(SUM(New_Customers) * 100.0 / NULLIF(SUM(Orders), 0), 1) AS new_share_pct
FROM campaign_performance
GROUP BY Audience_Type
ORDER BY roas DESC;

-- Q5.3 Creative-wise performance
SELECT Ad_Creative_Type,
       ROUND(SUM(Ad_Spend), 2) AS ad_spend,
       ROUND(SUM(Clicks) * 100.0 / NULLIF(SUM(Impressions), 0), 2) AS ctr_pct,
       ROUND(SUM(Ad_Spend) / NULLIF(SUM(Clicks), 0), 2)            AS cpc,
       ROUND(SUM(Orders) * 100.0 / NULLIF(SUM(Clicks), 0), 2)      AS cvr_pct,
       ROUND(SUM(Revenue) / NULLIF(SUM(Ad_Spend), 0), 2)           AS roas
FROM campaign_performance
GROUP BY Ad_Creative_Type
ORDER BY roas DESC;

-- Q5.4 Creative fatigue: CTR in first 4 weeks vs last 4 weeks per campaign
SELECT Campaign_Name,
       ROUND(SUM(CASE WHEN Date <  '2026-09-29' THEN Clicks END) * 100.0
           / NULLIF(SUM(CASE WHEN Date <  '2026-09-29' THEN Impressions END), 0), 2) AS ctr_first_4w,
       ROUND(SUM(CASE WHEN Date >= '2026-10-03' THEN Clicks END) * 100.0
           / NULLIF(SUM(CASE WHEN Date >= '2026-10-03' THEN Impressions END), 0), 2) AS ctr_last_4w
FROM campaign_performance
GROUP BY Campaign_Name
ORDER BY (SUM(CASE WHEN Date >= '2026-10-03' THEN Clicks END) * 1.0
          / NULLIF(SUM(CASE WHEN Date >= '2026-10-03' THEN Impressions END), 0))
       / NULLIF(SUM(CASE WHEN Date <  '2026-09-29' THEN Clicks END) * 1.0
          / NULLIF(SUM(CASE WHEN Date <  '2026-09-29' THEN Impressions END), 0), 0);


/* ---------------------------------------------------------------------
   SECTION 6. Budget optimisation support
   Assumptions (not data): contribution margin before ads = 55% for own
   website traffic (Meta/Google), 40% for marketplaces (E-commerce);
   target CAC = Rs 450; minimum 300 orders to trust a campaign.
   --------------------------------------------------------------------- */

-- Q6.1 Profitability after ad cost and a first-pass budget flag
WITH c AS (
  SELECT Campaign_Name, Marketing_Channel, Campaign_Objective, Audience_Type,
         SUM(Ad_Spend) AS spend, SUM(Revenue) AS revenue, SUM(Orders) AS orders,
         SUM(New_Customers) AS new_cust, SUM(Clicks) AS clicks,
         CASE WHEN Marketing_Channel = 'E-commerce' THEN 0.40 ELSE 0.55 END AS margin
  FROM campaign_performance
  GROUP BY Campaign_Name, Marketing_Channel, Campaign_Objective, Audience_Type
), k AS (
  SELECT *, revenue / NULLIF(spend, 0)  AS roas,
            spend / NULLIF(new_cust, 0) AS cac,
            1.0 / margin                AS breakeven_roas,
            revenue * margin - spend    AS contribution_after_ads
  FROM c
)
SELECT Campaign_Name, Marketing_Channel,
       ROUND(roas, 2) AS roas, ROUND(breakeven_roas, 2) AS breakeven_roas,
       ROUND(roas / breakeven_roas, 2) AS profit_index,
       ROUND(cac, 2) AS cac, orders,
       ROUND(new_cust * 100.0 / NULLIF(orders, 0), 1) AS new_share_pct,
       ROUND(contribution_after_ads, 2) AS contribution_after_ads,
       CASE
         WHEN Campaign_Objective = 'Awareness'                        THEN 'Cap & measure differently'
         WHEN Audience_Type = 'Brand Search'
           OR Audience_Type LIKE 'Retargeting%'                        THEN 'Maintain (limited headroom)'
         WHEN roas / breakeven_roas >= 1.3 AND cac <= 450 AND orders >= 300 THEN 'Scale test (+25%)'
         WHEN roas / breakeven_roas <  1.0 AND cac >  540              THEN 'Reduce (-30%) & fix'
         ELSE 'Optimise (hold budget)'
       END AS first_pass_decision
FROM k
ORDER BY profit_index DESC;
-- The Python notebook adds one more test (trend: ROAS change between first
-- and last 4 weeks) before confirming a scale decision.


/* ---------------------------------------------------------------------
   SECTION 7. A/B test summary
   --------------------------------------------------------------------- */

-- Q7.1 Side-by-side comparison of the two ads
SELECT Variation, Variation_Description,
       SUM(Impressions) AS impressions, SUM(Clicks) AS clicks, SUM(Orders) AS orders,
       ROUND(SUM(Revenue), 2)  AS revenue,
       ROUND(SUM(Ad_Spend), 2) AS ad_spend,
       ROUND(SUM(Clicks) * 100.0 / NULLIF(SUM(Impressions), 0), 2) AS ctr_pct,
       ROUND(SUM(Ad_Spend) / NULLIF(SUM(Clicks), 0), 2)            AS cpc,
       ROUND(SUM(Orders) * 100.0 / NULLIF(SUM(Clicks), 0), 2)      AS cvr_pct,
       ROUND(SUM(Revenue) / NULLIF(SUM(Ad_Spend), 0), 2)           AS roas
FROM ab_test
GROUP BY Variation, Variation_Description
ORDER BY Variation;
-- SQL shows WHAT happened. Whether the gap is real or luck needs a
-- significance test (z-test), done in the Python notebook and Excel AB_Test sheet.

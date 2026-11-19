# Power BI Dashboard: Build Guide (Beginner, Step by Step)

**Goal:** an interactive, one-page performance marketing dashboard with 11 KPI cards, 8 charts and 4 slicers.
**Time needed:** 2–3 hours the first time.
**Files you need:** `data/cleaned/campaign_performance_cleaned.csv`, `data/cleaned/ab_test_cleaned.csv`,
`powerbi/DAX_Measures.dax`, `powerbi/WishCare_PowerBI_Theme.json`
**Layout preview:** `charts/11_dashboard_mockup.png`

> **Why there is no .pbix file:** Power BI Desktop only runs on Windows and can't be scripted to produce a
> finished .pbix here. This guide, the DAX file and the theme give you everything you need to build it
> yourself in Power BI Desktop. Building it yourself is also what lets you explain it in an interview.

---

## Step 1. Load the data
1. Open **Power BI Desktop**. Choose **Home → Get data → Text/CSV** and select `campaign_performance_cleaned.csv`.
2. Click **Transform Data** (not Load). Power Query opens.
3. Rename the query to **Campaign** (right panel → Properties → Name).
4. Check the column types (the icon in each column header):
   - `Date`, `Week_Start` → **Date**
   - `Ad_Spend`, `Revenue` → **Fixed decimal number** (currency)
   - `Impressions`, `Reach`, `Clicks`, `Orders`, `New_Customers`, `Returning_Customers`, `Row_ID` → **Whole number**
   - all others → **Text**
   If a date shows as an error, right-click the column → **Change Type → Using Locale → Date, English (India)**.
5. **Close & Apply**. Repeat for `ab_test_cleaned.csv` and name it **AB_Test**.

*Why:* wrong data types are the most common beginner error. A number stored as text can't be summed.

## Step 2. Build the data model
1. **Modeling → New table** and paste the `Date = ADDCOLUMNS(CALENDAR(...))` code from `DAX_Measures.dax`.
2. Select the Date table → **Table tools → Mark as date table** → column `Date`.
3. Go to **Model view** (left sidebar, third icon). Drag `Date[Date]` onto `Campaign[Date]`.
   You should see a **1 → \*** (one-to-many) relationship with a single-direction arrow.
4. Leave `AB_Test` unconnected. It is a separate experiment and should not react to the main date slicer.

*Why a Date table?* It makes time functions (7-day average, week grouping) reliable and gives you one
date slicer that filters everything.

## Step 3. Create the measures
1. Right-click the **Campaign** table → **New measure**. Paste one measure at a time from `DAX_Measures.dax`
   (base sums first, then KPIs, then A/B measures; put the A/B measures on the AB_Test table).
2. Set formats (select the measure → Measure tools ribbon):
   | Measure | Format |
   |---|---|
   | Total Ad Spend, Total Revenue | Currency ₹, 0 decimals |
   | ROAS | Decimal, 2 places (add "x" in the card title) |
   | CAC, CPC, CPM, AOV | Currency ₹, 2 decimals |
   | CTR %, CVR % | Decimal, 2 places |
   | Orders, New Customers | Whole number, thousands separator |
3. **Create a display folder:** select several measures → Properties → Display folder `KPIs`. This keeps things tidy.

**Rule:** always use **measures**, never drag the `Ad_Spend` column into a card and choose "Average".
Measures recalculate correctly for every filter combination because they divide sums.

## Step 4. Apply the theme and page setup
1. **View → Themes → Browse for themes** → choose `WishCare_PowerBI_Theme.json`.
2. **Format page** (click empty canvas → paint-roller): Canvas settings → Type **16:9**; Canvas background is set by the theme (#F4F6FA).
3. Add a **Text box** at top-left: "WishCare-Inspired Performance Marketing Dashboard" (20 pt, bold) and a
   second line "Simulated data | Meta Ads, Google Ads, E-commerce".

## Step 5. Slicers (top-right row)
| Slicer | Field | Style |
|---|---|---|
| Date | `Date[Date]` | Slicer settings → Style **Between** (date range slider) |
| Marketing Channel | `Campaign[Marketing_Channel]` | **Dropdown**, multi-select with Ctrl |
| Campaign | `Campaign[Campaign_Name]` | Dropdown, add **Search** (… menu → Search) |
| Product | `Campaign[Product_Name]` | Dropdown |

Then **View → Sync slicers** if you add more pages. Put the four slicers in one row so the filters are easy to see.

## Step 6. KPI cards (one row of 11)
1. Insert a **Card (new)** visual. Drag the measures **Total Ad Spend, Total Revenue, Total Orders, New Customers,
   ROAS, CAC, CTR %, CPC, CPM, CVR %, AOV** into it. The new card shows several values in a row.
   (In older Power BI versions: create 11 separate **Card** visuals and align with Format → Align → Distribute horizontally.)
2. Format: callout value 20–24 pt, label 9 pt grey, rounded border (theme already sets this).
3. Optional: add a reference label on ROAS showing `Break-even ROAS` to give context.

## Step 7. The 8 charts
| # | Visual | Build steps | Why this chart |
|---|---|---|---|
| 1 | **Revenue by marketing channel** | Clustered **bar** chart. Y-axis = `Marketing_Channel`, X-axis = `Total Revenue`. Data labels on. Sort by revenue descending. | Bars compare sizes accurately; horizontal bars keep labels readable |
| 2 | **Ad Spend by campaign** | Clustered bar. Y = `Campaign_Name`, X = `Total Ad Spend`, Legend = `Marketing_Channel` (or colour bars by channel via fx). | Shows where the money goes |
| 3 | **ROAS by campaign** | Clustered bar. Y = `Campaign_Name`, X = `ROAS`. Add an **X-axis constant line** at 1.82 labelled "Break-even (website)" (Analytics pane → Constant line). Optional: Bars → Colour → fx → Field value → `ROAS Colour`. | Ranks efficiency; the break-even line shows which campaigns lose money |
| 4 | **Daily revenue trend** | **Line** chart. X = `Date[Date]` (set to Continuous, not hierarchy), Y = `Total Revenue`. Add `Revenue 7D Avg` as a second line. | Shows trend, spikes (festive sale) and seasonality |
| 5 | **Product-wise revenue** | Bar chart. Y = `Product_Name`, X = `Total Revenue`; tooltip adds ROAS and AOV. | Which products ads are really selling |
| 6 | **Campaign conversion comparison** | Bar chart. Y = `Campaign_Name`, X = `CVR %`. Add `CTR %` to tooltips. (Or a **scatter**: X = CTR %, Y = CVR %, Size = Total Orders, Details = Campaign_Name.) | Separates "ad gets clicks" from "visitor buys" problems |
| 7 | **New vs returning customers** | **Stacked bar**. Y = `Marketing_Channel`, X = `New Customers` and `Returning Customers`. | Shows which channels truly acquire new buyers |
| 8 | **A/B testing comparison** | **Clustered column**. X = `AB_Test[Variation_Description]`, Y = `AB CTR %`, `AB CVR %`, `AB ROAS` (or three small columns). Add a **Card** with `AB Verdict`. | Shows the test result *and* whether it's significant |

**Tooltips:** for every campaign chart, drag `ROAS`, `CAC`, `CVR %`, `Total Orders` into the Tooltips well.
Hover shows full context without cluttering the visual.

## Step 8. Interactions and polish
1. Click a chart → **Format → Edit interactions**. Make sure clicking a channel bar filters the other visuals
   (filter icon) and that the A/B visuals are set to **None**, so they don't react.
2. Add a **Reset filters** button: Insert → Buttons → Reset → Action → Bookmark (create a bookmark of the
   unfiltered page first via View → Bookmarks).
3. Use the same channel colour everywhere: Meta = blue #2A78D6, Google = orange #EB6834, E-commerce = green #1BAF7A
   (Format → Bars → Colours → per category).
4. Turn off gridlines you don't need, keep axis titles short, and remove chart titles that repeat the axis.
5. Add a small footnote text box: "Simulated data. Margin and target-CAC assumptions are planning inputs."

## Step 9. Check your numbers
With no filters, your cards must show the values below. They match Excel (`KPI_Summary` sheet), SQL (Q2.1) and Python:

| KPI | Expected |
|---|---|
| Total Ad Spend | ₹33,15,653 |
| Total Revenue | ₹72,67,424 |
| Total Orders | 12,359 |
| New Customers | 7,576 |
| ROAS | 2.19 |
| CAC | ₹437.65 |
| CTR % | 1.29 |
| CPC | ₹11.39 |
| CPM | ₹146.92 |
| CVR % | 4.25 |
| AOV | ₹588.03 |

Filter to **Google Ads** only → ROAS 3.13, CAC ₹318.32. Filter to **GGL_Search_Brand** → ROAS 12.41.

## Dashboard design choices (explain these in an interview)
- **Top-to-bottom story:** headline KPIs, then where revenue comes from (channel, product, trend), then campaign-level diagnosis, then customers and testing.
- **Bars over pies:** people compare bar lengths more accurately than slice angles.
- **One colour per channel, everywhere:** the viewer learns the colour once.
- **Context lines:** the break-even ROAS line turns "1.5x" into "losing money", which a number alone doesn't show.
- **Ratio-of-totals measures:** every number reconciles across Excel, SQL and Power BI.
- **A/B test isolated:** it's a separate experiment and should not be filtered by the main date slicer.

## Optional page 2: Budget optimisation
Import `data/cleaned/budget_recommendations.csv` as a table **Budget**. Add a **Table** visual with
`Campaign_Name, ROAS, Breakeven_ROAS, Profit_Index, CAC, Decision, Budget_Change_%`, and conditional-format the
Decision column (Format → Cell elements → Background colour → Rules).

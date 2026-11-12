# App Store Pricing → Apple Ads FP&A Model

An end‑to‑end finance project. It pulls live pricing data for **5,805 apps** from Apple's public App Store endpoints, turns that data into **demand drivers for an App Store search‑ads business**, and builds the FP&A deliverables on top of them:

- a driver‑based forecast
- a headcount / opex / capex plan
- a budget‑vs‑actual variance analysis
- an investment case
- an interactive dashboard
- an executive summary

**Live:** [Interactive FP&A dashboard](https://qichenz-0825.github.io/appstore-pricing-analysis/reports/dashboard.html) · [Pricing analysis report](https://qichenz-0825.github.io/appstore-pricing-analysis/reports/report.html) · [Executive summary](reports/Executive_Summary.md)

> Apple does not disclose Apple Ads revenue, advertisers or costs. All financial inputs are **illustrative** and are clearly marked (blue cells in Excel). The App Store data, the methods and the model structure are real.

## Deliverables

| File | What it is |
|---|---|
| [`reports/Executive_Summary.md`](reports/Executive_Summary.md) | One‑page leadership readout: results, variance drivers, risks, actions |
| [`reports/Apple_Ads_FPA_Model.xlsx`](reports/Apple_Ads_FPA_Model.xlsx) | Live‑formula Excel model (10 tabs, scenario switch, 0 formula errors) |
| [`reports/dashboard.html`](reports/dashboard.html) | Interactive dashboard: scenario toggle, driver sliders, drill‑down, sortable variance table, sensitivity heatmap |
| [`reports/report.html`](reports/report.html) | App Store pricing analysis (interactive charts) |
| [`reports/tableau/`](reports/tableau) | Tidy extracts ready for Tableau |
| [`reports/tables/`](reports/tables) | Every analysis table as CSV |

Both HTML pages are also served on GitHub Pages (see the **Live** links above).

## How the project maps to the Apple Ads Financial Analyst role

| Responsibility in the job description | Where it is in this project |
|---|---|
| Forecasting, annual planning | `Revenue_Forecast`: FY27 monthly plan by category = advertisers × spend/advertiser × seasonality, with Base / Upside / Downside scenarios |
| Headcount, cost, capital expenditures | `Opex_Headcount`: hiring plan by function, fully‑loaded cost, non‑HC opex, capex with straight‑line depreciation |
| Month‑end / quarter close, variance analysis | `Budget_vs_Actual`: Q1 variance split into **volume** (advertisers) and **rate** (spend/advertiser), with a reconciliation check and top drivers |
| Evaluate the financial impact of business decisions | `Market_Investment`: 3‑year NPV and payback for incremental sales investment by storefront, sized with a tax‑adjusted price index and an advertiser‑eligibility haircut for mainland China's approval requirement |
| Analyze costs, optimize profitability | `PnL_Summary`: quarterly P&L, operating margin, revenue per employee |
| Present findings to senior leadership | `Executive_Summary.md` and the `Dashboard` tab |
| Ad hoc analysis and decision support | `Sensitivity` grid plus the dashboard's live sliders |
| Excel, Tableau | Excel model built with formulas only (no pasted results); Tableau‑ready CSV extracts |
| Ability to work with ambiguity | Undisclosed financials handled with transparent, swappable assumptions; drivers come from observable data |

## Key findings from the data

- **99%** of Top Grossing apps are free to download. Ad budgets come from in‑app purchase and subscription economics, not from upfront prices.
- **98%** of paid prices end in .99. The median paid app costs **$3.99**. Category medians range from $1.99 (Social, Shopping) to $5.99 (Medical).
- Within the Top Paid charts, pricier apps tend to rank *higher* (negative Spearman correlation between price and rank in 20 of 22 categories).
- The same app costs **33% more in the UK** on the sticker price, but only **11% more** after removing VAT. US and Canadian prices exclude tax; UK, EU, Japan, Australia, India and Korea prices include it.
- Top Paid apps update every **133 days** (median), against **7 days** for Top Free.

## Model logic

```
Category advertisers  = FY26 advertisers × engagement share          (ratings volume of the category's top apps)
Spend / advertiser    = base spend × (0.4 + 0.6 × WTP index), normalized   (WTP index = category median paid price ÷ average)
Monthly revenue       = advertisers × (1+g_adv)^(m/12) × spend × (1+g_spend)^(m/12) × seasonality_m
Volume variance       = (Actual adv − Budget adv) × Budget rev/adv
Rate variance         = (Actual rev/adv − Budget rev/adv) × Actual adv
Market NPV            = −upfront + Σ (revenue_t × margin − opex) / (1+r)^t,
                        revenue = advertisers × eligibility × US rev/adv × tax‑adjusted price index × capture_t
```

The dashboard runs the same formulas in JavaScript. It was checked against the Excel model: the Base case matches to the cent.

## Data sources (no API key needed)

| Source | Used for |
|---|---|
| `itunes.apple.com/{cc}/rss/{topfree,toppaid,topgrossing}applications/limit=100/genre=G/json` | Top‑chart membership and rank, 22 categories × 3 charts |
| `itunes.apple.com/lookup?id=…&country=…` | Price, currency, ratings, release/update dates; local prices in 10 storefronts |
| `api.frankfurter.app` (ECB reference rates) | Converting local prices to USD |

## Reproduce

```bash
pip install -r requirements.txt
python3 src/collect.py          # ~5 min, rate-limited by Apple
python3 src/analyze.py          # pricing analysis -> reports/report.html, reports/tables/
python3 src/build_model.py      # Excel model + reports/model_inputs.json
python3 src/build_dashboard.py  # reports/dashboard.html
python3 src/export_tableau.py   # reports/tableau/*.csv
```

`data/raw/` (full API responses, about 57 MB) is git‑ignored; `collect.py` regenerates it.

## Limitations

- The data is a one‑day snapshot of the charts. Re‑run `collect.py` on a schedule to build a time series.
- In‑app purchase and subscription prices are not exposed by these endpoints.
- VAT/GST rates are standard headline rates; verify them against Apple's tax schedule.
- Q1 "actuals" are simulated around the budget (seeded) to demonstrate the variance workflow.
- All 9 storefronts are Apple Ads markets. Mainland China requires advertiser business and app approval ([Apple Ads Help](https://ads.apple.com/app-store/help/get-started/0074-advertise-in-mainland-china)); this is modeled as an illustrative 70% eligibility haircut, adjustable in Excel and on the dashboard.

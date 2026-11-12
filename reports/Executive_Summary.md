# Apple Ads — FY27 Plan: Executive Summary

*Prepared as a portfolio exercise. Data pulled 2026‑11‑11 from Apple's public App Store endpoints. Apple does not disclose Apple Ads financials, so every revenue, advertiser, cost and "actuals" figure below comes from **illustrative inputs**. The structure, drivers and method are the deliverable.*

## Bottom line

| | Downside | **Base** | Upside |
|---|---:|---:|---:|
| FY27 ad revenue ($mm) | 5,229 | **5,587** | 5,868 |
| YoY vs FY26 run‑rate ($5,040mm) | +3.7% | **+10.9%** | +16.4% |
| Operating income ($mm) | 3,469 | **3,875** | 4,169 |
| Operating margin | 66.3% | **69.3%** | 71.0% |

- **Q1 FY27 came in $25.0mm (−1.7%) below budget.** Volume was favorable: advertiser count added +$2.5mm. Rate was unfavorable: spend per advertiser cost −$27.5mm. The miss is a monetization issue rather than an acquisition issue. Largest swings: Finance (+$15.8mm), Games (−$12.0mm), Travel (−$10.9mm).
- **Investment priority:** at Base assumptions, the next sales dollar earns the most in **Japan** (3‑yr NPV $72.0mm), then the **UK** ($65.6mm) and **mainland China** ($50.5mm). All 9 markets are NPV‑positive. China is the largest addressable market before approval constraints, but advertisers there must get business and app approval (licenses, industry qualifications) before ads run. The model assumes 70% of advertisers clear it. At 100%, China would rank first ($87.3mm), so the approval rate is the swing factor.
- **Headcount:** the plan exits FY27 at 2,019 HC (+219), with $677mm of people cost and $175mm of capex.

## What the App Store data says about ad demand

1. **Nearly all revenue comes from free‑to‑download apps.** 99% of Top Grossing apps across 22 categories cost $0 upfront. Advertisers' willingness to pay is therefore driven by in‑app purchase and subscription lifetime value (LTV), not by upfront price. Advertiser growth will come mainly from freemium developers.
2. **Willingness to pay varies by category.** Median paid‑app price runs from $1.99 (Social, Shopping) to $5.99 (Medical). The model uses that spread as a willingness‑to‑pay (WTP) index that scales spend per advertiser.
3. **Engagement sizes the advertiser pool.** The ratings volume of each category's top apps sets its share of advertisers. Photo & Video, Finance and Games are the largest revenue categories in the plan.
4. **Paid apps are a long tail.** The median Top Paid app was last updated 133 days ago, against 7 days for Top Free. Top Paid apps also have far fewer ratings. These developers are a small part of the ad market.
5. **Compare prices across countries net of tax.** The same paid app's sticker price is 33% higher in the UK than in the US, but only 11% higher once the 20% VAT is removed (US prices exclude sales tax). Sizing international revenue on gross sticker prices would overstate UK, EU and APAC markets.

## Risks and sensitivities

- Each **±5 pts of advertiser growth** moves FY27 revenue by about **±$135mm**. Each **±4 pts of spend‑per‑advertiser growth** moves it by about **±$110mm**. See the Sensitivity tab.
- **Rate risk is the larger one:** Q1 shows that auction pricing and budget expansion can lag even when advertiser acquisition is on plan.
- **Seasonality:** December carries about 22% above‑average spend. A soft holiday quarter cannot be fully recovered in Q2, when January resets spend.

## Recommended actions

1. Rebase the Q2–Q4 rate assumption by holding spend per advertiser at the Q1 actual level. Show the resulting gap to the annual plan.
2. Run a deep dive on unfavorable rate variance in Games and Travel: cost per tap (CPT) trends, budget caps and the competitive set.
3. Build a full business case for Japan and the UK. For China, first validate the advertiser approval pass rate: it decides whether China ranks first or third.

*Model: `reports/Apple_Ads_FPA_Model.xlsx` (live formulas; change the scenario in Assumptions!B3). Interactive version: `reports/dashboard.html`.*

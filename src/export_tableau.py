"""Export tidy, Tableau-ready extracts to reports/tableau/.

apps_charts.csv  one row per (app, chart, category) with price, rating, dates
intl_prices.csv  one row per (app, storefront) with local price, USD price, gross & tax-adjusted index
category_kpis.csv one row per category with the model's demand drivers
"""
import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "reports" / "tableau"
VAT = {"gb": 0.20, "de": 0.19, "jp": 0.10, "au": 0.10, "in": 0.18, "kr": 0.10,
       "br": 0.0, "cn": 0.06, "ca": 0.0, "us": 0.0}


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    apps = pd.read_csv(ROOT / "data" / "apps.csv")
    charts = pd.read_csv(ROOT / "data" / "charts.csv")
    df = charts.merge(apps, on="trackId")
    df["chart"] = df.chart.map({"top_free": "Top Free", "top_paid": "Top Paid", "top_grossing": "Top Grossing"})
    df["is_free"] = df.price.eq(0)
    df["price_ends_99"] = (df.price.round(2) * 100 % 100).round().eq(99) & ~df.is_free
    df["release_date"] = pd.to_datetime(df.releaseDate).dt.date
    df["last_update"] = pd.to_datetime(df.currentVersionReleaseDate).dt.date
    cols = ["trackId", "trackName", "sellerName", "chart", "chart_genre", "rank", "price", "is_free",
            "price_ends_99", "averageUserRating", "userRatingCount", "contentAdvisoryRating",
            "release_date", "last_update", "fileSizeBytes", "trackViewUrl"]
    df[cols].rename(columns={"chart_genre": "category", "averageUserRating": "rating",
                             "userRatingCount": "rating_count"}).to_csv(OUT / "apps_charts.csv", index=False)

    intl = pd.read_csv(ROOT / "reports" / "tables" / "intl_prices_usd.csv")
    intl["vat_rate"] = intl.country.map(VAT)
    intl["index_tax_adjusted"] = intl["index"] / (1 + intl.vat_rate)
    intl = intl.merge(apps[["trackId", "trackName", "primaryGenreName"]], on="trackId", how="left")
    intl.rename(columns={"price": "local_price", "usd": "usd_price", "index": "index_gross",
                         "primaryGenreName": "category"}).to_csv(OUT / "intl_prices.csv", index=False)

    inputs = json.loads((ROOT / "reports" / "model_inputs.json").read_text())
    cat = pd.DataFrame(inputs["categories"])
    cat["engagement_share"] = cat.reviews / cat.reviews.sum()
    cat["wtp_index"] = cat.median_paid / cat.median_paid.mean()
    cat.rename(columns={"name": "category", "reviews": "rating_count_top_charts",
                        "free_share": "top_grossing_free_share", "median_paid": "median_paid_price"}) \
       .to_csv(OUT / "category_kpis.csv", index=False)
    print("Wrote", *[p.name for p in OUT.glob("*.csv")])


if __name__ == "__main__":
    main()

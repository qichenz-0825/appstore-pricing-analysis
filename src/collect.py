"""
Collect App Store pricing data from Apple's public endpoints (no API key needed).

1. Top Free / Top Paid / Top Grossing charts per category (iTunes RSS feeds)
2. Full app metadata for every charted app (iTunes Lookup API)
3. Local prices of the paid apps in several other storefronts (cross-country comparison)
4. FX rates to convert local prices into USD (ECB rates via frankfurter.app, static fallback)

Output: data/raw/*.json and data/apps.csv, data/intl_prices.csv
"""
from __future__ import annotations

import argparse
import json
import time
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
DATA = ROOT / "data"

# App Store genre IDs (software)
GENRES = {
    6000: "Business", 6001: "Weather", 6002: "Utilities", 6003: "Travel",
    6004: "Sports", 6005: "Social Networking", 6006: "Reference",
    6007: "Productivity", 6008: "Photo & Video", 6009: "News",
    6010: "Navigation", 6011: "Music", 6012: "Lifestyle",
    6013: "Health & Fitness", 6014: "Games", 6015: "Finance",
    6016: "Entertainment", 6017: "Education", 6018: "Books",
    6020: "Medical", 6023: "Food & Drink", 6024: "Shopping",
}

CHARTS = {
    "top_free": "topfreeapplications",
    "top_paid": "toppaidapplications",
    "top_grossing": "topgrossingapplications",
}

# Storefronts for the cross-country comparison
COUNTRIES = ["us", "gb", "de", "jp", "cn", "in", "br", "au", "ca", "kr"]

# Fallback FX (units of local currency per 1 USD) — used only if the live fetch fails
STATIC_FX = {
    "USD": 1.0, "GBP": 0.78, "EUR": 0.92, "JPY": 152.0, "CNY": 7.15,
    "INR": 84.0, "BRL": 5.7, "AUD": 1.52, "CAD": 1.39, "KRW": 1390.0,
}

SESSION = requests.Session()
SESSION.headers["User-Agent"] = "appstore-pricing-analysis/1.0"
PAUSE = 1.5  # seconds between calls; Apple throttles at roughly 20 req/min


def get_json(url: str, params: dict | None = None, retries: int = 4) -> dict:
    for attempt in range(retries):
        try:
            r = SESSION.get(url, params=params, timeout=30)
            if r.status_code == 200:
                time.sleep(PAUSE)
                return r.json()
            if r.status_code in (403, 429, 503):
                wait = 15 * (attempt + 1)
                print(f"  throttled ({r.status_code}), waiting {wait}s")
                time.sleep(wait)
                continue
            r.raise_for_status()
        except (requests.RequestException, json.JSONDecodeError) as e:
            print(f"  error {e!r}, retrying")
            time.sleep(5 * (attempt + 1))
    raise RuntimeError(f"Failed to fetch {url}")


def fetch_charts(country: str, limit: int) -> pd.DataFrame:
    rows = []
    for gid, gname in GENRES.items():
        for chart, feed in CHARTS.items():
            url = f"https://itunes.apple.com/{country}/rss/{feed}/limit={limit}/genre={gid}/json"
            try:
                entries = get_json(url).get("feed", {}).get("entry", [])
            except RuntimeError as e:
                print(f"  skip {gname}/{chart}: {e}")
                continue
            if isinstance(entries, dict):
                entries = [entries]
            for rank, e in enumerate(entries, 1):
                rows.append({
                    "trackId": int(e["id"]["attributes"]["im:id"]),
                    "chart": chart,
                    "chart_genre": gname,
                    "rank": rank,
                })
            print(f"  {gname:<18} {chart:<13} {len(entries):>3} apps")
    return pd.DataFrame(rows)


def lookup(ids: list[int], country: str) -> list[dict]:
    out = []
    for i in range(0, len(ids), 150):
        batch = ids[i:i + 150]
        res = get_json("https://itunes.apple.com/lookup",
                       {"id": ",".join(map(str, batch)), "country": country, "entity": "software"})
        out.extend(r for r in res.get("results", []) if r.get("wrapperType") == "software")
    return out


KEEP = [
    "trackId", "trackName", "artistName", "sellerName", "bundleId", "primaryGenreName",
    "price", "currency", "formattedPrice", "averageUserRating", "userRatingCount",
    "contentAdvisoryRating", "releaseDate", "currentVersionReleaseDate", "version",
    "fileSizeBytes", "minimumOsVersion", "trackViewUrl",
]


def fetch_fx() -> dict:
    try:
        r = SESSION.get("https://api.frankfurter.app/latest", params={"from": "USD"}, timeout=20)
        r.raise_for_status()
        rates = {"USD": 1.0, **r.json()["rates"]}
        print(f"  live ECB rates for {r.json()['date']}")
        return {k: rates.get(k, v) for k, v in STATIC_FX.items()}
    except Exception as e:  # noqa: BLE001
        print(f"  FX fetch failed ({e!r}); using static fallback rates")
        return STATIC_FX


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=100, help="apps per chart (max 200)")
    ap.add_argument("--intl-sample", type=int, default=400,
                    help="max paid apps to price in other countries")
    args = ap.parse_args()
    RAW.mkdir(parents=True, exist_ok=True)

    print("1/4 Fetching US top charts")
    charts = fetch_charts("us", args.limit)
    charts.to_csv(DATA / "charts.csv", index=False)

    print("2/4 Looking up app metadata")
    ids = sorted(charts.trackId.unique().tolist())
    meta = lookup(ids, "us")
    (RAW / "lookup_us.json").write_text(json.dumps(meta))
    apps = pd.DataFrame(meta)[KEEP]
    apps.to_csv(DATA / "apps.csv", index=False)
    print(f"  {len(apps)} unique apps")

    print("3/4 Pricing paid apps in other storefronts")
    paid = apps[apps.price > 0].sort_values("userRatingCount", ascending=False)
    paid_ids = paid.trackId.head(args.intl_sample).tolist()
    intl = []
    for c in COUNTRIES:
        res = lookup(paid_ids, c)
        intl += [{"trackId": r["trackId"], "country": c, "price": r.get("price"),
                  "currency": r.get("currency")} for r in res]
        print(f"  {c}: {len(res)} apps available")
    pd.DataFrame(intl).to_csv(DATA / "intl_prices.csv", index=False)

    print("4/4 FX rates")
    fx = fetch_fx()
    (DATA / "fx.json").write_text(json.dumps(fx, indent=2))
    print("Done.")


if __name__ == "__main__":
    main()

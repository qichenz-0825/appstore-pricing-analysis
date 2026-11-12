"""
Analyze the collected App Store data and build an interactive HTML report.

Reads  data/apps.csv, data/charts.csv, data/intl_prices.csv, data/fx.json
Writes reports/report.html and reports/tables/*.csv
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "reports"
TABLES = OUT / "tables"

COUNTRY_NAMES = {"us": "United States", "gb": "United Kingdom", "de": "Germany", "jp": "Japan",
                 "cn": "China", "in": "India", "br": "Brazil", "au": "Australia",
                 "ca": "Canada", "kr": "South Korea"}


def load():
    apps = pd.read_csv(DATA / "apps.csv")
    charts = pd.read_csv(DATA / "charts.csv")
    intl = pd.read_csv(DATA / "intl_prices.csv")
    fx = json.loads((DATA / "fx.json").read_text())
    for c in ("releaseDate", "currentVersionReleaseDate"):
        apps[c] = pd.to_datetime(apps[c], utc=True, errors="coerce")
    now = datetime.now(timezone.utc)
    apps["age_years"] = (now - apps.releaseDate).dt.days / 365.25
    apps["days_since_update"] = (now - apps.currentVersionReleaseDate).dt.days
    apps["size_mb"] = pd.to_numeric(apps.fileSizeBytes, errors="coerce") / 1e6
    df = charts.merge(apps, on="trackId", how="inner")
    return apps, charts, df, intl, fx


# ---------- analyses ----------

def freemium_share(df):
    """Share of Top Grossing apps that are free to download, per category."""
    g = df[df.chart == "top_grossing"]
    t = (g.assign(free=g.price == 0)
          .groupby("chart_genre")
          .agg(apps=("trackId", "size"), free_share=("free", "mean"))
          .sort_values("free_share"))
    return t


def price_points(df):
    p = df[(df.chart == "top_paid") & (df.price > 0)].drop_duplicates("trackId")
    counts = p.price.round(2).value_counts().sort_index()
    charm = (p.price.round(2) * 100 % 100).round().astype(int).eq(99).mean()
    return counts, charm, p


def category_prices(df):
    p = df[(df.chart == "top_paid") & (df.price > 0)]
    t = (p.groupby("chart_genre").price
          .agg(apps="size", median="median", mean="mean",
               p25=lambda s: s.quantile(.25), p75=lambda s: s.quantile(.75), max="max")
          .query("apps >= 5")
          .sort_values("median"))
    return t.round(2)


def rank_vs_price(df):
    """Within each category's Top Paid chart: do cheaper apps rank higher?"""
    p = df[(df.chart == "top_paid") & (df.price > 0)]
    rows = []
    for g, s in p.groupby("chart_genre"):
        if len(s) >= 15:
            rho = s["rank"].corr(s["price"], method="spearman")
            rows.append({"chart_genre": g, "spearman_rank_price": rho, "n": len(s)})
    # Pooled: price quartile vs median rank
    p = p.assign(price_band=pd.qcut(p.price, 4, duplicates="drop"))
    bands = p.groupby("price_band", observed=True).agg(
        median_rank=("rank", "median"), median_rating=("averageUserRating", "median"),
        median_reviews=("userRatingCount", "median"), apps=("trackId", "size"))
    bands.index = bands.index.astype(str)
    return pd.DataFrame(rows).sort_values("spearman_rank_price"), bands


def maintenance(df):
    """How recently are apps updated, by chart (free / paid / grossing)?"""
    d = df.drop_duplicates(["trackId", "chart"])
    return d.groupby("chart").agg(
        apps=("trackId", "size"),
        median_days_since_update=("days_since_update", "median"),
        share_updated_90d=("days_since_update", lambda s: (s <= 90).mean()),
        median_age_years=("age_years", "median"),
        median_rating=("averageUserRating", "median"),
        median_reviews=("userRatingCount", "median"),
    ).round(3)


def intl_index(intl, apps, fx):
    """Price of the same app in each storefront, converted to USD, indexed to US = 100."""
    intl = intl.dropna(subset=["price"]).copy()
    intl = intl[intl.price > 0]
    intl["usd"] = intl.apply(lambda r: r.price / fx.get(r.currency, np.nan), axis=1)
    us = intl[intl.country == "us"].set_index("trackId").usd
    intl = intl[intl.trackId.isin(us.index)]
    intl["index"] = intl.usd / intl.trackId.map(us) * 100
    t = (intl.groupby("country")
             .agg(apps=("trackId", "nunique"), median_index=("index", "median"),
                  p25=("index", lambda s: s.quantile(.25)), p75=("index", lambda s: s.quantile(.75)),
                  currency=("currency", "first"))
             .sort_values("median_index"))
    t["country_name"] = t.index.map(COUNTRY_NAMES)
    # availability: share of US paid apps also sold there
    t["availability"] = t.apps / len(us)
    return t.round(3), intl


# ---------- report ----------

def fig_bar_h(labels, values, name, fmt, hover, axis_title, color="--series-1", ref=None):
    trace = {"type": "bar", "orientation": "h", "y": labels, "x": values, "name": name,
             "marker": {"color": color}, "hovertemplate": hover + "<extra></extra>",
             "text": [fmt.format(v) for v in values], "textposition": "outside",
             "cliponaxis": False}
    layout = {"xaxis": {"title": axis_title}, "yaxis": {"automargin": True},
              "height": max(320, 26 * len(labels) + 90), "bargap": 0.35}
    if ref is not None:
        layout["shapes"] = [{"type": "line", "x0": ref, "x1": ref, "y0": -0.5,
                             "y1": len(labels) - 0.5, "line": {"dash": "dot", "width": 1.5,
                                                               "color": "--text-muted"}}]
    return {"data": [trace], "layout": layout}


def build_report(apps, charts, df, fm, counts, charm, paid, cat, rho, bands, maint, idx):
    n_apps = apps.trackId.nunique()
    n_paid = paid.trackId.nunique()
    fm_all = (df[df.chart == "top_grossing"].price == 0).mean()
    med_paid = paid.price.median()
    top_pt = counts.idxmax()

    figs = {}
    figs["freemium"] = fig_bar_h(
        fm.index.tolist(), (fm.free_share * 100).round(1).tolist(), "Free to download",
        "{:.0f}%", "%{y}: %{x:.0f}% of top‑grossing apps are free", "% of Top Grossing apps that are free",
        ref=float(fm_all * 100))
    figs["freemium"]["layout"]["xaxis"]["range"] = [0, 112]

    top = counts.sort_values(ascending=False).head(15).sort_index()
    figs["points"] = {"data": [{"type": "bar", "x": [f"${p:,.2f}" for p in top.index],
                                "y": top.values.tolist(), "marker": {"color": "--series-1"},
                                "hovertemplate": "%{x}: %{y} apps<extra></extra>"}],
                      "layout": {"xaxis": {"title": "Price point (USD)", "type": "category"},
                                 "yaxis": {"title": "Apps in Top Paid charts"}, "height": 360,
                                 "bargap": 0.3}}

    figs["category"] = {
        "data": [{"type": "bar", "orientation": "h", "y": cat.index.tolist(),
                  "x": (cat.p75 - cat.p25).tolist(), "base": cat.p25.tolist(),
                  "name": "Middle 50% (P25–P75)", "marker": {"color": "--series-1-soft"},
                  "customdata": cat[["p25", "p75", "apps"]].values.tolist(),
                  "hovertemplate": "%{y}<br>P25 $%{customdata[0]:.2f} – P75 $%{customdata[1]:.2f}"
                                   "<br>%{customdata[2]} apps<extra></extra>"},
                 {"type": "scatter", "mode": "markers", "y": cat.index.tolist(),
                  "x": cat["median"].tolist(), "name": "Median",
                  "marker": {"color": "--series-1", "size": 10, "line": {"width": 2, "color": "--surface"}},
                  "hovertemplate": "%{y}: median $%{x:.2f}<extra></extra>"}],
        "layout": {"xaxis": {"title": "Price (USD), log scale", "type": "log", "tickprefix": "$"},
                   "yaxis": {"automargin": True}, "height": max(340, 26 * len(cat) + 100),
                   "barmode": "overlay", "bargap": 0.45, "showlegend": True,
                   "legend": {"orientation": "h", "y": 1.06, "x": 0}}}

    figs["bands"] = {"data": [{"type": "bar", "x": bands.index.tolist(),
                               "y": bands.median_rank.tolist(), "marker": {"color": "--series-1"},
                               "customdata": bands[["apps", "median_rating"]].values.tolist(),
                               "hovertemplate": "Price %{x}<br>median rank %{y}<br>%{customdata[0]} "
                                                "listings · rating %{customdata[1]:.2f}<extra></extra>"}],
                     "layout": {"xaxis": {"title": "Price quartile (USD)", "type": "category"},
                                "yaxis": {"title": "Median chart rank (lower = better)"},
                                "height": 340, "bargap": 0.35}}

    figs["intl"] = fig_bar_h(
        [f"{r.country_name} ({r.currency})" for r in idx.itertuples()],
        idx.median_index.round(1).tolist(), "Price index", "{:.0f}",
        "%{y}: median %{x:.0f} (US = 100)", "Median price of the same app, USD‑converted (US = 100)",
        ref=100)

    def tbl(d, fmts):
        head = "".join(f"<th>{c}</th>" for c in [d.index.name or ""] + list(fmts))
        rows = "".join(
            "<tr><td>{}</td>{}</tr>".format(i, "".join(f"<td>{f.format(r[c])}</td>" for c, f in fmts.items()))
            for i, r in d.iterrows())
        return f"<div class='tw'><table><thead><tr>{head}</tr></thead><tbody>{rows}</tbody></table></div>"

    cheapest, priciest = idx.iloc[0], idx.iloc[-1]
    maint_named = maint.rename(index={"top_free": "Top Free", "top_paid": "Top Paid",
                                      "top_grossing": "Top Grossing"})
    pooled_rho = df[(df.chart == "top_paid") & (df.price > 0)][["rank", "price"]].corr("spearman").iloc[0, 1]

    insights = [
        f"<b>Freemium wins revenue.</b> {fm_all:.0%} of Top Grossing apps are free to download — "
        f"money comes from in‑app purchases and subscriptions, not the upfront price. "
        f"Lowest share: {fm.index[0]} ({fm.free_share.iloc[0]:.0%}).",
        f"<b>Charm pricing dominates.</b> {charm:.0%} of paid apps end in .99; the most common price "
        f"point is ${top_pt:.2f}, and the median paid app costs ${med_paid:.2f}.",
        f"<b>Category matters.</b> Median paid price ranges from ${cat['median'].iloc[0]:.2f} "
        f"({cat.index[0]}) to ${cat['median'].iloc[-1]:.2f} ({cat.index[-1]}).",
        f"<b>Price vs. rank.</b> Pooled Spearman correlation between price and chart rank is "
        f"{pooled_rho:+.2f} ({'cheaper apps tend to rank higher' if pooled_rho > 0.05 else 'little relationship' if abs(pooled_rho) <= 0.05 else 'pricier apps tend to rank higher'}).",
        f"<b>Regional pricing.</b> The same app is cheapest in {cheapest.country_name} "
        f"(index {cheapest.median_index:.0f}) and most expensive in {priciest.country_name} "
        f"(index {priciest.median_index:.0f}) after converting to USD. Much of the gap is tax: "
        f"UK/EU/JP/AU prices include VAT/GST while US prices do not.",
        "<b>Paid apps are a long tail.</b> Median Top Paid app is updated every "
        f"{maint.loc['top_paid','median_days_since_update']:.0f} days vs "
        f"{maint.loc['top_free','median_days_since_update']:.0f} for Top Free, with far fewer ratings — "
        "the developers with ad budgets are the freemium/subscription apps.",
    ]
    implications = [
        "<b>Who buys ads:</b> nearly all revenue-generating apps are free to download, so advertiser LTV comes from "
        "in-app purchases and subscriptions. Plan advertiser growth around freemium developers.",
        "<b>Category mix as a forecast driver:</b> engagement share (ratings volume) sizes each category's advertiser "
        "pool; median paid price is a willingness-to-pay index for spend per advertiser.",
        "<b>International prioritization:</b> compare markets on a tax-adjusted price index, not the sticker price, "
        "when sizing revenue per advertiser abroad.",
        "These drivers feed <code>reports/Apple_Ads_FPA_Model.xlsx</code> (forecast, opex/headcount, "
        "budget-vs-actual, market NPV, sensitivity).",
    ]

    stamp = datetime.now().strftime("%Y-%m-%d")
    html = TEMPLATE
    repl = {
        "{{DATE}}": stamp,
        "{{KPIS}}": "".join(f"<div class='kpi'><div class='v'>{v}</div><div class='l'>{l}</div></div>" for v, l in [
            (f"{n_apps:,}", "unique apps analyzed"), (f"{fm_all:.0%}", "of top‑grossing apps are free"),
            (f"${med_paid:.2f}", "median paid‑app price"), (f"{charm:.0%}", "of paid prices end in .99")]),
        "{{INSIGHTS}}": "".join(f"<li>{i}</li>" for i in insights),
        "{{IMPLICATIONS}}": "".join(f"<li>{i}</li>" for i in implications),
        "{{FIGS}}": json.dumps(figs),
        "{{T_FREEMIUM}}": tbl(fm.sort_values("free_share", ascending=False),
                             {"apps": "{:.0f}", "free_share": "{:.0%}"}),
        "{{T_CAT}}": tbl(cat.sort_values("median", ascending=False),
                        {"apps": "{:.0f}", "median": "${:.2f}", "p25": "${:.2f}", "p75": "${:.2f}", "max": "${:.2f}"}),
        "{{T_BANDS}}": tbl(bands, {"apps": "{:.0f}", "median_rank": "{:.0f}", "median_rating": "{:.2f}",
                                   "median_reviews": "{:,.0f}"}),
        "{{T_MAINT}}": tbl(maint_named, {"apps": "{:.0f}", "median_days_since_update": "{:.0f}",
                                         "share_updated_90d": "{:.0%}", "median_age_years": "{:.1f}",
                                         "median_rating": "{:.2f}", "median_reviews": "{:,.0f}"}),
        "{{T_INTL}}": tbl(idx.set_index("country_name").sort_values("median_index"),
                         {"currency": "{}", "apps": "{:.0f}", "availability": "{:.0%}",
                          "median_index": "{:.0f}", "p25": "{:.0f}", "p75": "{:.0f}"}),
        "{{N_PAID}}": f"{n_paid:,}",
        "{{N_CATS}}": str(charts.chart_genre.nunique()),
    }
    for k, v in repl.items():
        html = html.replace(k, v)
    return html


TEMPLATE = (Path(__file__).parent / "report_template.html").read_text()


def main():
    OUT.mkdir(exist_ok=True)
    TABLES.mkdir(exist_ok=True)
    apps, charts, df, intl, fx = load()
    fm = freemium_share(df)
    counts, charm, paid = price_points(df)
    cat = category_prices(df)
    rho, bands = rank_vs_price(df)
    maint = maintenance(df)
    idx, intl_detail = intl_index(intl, apps, fx)

    for name, t in {"freemium_share": fm, "price_points": counts.rename("apps"),
                    "category_prices": cat, "rank_price_corr": rho, "price_bands": bands,
                    "maintenance": maint, "intl_price_index": idx}.items():
        t.to_csv(TABLES / f"{name}.csv")
    intl_detail.to_csv(TABLES / "intl_prices_usd.csv", index=False)

    html = build_report(apps, charts, df, fm, counts, charm, paid, cat, rho, bands, maint, idx)
    (OUT / "report.html").write_text(html)
    print(f"Wrote {OUT / 'report.html'}")
    print("\nFreemium share (top grossing):\n", fm.tail(5))
    print("\nCategory prices:\n", cat)
    print("\nPrice bands:\n", bands)
    print("\nMaintenance:\n", maint)
    print("\nIntl index:\n", idx)


if __name__ == "__main__":
    main()

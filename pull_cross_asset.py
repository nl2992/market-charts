#!/usr/bin/env python3
"""
Pull the cross-asset price history that the charts workbook rebases.

History comes from two public, unauthenticated sources:

    Yahoo Finance   daily closes for indices, ETFs and futures
    FRED            constant-maturity Treasury yields

Neither is a substitute for a professional data terminal. The series table
below carries the corresponding vendor identifier for each line, so the
workbook can be repointed at a paid feed without changing anything else.

Output is a cached CSV so the build is reproducible and runs offline:

    data/cross_asset_daily.csv    one row per weekday, one column per series

Run:  python3 data-puller/pull_cross_asset.py
"""

import csv
import datetime as dt
import json
import os
import time
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "data")
OUT_CSV = os.path.join(OUT_DIR, "cross_asset_daily.csv")

# Window. The pull starts a month early so every series has a print to carry
# into the first spine date — the base has to be a real close, not a blank.
SPINE_START = dt.date(2023, 1, 2)
SPINE_END = dt.date(2026, 7, 31)
PULL_START = dt.date(2022, 11, 25)

# --------------------------------------------------------------------------
# The universe
# --------------------------------------------------------------------------
# (key, source symbol, label, block, unit note, Bloomberg ticker)
#
# Blocks are the four things a cross-asset book actually allocates between.
# `unit` records what the series is in, because the rebasing hides it: the
# equity indices are local-currency price returns, the bond and commodity
# proxies are USD total return, and FX is USD per unit of the pair as quoted.
#
# Where a benchmark is not free, a liquid ETF stands in and says so. The ETF
# series use Yahoo's adjusted close, so distributions are reinvested — which
# is the right basis for rebasing a bond sleeve, and the wrong one to compare
# against a headline price index without saying so.

SERIES = [
    # key        symbol         label                        block  unit                         BBG ticker
    ("ACWI",     "ACWI",        "MSCI ACWI (ETF)",           "EQ",  "USD total return",          "MXWD Index"),
    ("SPX",      "^GSPC",       "S&P 500",                   "EQ",  "USD price",                 "SPX Index"),
    ("SX5E",     "^STOXX50E",   "Euro Stoxx 50",             "EQ",  "EUR price",                 "SX5E Index"),
    ("AXJ",      "AAXJ",        "MSCI Asia ex-Japan (ETF)",  "EQ",  "USD total return",          "MXASJ Index"),
    ("NKY",      "^N225",       "Nikkei 225",                "EQ",  "JPY price",                 "NKY Index"),
    ("HSI",      "^HSI",        "Hang Seng",                 "EQ",  "HKD price",                 "HSI Index"),
    ("KOSPI",    "^KS11",       "KOSPI",                     "EQ",  "KRW price",                 "KOSPI Index"),
    ("NIFTY",    "^NSEI",       "NIFTY 50",                  "EQ",  "INR price",                 "NIFTY Index"),
    ("TWSE",     "^TWII",       "TAIEX",                     "EQ",  "TWD price",                 "TWSE Index"),

    ("UST710",   "IEF",         "UST 7-10y (ETF)",           "FI",  "USD total return",          "LUATTRUU Index"),
    ("UST20",    "TLT",         "UST 20y+ (ETF)",            "FI",  "USD total return",          "LUTLTRUU Index"),
    ("USAGG",    "AGG",         "US Aggregate (ETF)",        "FI",  "USD total return",          "LBUSTRUU Index"),
    ("USIG",     "LQD",         "US IG credit (ETF)",        "FI",  "USD total return",          "LUACTRUU Index"),
    ("USHY",     "HYG",         "US high yield (ETF)",       "FI",  "USD total return",          "LF98TRUU Index"),
    ("EMSOV",    "EMB",         "EM USD sovereign (ETF)",    "FI",  "USD total return",          "JPEIDIVR Index"),
    ("EXUSGOV",  "IGOV",        "Non-US govt (ETF)",         "FI",  "USD total return, unhedged", "LGTRTRUU Index"),

    ("BRENT",    "BZ=F",        "Brent crude",               "CM",  "USD/bbl, front future",     "CO1 Comdty"),
    ("GOLD",     "GC=F",        "Gold",                      "CM",  "USD/oz",                    "XAU Curncy"),
    ("COPPER",   "HG=F",        "Copper",                    "CM",  "USD/lb, front future",      "HG1 Comdty"),
    ("CMDTY",    "DBC",         "Broad commodity (ETF)",     "CM",  "USD total return",          "BCOMTR Index"),
    ("GSCI",     "^SPGSCI",     "S&P GSCI",                  "CM",  "USD index, spot",           "SPGSCI Index"),

    ("DXY",      "DX-Y.NYB",    "Dollar index",              "FX",  "index",                     "DXY Curncy"),
    ("USDJPY",   "JPY=X",       "USD/JPY",                   "FX",  "JPY per USD",               "USDJPY Curncy"),
    ("USDCNY",   "CNY=X",       "USD/CNY",                   "FX",  "CNY per USD",               "USDCNY Curncy"),
]

# Context levels. These are rates and a volatility index — rebasing a yield to
# 100 would be meaningless, so they sit alongside the rebased block rather than
# inside it.
LEVELS = [
    ("VIX",    "yahoo", "^VIX",  "VIX",              "LV", "index points",   "VIX Index"),
    ("UST10Y", "fred",  "DGS10", "UST 10y yield",    "LV", "per cent",       "USGG10YR Index"),
    ("UST2Y",  "fred",  "DGS2",  "UST 2y yield",     "LV", "per cent",       "USGG2YR Index"),
]

# Yahoo returns an empty body without a browser user agent. FRED does the
# opposite — it stops responding to one — so the header is per-source.
UA_YAHOO = {"User-Agent": "Mozilla/5.0"}
UA_FRED = {}


def _get(url, headers, tries=3):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.read()
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(2 * (i + 1))


def yahoo(symbol):
    """{date: close} for one symbol, using the adjusted close where one exists.

    Adjusted close is what makes the ETF proxies total-return; for an index or
    a futures contract Yahoo returns it unadjusted, so the two agree.
    """
    p1 = int(dt.datetime.combine(PULL_START, dt.time()).timestamp())
    p2 = int(dt.datetime.combine(SPINE_END + dt.timedelta(days=2), dt.time()).timestamp())
    url = ("https://query1.finance.yahoo.com/v8/finance/chart/"
           + urllib.parse.quote(symbol)
           + "?period1=%d&period2=%d&interval=1d" % (p1, p2))
    res = json.loads(_get(url, UA_YAHOO))["chart"]
    if not res.get("result"):
        raise RuntimeError("no data for %s: %s" % (symbol, res.get("error")))
    d = res["result"][0]
    stamps = d["timestamp"]
    ind = d["indicators"]
    vals = ind.get("adjclose", [{}])[0].get("adjclose") or ind["quote"][0]["close"]
    out = {}
    for t, v in zip(stamps, vals):
        if v is None:
            continue
        out[dt.datetime.fromtimestamp(t, dt.timezone.utc).date()] = float(v)
    return out


def fred(series_id):
    """{date: value} for a FRED series. Missing prints come back as '.'."""
    url = ("https://fred.stlouisfed.org/graph/fredgraph.csv?id=%s&cosd=%s&coed=%s"
           % (series_id, PULL_START.isoformat(), SPINE_END.isoformat()))
    rows = _get(url, UA_FRED).decode("utf-8").splitlines()
    out = {}
    for row in csv.reader(rows[1:]):
        if len(row) < 2 or row[1] in (".", ""):
            continue
        out[dt.date.fromisoformat(row[0])] = float(row[1])
    return out


def weekday_spine(start, end):
    """Every Monday-to-Friday date in the window.

    A union-of-calendars spine would privilege whichever market happened to be
    open; a single market's calendar would privilege that market. Weekdays are
    neutral, and the carry-forward below is what every cross-asset comparison
    does with a foreign holiday anyway.
    """
    days, d = [], start
    while d <= end:
        if d.weekday() < 5:
            days.append(d)
        d += dt.timedelta(days=1)
    return days


def carry(raw, spine):
    """Project a series onto the spine, carrying the last print forward.

    Returns (values, traded) — `traded` marks the dates the series actually
    printed on, which is how the workbook keeps holiday carry out of its
    volatility calculation.
    """
    dates = sorted(raw)
    vals, traded, last, i = [], [], None, 0
    for day in spine:
        while i < len(dates) and dates[i] <= day:
            last = raw[dates[i]]
            i += 1
        vals.append(last)
        traded.append(day in raw)
    return vals, traded


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    spine = weekday_spine(SPINE_START, SPINE_END)
    print("spine: %d weekdays, %s to %s" % (len(spine), spine[0], spine[-1]))

    cols, flags = {}, {}
    for key, symbol, label, _blk, _unit, _bbg in SERIES:
        raw = yahoo(symbol)
        cols[key], flags[key] = carry(raw, spine)
        print("  %-8s %-12s %5d prints, %s to %s"
              % (key, symbol, len(raw), min(raw), max(raw)))

    for key, src, symbol, label, _blk, _unit, _bbg in LEVELS:
        raw = yahoo(symbol) if src == "yahoo" else fred(symbol)
        cols[key], flags[key] = carry(raw, spine)
        print("  %-8s %-12s %5d prints, %s to %s"
              % (key, symbol, len(raw), min(raw), max(raw)))

    keys = [s[0] for s in SERIES] + [l[0] for l in LEVELS]
    missing = [(k, i) for k in keys for i, v in enumerate(cols[k]) if v is None]
    if missing:
        raise RuntimeError("gap at the front of the spine: %s"
                           % sorted({k for k, _ in missing}))

    with open(OUT_CSV, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["date"] + keys + [k + "_traded" for k in keys])
        for i, day in enumerate(spine):
            w.writerow([day.isoformat()]
                       + ["%.6f" % cols[k][i] for k in keys]
                       + [1 if flags[k][i] else 0 for k in keys])

    print("\nwrote %s — %d rows x %d series" % (OUT_CSV, len(spine), len(keys)))


if __name__ == "__main__":
    main()

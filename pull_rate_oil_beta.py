#!/usr/bin/env python3
"""
Pull 10y government yields and Brent, and regress one on the other.

The output is a sensitivity: how many basis points a country's 10y moves for
a one per cent move in Brent, estimated on weekly changes since February.
It answers which rates markets are trading the oil shock and which are not.

    Riksbank SWEA   GB, EU (Bund), SE, NO, JP
    FRED            US
    RBA             AU
    Yahoo Finance   Brent front future

Three of the G10 are missing and there is no free daily source for them:
Canada's Bank of Canada benchmark series has stopped returning observations,
Switzerland's SNB bond cube is monthly and stale, and the RBNZ blocks
automated access. They are left out rather than filled with a proxy.

    rate_oil_weekly.csv   the weekly panel the regression runs on
    rate_oil_betas.csv    one beta per country, sorted

Run:  python3 data-puller/pull_rate_oil_beta.py
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

# Betas are estimated from the first Friday of February onwards; the pull
# starts earlier so that first weekly change has a prior observation.
BETA_START = dt.date(2026, 2, 1)
PULL_START = dt.date(2026, 1, 2)

RIKSBANK = "https://api.riksbank.se/swea/v1/Observations/%s/%s"
FRED = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=%s&cosd=%s&coed=%s"
RBA = "https://www.rba.gov.au/statistics/tables/csv/f2-data.csv"
YAHOO = ("https://query1.finance.yahoo.com/v8/finance/chart/%s"
         "?period1=%d&period2=%d&interval=1d")

# (code, label, source, identifier)
COUNTRIES = [
    ("GB", "GB 10y", "riksbank", "GBGVB10Y"),
    ("EU", "EU 10y", "riksbank", "DEGVB10Y"),
    ("SE", "SE 10y", "riksbank", "SEGVB10YC"),
    ("NO", "NO 10y", "riksbank", "NOGVB10Y"),
    ("JP", "JP 10y", "riksbank", "JPGVB10Y"),
    ("US", "US 10y", "fred", "DGS10"),
    ("AU", "AU 10y", "rba", "FCMYGBAG10D"),
]


def _get(url, headers=None, tries=4, timeout=60):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=headers or {})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read()
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(4 * (i + 1))


def riksbank(series_id):
    body = _get(RIKSBANK % (series_id, PULL_START.isoformat()))
    return {dt.date.fromisoformat(o["date"]): float(o["value"])
            for o in json.loads(body) if o.get("value") is not None}


def fred(series_id):
    # FRED refuses a browser user agent here; the default is fine.
    body = _get(FRED % (series_id, PULL_START.isoformat(),
                        dt.date.today().isoformat())).decode("utf-8")
    out = {}
    for row in csv.reader(body.splitlines()[1:]):
        if len(row) > 1 and row[1] not in (".", ""):
            out[dt.date.fromisoformat(row[0])] = float(row[1])
    return out


def rba(series_id):
    body = _get(RBA, headers={"User-Agent": "Mozilla/5.0"}).decode("utf-8-sig")
    rows = list(csv.reader(body.splitlines()))
    header = next(r for r in rows if r and r[0] == "Series ID")
    col = header.index(series_id)
    out = {}
    for row in rows:
        if not row or not row[0]:
            continue
        try:
            d = dt.datetime.strptime(row[0], "%d-%b-%Y").date()
        except ValueError:
            continue
        if len(row) > col and row[col] not in ("", None):
            out[d] = float(row[col])
    return out


def brent():
    p1 = int(dt.datetime.combine(PULL_START - dt.timedelta(days=10),
                                 dt.time()).timestamp())
    p2 = int(dt.datetime.combine(dt.date.today() + dt.timedelta(days=1),
                                 dt.time()).timestamp())
    body = _get(YAHOO % (urllib.parse.quote("BZ=F"), p1, p2),
                headers={"User-Agent": "Mozilla/5.0"})
    d = json.loads(body)["chart"]["result"][0]
    closes = d["indicators"]["quote"][0]["close"]
    out = {}
    for t, v in zip(d["timestamp"], closes):
        if v is not None:
            out[dt.datetime.fromtimestamp(t, dt.timezone.utc).date()] = float(v)
    return out


def fridays(start, end):
    d = start + dt.timedelta(days=(4 - start.weekday()) % 7)
    out = []
    while d <= end:
        out.append(d)
        d += dt.timedelta(days=7)
    return out


def asof(series, day, window=6):
    """Last observation on or before `day`. Markets close on different
    holidays, so an exact-date join would silently drop weeks."""
    for back in range(window + 1):
        v = series.get(day - dt.timedelta(days=back))
        if v is not None:
            return v
    return None


def ols(xs, ys):
    """Slope and R-squared of y on x. Plain least squares — with 25-odd
    weekly points there is nothing a heavier estimator would add."""
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    if sxx == 0:
        return 0.0, 0.0
    beta = sxy / sxx
    alpha = my - beta * mx
    ss_res = sum((y - (alpha + beta * x)) ** 2 for x, y in zip(xs, ys))
    ss_tot = sum((y - my) ** 2 for y in ys)
    return beta, (1 - ss_res / ss_tot) if ss_tot else 0.0


def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    print("yields:")
    series = {}
    for code, label, src, ident in COUNTRIES:
        fn = {"riksbank": riksbank, "fred": fred, "rba": rba}[src]
        s = fn(ident)
        series[code] = s
        print("  %-3s %-10s %-12s %4d obs  %s to %s"
              % (code, src, ident, len(s), min(s), max(s)))
        time.sleep(1)

    oil = brent()
    print("  %-3s %-10s %-12s %4d obs  %s to %s"
          % ("--", "yahoo", "BZ=F", len(oil), min(oil), max(oil)))

    last = min(max(s) for s in list(series.values()) + [oil])
    weeks = [d for d in fridays(PULL_START, last)]
    print("\nweekly spine: %d Fridays, %s to %s" % (len(weeks), weeks[0], weeks[-1]))

    # Weekly levels, then changes. Yield changes in bps, Brent in per cent.
    codes = [c[0] for c in COUNTRIES]
    panel, prev = [], None
    for day in weeks:
        row = {"date": day, "brent": asof(oil, day)}
        for c in codes:
            row[c] = asof(series[c], day)
        if row["brent"] is None or any(row[c] is None for c in codes):
            continue
        if prev is not None:
            rec = {"date": day.isoformat(),
                   "brent_pct": 100.0 * (row["brent"] / prev["brent"] - 1)}
            for c in codes:
                rec[c] = 100.0 * (row[c] - prev[c])       # pp -> bps
            panel.append(rec)
        prev = row

    panel = [r for r in panel if dt.date.fromisoformat(r["date"]) >= BETA_START]
    print("regression sample: %d weekly changes, %s to %s"
          % (len(panel), panel[0]["date"], panel[-1]["date"]))

    with open(os.path.join(OUT_DIR, "rate_oil_weekly.csv"), "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["date", "brent_pct"] + codes)
        for r in panel:
            w.writerow([r["date"], "%.4f" % r["brent_pct"]]
                       + ["%.2f" % r[c] for c in codes])

    xs = [r["brent_pct"] for r in panel]
    betas = []
    for code, label, src, ident in COUNTRIES:
        b, r2 = ols(xs, [r[code] for r in panel])
        betas.append((code, label, b, r2, src, ident))
    betas.sort(key=lambda t: -t[2])

    print("\n%-4s %-9s %10s %7s" % ("", "label", "beta", "R2"))
    for code, label, b, r2, src, ident in betas:
        print("  %-3s %-9s %9.3f %7.2f" % (code, label, b, r2))

    with open(os.path.join(OUT_DIR, "rate_oil_betas.csv"), "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["code", "label", "beta_bps_per_pct_brent", "r_squared",
                    "source", "identifier"])
        for code, label, b, r2, src, ident in betas:
            w.writerow([code, label, "%.4f" % b, "%.4f" % r2, src, ident])
    print("\nwrote rate_oil_weekly.csv and rate_oil_betas.csv")


if __name__ == "__main__":
    main()

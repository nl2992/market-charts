# Market charts

Five cross-asset and US macro charts, built from public data as native Excel
chart objects. Every chart points at a data sheet in the same workbook, so any
number on a chart can be traced back to the series it came from and the
identifier it was pulled with. Nothing is pasted in as a picture.

**[`Market_Charts.xlsx`](Market_Charts.xlsx)** — open the `Charts` sheet.

| # | Chart | What it shows |
|---|---|---|
| 1 | Equities, bonds and commodities | MSCI ACWI, UST 7–10y and the S&P GSCI rebased to 100 at 2 January 2023 |
| 2 | Initial claims by week of year | US initial jobless claims, 2023–2026, against the 2023–25 average |
| 3 | Continuing claims and the unemployment rate | This cycle against the same months a cycle earlier |
| 4 | The 10y through equity bubbles | US 10y Treasury measured from the start of six equity episodes |
| 5 | Construction spending | US private non-residential: data centres against everything else |

## Sources

All free and unauthenticated.

| Source | Series |
|---|---|
| Yahoo Finance | `ACWI`, `IEF`, `^SPGSCI` daily closes |
| FRED | `ICSA`, `CCSA`, `UNRATE`, `DGS10`, `WPUSI012011` |
| US Census Bureau | C30 value of construction put in place, private, seasonally adjusted |

The data centre construction series is **not carried on FRED** — the Census
release has 144 series there and none of them is it. It comes from the Census
C30 file directly, where data centres are a line under Office.

## Rebuilding

```bash
python3 pull_cross_asset.py      # -> data/cross_asset_daily.csv
python3 pull_duration_case.py    # -> data/claims_*.csv, bubbles.csv, construction.csv
python3 build_public_charts.py   # -> Market_Charts.xlsx
```

Requires `openpyxl`. The cached CSVs are committed, so the workbook rebuilds
without a network round trip.

## Notes on the data

- **Rebasing** divides each series by its own first close and multiplies by 100,
  so every line starts together and the chart reads as relative performance.
  The workbook does this in live formulas, not in Python — move the base row and
  it recalculates.
- **The cross-asset spine is every weekday**, with the previous close carried
  through a market holiday. A union-of-calendars spine would privilege whichever
  market happened to be open.
- **Chart 1 mixes bases.** ACWI and IEF are total return (they use the adjusted
  close, so distributions are reinvested); the S&P GSCI is a spot index carrying
  neither roll nor collateral return. The commodity line understates what the
  asset class actually paid.
- **Chart 4 measures each episode from its own start date.** Yields rose in four
  of the six and fell in two — the two they fell through ended in recession. The
  current episode runs from 26 June 2025 and is not complete.
- **Chart 5 is nominal**, which is the conventional way to draw it. Columns E–G
  of the `Construction` sheet repeat the same series in January 2024 prices;
  deflated, the divergence is wider, not narrower.
- **The dashed line on chart 2** is the 2023–25 average across all weeks,
  computed live on the `Claims` sheet. It is a reference to read the current year
  against, not a threshold with any meaning of its own.

## Licence

Code is MIT. The underlying data belongs to its publishers (BLS, the Federal
Reserve Board, the US Census Bureau and Yahoo Finance) and is subject to their
terms.

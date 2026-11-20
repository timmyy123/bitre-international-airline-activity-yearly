# International Airline Activity (Australia): yearly aggregation

BITRE publishes *International Airline Activity* as one Excel file per month, and each file's Tables 1–5 cover only that month. This repo adds them up by **calendar year**, for July 2017 to July 2026 (109 monthly files). It also builds yearly versions of BITRE's `High_YTD` and `High_Month` highlight sheets.

## What's in the repo

| Path | What it is |
|---|---|
| `output/International_airline_activity_yearly_ALL.xlsx` | All years in one workbook. Sheets: **README** (method, checks, source-data problems), **Highlights** (charts, top city pairs / countries / airports / airlines, check against BITRE's published annual figures), **Monthly_series** (all 109 months), **Summary**, and Tables 1–5 with a `Year` column (filterable) |
| `output/by_year/International_airline_activity_YYYY.xlsx` | One workbook per year with the same layout as BITRE's monthly files: **Notes**, **High_Year** (like High_YTD), **High_Monthly** (like High_Month), **Table_1 … Table_5** |
| `source_monthly/` | The 109 BITRE monthly files, unchanged, plus `manifest.txt` (month → download link) |
| `build_yearly.py` | Reads the monthly files, checks them and writes everything in `output/` |
| `highlights.py` | Builds the highlight sheets and charts |
| `download_sources.py` | Downloads any monthly files listed in the manifest that are missing locally |

The workbooks open in Excel 2016 or later (including Excel 2021). They contain plain values and standard charts, with no formulas.

## Passengers by year

| Year | Months | Passengers |
|---|---|---:|
| 2017 | Jul–Dec | 20,514,476 |
| 2018 | 12 | 41,578,174 |
| 2019 | 12 | 42,504,908 |
| 2020 | 12 | 9,296,335 |
| 2021 | 12 | 1,554,763 |
| 2022 | 12 | 18,927,259 |
| 2023 | 12 | 35,638,191 |
| 2024 | 12 | 41,329,285 |
| 2025 | 12 | 44,841,900 |
| 2026 | Jan–Jul | 26,143,441 |

## Method

- **Detail rows only:** operator × country, airport × measure, and foreign port × Australian port are summed over each year's months.
- **Subtotals and totals rebuilt:** `ALL SERVICES` subtotals and `TOTAL` rows are recalculated from the yearly sums. Monthly subtotals aren't added up, because an operator can fly one route in some months and several in others.
- **Ratios recalculated:** seat utilisation, market share and % change are worked out from the yearly totals, not averaged across months.
- **Symbols:** `..` means not available (for example passengers on freight-only airlines). `-` means nil.
- **Comparisons with the previous year use the same months.** This matters for the part-years 2017 (Jul–Dec) and 2026 (Jan–Jul).
- **Figures as first published:** each month is taken as BITRE first published it. BITRE later revises earlier months, so its own year-ended totals differ slightly. The Highlights sheet compares the two: 2025 passengers differ by 0.001%, and earlier years by up to about 1%.

## Checks

- **TOTAL rows:** for every month and table, the detail rows must add up to BITRE's published `TOTAL` row. All 109 months reconcile.
- **Table 4 vs Table 1:** every month, Table 4's total passengers must equal Table 1's.
- **Same total in every view:** each year's passengers come out the same whether added up by city pair, country, airport or airline.
- **Unreadable cells:** the build stops if any cell can't be read as a number.

## Problems in the source data (handled in the build)

The full list, with months and cell counts, is on the README sheet of the combined workbook. In brief:

- **December 2024:** the BITRE page links to a different publication. The real Tables file was downloaded from BITRE's usual address for that month.
- **December 2022:** numbers are stored as text with spaces ("1 223 793"). They're converted to numbers.
- **December 2022 and January 2023:** ditto marks (`"`) are used instead of blank cells. They're read as "same as above".
- **October 2025, Table 5:** the Abu Dhabi row sits above its "Sydney" label. Each block's port is taken from its closing `Total` row instead.
- **Table 4 headings, several months:** the column years are wrong. Years are taken from the file's month instead; the data itself is correct.
- **Table 3, a few months:** leftover working columns are ignored.
- **Table 5:** every city pair is listed twice (by port, then a national block). The national block is left out of the sums and rebuilt.
- **Broome, June–September 2024:** missing from the monthly tables. BITRE added it later as revisions, which only affects Broome's 2024 figure.

## Rebuild

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python download_sources.py   # only needed if source_monthly/ is missing files
python build_yearly.py
```

To add a new month, append `YYYY-MM <link>` to `source_monthly/manifest.txt`, then run both scripts.

## Source and attribution

Data source: Bureau of Infrastructure and Transport Research Economics (BITRE), *International Airline Activity — monthly publications*, https://www.bitre.gov.au/resource/aviation/international-airline-activity-monthly-publications. © Commonwealth of Australia.

The files in `source_monthly/` are BITRE's originals, unchanged. Reuse is subject to the terms at https://www.infrastructure.gov.au/copyright. The yearly figures in `output/` are derived from them by the scripts in this repo; they are not official BITRE statistics.

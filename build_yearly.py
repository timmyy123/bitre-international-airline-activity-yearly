"""
Aggregate BITRE "International Airline Activity" monthly tables (Tables 1-5) into calendar years.

Input : source_monthly/YYYY-MM.xls(x)   (one BITRE monthly publication per file)
Output: output/International_airline_activity_yearly_ALL.xlsx   (tidy, one sheet per table, Year column)
        output/by_year/International_airline_activity_YYYY.xlsx (same layout as the monthly files, one per year)

Method
- Only base rows are read (operator x country, airport x measure, foreign port x Australian port).
  "ALL SERVICES" / "TOTAL" / "Total" rows are dropped and rebuilt from the yearly sums, because an
  operator can have one route in some months and several in others.
- "-" = zero, ".." = not available/not applicable (kept as not available unless some month has a value).
- Ratios (seat utilisation, market share, % change) are recomputed from the yearly sums.
- Prior-year columns (Tables 2, 4, 5) are the sum of the prior-year same-month figures published in
  each monthly file, so a partial year is compared with the same months a year earlier.
"""
import glob, os, re, sys, warnings
from collections import defaultdict

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import highlights as hl

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "source_monthly")
OUT = os.path.join(HERE, "output")

FOOTNOTE = re.compile(r"\s*\([a-z]\)\s*$")
WARN = []
UNPARSED = []
# Source-data issues found while reading, reported in the README sheet
CTX = {"ym": None, "t": None}
TEXTNUM = defaultdict(int)     # (ym, table) -> count of numbers stored as text
DITTO = defaultdict(int)       # (ym, table) -> count of ditto marks used instead of blank cells
HDR_YEARS = defaultdict(list)  # table -> [(ym, printed years)]
NO_LABEL = []                  # (ym, foreign port, assigned port)
STRAY_T3 = []                  # months where Table 3 has extra helper columns
SMALL_DIFFS = []               # (ym, table, column, rows sum, published total) within rounding tolerance


def clean_name(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    s = re.sub(r"\s+", " ", str(v)).strip()
    s = FOOTNOTE.sub("", s).strip()
    if s in ('"', "''", "〃", "”", "“"):  # ditto mark = same as above
        DITTO[(CTX["ym"], CTX["t"])] += 1
        return None
    return s or None


def num(v):
    """'-' -> 0, '..' / blank -> NaN, numbers -> float."""
    if v is None:
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = re.sub(r"[\s   ,]", "", str(v))  # some months store numbers as text, e.g. "1 223 793"
    if s in ("-", "–", "—"):
        return 0.0
    if s in ("..", "", "n.a.", "na", "np"):
        return np.nan
    try:
        f = float(s)
        TEXTNUM[(CTX["ym"], CTX["t"])] += 1
        return f
    except ValueError:
        UNPARSED.append(str(v))
        return np.nan


def is_note(s):
    s = (s or "").lower()
    return s.startswith(("please", "note", "source", "(")) or len(s) > 60


def data_rows(df, start):
    """Yield (row_index, row) from `start` until the notes section."""
    for i in range(start, len(df)):
        yield i, df.iloc[i]


def check_header(df, row, expected, ym, t):
    got = [clean_name(v) for v in df.iloc[row, : len(expected)].tolist()]
    for g, e in zip(got, expected):
        if e is not None and (g or "").lower() != e.lower():
            raise SystemExit(f"{ym} Table {t}: unexpected header {got} (expected {expected})")


# ---------------------------------------------------------------- parsers
def parse_t1(df, ym):
    check_header(df, 2, ["Scheduled Operator", None, "Passengers", "Freight", "Mail", None, "Passengers", "Freight", "Mail"], ym, 1)
    rows, total, op = [], None, None
    for i, r in data_rows(df, 5):
        a, b = clean_name(r.iloc[0]), clean_name(r.iloc[1])
        if a and a.upper() == "TOTAL":
            total = [num(r.iloc[c]) for c in (2, 3, 4, 6, 7, 8)]
            break
        if a:
            op = a
        if not b or b.upper() == "ALL SERVICES":
            continue
        rows.append(dict(Operator=op, Country=b, In_Passengers=num(r.iloc[2]), In_Freight_t=num(r.iloc[3]), In_Mail_t=num(r.iloc[4]),
                         Out_Passengers=num(r.iloc[6]), Out_Freight_t=num(r.iloc[7]), Out_Mail_t=num(r.iloc[8])))
    return rows, total


def parse_t2(df, ym):
    check_header(df, 3, ["Scheduled Operators", "Country to/from"], ym, 2)
    y = int(ym[:4])
    hy = df.iloc[3].tolist()
    if (hy[2], hy[3]) != (y - 1, y):
        WARN.append(f"{ym} Table 2 header years {hy[2]}/{hy[3]} (expected {y-1}/{y}); used file month")
        HDR_YEARS[2].append((ym, f"{hy[2]}/{hy[3]}"))
    rows, total, op = [], None, None
    for i, r in data_rows(df, 5):
        a, b = clean_name(r.iloc[0]), clean_name(r.iloc[1])
        if a and a.upper() == "TOTAL":
            total = [num(r.iloc[c]) for c in (2, 3, 7, 8, 12, 13)]
            break
        if a:
            op = a
        if not b or b.upper() == "ALL SERVICES":
            continue
        rows.append(dict(Operator=op, Country=b, Pax_prev=num(r.iloc[2]), Pax=num(r.iloc[3]), Freight_prev=num(r.iloc[7]),
                         Freight=num(r.iloc[8]), Mail_prev=num(r.iloc[12]), Mail=num(r.iloc[13])))
    return rows, total


def parse_t3(df, ym):
    check_header(df, 2, ["Scheduled Operator", None, "No. of", "Pax", "Seats", "Seat", None, "No. of", "Pax", "Seats", "Seat"], ym, 3)
    rows, total, op = [], None, None
    for i, r in data_rows(df, 5):
        a, b = clean_name(r.iloc[0]), clean_name(r.iloc[1])
        if a and a.upper() == "TOTAL":
            total = [num(r.iloc[c]) for c in (2, 3, 4, 7, 8, 9)]
            break
        if a:
            op = a
        if not b or b.upper() == "ALL SERVICES":
            continue
        rows.append(dict(Operator=op, Service=b, In_Flights=num(r.iloc[2]), In_Pax=num(r.iloc[3]), In_Seats=num(r.iloc[4]),
                         Out_Flights=num(r.iloc[7]), Out_Pax=num(r.iloc[8]), Out_Seats=num(r.iloc[9])))
    return rows, total


def parse_t4(df, ym):
    y = int(ym[:4])
    hy = df.iloc[2].tolist()
    if (hy[2], hy[3]) != (y - 1, y):
        WARN.append(f"{ym} Table 4 header years {hy[2]}/{hy[3]} (expected {y-1}/{y}); used file month")
        HDR_YEARS[4].append((ym, f"{hy[2]}/{hy[3]}"))
    rows, total, ap = [], [], None
    for i, r in data_rows(df, 4):
        a, b = clean_name(r.iloc[0]), clean_name(r.iloc[1])
        if a and is_note(a):
            break
        if a:
            ap = a
        if not b:
            continue
        rec = dict(Airport=ap, Measure=b, In_prev=num(r.iloc[2]), In=num(r.iloc[3]), Out_prev=num(r.iloc[6]), Out=num(r.iloc[7]))
        (total if ap.upper() == "TOTAL" else rows).append(rec)
    return rows, total


def parse_t5(df, ym):
    check_header(df, 3, [None, None, "Inbound", "Outbound", "Total"], ym, 5)
    # Each Australian-port block ends with a "Total | <port>" row. The port is taken from that closing row,
    # because the opening label is occasionally misplaced (e.g. Oct 2025 Abu Dhabi-Sydney sits above the label).
    rows, total, block = [], None, []
    for i, r in data_rows(df, 5):
        a, b = clean_name(r.iloc[0]), clean_name(r.iloc[1])
        if a and is_note(a):
            break
        if a and a.lower() == "total":
            if b and b.lower() == "australia":
                total = [num(r.iloc[c]) for c in (2, 3, 6, 7, 10, 11, 14, 15)]
                block = []  # national summary by foreign port; rebuilt later from the port blocks
                continue
            for rec in block:
                if rec["_label"] is None:
                    WARN.append(f"{ym} Table 5: '{rec['Foreign_Port']}' had no Australian-port label; assigned to '{b}' (its block's Total row)")
                    NO_LABEL.append((ym, rec["Foreign_Port"], b))
                elif rec["_label"] != b:
                    WARN.append(f"{ym} Table 5: '{rec['Foreign_Port']}' labelled '{rec['_label']}' but sits in the '{b}' block; used '{b}'")
                rec.pop("_label"); rec["Australian_Port"] = b
            rows.extend(block); block = []
            continue
        if not a:
            continue
        lbl = b if b else (block[-1]["_label"] if block else None)
        block.append(dict(_label=lbl, Australian_Port=None, Foreign_Port=a, Pax_In_prev=num(r.iloc[2]), Pax_Out_prev=num(r.iloc[3]),
                          Pax_In=num(r.iloc[6]), Pax_Out=num(r.iloc[7]), Frt_In_prev=num(r.iloc[10]), Frt_Out_prev=num(r.iloc[11]),
                          Frt_In=num(r.iloc[14]), Frt_Out=num(r.iloc[15])))
    if block:
        WARN.append(f"{ym} Table 5: {len(block)} rows after the last Total row were ignored")
    return rows, total


PARSERS = {1: parse_t1, 2: parse_t2, 3: parse_t3, 4: parse_t4, 5: parse_t5}
KEYS = {1: ["Operator", "Country"], 2: ["Operator", "Country"], 3: ["Operator", "Service"],
        4: ["Airport", "Measure"], 5: ["Australian_Port", "Foreign_Port"]}


# ---------------------------------------------------------------- load
def load():
    files = sorted(f for f in glob.glob(os.path.join(SRC, "20??-??.xls*")))
    data = {t: [] for t in PARSERS}
    pub_totals = {t: defaultdict(lambda: None) for t in PARSERS}
    months = defaultdict(list)
    for f in files:
        ym = os.path.basename(f)[:7]
        x = pd.ExcelFile(f)
        months[int(ym[:4])].append(int(ym[5:]))
        for t, p in PARSERS.items():
            CTX["ym"], CTX["t"] = ym, t
            df = pd.read_excel(x, f"Table_{t}", header=None, dtype=object)
            if t == 3 and df.shape[1] > 11 and df.iloc[5:, 11:].notna().any().any():
                STRAY_T3.append(ym)
            title = " ".join(str(v) for v in df.iloc[0].tolist() if isinstance(v, str))
            if f"TABLE {t}" not in title.upper():
                raise SystemExit(f"{ym}: Table_{t} title not found")
            rows, total = p(df, ym)
            for r in rows:
                r["Year"], r["Month"] = int(ym[:4]), int(ym[5:])
            data[t].extend(rows)
            pub_totals[t][ym] = total
    return {t: pd.DataFrame(v) for t, v in data.items()}, pub_totals, months


def checks(data, pub, months):
    """Each month: base rows must add up to the published TOTAL row."""
    problems = []
    cols = {1: ["In_Passengers", "In_Freight_t", "In_Mail_t", "Out_Passengers", "Out_Freight_t", "Out_Mail_t"],
            2: ["Pax_prev", "Pax", "Freight_prev", "Freight", "Mail_prev", "Mail"],
            3: ["In_Flights", "In_Pax", "In_Seats", "Out_Flights", "Out_Pax", "Out_Seats"],
            5: ["Pax_In_prev", "Pax_Out_prev", "Pax_In", "Pax_Out", "Frt_In_prev", "Frt_Out_prev", "Frt_In", "Frt_Out"]}
    for t, cs in cols.items():
        g = data[t].groupby(["Year", "Month"])[cs].sum(min_count=1)
        for (y, m), row in g.iterrows():
            ym = f"{y}-{m:02d}"
            tot = pub[t][ym]
            if tot is None:
                problems.append(f"{ym} T{t}: no TOTAL row"); continue
            for c, pv in zip(cs, tot):
                sv = row[c]
                if np.isnan(pv) and (np.isnan(sv) or sv == 0):
                    continue
                d = abs((0 if np.isnan(sv) else sv) - pv)
                if d > max(1.0, abs(pv) * 1e-4):
                    problems.append(f"{ym} T{t} {c}: rows sum {sv:,.3f} vs published TOTAL {pv:,.3f}")
                elif d > 0.05:
                    SMALL_DIFFS.append((ym, t, c, sv, pv))
    # Table 4: airports must add up to TOTAL, and TOTAL passengers must match Table 1
    t4 = data[4]
    for ym, tot in pub[4].items():
        y, m = int(ym[:4]), int(ym[5:])
        sub = t4[(t4.Year == y) & (t4.Month == m)]
        for rec in tot:
            s = sub[sub.Measure == rec["Measure"]][["In", "Out"]].sum()
            for c in ("In", "Out"):
                if abs(s[c] - rec[c]) > max(1.0, abs(rec[c]) * 1e-4):
                    problems.append(f"{ym} T4 {rec['Measure']} {c}: airports {s[c]:,.0f} vs TOTAL {rec[c]:,.0f}")
            if rec["Measure"].lower() == "passengers":
                t1 = pub[1][ym]
                if abs(rec["In"] - t1[0]) > 1 or abs(rec["Out"] - t1[3]) > 1:
                    problems.append(f"{ym} T4 TOTAL passengers {rec['In']:,.0f}/{rec['Out']:,.0f} != T1 {t1[0]:,.0f}/{t1[3]:,.0f}")
    return problems


# ---------------------------------------------------------------- aggregate
def pct(a, b):
    with np.errstate(divide="ignore", invalid="ignore"):
        return np.where((b > 0) & ~np.isnan(a), a / b * 100, np.nan)


def chg(cur, prev):
    with np.errstate(divide="ignore", invalid="ignore"):
        return np.where((prev > 0) & ~np.isnan(cur), (cur - prev) / prev * 100, np.nan)


def yearly(df, keys):
    vals = [c for c in df.columns if c not in keys + ["Year", "Month"]]
    return df.groupby(["Year"] + keys, sort=True)[vals].sum(min_count=1).reset_index()


def with_subtotals(y, keys, vals, sub_label="ALL SERVICES", total_label="TOTAL"):
    """Insert operator subtotal rows (only for operators with 2+ rows) and a grand TOTAL, per year."""
    out = []
    for yr, g in y.groupby("Year", sort=True):
        g = g.sort_values(keys, key=lambda s: s.str.lower())
        for k0, gg in g.groupby(keys[0], sort=False):
            out.append(gg.assign(_kind="row"))
            if len(gg) > 1:
                s = gg[vals].sum(min_count=1).to_frame().T
                s["Year"], s[keys[0]], s[keys[1]], s["_kind"] = yr, k0, sub_label, "sub"
                out.append(s)
        t = g[vals].sum(min_count=1).to_frame().T
        t["Year"], t[keys[0]], t[keys[1]], t["_kind"] = yr, total_label, None, "total"
        out.append(t)
    r = pd.concat(out, ignore_index=True)
    r[vals] = r[vals].astype(float)
    return r


def build(data, months):
    res = {}
    # T1
    k = KEYS[1]; y = yearly(data[1], k); v = [c for c in y.columns if c not in k + ["Year"]]
    t1 = with_subtotals(y, k, v)
    t1["Total_Passengers"] = t1[["In_Passengers", "Out_Passengers"]].sum(axis=1, min_count=1)
    t1["Total_Freight_t"] = t1[["In_Freight_t", "Out_Freight_t"]].sum(axis=1, min_count=1)
    t1["Total_Mail_t"] = t1[["In_Mail_t", "Out_Mail_t"]].sum(axis=1, min_count=1)
    res[1] = t1
    # T2
    k = KEYS[2]; y = yearly(data[2], k); v = [c for c in y.columns if c not in k + ["Year"]]
    t2 = with_subtotals(y, k, v)
    for m in ("Pax", "Freight", "Mail"):
        tot = t2[t2._kind == "total"].set_index("Year")[m]
        t2[f"{m}_share_pct"] = pct(t2[m].values, t2.Year.map(tot).values)
        t2[f"{m}_change_pct"] = chg(t2[m].values, t2[f"{m}_prev"].values)
    res[2] = t2
    # T3
    k = KEYS[3]; y = yearly(data[3], k); v = [c for c in y.columns if c not in k + ["Year"]]
    t3 = with_subtotals(y, k, v)
    t3["In_Seat_Util_pct"] = pct(t3.In_Pax.values, t3.In_Seats.values)
    t3["Out_Seat_Util_pct"] = pct(t3.Out_Pax.values, t3.Out_Seats.values)
    res[3] = t3
    # T4 (airports incl. TOTAL computed from airports)
    d4 = data[4]
    y = d4.groupby(["Year", "Airport", "Measure"], sort=False)[["In_prev", "In", "Out_prev", "Out"]].sum(min_count=1).reset_index()
    tot = d4.groupby(["Year", "Measure"], sort=False)[["In_prev", "In", "Out_prev", "Out"]].sum(min_count=1).reset_index()
    tot["Airport"] = "TOTAL"
    order = {m: i for i, m in enumerate(d4.Measure.drop_duplicates())}
    t4 = pd.concat([y.sort_values(["Year", "Airport"]), tot], ignore_index=True)
    t4["_mo"] = t4.Measure.map(order); t4["_tot"] = (t4.Airport == "TOTAL").astype(int)
    t4 = t4.sort_values(["Year", "_tot", "Airport", "_mo"]).drop(columns=["_mo", "_tot"]).reset_index(drop=True)
    t4["Total_prev"] = t4[["In_prev", "Out_prev"]].sum(axis=1, min_count=1)
    t4["Total"] = t4[["In", "Out"]].sum(axis=1, min_count=1)
    for c in ("In", "Out", "Total"):
        t4[f"{c}_change_pct"] = chg(t4[c].values, t4[f"{c}_prev"].values)
    grand = t4[t4.Airport == "TOTAL"].set_index(["Year", "Measure"]).Total
    t4["Share_of_total_pct"] = pct(t4.Total.values, pd.Series(list(zip(t4.Year, t4.Measure))).map(grand).values)
    res[4] = t4
    # T5 (Australian-port subtotals labelled "Total", grand "Australia")
    k = KEYS[5]; y = yearly(data[5], k); v = [c for c in y.columns if c not in k + ["Year"]]
    out = []
    for yr, g in y.groupby("Year"):
        for au, gg in g.groupby("Australian_Port", sort=True):
            out.append(gg.sort_values("Foreign_Port").assign(_kind="row"))
            s = gg[v].sum(min_count=1).to_frame().T
            s["Year"], s["Australian_Port"], s["Foreign_Port"], s["_kind"] = yr, au, "Total", "sub"
            out.append(s)
        nat = g.groupby("Foreign_Port", sort=True)[v].sum(min_count=1).reset_index()
        nat["Year"], nat["Australian_Port"], nat["_kind"] = yr, "Australia", "aus"
        out.append(nat)
        s = g[v].sum(min_count=1).to_frame().T
        s["Year"], s["Australian_Port"], s["Foreign_Port"], s["_kind"] = yr, "Australia", "Total", "total"
        out.append(s)
    t5 = pd.concat(out, ignore_index=True); t5[v] = t5[v].astype(float)
    for p in ("Pax", "Frt"):
        for s in ("_prev", ""):
            t5[f"{p}_Total{s}"] = t5[[f"{p}_In{s}", f"{p}_Out{s}"]].sum(axis=1, min_count=1)
    res[5] = t5
    for t in res:
        res[t]["Months_included"] = res[t].Year.map(lambda yr: len(months[yr]))
    return res


# ---------------------------------------------------------------- write
MON = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
HDR_FILL = PatternFill("solid", fgColor="DCE6F1")
SUB_FILL = PatternFill("solid", fgColor="F2F2F2")
TOT_FILL = PatternFill("solid", fgColor="FFF2CC")
BOLD = Font(bold=True)
NUM0 = '#,##0;-#,##0;"-"'
NUM1 = '#,##0.0;-#,##0.0;"-"'
NUM3 = '#,##0.000;-#,##0.000;"-"'
THIN = Border(bottom=Side(style="thin", color="7F7F7F"))


def period(yr, months):
    ms = sorted(months[yr])
    return f"{yr}" if len(ms) == 12 else f"{yr} ({MON[ms[0]-1]}–{MON[ms[-1]-1]}, {len(ms)} months)"


def put(ws, r, c, v, fmt=None):
    if isinstance(v, (float, np.floating)) and np.isnan(v):
        v = ".."
    elif isinstance(v, (np.floating, np.integer)):
        v = float(v)
    cell = ws.cell(r, c, v)
    if fmt and not isinstance(v, str):
        cell.number_format = fmt
    if isinstance(v, str) and v == "..":
        cell.alignment = Alignment(horizontal="right")
    return cell


def headers(ws, r0, rows):
    for i, row in enumerate(rows):
        for j, h in enumerate(row, 1):
            if h is not None:
                c = ws.cell(r0 + i, j, h); c.font = BOLD; c.fill = HDR_FILL
                c.alignment = Alignment(horizontal="center", wrap_text=True)
    for j in range(1, max(len(x) for x in rows) + 1):
        ws.cell(r0 + len(rows) - 1, j).border = THIN
        for i in range(len(rows)):
            ws.cell(r0 + i, j).fill = HDR_FILL
    return r0 + len(rows)


def style_row(ws, r, ncol, kind):
    if kind in ("sub", "total"):
        for j in range(1, ncol + 1):
            ws.cell(r, j).fill = SUB_FILL if kind == "sub" else TOT_FILL
            ws.cell(r, j).font = BOLD


def write_layout(ws, t, df, yr, months):
    """Write one table for one year in the same layout as the BITRE monthly file."""
    p = period(yr, months); py = period(yr - 1, months) if yr - 1 in months else f"{yr-1}"
    prev_lbl = f"{yr-1} (same months)"
    widths = {}
    if t == 1:
        ws.cell(1, 1, f"TABLE 1   SCHEDULED INTERNATIONAL AIR TRAFFIC TO AND FROM AUSTRALIA: {p}").font = Font(bold=True, size=12)
        r = headers(ws, 3, [[None, None, "Inbound", None, None, None, "Outbound", None, None, None, "Total", None, None],
                            ["Scheduled Operator", "Country to/from", "Passengers", "Freight (tonnes)", "Mail (tonnes)", None,
                             "Passengers", "Freight (tonnes)", "Mail (tonnes)", None, "Passengers", "Freight (tonnes)", "Mail (tonnes)"]])
        prev = None
        for _, x in df.iterrows():
            put(ws, r, 1, x.Operator if (x.Operator != prev or x._kind != "row") else None); prev = x.Operator
            put(ws, r, 2, x.Country if x.Country else None)
            for c, col, f in [(3, "In_Passengers", NUM0), (4, "In_Freight_t", NUM3), (5, "In_Mail_t", NUM3), (7, "Out_Passengers", NUM0),
                              (8, "Out_Freight_t", NUM3), (9, "Out_Mail_t", NUM3), (11, "Total_Passengers", NUM0),
                              (12, "Total_Freight_t", NUM3), (13, "Total_Mail_t", NUM3)]:
                put(ws, r, c, x[col], f)
            style_row(ws, r, 13, x._kind); r += 1
        widths = {1: 34, 2: 24}
    elif t == 2:
        ws.cell(1, 1, f"TABLE 2   SCHEDULED OPERATOR MARKET SHARES AND GROWTH: {p}").font = Font(bold=True, size=12)
        r = headers(ws, 3, [[None, None, "Total Passengers", None, None, None, None, "Total Freight (tonnes)", None, None, None, None, "Total Mail (tonnes)", None, None, None],
                            ["Scheduled Operators", "Country to/from", prev_lbl, str(yr), "% of TOTAL", "% Change", None,
                             prev_lbl, str(yr), "% of TOTAL", "% Change", None, prev_lbl, str(yr), "% of TOTAL", "% Change"]])
        prev = None
        for _, x in df.iterrows():
            put(ws, r, 1, x.Operator if (x.Operator != prev or x._kind != "row") else None); prev = x.Operator
            put(ws, r, 2, x.Country if x.Country else None)
            for c0, m, f in [(3, "Pax", NUM0), (8, "Freight", NUM3), (13, "Mail", NUM3)]:
                put(ws, r, c0, x[f"{m}_prev"], f); put(ws, r, c0 + 1, x[m], f)
                put(ws, r, c0 + 2, x[f"{m}_share_pct"], "0.00"); put(ws, r, c0 + 3, x[f"{m}_change_pct"], "0.0")
            style_row(ws, r, 16, x._kind); r += 1
        widths = {1: 34, 2: 24}
    elif t == 3:
        ws.cell(1, 1, f"TABLE 3   AIRLINE PASSENGER CAPACITY AND UTILISATION TO AND FROM AUSTRALIA BY OPERATOR: {p}").font = Font(bold=True, size=12)
        r = headers(ws, 3, [[None, None, "Inbound", None, None, None, None, "Outbound", None, None, None],
                            ["Scheduled Operator", "Service to/from", "No. of Flights", "Pax Carried", "Seats Available", "Seat Utilisation %", None,
                             "No. of Flights", "Pax Carried", "Seats Available", "Seat Utilisation %"]])
        prev = None
        for _, x in df.iterrows():
            put(ws, r, 1, x.Operator if (x.Operator != prev or x._kind != "row") else None); prev = x.Operator
            put(ws, r, 2, x.Service if x.Service else None)
            for c, col, f in [(3, "In_Flights", NUM0), (4, "In_Pax", NUM0), (5, "In_Seats", NUM0), (6, "In_Seat_Util_pct", NUM1),
                              (8, "Out_Flights", NUM0), (9, "Out_Pax", NUM0), (10, "Out_Seats", NUM0), (11, "Out_Seat_Util_pct", NUM1)]:
                put(ws, r, c, x[col], f)
            style_row(ws, r, 11, x._kind); r += 1
        widths = {1: 34, 2: 24}
    elif t == 4:
        ws.cell(1, 1, f"TABLE 4   SCHEDULED INTERNATIONAL AIRPORT TRAFFIC AND AIRCRAFT MOVEMENTS: {p}").font = Font(bold=True, size=12)
        r = headers(ws, 3, [[None, None, "Inbound", None, None, None, "Outbound", None, None, None, "Total", None, None, None],
                            ["Airport", "Measure", prev_lbl, str(yr), "% Change", None, prev_lbl, str(yr), "% Change", None,
                             prev_lbl, str(yr), "% Change", "% of TOTAL"]])
        prev = None
        for _, x in df.iterrows():
            if prev is not None and x.Airport != prev:
                r += 1
            put(ws, r, 1, x.Airport if x.Airport != prev else None); prev = x.Airport
            put(ws, r, 2, x.Measure)
            f = NUM3 if "freight" in x.Measure.lower() else NUM0
            for c0, m in [(3, "In"), (7, "Out"), (11, "Total")]:
                put(ws, r, c0, x[f"{m}_prev"], f); put(ws, r, c0 + 1, x[m], f); put(ws, r, c0 + 2, x[f"{m}_change_pct"], "0.0")
            put(ws, r, 14, x.Share_of_total_pct, "0.00")
            style_row(ws, r, 14, "total" if x.Airport == "TOTAL" else "row"); r += 1
        widths = {1: 22, 2: 22}
    elif t == 5:
        ws.cell(1, 1, f"TABLE 5   SCHEDULED INTERNATIONAL TRAFFIC BY CITY PAIRS: {p}").font = Font(bold=True, size=12)
        r = headers(ws, 3, [[None, None, "Passengers", None, None, None, None, None, None, None, "Freight (tonnes)", None, None, None, None, None, None],
                            ["Foreign", "Australian", prev_lbl, None, None, None, str(yr), None, None, None, prev_lbl, None, None, None, str(yr), None, None],
                            ["Port", "Port", "Inbound", "Outbound", "Total", None, "Inbound", "Outbound", "Total", None,
                             "Inbound", "Outbound", "Total", None, "Inbound", "Outbound", "Total"]])
        prev = None
        for _, x in df.iterrows():
            put(ws, r, 1, x.Foreign_Port)
            put(ws, r, 2, x.Australian_Port if (x.Australian_Port != prev or x._kind in ("sub", "total")) else None)
            prev = x.Australian_Port if x._kind in ("row", "aus") else None
            for c0, p_, f in [(3, "Pax", NUM0), (11, "Frt", NUM3)]:
                for off, s in [(0, "_prev"), (4, "")]:
                    put(ws, r, c0 + off, x[f"{p_}_In{s}"], f); put(ws, r, c0 + off + 1, x[f"{p_}_Out{s}"], f); put(ws, r, c0 + off + 2, x[f"{p_}_Total{s}"], f)
            style_row(ws, r, 17, x._kind)
            r += 1
            if x._kind == "sub":
                r += 1
        widths = {1: 22, 2: 22}
    for j in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(j, 13)
    ws.freeze_panes = ws.cell(6 if t == 5 else 5, 3)
    ws.row_dimensions[4].height = 30


TIDY_COLS = {
    1: ["Year", "Months_included", "Operator", "Country", "In_Passengers", "In_Freight_t", "In_Mail_t", "Out_Passengers", "Out_Freight_t", "Out_Mail_t",
        "Total_Passengers", "Total_Freight_t", "Total_Mail_t"],
    2: ["Year", "Months_included", "Operator", "Country", "Pax_prev", "Pax", "Pax_share_pct", "Pax_change_pct", "Freight_prev", "Freight",
        "Freight_share_pct", "Freight_change_pct", "Mail_prev", "Mail", "Mail_share_pct", "Mail_change_pct"],
    3: ["Year", "Months_included", "Operator", "Service", "In_Flights", "In_Pax", "In_Seats", "In_Seat_Util_pct", "Out_Flights", "Out_Pax", "Out_Seats", "Out_Seat_Util_pct"],
    4: ["Year", "Months_included", "Airport", "Measure", "In_prev", "In", "In_change_pct", "Out_prev", "Out", "Out_change_pct", "Total_prev", "Total",
        "Total_change_pct", "Share_of_total_pct"],
    5: ["Year", "Months_included", "Australian_Port", "Foreign_Port", "Pax_In_prev", "Pax_Out_prev", "Pax_Total_prev", "Pax_In", "Pax_Out", "Pax_Total",
        "Frt_In_prev", "Frt_Out_prev", "Frt_Total_prev", "Frt_In", "Frt_Out", "Frt_Total"],
}
TIDY_NAMES = {1: "T1_Traffic_by_operator", 2: "T2_Market_share", 3: "T3_Capacity_utilisation", 4: "T4_Airports", 5: "T5_City_pairs"}


def write_tidy(wb, t, df):
    ws = wb.create_sheet(TIDY_NAMES[t])
    cols = TIDY_COLS[t] + (["Row_type"] if "_kind" in df.columns else [])
    for j, c in enumerate(cols, 1):
        cell = ws.cell(1, j, c); cell.font = BOLD; cell.fill = HDR_FILL
    kindmap = {"row": "detail", "sub": "subtotal", "total": "total", "aus": "all Australian ports"}
    for i, (_, x) in enumerate(df.iterrows(), 2):
        for j, c in enumerate(cols, 1):
            if c == "Row_type":
                v = kindmap[x._kind]
            else:
                v = x[c]
            if isinstance(v, (float, np.floating)) and np.isnan(v):
                v = None
            elif isinstance(v, (np.floating, np.integer)):
                v = float(v) if not float(v).is_integer() or "pct" in c or "_t" in c or "Frt" in c or "Freight" in c or "Mail" in c else int(v)
            cell = ws.cell(i, j, v)
            if "pct" in c:
                cell.number_format = "0.00"
            elif isinstance(v, float):
                cell.number_format = "#,##0.000"
            elif isinstance(v, int) and c not in ("Year", "Months_included"):
                cell.number_format = "#,##0"
        if t == 4 and x.Airport == "TOTAL":
            for j in range(1, len(cols) + 1):
                ws.cell(i, j).font = BOLD
    for j, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = 28 if c in ("Operator", "Country", "Service", "Airport", "Measure", "Australian_Port", "Foreign_Port") else 15
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = ws.dimensions


def summary(res, months):
    t1 = res[1][res[1]._kind == "total"].set_index("Year")
    t3 = res[3][res[3]._kind == "total"].set_index("Year")
    t2 = res[2][res[2]._kind == "total"].set_index("Year")
    s = pd.DataFrame({
        "Months_included": [len(months[y]) for y in t1.index],
        "Passengers_in": t1.In_Passengers, "Passengers_out": t1.Out_Passengers, "Passengers_total": t1.Total_Passengers,
        "Passengers_same_months_prev_year": t2.Pax_prev, "Passenger_change_pct": t2.Pax_change_pct,
        "Freight_t_total": t1.Total_Freight_t, "Mail_t_total": t1.Total_Mail_t,
        "Flights_total": t3.In_Flights + t3.Out_Flights, "Seats_total": t3.In_Seats + t3.Out_Seats,
        "Seat_util_pct": pct((t3.In_Pax + t3.Out_Pax).values, (t3.In_Seats + t3.Out_Seats).values),
        "Operators_with_traffic": res[1][res[1]._kind == "row"].groupby("Year").Operator.nunique(),
    })
    s.index.name = "Year"
    return s.reset_index()


def fmt_months(yms):
    """['2024-01','2024-02','2025-01'] -> 'Jan–Feb 2024, Jan 2025'"""
    yms = sorted(set(yms)); out = []; i = 0
    while i < len(yms):
        j = i
        while j + 1 < len(yms):
            y0, m0 = int(yms[j][:4]), int(yms[j][5:]); y1, m1 = int(yms[j + 1][:4]), int(yms[j + 1][5:])
            if (y1 * 12 + m1) - (y0 * 12 + m0) == 1 and y1 == int(yms[i][:4]):
                j += 1
            else:
                break
        a, b = yms[i], yms[j]
        out.append(f"{MON[int(a[5:])-1]} {a[:4]}" if a == b else f"{MON[int(a[5:])-1]}–{MON[int(b[5:])-1]} {a[:4]}")
        i = j + 1
    return ", ".join(out)


def source_issues():
    """Problems found in the BITRE files, each with what was done about it."""
    iss = []

    def add(issue, yms, tables, handled, effect):
        iss.append(dict(issue=issue, yms=sorted(set(yms)), tables=tables, handled=handled, effect=effect,
                        years={int(x[:4]) for x in yms}))

    man = os.path.join(SRC, "manifest.txt")
    if os.path.exists(man) and "not linked on page" in open(man).read():
        add("Wrong file linked on the BITRE page: 'December 2024' points to the 'Monthly airline performance' "
            "spreadsheet (a 1985-onwards time series), not the monthly Tables 1–6 file.",
            ["2024-12"], "All",
            "Downloaded the Tables file from BITRE's usual address for that month "
            "(…/documents/International_airline_activity_1224.xlsx) and confirmed its title reads 'December 2024'.",
            "None — December 2024 is included normally.")
    tn = defaultdict(list)
    for (ym, t), n in TEXTNUM.items():
        tn[ym].append(f"T{t}: {n:,} cells")
    if tn:
        add("Numbers stored as text with spaces as thousands separators (e.g. '1 223 793'), including the TOTAL rows.",
            list(tn), {k: ", ".join(sorted(v)) for k, v in tn.items()},
            "Spaces removed and the text converted to numbers. The parser now stops the build if any cell can't be read as a number "
            f"(cells that could not be read: {len(UNPARSED)}).",
            "None after the fix. (Before it, most of that month was silently read as blank.)")
    dt = defaultdict(list)
    for (ym, t), n in DITTO.items():
        dt[ym].append(f"T{t}: {n:,} cells")
    if dt:
        add("Ditto marks ('\"') used in the name columns (operator, Australian port) instead of leaving the cell blank.",
            list(dt), {k: ", ".join(sorted(v)) for k, v in dt.items()},
            "A ditto mark is read as 'same as the row above'.",
            "None after the fix. (Before it, Table 5's national summary block was double-counted for January 2023.)")
    for ym, fp, au in NO_LABEL:
        add(f"Row out of place: the '{fp}' row sits above its Australian-port label, so it has no Australian port.",
            [ym], "T5",
            f"Each block's Australian port is taken from its closing 'Total | <port>' row, so '{fp}' goes to {au}.",
            f"None after the fix. (Before it, {fp}–{au} traffic was missing from the yearly table.)")
    for t, lst in sorted(HDR_YEARS.items()):
        add("Wrong years printed in the column headings (the previous publication's years were left in).",
            [ym for ym, _ in lst], f"T{t}",
            "Years are taken from the file's month (current year, and the year before for the comparison columns), not from the heading. "
            "Each month's Table 4 TOTAL passengers were checked against Table 1 and match, so only the headings are wrong.",
            "None — the data in those columns is correct.")
    if STRAY_T3:
        add("Extra working columns to the right of Table 3 (fill-in values and lookup keys left in the file).",
            STRAY_T3, "T3", "Ignored; only the published Inbound/Outbound columns are read.", "None.")
    if SMALL_DIFFS:
        worst = max(SMALL_DIFFS, key=lambda d: abs(d[3] - d[4]))
        add("Published TOTAL row differs slightly from the sum of its detail rows (rounding in the published TOTAL).",
            [d[0] for d in SMALL_DIFFS], ", ".join(sorted({f"T{d[1]}" for d in SMALL_DIFFS})),
            "Totals are rebuilt from the detail rows, so the yearly TOTAL is the sum of the rows shown.",
            f"Negligible: largest gap {abs(worst[3]-worst[4]):,.3f} {'tonnes' if any(k in worst[2] for k in ('Frt', 'Freight', 'Mail')) else ''} "
            f"({worst[2]}, {fmt_months([worst[0]])}).")
    add("Airport missing from early monthly tables: Broome's international services (from June 2024) first appear in the "
        "October 2024 tables. BITRE later added June–September 2024 in the prior-year columns of the 2025 files (391, 1,638, 1,327, 1,730 passengers).",
        ["2024-06", "2024-07", "2024-08", "2024-09"], "T4, T5",
        "Left as first published, so this report matches the 2024 monthly tables. The revised figures show up only as the "
        "'same months a year earlier' comparison for 2025.",
        "Broome 2024 = 1,923 passengers here vs 7,009 in BITRE's later year-ended figure; Broome's 2025 % change is overstated. "
        "Australia-wide totals are unaffected because the monthly TOTAL rows reconcile.")
    add("Table 5 lists every city pair twice: once under each Australian port, then again in a national 'Australia' "
        "block summed by foreign port. (This is how BITRE lays out the table, not an error, but adding all rows would double the totals.)",
        [], "T5", "The 'Australia' block is left out of the sums and rebuilt from the yearly Australian-port figures.",
        "None — the yearly table keeps the same layout.")
    add("Footnote markers on names change from month to month (e.g. 'Jin Air (a)' in one month, 'Jin Air (b)' in another).",
        [], "T1–T3", "Markers like '(a)' are removed so the same operator joins up across months.",
        "None. The footnote texts (e.g. 'Services recommenced December 2017') are not carried into the yearly files.")
    return iss


def write_issues(ws, r0, issues, year=None):
    hdr = ["Problem", "Months affected", "Tables", "How it was handled", "Effect on the yearly figures"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(r0, j, h); c.font = BOLD; c.fill = HDR_FILL; c.border = THIN
    for i, x in enumerate(issues, r0 + 1):
        yms = [m for m in x["yms"] if year is None or int(m[:4]) == year]
        tb = x["tables"]
        if isinstance(tb, dict):  # per-month detail, e.g. {'2022-12': 'T1: 178 cells, T5: 527 cells'}
            tb = "; ".join(f"{fmt_months([k])} — {v}" if len(yms) > 1 else v for k, v in sorted(tb.items()) if k in yms)
        vals = [x["issue"], fmt_months(yms) if yms else "All months", tb, x["handled"], x["effect"]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v); c.alignment = Alignment(wrap_text=True, vertical="top")
    for j, w in enumerate([55, 26, 22, 60, 45], 1):
        ws.column_dimensions[get_column_letter(j)].width = w


def main():
    data, pub, months = load()
    problems = checks(data, pub, months)
    res = build(data, months)
    os.makedirs(os.path.join(OUT, "by_year"), exist_ok=True)

    # one workbook per year, same layout as the monthly publication, plus highlight sheets
    cx = hl.Ctx(data, pub, months)
    issues = source_issues()
    for yr in sorted(months):
        wb = Workbook(); wb.remove(wb.active)
        wn = wb.create_sheet("Notes")
        lines = [
            (f"International Airline Activity — calendar year {period(yr, months)}", True),
            ("Source: BITRE monthly publications (bitre.gov.au). Built by build_yearly.py.", False),
            ("Detail rows are summed over the year; ALL SERVICES / TOTAL rows and all percentages are recalculated from the yearly sums.", False),
            ("Prior-year columns (Tables 2, 4, 5) cover the same months a year earlier. '..' = not available, '-' = nil.", False),
            ("Sheets: High_Year (like BITRE's High_YTD), High_Monthly (like High_Month), Table_1–Table_5.", False),
            ("", False),
        ]
        mine = [x for x in issues if yr in x["years"] or not x["years"]]
        lines.append((f"Problems in the source data affecting {yr}" if mine else f"No source-data problems affected {yr}.", True))
        for i, (txt, b) in enumerate(lines, 1):
            c = wn.cell(i, 1, txt); c.font = Font(bold=b, size=13 if i == 1 else 11)
        if mine:
            write_issues(wn, len(lines) + 1, mine, year=yr)
        hl.high_year(wb.create_sheet("High_Year"), cx, yr)
        hl.high_monthly(wb.create_sheet("High_Monthly"), cx, yr)
        for t in range(1, 6):
            write_layout(wb.create_sheet(f"Table_{t}"), t, res[t][res[t].Year == yr].reset_index(drop=True), yr, months)
        wb.save(os.path.join(OUT, "by_year", f"International_airline_activity_{yr}.xlsx"))

    # combined tidy workbook
    wb = Workbook(); ws = wb.active; ws.title = "README"
    notes = [
        ("International Airline Activity — Tables 1–5 aggregated by calendar year", True),
        ("Source: BITRE monthly publications (bitre.gov.au), one file per month, " + f"{sum(len(v) for v in months.values())} months.", False),
        ("", False), ("Months included per year", True),
    ] + [(f"  {period(y, months)}", False) for y in sorted(months)] + [
        ("", False), ("Method", True),
        ("  • Detail rows (operator × country, airport × measure, foreign × Australian port) are summed over the months of each year.", False),
        ("  • ALL SERVICES subtotals and TOTAL rows are rebuilt from the yearly sums (not summed from monthly subtotals).", False),
        ("  • Seat utilisation, market share and % change are recalculated from yearly totals, not averaged.", False),
        ("  • '..' = not available (e.g. passengers for freight-only carriers); '-' = nil.", False),
        ("  • Footnote markers such as '(a)' are removed from operator names so months join up.", False),
        ("  • Prior-year columns in Tables 2, 4, 5 = sum of the prior-year same-month figures printed in each monthly file.", False),
        ("    For a partial year (2017, 2026) these compare the same months a year earlier.", False),
        ("  • Operator names are used as published; a rename (e.g. a rebrand) appears as a separate operator.", False),
        ("", False), ("Checks", True),
        (f"  • Every month and table: detail rows were checked against BITRE's published TOTAL row. "
         f"{'All reconcile.' if not problems else f'{len(problems)} did NOT reconcile — see below.'}", False),
        ("  • Every month: Table 4 TOTAL passengers were checked against Table 1 TOTAL passengers.", False),
    ] + [(f"  • NOT RECONCILED: {p}", False) for p in problems] + [
        ("", False), ("Sheets", True),
        ("  Highlights — like BITRE's High_YTD, across all years: charts, Table I traffic, top city pairs / countries / airports / airlines,", False),
        ("               and a check against BITRE's own published year-ended-December figures", False),
        ("  Monthly_series — every month: passengers in/out, freight, mail, seats, flights, aircraft movements, seat utilisation", False),
        ("  Summary — headline totals by year", False),
    ] + [(f"  {TIDY_NAMES[t]} — Table {t}, all years, one row per item (Row_type = detail / subtotal / total)", False) for t in range(1, 6)] + [
        ("  Per-year workbooks in output/by_year/: Notes, High_Year (like High_YTD), High_Monthly (like High_Month), Table_1–Table_5.", False),
        ("", False), ("Problems in the source data and how they were handled", True),
    ]
    for i, (txt, b) in enumerate(notes, 1):
        c = ws.cell(i, 1, txt); c.font = Font(bold=b, size=13 if i == 1 else 11)
    write_issues(ws, len(notes) + 1, issues)

    s = summary(res, months)
    ws = wb.create_sheet("Summary")
    for j, c in enumerate(s.columns, 1):
        cell = ws.cell(1, j, c); cell.font = BOLD; cell.fill = HDR_FILL; cell.alignment = Alignment(wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = 16
    for i, row in enumerate(s.itertuples(index=False), 2):
        for j, (c, v) in enumerate(zip(s.columns, row), 1):
            v = None if isinstance(v, float) and np.isnan(v) else (float(v) if isinstance(v, (np.floating, float)) else int(v))
            cell = ws.cell(i, j, v)
            cell.number_format = "0.0" if "pct" in c else ("#,##0.0" if "_t_" in c else "#,##0") if c not in ("Year", "Months_included") else "0"
    ws.freeze_panes = "B2"
    for t in range(1, 6):
        write_tidy(wb, t, res[t])
    hl.highlights_all(wb, cx, SRC)
    hl.monthly_sheet(wb, cx)
    wb.save(os.path.join(OUT, "International_airline_activity_yearly_ALL.xlsx"))

    print("Months per year:", {y: len(m) for y, m in sorted(months.items())})
    print("Header warnings:", WARN or "none")
    print(f"Unparseable cells: {len(UNPARSED)}", sorted(set(UNPARSED))[:20])
    print(f"Consistency checks: {len(problems)} problem(s)")
    for p in problems[:40]:
        print("  ", p)
    print(s.to_string(index=False))


if __name__ == "__main__":
    main()

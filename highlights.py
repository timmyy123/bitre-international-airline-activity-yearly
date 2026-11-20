"""
Highlights sheets (the yearly equivalent of BITRE's High_YTD and High_Month sheets).

Per-year workbook:
  High_Year     — like High_YTD: monthly passenger chart, airline pie, Table I traffic, top 10 city pairs
                  (passengers / freight), top 10 countries, airports, top 10 airlines.
  High_Monthly  — like High_Month: month-by-month traffic for the year, and passenger / freight charts
                  comparing the months with earlier years.
Combined workbook:
  Highlights    — all years side by side, with charts.
  Monthly_series— every month: passengers, freight, mail, seats, flights, aircraft movements, seat utilisation.

Comparisons with the previous year use the SAME MONTHS of the previous year (matters for part-years).
Where the previous year's months are not in the downloaded files (2016), the prior-year figures printed
in the monthly tables are used (passengers, freight, mail, movements, city pairs, airports); otherwise '..'.
"""
import os
import re

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MON = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
NAVY = "1F2D86"
TITLE = Font(bold=True, size=16, color="1F2DB8")
H2 = Font(bold=True, size=11, color=NAVY)
BOLD = Font(bold=True)
HDR_FILL = PatternFill("solid", fgColor="DCE6F1")
TOT_FILL = PatternFill("solid", fgColor="FFF2CC")
SUB_FILL = PatternFill("solid", fgColor="F2F2F2")
RULE = Border(bottom=Side(style="medium", color="8EA9C1"))
THIN = Border(bottom=Side(style="thin", color="7F7F7F"))
INT = '#,##0;-#,##0;"-"'
DEC1 = '#,##0.0;-#,##0.0;"-"'
DEC3 = '#,##0.000;-#,##0.000;"-"'
PCT = '0.0%;-0.0%;0.0%'
SERIES_COLORS = ["1F3FD9", "E46C0A", "7F7F7F", "2E8B57", "8E44AD", "C0392B", "17A2B8", "B8860B", "5D6D7E", "D35400"]


# ------------------------------------------------------------------ data
def monthly_series(data, pub):
    d1 = data[1].groupby(["Year", "Month"])[["In_Passengers", "Out_Passengers", "In_Freight_t", "Out_Freight_t", "In_Mail_t", "Out_Mail_t"]].sum(min_count=1)
    d3 = data[3].groupby(["Year", "Month"])[["In_Flights", "Out_Flights", "In_Seats", "Out_Seats", "In_Pax", "Out_Pax"]].sum(min_count=1)
    d4 = data[4][data[4].Measure.str.lower().str.contains("movement")].groupby(["Year", "Month"])[["In", "Out", "In_prev", "Out_prev"]].sum(min_count=1)
    m = pd.DataFrame(index=d1.index)
    m["Pax_In"], m["Pax_Out"] = d1.In_Passengers, d1.Out_Passengers
    m["Passengers"] = m.Pax_In + m.Pax_Out
    m["Freight_In"], m["Freight_Out"] = d1.In_Freight_t, d1.Out_Freight_t
    m["Freight"] = d1[["In_Freight_t", "Out_Freight_t"]].sum(axis=1, min_count=1)
    m["Mail"] = d1[["In_Mail_t", "Out_Mail_t"]].sum(axis=1, min_count=1)
    m["Seats"] = d3.In_Seats + d3.Out_Seats
    m["Flights"] = d3.In_Flights + d3.Out_Flights
    m["Movements"] = d4.In + d4.Out
    m["Seat_util"] = (d3.In_Pax + d3.Out_Pax) / m.Seats
    # prior-year same-month figures as printed in each monthly publication
    p2 = pd.DataFrame({(int(k[:4]), int(k[5:])): v for k, v in pub[2].items()}, index=["Pax_prev", "_p", "Freight_prev", "_f", "Mail_prev", "_m"]).T
    m["Passengers_prev"], m["Freight_prev"], m["Mail_prev"] = p2.Pax_prev, p2.Freight_prev, p2.Mail_prev
    m["Movements_prev"] = d4.In_prev + d4.Out_prev
    return m.sort_index()


METRICS = [("Passengers Carried", "Passengers", INT), ("Freight (tonnes)", "Freight", DEC1), ("Mail (tonnes)", "Mail", DEC1),
           ("Available Seats", "Seats", INT), ("Flights", "Flights", INT), ("Aircraft Movements", "Movements", INT)]


class Ctx:
    def __init__(self, data, pub, months):
        self.data, self.months = data, {y: sorted(v) for y, v in months.items()}
        self.ms = monthly_series(data, pub)
        self.years = sorted(self.months)
        self.full = [y for y in self.years if len(self.months[y]) == 12]
        d = data
        both = lambda df, a, b: df[[a, b]].sum(axis=1, min_count=1)  # inbound + outbound; '..' in one direction doesn't blank the total
        self.cp = d[5].assign(P=both(d[5], "Pax_In", "Pax_Out"), P_prev=both(d[5], "Pax_In_prev", "Pax_Out_prev"),
                              F=d[5][["Frt_In", "Frt_Out"]].sum(axis=1, min_count=1), F_prev=d[5][["Frt_In_prev", "Frt_Out_prev"]].sum(axis=1, min_count=1))
        self.ops = d[1].assign(P=d[1][["In_Passengers", "Out_Passengers"]].sum(axis=1, min_count=1),
                               F=d[1][["In_Freight_t", "Out_Freight_t"]].sum(axis=1, min_count=1))
        ap = d[4][d[4].Measure == "Passengers"]
        self.ap = ap.assign(P=both(ap, "In", "Out"), P_prev=both(ap, "In_prev", "Out_prev"))
        af = d[4][d[4].Measure.str.lower().str.startswith("freight")]
        self.apf = af.assign(F=both(af, "In", "Out"), F_prev=both(af, "In_prev", "Out_prev"))

    def label(self, y):
        ms = self.months[y]
        return str(y) if len(ms) == 12 else f"{y} ({MON[ms[0]-1]}–{MON[ms[-1]-1]})"

    def has_months(self, y, ms):
        return y in self.months and set(ms) <= set(self.months[y])

    # yearly total of a monthly-series metric
    def total(self, y, col):
        return self.ms.loc[y, col].sum(min_count=1) if y in self.months else np.nan

    def prev_same(self, y, col):
        """Previous year, same months as year y."""
        ms = self.months[y]
        if self.has_months(y - 1, ms):
            return self.ms.loc[[(y - 1, m) for m in ms], col].sum(min_count=1)
        if f"{col}_prev" in self.ms.columns:
            return self.ms.loc[y, f"{col}_prev"].sum(min_count=1)
        return np.nan

    def by_key(self, df, keys, col, y, ms=None, prev_col=None):
        """Sum of `col` by keys for year y (optionally only months ms). Uses prev_col of year y+1 if y is missing."""
        if y in self.months and (ms is None or self.has_months(y, ms)):
            sub = df[df.Year == y] if ms is None else df[(df.Year == y) & df.Month.isin(ms)]
            return sub.groupby(keys)[col].sum(min_count=1)
        if prev_col and ms is not None and y + 1 in self.months:
            sub = df[(df.Year == y + 1) & df.Month.isin(ms)]
            return sub.groupby(keys)[prev_col].sum(min_count=1)
        return pd.Series(dtype=float)


def chg(cur, prev):
    if prev is None or cur is None or pd.isna(prev) or pd.isna(cur) or prev == 0:
        return np.nan
    return cur / prev - 1


# ------------------------------------------------------------------ cell helpers
def put(ws, r, c, v, fmt=None, bold=False, fill=None, align=None):
    if isinstance(v, (float, np.floating)) and np.isnan(v):
        v = ".."
    elif isinstance(v, np.integer):
        v = int(v)
    elif isinstance(v, np.floating):
        v = float(v)
    cell = ws.cell(r, c, v)
    if fmt and not isinstance(v, str):
        cell.number_format = fmt
    if v == "..":
        cell.alignment = Alignment(horizontal="right")
    if bold:
        cell.font = BOLD
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = Alignment(horizontal=align, wrap_text=True)
    return cell


def hdr(ws, r, c, v, width_cols=1):
    cell = ws.cell(r, c, v)
    cell.font = BOLD
    cell.fill = HDR_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return cell


def title(ws, r, text, c=2):
    ws.cell(r, c, text).font = H2
    return r + 1


def style_chart(ch, w=24, h=8):
    ch.width, ch.height = w, h
    for ax in (ch.x_axis, ch.y_axis):
        ax.delete = False
    ch.legend.position = "b"


def color_series(ch, colors=SERIES_COLORS, width=22000):
    for i, s in enumerate(ch.series):
        s.graphicalProperties.line.solidFill = colors[i % len(colors)]
        s.graphicalProperties.line.width = width
        s.smooth = False


# ------------------------------------------------------------------ table blocks
def table_traffic(ws, r, cx, years, cmp_year=None, title_text="Table I       International Traffic"):
    """Table I — traffic by year. cmp_year: show one 'Change' column for cmp_year vs previous year (same months)."""
    r = title(ws, r, title_text)
    hdr(ws, r, 2, "")
    for j, y in enumerate(years):
        hdr(ws, r, 4 + j, cx.label(y))
    cc = 4 + len(years)
    if cmp_year:
        hdr(ws, r, cc, f"Change {cmp_year}/{cmp_year-1}" + ("" if len(cx.months[cmp_year]) == 12 else " (same months)"))
    for j in range(2, cc + (1 if cmp_year else 0)):
        ws.cell(r, j).border = THIN
    r += 1
    for lab, col, fmt in METRICS:
        put(ws, r, 2, lab)
        for j, y in enumerate(years):
            put(ws, r, 4 + j, cx.total(y, col), fmt)
        if cmp_year:
            put(ws, r, cc, chg(cx.total(cmp_year, col), cx.prev_same(cmp_year, col)), PCT)
        r += 1
    put(ws, r, 2, "Seat utilisation %")
    for j, y in enumerate(years):
        t3 = cx.data[3][cx.data[3].Year == y]
        put(ws, r, 4 + j, (t3.In_Pax.sum() + t3.Out_Pax.sum()) / (t3.In_Seats.sum() + t3.Out_Seats.sum()), PCT)
    return r + 2


def table_ranked(ws, r, cx, title_text, df, keys, key_labels, col, prev_col, years, rank_year, top=10, fmt=INT,
                 all_label="ALL", top_label="Top 10", other_label="Other", cmp_year=None, show_all=False):
    """Top-N table: one row per key, one column per year, ranked by rank_year; then Top N / Other / ALL rows."""
    r = title(ws, r, title_text)
    for i, kl in enumerate(key_labels):
        hdr(ws, r, 2 + i, kl)
    if len(key_labels) == 1:
        hdr(ws, r, 3, "")
    for j, y in enumerate(years):
        hdr(ws, r, 4 + j, cx.label(y))
    cc = 4 + len(years)
    if cmp_year:
        hdr(ws, r, cc, f"Change {cmp_year}/{cmp_year-1}" + ("" if len(cx.months[cmp_year]) == 12 else " (same months)"))
    r += 1
    vals = {y: cx.by_key(df, keys, col, y) for y in years}
    rank = vals[rank_year].dropna().sort_values(ascending=False)
    rank = rank[rank > 0]
    chosen = list(rank.index if show_all else rank.index[:top])
    cur = prv = None
    if cmp_year:
        ms = cx.months[cmp_year]
        cur = cx.by_key(df, keys, col, cmp_year)
        prv = cx.by_key(df, keys, col, cmp_year - 1, ms=ms, prev_col=prev_col)
    for k in chosen:
        kk = k if isinstance(k, tuple) else (k,)
        for i, v in enumerate(kk):
            put(ws, r, 2 + i, v)
        for j, y in enumerate(years):
            put(ws, r, 4 + j, vals[y].get(k, np.nan), fmt)
        if cmp_year:
            put(ws, r, cc, chg(cur.get(k, np.nan), prv.get(k, np.nan)), PCT)
        r += 1
    rows = [] if show_all else [(f"{top_label}", lambda s: s.reindex(chosen).sum(min_count=1)),
                                (f"{other_label}", lambda s: s.drop([c for c in chosen if c in s.index]).sum(min_count=1))]
    rows.append((all_label, lambda s: s.sum(min_count=1)))
    for lab, fn in rows:
        put(ws, r, 2, lab, bold=True)
        for j, y in enumerate(years):
            put(ws, r, 4 + j, fn(vals[y]), fmt, bold=True)
        if cmp_year:
            put(ws, r, cc, chg(fn(cur), fn(prv)), PCT, bold=True)
        fill = TOT_FILL if lab == all_label else SUB_FILL
        for j in range(2, cc + (1 if cmp_year else 0)):
            ws.cell(r, j).fill = fill
        r += 1
    return r + 2, chosen


def table_airline_share(ws, r, cx, years, rank_year, top=10):
    """Top ten airlines: share of passengers and of freight, by year."""
    r = title(ws, r, f"Table VI       Top ten airlines — share of passengers and freight carried (ranked by {cx.label(rank_year)})")
    hdr(ws, r, 2, "Share of passengers carried")
    for j, y in enumerate(years):
        hdr(ws, r, 4 + j, cx.label(y))
    c2 = 4 + len(years) + 1
    hdr(ws, r, c2, "Share of freight carried")
    for j, y in enumerate(years):
        hdr(ws, r, c2 + 2 + j, cx.label(y))
    r += 1
    out = r
    for col, c0 in (("P", 2), ("F", c2)):
        shares = {y: (lambda s: s / s.sum())(cx.by_key(cx.ops, ["Operator"], col, y)) for y in years}
        top_ops = shares[rank_year].sort_values(ascending=False).index[:top]
        rr = r
        for op in top_ops:
            put(ws, rr, c0, op)
            for j, y in enumerate(years):
                put(ws, rr, c0 + 2 + j, shares[y].get(op, np.nan), PCT)
            rr += 1
        put(ws, rr, c0, "Top ten total", bold=True)
        for j, y in enumerate(years):
            put(ws, rr, c0 + 2 + j, shares[y].reindex(top_ops).sum(min_count=1), PCT, bold=True)
        out = max(out, rr + 1)
    return out + 2


def table_monthly(ws, r, cx, y):
    """Month-by-month traffic for one year, with the same month a year earlier."""
    r = title(ws, r, f"Table I       Monthly traffic — {cx.label(y)}")
    cols = [("Month", None, None), ("Passengers inbound", "Pax_In", INT), ("Passengers outbound", "Pax_Out", INT),
            ("Passengers total", "Passengers", INT), (f"Passengers same month {y-1}", "_prev", INT), ("Change", "_chg", PCT),
            ("Freight (tonnes)", "Freight", DEC1), ("Mail (tonnes)", "Mail", DEC1), ("Available seats", "Seats", INT),
            ("Flights", "Flights", INT), ("Aircraft movements", "Movements", INT), ("Seat utilisation", "Seat_util", PCT)]
    for j, (h, _, _) in enumerate(cols):
        hdr(ws, r, 2 + j, h)
    ws.row_dimensions[r].height = 30
    r += 1
    first = r
    for m in cx.months[y]:
        row = cx.ms.loc[(y, m)]
        prev = cx.ms.loc[(y - 1, m), "Passengers"] if (y - 1, m) in cx.ms.index else row.Passengers_prev
        for j, (h, col, fmt) in enumerate(cols):
            if col is None:
                put(ws, r, 2 + j, MON[m - 1])
            elif col == "_prev":
                put(ws, r, 2 + j, prev, fmt)
            elif col == "_chg":
                put(ws, r, 2 + j, chg(row.Passengers, prev), fmt)
            else:
                put(ws, r, 2 + j, row[col], fmt)
        r += 1
    # total row
    put(ws, r, 2, "Year" if len(cx.months[y]) == 12 else f"Total ({len(cx.months[y])} months)", bold=True)
    yr = cx.ms.loc[y]
    prev_tot = cx.prev_same(y, "Passengers")
    for j, (h, col, fmt) in enumerate(cols[1:], 1):
        if col == "_prev":
            v = prev_tot
        elif col == "_chg":
            v = chg(yr.Passengers.sum(), prev_tot)
        elif col == "Seat_util":
            t3 = cx.data[3][cx.data[3].Year == y]
            v = (t3.In_Pax.sum() + t3.Out_Pax.sum()) / (t3.In_Seats.sum() + t3.Out_Seats.sum())
        else:
            v = yr[col].sum(min_count=1)
        put(ws, r, 2 + j, v, fmt, bold=True)
    for j in range(len(cols)):
        ws.cell(r, 2 + j).fill = TOT_FILL
    # busiest / quietest month
    r += 2
    busiest = yr.Passengers.idxmax(); quietest = yr.Passengers.idxmin()
    ws.cell(r, 2, f"Busiest month: {MON[busiest-1]} ({yr.Passengers.max():,.0f} passengers).   "
                  f"Quietest month: {MON[quietest-1]} ({yr.Passengers.min():,.0f} passengers).").font = Font(italic=True)
    return r + 2


# ------------------------------------------------------------------ chart data + charts
def chart_data_block(ws, r, c, header, rows):
    """Write a small table used as chart source; returns (first_row, last_row)."""
    for j, h in enumerate(header):
        ws.cell(r, c + j, h).font = BOLD
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row):
            put(ws, r + i, c + j, v, "0.000" if isinstance(v, (float, np.floating)) else None)
    return r, r + len(rows)


def line_chart(ws, anchor, title_text, y_title, data_ref, cats_ref, w=24, h=8, colors=SERIES_COLORS, y_fmt=None):
    ch = LineChart()
    ch.title, ch.y_axis.title = title_text, y_title
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)
    style_chart(ch, w, h)
    color_series(ch, colors)
    if y_fmt:
        ch.y_axis.numFmt = y_fmt
    ch.y_axis.majorGridlines = ch.y_axis.majorGridlines
    ch.x_axis.tickLblSkip = None
    ws.add_chart(ch, anchor)
    return ch


def pie_chart(ws, anchor, title_text, data_ref, cats_ref, w=18, h=10):
    ch = PieChart()
    ch.title = title_text
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showCatName = True
    ch.dataLabels.showPercent = True
    ch.dataLabels.showVal = False
    ch.dataLabels.showSerName = False
    ch.dataLabels.showLeaderLines = True
    ch.dataLabels.position = "outEnd"
    ch.width, ch.height = w, h
    ch.legend = None
    ws.add_chart(ch, anchor)


def bar_chart(ws, anchor, title_text, y_title, data_ref, cats_ref, w=24, h=8):
    ch = BarChart()
    ch.type = "col"
    ch.title, ch.y_axis.title = title_text, y_title
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)
    style_chart(ch, w, h)
    ch.legend = None
    ch.series[0].graphicalProperties.solidFill = NAVY
    ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
    ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = False
    ws.add_chart(ch, anchor)


def airline_pie_rows(cx, y, top=10):
    s = cx.by_key(cx.ops, ["Operator"], "P", y).dropna().sort_values(ascending=False)
    tot = s.sum()
    rows = [(op, v / tot) for op, v in s.iloc[:top].items()]
    rows.append(("Others", s.iloc[top:].sum() / tot))
    return rows


def setup(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 18
    for j in range(4, 30):
        ws.column_dimensions[get_column_letter(j)].width = 15


# ------------------------------------------------------------------ per-year sheets
def high_year(ws, cx, y):
    setup(ws)
    ms = cx.months[y]
    head = f"YEAR ENDED DECEMBER {y}" if len(ms) == 12 else f"{y} — {MON[ms[0]-1].upper()} TO {MON[ms[-1]-1].upper()} ({len(ms)} MONTHS)"
    c = ws.cell(2, 2, head); c.font = TITLE
    for j in range(2, 10):
        ws.cell(2, j).border = RULE
    window = [yy for yy in range(y - 3, y + 1) if yy in cx.months]
    DATA_COL = 14  # chart data lives to the right (column N onwards)

    # Chart I: monthly passengers inbound / outbound, up to 3 years ending with y
    pts = [(yy, m) for yy in window[-3:] for m in cx.months[yy]]
    rows = [(f"{MON[m-1]}-{str(yy)[2:]}", cx.ms.loc[(yy, m), "Pax_In"] / 1e6, cx.ms.loc[(yy, m), "Pax_Out"] / 1e6) for yy, m in pts]
    ws.cell(4, 2, f"Chart I       International Passengers Carried (millions) — {MON[pts[0][1]-1]} {pts[0][0]} to {MON[pts[-1][1]-1]} {pts[-1][0]}").font = H2
    r0, r1 = chart_data_block(ws, 4, DATA_COL, ["Month", "INBOUND", "OUTBOUND"], rows)
    ws.cell(3, DATA_COL, "Data for charts").font = H2
    line_chart(ws, "B5", None, "Passengers (millions)", Reference(ws, min_col=DATA_COL + 1, max_col=DATA_COL + 2, min_row=r0, max_row=r1),
               Reference(ws, min_col=DATA_COL, min_row=r0 + 1, max_row=r1), w=26, h=8.5)

    # Chart II: airline pie
    ws.cell(23, 2, f"Chart II       International Passengers by Major Airlines — {cx.label(y)}").font = H2
    prow = airline_pie_rows(cx, y)
    p0, p1 = chart_data_block(ws, r1 + 3, DATA_COL, ["Airline", "Share"], prow)
    for rr in range(p0 + 1, p1 + 1):
        ws.cell(rr, DATA_COL + 1).number_format = PCT
    pie_chart(ws, "B24", None, Reference(ws, min_col=DATA_COL + 1, min_row=p0, max_row=p1), Reference(ws, min_col=DATA_COL, min_row=p0 + 1, max_row=p1), w=22, h=11)

    r = 47
    r = table_traffic(ws, r, cx, window, cmp_year=y)
    r, _ = table_ranked(ws, r, cx, "Table II       International Passengers by Uplift/Discharge City Pairs", cx.cp, ["Foreign_Port", "Australian_Port"],
                        ["Foreign Port", "Australian Port"], "P", "P_prev", window, y, all_label="ALL CITY PAIRS", top_label="Top 10 City Pairs",
                        other_label="Other City Pairs", cmp_year=y)
    r, _ = table_ranked(ws, r, cx, "Table III       International Freight (tonnes) by Uplift/Discharge City Pairs", cx.cp, ["Foreign_Port", "Australian_Port"],
                        ["Foreign Port", "Australian Port"], "F", "F_prev", window, y, fmt=DEC1, all_label="ALL CITY PAIRS",
                        top_label="Top 10 City Pairs", other_label="Other City Pairs", cmp_year=y)
    r, _ = table_ranked(ws, r, cx, "Table IV       International Passengers by Uplift/Discharge Country", cx.ops, ["Country"], ["Country"], "P", None,
                        window, y, all_label="ALL COUNTRIES", top_label="Top 10 Countries", other_label="Other Countries", cmp_year=y)
    r, _ = table_ranked(ws, r, cx, "Table V       International Passenger Traffic through Australian International Airports", cx.ap, ["Airport"],
                        ["Airport"], "P", "P_prev", window, y, all_label="ALL AIRPORTS", cmp_year=y, show_all=True)
    r = table_airline_share(ws, r, cx, window[-2:], y)
    ws.cell(r, 2, "Comparisons with the previous year use the same months of that year. '..' = not available.").font = Font(italic=True, size=9)
    ws.freeze_panes = "A3"


def high_monthly(ws, cx, y):
    setup(ws)
    c = ws.cell(2, 2, f"MONTHLY PROFILE — {cx.label(y)}"); c.font = TITLE
    for j in range(2, 14):
        ws.cell(2, j).border = RULE
    r = table_monthly(ws, 4, cx, y)
    DATA_COL = 16
    window = [yy for yy in range(y - 3, y + 1) if yy in cx.months]
    # chart data: rows = months, cols = years
    ws.cell(r, 2, "Chart I       Passengers carried by month (millions)").font = H2
    ws.cell(r, 10, "Chart II       Freight carried by month (thousand tonnes)").font = H2
    hdr_row = r + 22
    ws.cell(hdr_row - 1, DATA_COL, "Data for charts").font = H2
    for k, (col, scale) in enumerate((("Passengers", 1e6), ("Freight", 1e3))):
        c0 = DATA_COL + k * (len(window) + 2)
        rows = []
        for m in range(1, 13):
            rows.append([MON[m - 1]] + [cx.ms.loc[(yy, m), col] / scale if (yy, m) in cx.ms.index else None for yy in window])
        a, b = chart_data_block(ws, hdr_row, c0, ["Month"] + [str(yy) for yy in window], rows)
        line_chart(ws, f"B{r+1}" if k == 0 else f"J{r+1}", None, "Passengers (millions)" if k == 0 else "Freight ('000 tonnes)",
                   Reference(ws, min_col=c0 + 1, max_col=c0 + len(window), min_row=a, max_row=b),
                   Reference(ws, min_col=c0, min_row=a + 1, max_row=b), w=16, h=9,
                   colors=["BFBFBF", "8EA9C1", "E46C0A", "1F3FD9"][-len(window):])
    ws.freeze_panes = "A3"


# ------------------------------------------------------------------ combined workbook
def bitre_published(src):
    """BITRE's own year-ended-December figures, taken from the latest December file that reports each year."""
    out = {}
    for y in range(2017, 2100):
        f = os.path.join(src, f"{y}-12.xlsx")
        if not os.path.exists(f):
            continue
        ws = openpyxl.load_workbook(f, data_only=True, read_only=True)["High_YTD"]
        rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=200, values_only=True)]
        year_row = None
        seen_all = 0
        for r in rows:
            # year headings are ints (2022 …) in newer files and dates (Dec-2016 …) in older ones
            nums = [(j, v.year if hasattr(v, "year") else v) for j, v in enumerate(r)
                    if (isinstance(v, int) and 2000 < v < 2100) or (hasattr(v, "year") and getattr(v, "month", 0) == 12)]
            vals = [v for _, v in nums]
            if len(nums) >= 3 and vals == list(range(vals[0], vals[0] + len(vals))):  # a header row of consecutive years
                year_row = dict(nums)
            lab = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ""
            key = None
            if lab in ("Passengers Carried", "Freight (tonnes)", "Mail (tonnes)", "Available Seats", "Flights", "Aircraft Movements"):
                key = lab
            elif lab == "ALL CITY PAIRS":
                key = ["Passengers Carried", "Freight (tonnes)"][min(seen_all, 1)]; seen_all += 1
            if key and year_row:
                for j, yy in year_row.items():
                    v = r[j] if j < len(r) else None
                    if isinstance(v, (int, float)):
                        out[(yy, key)] = (float(v), y)
    return out


def highlights_all(wb, cx, src):
    ws = wb.create_sheet("Highlights", 1)
    setup(ws)
    yrs = cx.years
    last_full = cx.full[-1]
    ws.cell(2, 2, f"HIGHLIGHTS — {cx.label(yrs[0])} to {cx.label(yrs[-1])}").font = TITLE
    for j in range(2, 4 + len(yrs) + 1):
        ws.cell(2, j).border = RULE
    ws.cell(3, 2, f"Part-years: {', '.join(cx.label(y) for y in yrs if y not in cx.full)}. "
                  "Changes compare the same months of the previous year.").font = Font(italic=True, size=9)
    DATA_COL = 2 * len(yrs) + 12  # to the right of every table (Table VI uses two blocks of year columns)

    # Chart I: annual passengers (bar)
    ws.cell(5, 2, "Chart I       International passengers carried by year (millions)").font = H2
    a, b = chart_data_block(ws, 5, DATA_COL, ["Year", "Passengers (m)"], [(cx.label(y), cx.total(y, "Passengers") / 1e6) for y in yrs])
    ws.cell(4, DATA_COL, "Data for charts").font = H2
    bar_chart(ws, "B6", None, "Passengers (millions)", Reference(ws, min_col=DATA_COL + 1, min_row=a, max_row=b),
              Reference(ws, min_col=DATA_COL, min_row=a + 1, max_row=b), w=26, h=8)

    # Chart II: monthly passengers in/out across all months
    ws.cell(23, 2, f"Chart II       International passengers carried by month (millions) — all {len(cx.ms)} months").font = H2
    rows = [(f"{MON[m-1]}-{str(y)[2:]}", r.Pax_In / 1e6, r.Pax_Out / 1e6) for (y, m), r in cx.ms.iterrows()]
    a2, b2 = chart_data_block(ws, b + 3, DATA_COL, ["Month", "INBOUND", "OUTBOUND"], rows)
    line_chart(ws, "B24", None, "Passengers (millions)", Reference(ws, min_col=DATA_COL + 1, max_col=DATA_COL + 2, min_row=a2, max_row=b2),
               Reference(ws, min_col=DATA_COL, min_row=a2 + 1, max_row=b2), w=26, h=8)

    # Chart III: seat utilisation by year; Chart IV: airline pie for last full year
    ws.cell(41, 2, "Chart III       Seat utilisation by year").font = H2
    ws.cell(41, 9, f"Chart IV       Passengers by major airline — {cx.label(last_full)}").font = H2
    su = []
    for y in yrs:
        t3 = cx.data[3][cx.data[3].Year == y]
        su.append((cx.label(y), (t3.In_Pax.sum() + t3.Out_Pax.sum()) / (t3.In_Seats.sum() + t3.Out_Seats.sum())))
    a3, b3 = chart_data_block(ws, b2 + 3, DATA_COL, ["Year", "Seat utilisation"], su)
    for rr in range(a3 + 1, b3 + 1):
        ws.cell(rr, DATA_COL + 1).number_format = PCT
    ch = line_chart(ws, "B42", None, "Seat utilisation", Reference(ws, min_col=DATA_COL + 1, min_row=a3, max_row=b3),
                    Reference(ws, min_col=DATA_COL, min_row=a3 + 1, max_row=b3), w=13, h=9, y_fmt="0%")
    ch.legend = None
    prow = airline_pie_rows(cx, last_full)
    a4, b4 = chart_data_block(ws, b3 + 3, DATA_COL, ["Airline", "Share"], prow)
    for rr in range(a4 + 1, b4 + 1):
        ws.cell(rr, DATA_COL + 1).number_format = PCT
    pie_chart(ws, "I42", None, Reference(ws, min_col=DATA_COL + 1, min_row=a4, max_row=b4), Reference(ws, min_col=DATA_COL, min_row=a4 + 1, max_row=b4), w=17, h=9)

    r = 62
    r = table_traffic(ws, r, cx, yrs)
    # change row block
    r -= 1
    r = title(ws, r, "Change on previous year (same months)")
    for j, y in enumerate(yrs):
        hdr(ws, r, 4 + j, cx.label(y))
    r += 1
    for lab, col, _ in METRICS:
        put(ws, r, 2, lab)
        for j, y in enumerate(yrs):
            put(ws, r, 4 + j, chg(cx.total(y, col), cx.prev_same(y, col)), PCT)
        r += 1
    r += 2
    kw = dict(years=yrs, rank_year=last_full)
    r, _ = table_ranked(ws, r, cx, f"Table II       Passengers by Uplift/Discharge City Pairs (top 10 in {last_full})", cx.cp, ["Foreign_Port", "Australian_Port"],
                        ["Foreign Port", "Australian Port"], "P", "P_prev", all_label="ALL CITY PAIRS", top_label="Top 10 City Pairs",
                        other_label="Other City Pairs", **kw)
    r, _ = table_ranked(ws, r, cx, f"Table III       Freight (tonnes) by Uplift/Discharge City Pairs (top 10 in {last_full})", cx.cp,
                        ["Foreign_Port", "Australian_Port"], ["Foreign Port", "Australian Port"], "F", "F_prev", fmt=DEC1, all_label="ALL CITY PAIRS",
                        top_label="Top 10 City Pairs", other_label="Other City Pairs", **kw)
    r, _ = table_ranked(ws, r, cx, f"Table IV       Passengers by Uplift/Discharge Country (top 10 in {last_full})", cx.ops, ["Country"], ["Country"], "P", None,
                        all_label="ALL COUNTRIES", top_label="Top 10 Countries", other_label="Other Countries", **kw)
    r, _ = table_ranked(ws, r, cx, "Table V       Passenger Traffic through Australian International Airports", cx.ap, ["Airport"], ["Airport"], "P", "P_prev",
                        all_label="ALL AIRPORTS", show_all=True, **kw)
    r, _ = table_ranked(ws, r, cx, f"Table VI       Passengers by Airline (top 10 in {last_full})", cx.ops, ["Operator"], ["Airline"], "P", None,
                        all_label="ALL AIRLINES", top_label="Top 10 Airlines", other_label="Other Airlines", **kw)
    r = table_airline_share(ws, r, cx, yrs, last_full)

    # reconciliation with BITRE's published year-ended-December figures
    pub = bitre_published(src)
    r = title(ws, r, "Check against BITRE's published year-ended-December figures (from the latest December file reporting each year)")
    heads = ["Measure", "", "Year", "This report", "BITRE published", "Difference", "Difference %", "Source file"]
    for j, h in enumerate(heads):
        hdr(ws, r, 2 + j, h)
    r += 1
    col_of = {lab: col for lab, col, _ in METRICS}
    for lab, col, fmt in METRICS:
        for y in cx.full:
            if (y, lab) not in pub:
                continue
            pv, f = pub[(y, lab)]
            mine = cx.total(y, col)
            put(ws, r, 2, lab); put(ws, r, 4, y); put(ws, r, 5, mine, fmt); put(ws, r, 6, pv, fmt)
            put(ws, r, 7, mine - pv, fmt); put(ws, r, 8, chg(mine, pv), "0.00%"); put(ws, r, 9, f"Dec {f} High_YTD")
            r += 1
    ws.cell(r, 2, "Small differences are expected: BITRE revises earlier months in later publications, while this report sums the figures "
                  "as first published each month.").font = Font(italic=True, size=9)
    ws.freeze_panes = "A4"
    return pub


def monthly_sheet(wb, cx):
    ws = wb.create_sheet("Monthly_series", 2)
    cols = [("Year", None), ("Month", None), ("Passengers_in", "Pax_In"), ("Passengers_out", "Pax_Out"), ("Passengers_total", "Passengers"),
            ("Passengers_same_month_prev_year", "_prev"), ("Passenger_change_pct", "_chg"), ("Freight_t", "Freight"), ("Mail_t", "Mail"),
            ("Available_seats", "Seats"), ("Flights", "Flights"), ("Aircraft_movements", "Movements"), ("Seat_util_pct", "Seat_util")]
    for j, (h, _) in enumerate(cols, 1):
        c = ws.cell(1, j, h); c.font = BOLD; c.fill = HDR_FILL; c.alignment = Alignment(wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = 15
    for i, ((y, m), row) in enumerate(cx.ms.iterrows(), 2):
        prev = cx.ms.loc[(y - 1, m), "Passengers"] if (y - 1, m) in cx.ms.index else row.Passengers_prev
        for j, (h, col) in enumerate(cols, 1):
            if col is None:
                v, fmt = (y, "0") if h == "Year" else (MON[m - 1], None)
            elif col == "_prev":
                v, fmt = prev, INT
            elif col == "_chg":
                v, fmt = chg(row.Passengers, prev), PCT
            else:
                v = row[col]; fmt = PCT if col == "Seat_util" else (DEC1 if col in ("Freight", "Mail") else INT)
            put(ws, i, j, v, fmt)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

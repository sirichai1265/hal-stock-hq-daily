"""
HAL Stock HQ - Daily report builder.

Reads the three same-day source files and fills the Daily sheet (plus the
RF SEASONAL rows and the NEW FORMAT zero-hiding / balance colours) of the
template workbook, producing  "HAL Stock HQ - Daily <YYYY-MM-DD>.xlsx".

Usage:
    python build_stock_hq.py                     # auto-detect the 3 files, report date = today
    python build_stock_hq.py 2026-09-10          # force the report date
    python build_stock_hq.py --actual 9-10-ACTUAL.xls --bkg 9-10-BKG+PD.xls \
                             --staying 9-10-STAYING.xls --date 2026-09-10

Source files (same business day):
    * ACTUAL   - snapshot of FULL containers on hand      -> FULL INBOUND
    * STAYING  - per-container empty-stock snapshot        -> CURRENT STOCK / RF SEASONAL
    * BKG+PD   - live pickup list of not-yet-collected BKs -> BOOKING wk1 / wk2
"""

import argparse
import datetime as dt
import glob
import os
import re
import shutil
import subprocess
import sys

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill

# --------------------------------------------------------------------------- #
#  Static configuration                                                        #
# --------------------------------------------------------------------------- #
HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "(HAL)Stock form.xlsx")

TEMPLATE_TYPES = ["22GP", "42GP", "45GP", "22RE", "45RE", "22UT", "42UT", "22PC", "42PC"]

# template column letters, in TEMPLATE_TYPES order
BKK_COLS_INPUT = ["C", "D", "E", "F", "G", "H", "I", "J", "K"]      # FULL INBOUND + BOOKING (BKK)
BKK_COLS_STOCK = ["L", "M", "N", "O", "P", "Q", "R", "S", "T"]      # CURRENT STOCK + STOCK END WK (BKK)
LCH_COLS_INPUT = ["X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF"]  # FULL INBOUND + BOOKING (LCH)
LCH_COLS_STOCK = ["AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO"]  # CURRENT STOCK + STOCK END WK (LCH)

# location group -> (business type, [location codes])
GROUP_TYPE = {
    "PAT": "TERMINAL", "UNITHAI": "TERMINAL", "ESCO #2": "ICD", "BMTP": "ICD",
    "SMART": "DEPOT", "BC2": "DEPOT",
    "LCMT": "TERMINAL", "ESCO": "TERMINAL", "LCIT": "TERMINAL", "HPT": "TERMINAL",
    "HAST": "DEPOT", "CELLO": "DEPOT", "PW": "DEPOT",
}
GROUP_CODES = {
    "PAT": ["BKK01", "BKK04"], "UNITHAI": ["BKK02"], "ESCO #2": ["LCH55"],
    "BMTP": ["BKKY4"], "SMART": ["BKK25"], "BC2": ["BKK27"],
    "LCMT": ["LCH04", "LCH07"], "ESCO": ["LCH01"], "LCIT": ["LCH05", "LCH10"],
    "HPT": ["LCH02", "LCH06", "LCH08", "LCH09"],
    "HAST": ["LCH27"], "CELLO": ["LCH28"], "PW": ["LCHY5"],
}

# template row numbers per group
BKK_STOCK_ROW = {"PAT": 4, "UNITHAI": 5, "ESCO #2": 6, "BMTP": 7, "SMART": 8, "BC2": 9}
LCH_STOCK_ROW = {"LCMT": 4, "ESCO": 5, "LCIT": 6, "HPT": 7, "HAST": 8, "CELLO": 9, "PW": 10}
BKK_BKG_ROW_WK1 = {"PAT": 31, "UNITHAI": 32, "ESCO #2": 33, "BMTP": 34, "SMART": 35, "BC2": 36}
LCH_BKG_ROW_WK1 = {"LCMT": 31, "ESCO": 32, "LCIT": 33, "HPT": 34, "HAST": 35, "CELLO": 36, "PW": 37}
BKK_BKG_ROW_WK2 = {g: r + 14 for g, r in BKK_BKG_ROW_WK1.items()}   # 45-50
LCH_BKG_ROW_WK2 = {g: r + 14 for g, r in LCH_BKG_ROW_WK1.items()}   # 45-51

BKK_GROUPS = ["PAT", "UNITHAI", "ESCO #2", "BMTP", "SMART", "BC2"]
LCH_GROUPS = ["LCMT", "ESCO", "LCIT", "HPT", "HAST", "CELLO", "PW"]

# BKG+PD container-count column -> template type
BKG_TYPE_COL = {
    "GP22": "22GP", "GP42": "42GP", "GP45": "45GP",
    "RE22": "22RE", "RE45": "45RE",
    "UT22": "22UT", "UT42": "42UT",
    "PC22": "22PC", "PC42": "42PC",
}

# RF SEASONAL sheet
RF_SEASONAL_YEARS = [2025, 2023, 2021, 2020]
RF_SEASONAL_YEAR_ROW = {2025: 4, 2023: 5, 2021: 6, 2020: 7}
RF_BRAND_TO_COL = {"CARRIER": "CARRIER", "DAIKIN": "DAIKIN", "THERMOKING": "TRMK", "TRMK": "TRMK"}
RF_COL = {
    "BKK": {"CARRIER": "B", "DAIKIN": "C", "TRMK": "D"},
    "LCH": {"CARRIER": "G", "DAIKIN": "H", "TRMK": "I"},
}

BLACK, BLUE, RED = "FF000000", "FF0000FF", "FFFF0000"
PINK_FILL = PatternFill("solid", fgColor="FFC7CE")
ZERO_HIDE_FORMAT = "0;-0;;@"

# NEW FORMAT sheet: rows where a single 0 must be hidden (except Current Stock / Balance)
NF_HIDE_ROWS = [4, 7, 8, 9, 15, 18, 19, 20]
NF_COLS = ["C", "D", "E", "F", "G", "H", "I", "J", "K"]
NF_BALANCE = {10: [f"{c}10" for c in "CDEFGHIJK"], 21: [f"{c}21" for c in "CDEFGHIJK"]}


# --------------------------------------------------------------------------- #
#  Source-file loading                                                         #
# --------------------------------------------------------------------------- #
def load_data_sheet(path, header=0):
    """Always open the raw per-container tab by NAME, never by position."""
    xls = pd.ExcelFile(path)
    sheet = "Sheet" if "Sheet" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(path, sheet_name=sheet, header=header)


def load_booking(path):
    df = load_data_sheet(path, header=0)
    if "Pickup" not in df.columns:                       # a title row sat above the header
        df = load_data_sheet(path, header=1)
    df = df[df["BK No"].notna()].copy()                  # drop the trailing total row
    df["_date"] = pd.to_datetime(
        df["TRAN DT"].astype("Int64").astype(str), format="%Y%m%d", errors="coerce"
    )
    return df


def load_staying(path):
    df = load_data_sheet(path, header=0)
    return df[df["Location"].notna()].copy()


def load_actual(path):
    df = load_data_sheet(path, header=0)
    df = df[df["Location"].notna()].copy()
    if "Move Code" not in df.columns or "Size/Type" not in df.columns:
        raise SystemExit(
            f"{os.path.basename(path)}: expected raw per-container columns "
            f"(Move Code / Size/Type) not found - sheets: {pd.ExcelFile(path).sheet_names}"
        )
    return df


# --------------------------------------------------------------------------- #
#  Aggregation                                                                 #
# --------------------------------------------------------------------------- #
def _empty_counts():
    return {t: 0 for t in TEMPLATE_TYPES}


def current_stock(stay):
    """CURRENT STOCK: plain snapshot count of STAYING rows by group + type."""
    out = {}
    for group, codes in GROUP_CODES.items():
        sub = stay[stay["Location"].isin(codes)]
        c = _empty_counts()
        for t, n in sub["Size/Type"].value_counts().items():
            if t in c:
                c[t] = int(n)
        out[group] = c
    return out


def full_inbound(actual):
    """FULL INBOUND: ACTUAL rows by group + type where Move Code != 'OFD'."""
    inb = actual[actual["Move Code"].astype(str).str.upper() != "OFD"]
    out = {}
    for group, codes in GROUP_CODES.items():
        sub = inb[inb["Location"].isin(codes)]
        c = _empty_counts()
        for t, n in sub["Size/Type"].value_counts().items():
            if t in c:
                c[t] = int(n)
        out[group] = c
    return out


def booking(bkg, lo, hi):
    """Sum BKG+PD container-count columns by group for lo <= TRAN DT <= hi.

    lo=None means 'no lower bound' (wk1: overdue backlog is still outstanding).
    """
    m = bkg["_date"] <= pd.Timestamp(hi)
    if lo is not None:
        m &= bkg["_date"] >= pd.Timestamp(lo)
    win = bkg[m]
    out = {}
    for group, codes in GROUP_CODES.items():
        sub = win[win["Pickup"].isin(codes)]
        c = _empty_counts()
        for src, tgt in BKG_TYPE_COL.items():
            if src in sub.columns:
                c[tgt] = int(pd.to_numeric(sub[src], errors="coerce").fillna(0).sum())
        out[group] = c
    return out


def rf_seasonal_counts(stay):
    """45RE reefer units by Area x Built Year x brand, selected years only."""
    out = {(a, y, b): 0
           for a in ("BKK", "LCH") for y in RF_SEASONAL_YEARS
           for b in ("CARRIER", "DAIKIN", "TRMK")}
    if "Built Year" not in stay.columns or "RF Brand" not in stay.columns:
        raise SystemExit("STAYING file lacks 'Built Year' / 'RF Brand' - cannot build RF SEASONAL.")
    d = stay.copy()
    d["_year"] = pd.to_numeric(d["Built Year"], errors="coerce")
    d = d[(d["Size/Type"] == "45RE") & (d["_year"].isin(RF_SEASONAL_YEARS))]
    for (area, year, brand), n in d.groupby(["Area", "_year", "RF Brand"]).size().items():
        col = RF_BRAND_TO_COL.get(str(brand).strip().upper())
        if col is None:
            print(f"  WARNING: RF SEASONAL - unmapped RF Brand '{brand}' "
                  f"({area}/{int(year)}, {n} cntr) not counted")
            continue
        if area in ("BKK", "LCH"):
            out[(area, int(year), col)] += int(n)
    return out


# --------------------------------------------------------------------------- #
#  Date-range label                                                            #
# --------------------------------------------------------------------------- #
def fmt_range(d1, d2):
    if d1.year != d2.year:
        return f"{d1:%d/%m/%Y}-{d2:%d/%m/%Y}"
    if d1.month != d2.month:
        return f"{d1:%d/%m}-{d2:%d/%m/%Y}"
    return f"{d1.day:02d}-{d2.day:02d}/{d2:%m/%Y}"


# --------------------------------------------------------------------------- #
#  Workbook writing                                                            #
# --------------------------------------------------------------------------- #
def write_row(ws, row, col_letters, values):
    """Write the 9 type values into one row; a plain 0 is written blank."""
    for col, t in zip(col_letters, TEMPLATE_TYPES):
        v = int(values[t])
        ws[f"{col}{row}"] = v if v != 0 else None


def colour_row(ws, row, col_letters, group):
    rgb = BLUE if GROUP_TYPE[group] == "DEPOT" else BLACK
    for col in col_letters:
        c = ws[f"{col}{row}"]
        f = c.font
        c.font = Font(color=rgb, bold=f.bold, name=f.name, size=f.size, italic=f.italic)


def style_negative(ws, row, col_letters, computed):
    """Bold + pink fill on STOCK END WK cells whose computed value is negative."""
    for col, t in zip(col_letters, TEMPLATE_TYPES):
        if computed[t] < 0:
            c = ws[f"{col}{row}"]
            f = c.font
            c.font = Font(color=f.color, bold=True, name=f.name, size=f.size, italic=f.italic)
            c.fill = PINK_FILL


def hide_zero_format(ws, row, col_letters):
    for col in col_letters:
        ws[f"{col}{row}"].number_format = ZERO_HIDE_FORMAT


def build(actual_path, bkg_path, staying_path, report_date, out_path):
    actual = load_actual(actual_path)
    stay = load_staying(staying_path)
    bkg = load_booking(bkg_path)

    monday = report_date - dt.timedelta(days=report_date.weekday())
    sunday = monday + dt.timedelta(days=6)
    next_monday = monday + dt.timedelta(days=7)
    next_sunday = next_monday + dt.timedelta(days=6)

    fi = full_inbound(actual)
    cs = current_stock(stay)
    bk1 = booking(bkg, None, sunday)                       # TODAY+WK1ST: no lower bound
    bk2 = booking(bkg, next_monday, next_sunday)           # WK2ND
    rf = rf_seasonal_counts(stay)

    # STOCK END WK, computed exactly like the template's own subtraction
    sew1 = {g: {t: cs[g][t] - bk1[g][t] for t in TEMPLATE_TYPES} for g in GROUP_CODES}
    sew2 = {g: {t: sew1[g][t] - bk2[g][t] for t in TEMPLATE_TYPES} for g in GROUP_CODES}

    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Daily"]

    sides = [
        ("BKK", BKK_GROUPS, BKK_COLS_INPUT, BKK_COLS_STOCK,
         BKK_STOCK_ROW, BKK_BKG_ROW_WK1, BKK_BKG_ROW_WK2),
        ("LCH", LCH_GROUPS, LCH_COLS_INPUT, LCH_COLS_STOCK,
         LCH_STOCK_ROW, LCH_BKG_ROW_WK1, LCH_BKG_ROW_WK2),
    ]

    for _side, groups, cin, cstock, srow, r1, r2 in sides:
        for g in groups:
            # FULL INBOUND + CURRENT STOCK (rows 4-9/4-10)
            write_row(ws, srow[g], cin, fi[g])
            write_row(ws, srow[g], cstock, cs[g])
            colour_row(ws, srow[g], cin, g)
            colour_row(ws, srow[g], cstock, g)
            # BOOKING wk1 (rows 31-36/31-37) + its STOCK END WK formula cells
            write_row(ws, r1[g], cin, bk1[g])
            colour_row(ws, r1[g], cin, g)
            colour_row(ws, r1[g], cstock, g)
            style_negative(ws, r1[g], cstock, sew1[g])
            # BOOKING wk2 (rows 45-50/45-51) + its STOCK END WK formula cells
            write_row(ws, r2[g], cin, bk2[g])
            colour_row(ws, r2[g], cin, g)
            colour_row(ws, r2[g], cstock, g)
            style_negative(ws, r2[g], cstock, sew2[g])

    # hide a bare 0 on every STOCK END WK formula cell, full block height
    for row in list(range(31, 41)) + list(range(45, 55)):
        hide_zero_format(ws, row, BKK_COLS_STOCK)
        hide_zero_format(ws, row, LCH_COLS_STOCK)

    # report date + d/m/yyyy display
    ws["AL1"] = dt.datetime(report_date.year, report_date.month, report_date.day)
    ws["R2"].number_format = "d/m/yyyy"
    ws["AM2"].number_format = "d/m/yyyy"

    # bare date-range labels (no "TODAY+WK 1ST" / "WK 2ND" text)
    wk1_lbl = fmt_range(report_date, sunday)
    wk2_lbl = fmt_range(next_monday, next_sunday)
    ws["I29"] = ws["AD29"] = wk1_lbl
    ws["I43"] = ws["AD43"] = wk2_lbl

    # --- RF SEASONAL rows 4-7 (row 8 TOTAL formula untouched) ---
    rfs = wb["RF SEASONAL"]
    for area in ("BKK", "LCH"):
        for year, row in RF_SEASONAL_YEAR_ROW.items():
            for brand, col in RF_COL[area].items():
                n = rf[(area, year, brand)]
                rfs[f"{col}{row}"] = n if n != 0 else None

    # --- NEW FORMAT: hide single 0s (except Current Stock + Balance) ---
    nf = wb["NEW FORMAT"]
    for row in NF_HIDE_ROWS:
        for col in NF_COLS:
            nf[f"{col}{row}"].number_format = ZERO_HIDE_FORMAT

    # NEW FORMAT Balance colour: red iff genuinely negative, else blue.
    # Balance = CURRENT STOCK total + FULL INBOUND total - BOOKING wk1 total - BOOKING wk2 total
    bal = {
        "BKK": {t: sum(cs[g][t] for g in BKK_GROUPS) + sum(fi[g][t] for g in BKK_GROUPS)
                - sum(bk1[g][t] for g in BKK_GROUPS) - sum(bk2[g][t] for g in BKK_GROUPS)
                for t in TEMPLATE_TYPES},
        "LCH": {t: sum(cs[g][t] for g in LCH_GROUPS) + sum(fi[g][t] for g in LCH_GROUPS)
                - sum(bk1[g][t] for g in LCH_GROUPS) - sum(bk2[g][t] for g in LCH_GROUPS)
                for t in TEMPLATE_TYPES},
    }
    for row, side in ((10, "BKK"), (21, "LCH")):
        for col, t in zip("CDEFGHIJK", TEMPLATE_TYPES):
            cell = nf[f"{col}{row}"]
            f = cell.font
            cell.font = Font(color=RED if bal[side][t] < 0 else BLUE,
                             bold=f.bold, name=f.name, size=f.size, italic=f.italic)

    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)

    recalc(out_path)

    # ---- HTML dashboard (index.html for GitHub Pages) ------------------- #
    dash = os.path.join(os.path.dirname(out_path) or ".", "index.html")
    write_dashboard(dash, report_date, wk1_lbl, wk2_lbl,
                    fi, cs, bk1, bk2, sew1, sew2, bal, rf,
                    os.path.basename(out_path))
    print(f"Dashboard   : {os.path.basename(dash)}")

    # ---- console report -------------------------------------------------- #
    print(f"\nReport date : {report_date:%Y-%m-%d} ({report_date:%A})")
    print(f"WK 1ST      : {wk1_lbl}   (TRAN DT <= {sunday:%d/%m})")
    print(f"WK 2ND      : {wk2_lbl}   ({next_monday:%d/%m} .. {next_sunday:%d/%m})")
    print(f"Saved       : {os.path.basename(out_path)}")

    report_negatives("WK1ST", sew1, fi)
    report_negatives("WK2ND", sew2, fi)
    for side in ("BKK", "LCH"):
        neg = {t: v for t, v in bal[side].items() if v < 0}
        if neg:
            print(f"  NEW FORMAT balance negative ({side}): " +
                  ", ".join(f"{t} {v}" for t, v in neg.items()))
    return out_path


def report_negatives(week, sew, fi):
    any_neg = False
    for g, byt in sew.items():
        for t, v in byt.items():
            if v < 0:
                any_neg = True
                side = "BKK" if g in BKK_GROUPS else "LCH"
                if GROUP_TYPE[g] == "DEPOT":
                    note = "DEPOT - genuine shortage (no FULL INBOUND buffer), needs lease-in / repo"
                else:
                    buf = fi[g][t]
                    note = (f"FULL INBOUND holds {buf} x {t} awaiting devanning - covers it"
                            if buf >= -v else
                            f"only {buf} x {t} in FULL INBOUND - short by {-v - buf}")
                print(f"  NEG {week} {side}/{g} {t}: {v}  ->  {note}")
    if not any_neg:
        print(f"  {week}: no negative STOCK END WK cells")


def neg_note(g, t, v, fi):
    if GROUP_TYPE[g] == "DEPOT":
        return "DEPOT – ไม่มีตู้ FULL รองรับ ต้อง lease-in / repo"
    buf = fi[g][t]
    if buf >= -v:
        return f"มีตู้ FULL {buf}×{t} รอ devanning – ครอบคลุมได้"
    return f"ตู้ FULL มีแค่ {buf}×{t} – ยังขาดอีก {-v - buf}"


# --------------------------------------------------------------------------- #
#  HTML dashboard                                                              #
# --------------------------------------------------------------------------- #
DASH_CSS = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#1c2530;--mut:#667085;--line:#e3e6ea;
--depot:#1a56db;--neg:#b42318;--negbg:#fff0ef;--head:#0f2942;--accent:#0f6fff}
@media (prefers-color-scheme:dark){:root{--bg:#0e1116;--card:#161b22;--ink:#e6edf3;
--mut:#9aa4b2;--line:#2a313c;--depot:#6ea8ff;--neg:#ff8079;--negbg:#3a1c1a;--head:#c9d6e5;--accent:#4f9dff}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);
font:14px/1.5 -apple-system,Segoe UI,Roboto,"Noto Sans Thai",sans-serif}
.wrap{max-width:1180px;margin:0 auto;padding:28px 20px 60px}
h1{font-size:21px;margin:0 0 2px}.sub{color:var(--mut);margin:0 0 18px}
.meta{display:flex;gap:22px;flex-wrap:wrap;margin:0 0 22px;font-size:13px;color:var(--mut)}
.meta b{color:var(--ink)}
.alert{background:var(--card);border:1px solid var(--line);border-left:4px solid var(--neg);
border-radius:10px;padding:14px 16px;margin:0 0 24px}
.alert h2{font-size:14px;margin:0 0 8px;color:var(--neg)}
.alert ul{margin:0;padding-left:18px}.alert li{margin:3px 0}
.alert .ok{color:var(--mut)}
.grid{display:grid;gap:22px}
@media(min-width:900px){.grid{grid-template-columns:1fr 1fr}}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;overflow-x:auto}
.card h3{font-size:13px;letter-spacing:.04em;text-transform:uppercase;color:var(--head);margin:0 0 10px}
table{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums}
th,td{padding:5px 7px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap}
th:first-child,td:first-child{text-align:left}
thead th{font-size:11px;color:var(--mut);font-weight:600;border-bottom:1px solid var(--line)}
tbody tr:last-child td{border-bottom:none}
tr.total td{font-weight:700;border-top:2px solid var(--line)}
td.depot,th.depot{color:var(--depot)}
td.neg{color:var(--neg);font-weight:700;background:var(--negbg)}
.zero{color:var(--line)}
footer{margin-top:30px;color:var(--mut);font-size:12px;text-align:center}
"""


def _cell(v, neg_ok):
    if v == 0:
        return '<td class="zero">·</td>'
    if neg_ok and v < 0:
        return f'<td class="neg">{v}</td>'
    return f"<td>{v}</td>"


def _table(title, groups, data, neg_ok=False):
    th = "".join(f"<th>{t}</th>" for t in TEMPLATE_TYPES)
    rows = []
    tot = {t: 0 for t in TEMPLATE_TYPES}
    for g in groups:
        cls = ' class="depot"' if GROUP_TYPE[g] == "DEPOT" else ""
        cells = "".join(_cell(int(data[g][t]), neg_ok) for t in TEMPLATE_TYPES)
        for t in TEMPLATE_TYPES:
            tot[t] += int(data[g][t])
        rows.append(f"<tr><td{cls}>{g}</td>{cells}</tr>")
    tcells = "".join(f"<td>{tot[t] or ''}</td>" for t in TEMPLATE_TYPES)
    rows.append(f'<tr class="total"><td>TOTAL</td>{tcells}</tr>')
    return (f'<div class="card"><h3>{title}</h3><table><thead><tr><th>Group</th>{th}</tr>'
            f"</thead><tbody>{''.join(rows)}</tbody></table></div>")


def write_dashboard(path, report_date, wk1_lbl, wk2_lbl,
                    fi, cs, bk1, bk2, sew1, sew2, bal, rf, xlsx_name):
    alerts = []
    for wk, sew in (("WK1", sew1), ("WK2", sew2)):
        for g in list(BKK_GROUPS) + list(LCH_GROUPS):
            side = "BKK" if g in BKK_GROUPS else "LCH"
            for t in TEMPLATE_TYPES:
                v = sew[g][t]
                if v < 0:
                    alerts.append(f"<li><b>{side} / {g} / {t}</b> ({wk}): "
                                  f"<span style='color:var(--neg);font-weight:700'>{v}</span> "
                                  f"&mdash; {neg_note(g, t, v, fi)}</li>")
    for side in ("BKK", "LCH"):
        neg = {t: v for t, v in bal[side].items() if v < 0}
        if neg:
            alerts.append("<li><b>NEW FORMAT balance " + side + "</b>: "
                          + ", ".join(f"{t} {v}" for t, v in neg.items()) + "</li>")
    alert_html = ("<ul>" + "".join(alerts) + "</ul>") if alerts else \
        '<p class="ok">ไม่มีช่อง STOCK END WK ติดลบ</p>'

    def side_block(name, groups):
        return (f"<h2 style='font-size:16px;margin:26px 0 12px'>{name}</h2><div class='grid'>"
                + _table("FULL INBOUND", groups, fi)
                + _table("CURRENT STOCK", groups, cs)
                + _table(f"BOOKING &nbsp;{wk1_lbl}", groups, bk1)
                + _table(f"BOOKING &nbsp;{wk2_lbl}", groups, bk2)
                + _table(f"STOCK END WK &nbsp;{wk1_lbl}", groups, sew1, neg_ok=True)
                + _table(f"STOCK END WK &nbsp;{wk2_lbl}", groups, sew2, neg_ok=True)
                + "</div>")

    # RF SEASONAL mini-table
    rf_rows = []
    for y in RF_SEASONAL_YEARS:
        b = [rf[("BKK", y, c)] for c in ("CARRIER", "DAIKIN", "TRMK")]
        l = [rf[("LCH", y, c)] for c in ("CARRIER", "DAIKIN", "TRMK")]
        rf_rows.append("<tr><td>{}</td>{}{}</tr>".format(
            y, "".join(f"<td>{x or ''}</td>" for x in b),
            "".join(f"<td>{x or ''}</td>" for x in l)))
    rf_html = (
        "<div class='card' style='margin-top:22px'><h3>RF SEASONAL &mdash; 45RE by year built</h3>"
        "<table><thead><tr><th>Year</th><th>BKK CAR</th><th>BKK DAI</th><th>BKK TRMK</th>"
        "<th>LCH CAR</th><th>LCH DAI</th><th>LCH TRMK</th></tr></thead><tbody>"
        + "".join(rf_rows) + "</tbody></table></div>")

    html = f"""<!doctype html><html lang="th"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>HAL Stock HQ &ndash; Daily {report_date:%Y-%m-%d}</title>
<style>{DASH_CSS}</style></head><body><div class="wrap">
<h1>HAL Stock HQ &ndash; Daily Container Stock</h1>
<p class="sub">รายงานประจำวันที่ {report_date:%d/%m/%Y} ({report_date:%A})</p>
<div class="meta">
<span>WK 1ST: <b>{wk1_lbl}</b></span>
<span>WK 2ND: <b>{wk2_lbl}</b></span>
<span>Excel: <b>{xlsx_name}</b></span>
<span>สร้างเมื่อ <b>{dt.datetime.now():%Y-%m-%d %H:%M}</b></span>
</div>
<div class="alert"><h2>&#9888; ช่อง STOCK END WK ที่ติดลบ</h2>{alert_html}</div>
{side_block("BKK side", BKK_GROUPS)}
{side_block("LCH side", LCH_GROUPS)}
{rf_html}
<footer>สร้างอัตโนมัติจาก build_stock_hq.py &middot; ตัวเลขเป็นยอดรวมรายกลุ่มสถานที่ (ไม่มีชื่อลูกค้า/เลขบุ๊คกิ้ง)</footer>
</div></body></html>"""
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)


# --------------------------------------------------------------------------- #
#  LibreOffice head-less recalculation (optional)                              #
# --------------------------------------------------------------------------- #
def _find_soffice():
    for c in ("soffice", "libreoffice"):
        p = shutil.which(c)
        if p:
            return p
    for p in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.exists(p):
            return p
    return None


def recalc(path):
    """Refresh cached formula values in-place via LibreOffice, if available."""
    soffice = _find_soffice()
    if not soffice:
        print("  (LibreOffice not found - formulas will recalculate when Excel opens the file)")
        return
    outdir = os.path.dirname(path) or "."
    subprocess.run(
        [soffice, "--headless", "--convert-to",
         "xlsx:Calc MS Excel 2007 XML", "--outdir", outdir, path],
        check=True, capture_output=True,
    )
    print("  LibreOffice recalc done")


# --------------------------------------------------------------------------- #
#  File auto-detection / CLI                                                    #
# --------------------------------------------------------------------------- #
def _autodetect():
    found = {"actual": None, "bkg": None, "staying": None}
    for f in glob.glob(os.path.join(HERE, "*.xls")) + glob.glob(os.path.join(HERE, "*.xlsx")):
        name = os.path.basename(f).upper()
        if "STOCK HQ" in name or name.startswith("(HAL)"):
            continue
        if "ACTUAL" in name:
            found["actual"] = found["actual"] or f
        elif "STAY" in name:
            found["staying"] = found["staying"] or f
        elif "BKG" in name or "BKGPD" in name or "PD" in name:
            found["bkg"] = found["bkg"] or f
    return found


def main(argv=None):
    ap = argparse.ArgumentParser(description="Build the HAL Stock HQ daily report.")
    ap.add_argument("date", nargs="?", help="report date YYYY-MM-DD (default: today)")
    ap.add_argument("--actual"); ap.add_argument("--bkg"); ap.add_argument("--staying")
    ap.add_argument("--date", dest="date_opt")
    ap.add_argument("--out")
    ap.add_argument("--dashboard-dir", default=os.path.join(HERE, "dashboard-public"),
                    help="local clone of the public dashboard repo (default: ./dashboard-public)")
    ap.add_argument("--publish", action="store_true",
                    help="copy index.html into --dashboard-dir and git commit + push it")
    a = ap.parse_args(argv)

    det = _autodetect()
    actual = a.actual or det["actual"]
    bkg = a.bkg or det["bkg"]
    staying = a.staying or det["staying"]
    missing = [n for n, v in (("ACTUAL", actual), ("BKG+PD", bkg), ("STAYING", staying)) if not v]
    if missing:
        raise SystemExit("Missing source file(s): " + ", ".join(missing) +
                         "  - pass them with --actual / --bkg / --staying")

    date_str = a.date_opt or a.date
    report_date = (dt.datetime.strptime(date_str, "%Y-%m-%d").date()
                   if date_str else dt.date.today())

    out = a.out or os.path.join(HERE, f"HAL Stock HQ - Daily {report_date:%Y-%m-%d}.xlsx")

    print("Sources:")
    print("  ACTUAL :", os.path.basename(actual))
    print("  BKG+PD :", os.path.basename(bkg))
    print("  STAYING:", os.path.basename(staying))
    build(actual, bkg, staying, report_date, out)

    if a.publish:
        publish_dashboard(os.path.join(HERE, "index.html"), a.dashboard_dir, report_date)


def publish_dashboard(src_html, dash_dir, report_date):
    """Copy the freshly built index.html into the public dashboard repo and push."""
    if not os.path.isdir(os.path.join(dash_dir, ".git")):
        print(f"  publish skipped: {dash_dir} is not a git repo "
              f"(clone https://github.com/sirichai1265/hal-stock-hq-dashboard there first)")
        return
    shutil.copyfile(src_html, os.path.join(dash_dir, "index.html"))
    subprocess.run(["git", "-C", dash_dir, "add", "index.html"], check=True)
    r = subprocess.run(["git", "-C", dash_dir, "commit", "-m",
                        f"dashboard {report_date:%Y-%m-%d}"], capture_output=True, text=True)
    if r.returncode and "nothing to commit" in (r.stdout + r.stderr):
        print("  publish: no dashboard change")
        return
    subprocess.run(["git", "-C", dash_dir, "push", "-q"], check=True)
    print("  published -> https://sirichai1265.github.io/hal-stock-hq-dashboard/")


if __name__ == "__main__":
    main()

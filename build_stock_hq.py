"""
HAL Stock HQ - Daily report builder.

Reads the three same-day source files and fills the Daily sheet (plus the
RF SEASONAL rows and the NEW FORMAT zero-hiding / balance colours) of the
template workbook, producing  "Stock Daily <YYYY-MM-DD>.xlsx".

Usage:
    python build_stock_hq.py                     # auto-detect the 3 files, report date = today
    python build_stock_hq.py 2026-09-10          # force the report date
    python build_stock_hq.py --actual 9-10-ACTUAL.xls --bkg 9-10-BKG+PD.xls \
                             --staying 9-10-STAYING.xls --date 2026-09-10

Source files (same business day):
    * ACTUAL   - snapshot of FULL containers on hand      -> FULL INBOUND
    * STAYING  - per-container empty-stock snapshot        -> CURRENT STOCK / RF SEASONAL
    * BKG+PD   - live pickup list of not-yet-collected BKs -> BOOKING wk1 / wk2
    * EP2      - optional P.O.D vessel arrivals list       -> REPO (E/P) row (row 14),
                 grouped by P.O.D (THBKK/THLCH) x container type. Not every day has one;
                 the row is simply left blank when no EP2 file is found/passed.
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
RF_SEASONAL_YEARS = [2026, 2025, 2024, 2023, 2022, 2021, 2020]
RF_SEASONAL_YEAR_ROW = {2026: 4, 2025: 5, 2024: 6, 2023: 7, 2022: 8, 2021: 9, 2020: 10}
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
    """Outstanding-pickup list. Two schema variants seen so far, both handled here:
    - BKG+PD.xls: columns GP22/GP42/.., TRAN DT as a YYYYMMDD float, footer row has
      every field (including BK No) blank.
    - *PENDING*.xlsx ("Daily Booking" sheet): columns "GP22 Remaining"/.., TRAN DT
      already a real date, footer row has BK No == 'TOTAL' but Pickup still blank.
    Never hardcode the exact filename/schema - match by content (see module docstring)."""
    df = load_data_sheet(path, header=0)
    if "Pickup" not in df.columns:                       # a title row sat above the header
        df = load_data_sheet(path, header=1)
    df = df[df["Pickup"].notna()].copy()                 # drop the trailing total/footer row
                                                          # (works whether or not BK No is blank there)
    rename = {f"{src} Remaining": src for src in BKG_TYPE_COL
              if f"{src} Remaining" in df.columns and src not in df.columns}
    if rename:
        df = df.rename(columns=rename)
    if pd.api.types.is_numeric_dtype(df["TRAN DT"]):
        df["_date"] = pd.to_datetime(
            df["TRAN DT"].astype("Int64").astype(str), format="%Y%m%d", errors="coerce"
        )
    else:
        df["_date"] = pd.to_datetime(df["TRAN DT"], errors="coerce")
    return df


def load_booking_all(paths):
    """Load and combine one or more outstanding-pickup-list files. Starting
    2026-09-15 the export comes split in two: a backlog/overdue view
    (*PENDING*.xlsx, every TRAN DT before today) plus a forward view covering
    the next ~3 weeks (*BKG-3WK*.xls, every TRAN DT from today on) - together
    they cover the same ground the old single BKG+PD.xls file did. A BK No
    appearing in more than one file (seen rarely, e.g. a rescheduled pickup)
    keeps the version from whichever file was passed LAST."""
    frames = [load_booking(p) for p in paths]
    df = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]
    dupes = df.loc[df["BK No"].duplicated(keep=False), "BK No"].unique()
    if len(dupes):
        print(f"  NOTE: {len(dupes)} BK No appear in more than one booking file - "
              f"keeping the later file's version: {list(dupes)[:5]}")
        df = df.drop_duplicates(subset="BK No", keep="last")
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


def load_ep2(path):
    """P.O.D vessel-arrivals list -> REPO (E/P) row. Container-type columns already
    match TEMPLATE_TYPES order 1:1 (22GP, 42GP, 45GP, 22RE, 45RE, 22UT, 42UT, 22PC, 42PC)."""
    df = load_data_sheet(path, header=0)
    if "P.O.D" not in df.columns:
        raise SystemExit(
            f"{os.path.basename(path)}: expected a 'P.O.D' column - "
            f"columns found: {list(df.columns)}"
        )
    return df[df["P.O.D"].notna()].copy()


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


def repo_ep_totals(ep2):
    """REPO (E/P) row: EP2 rows summed by P.O.D (THBKK/THLCH) x container type.
    ep2=None (no file that day) -> all zeros, row 14 just stays blank."""
    out = {"BKK": _empty_counts(), "LCH": _empty_counts()}
    if ep2 is None:
        return out
    pod_side = {"THBKK": "BKK", "THLCH": "LCH"}
    for pod, side in pod_side.items():
        sub = ep2[ep2["P.O.D"] == pod]
        for t in TEMPLATE_TYPES:
            if t in sub.columns:
                out[side][t] = int(pd.to_numeric(sub[t], errors="coerce").fillna(0).sum())
    unknown = sorted(set(ep2["P.O.D"].astype(str)) - set(pod_side))
    if unknown:
        print(f"  WARNING: EP2 - unmapped P.O.D value(s) {unknown}, not counted in REPO (E/P)")
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


def build(actual_path, bkg_paths, staying_path, report_date, out_path, ep2_path=None):
    if isinstance(bkg_paths, str):
        bkg_paths = [bkg_paths]
    actual = load_actual(actual_path)
    stay = load_staying(staying_path)
    bkg = load_booking_all(bkg_paths)
    ep2 = load_ep2(ep2_path) if ep2_path else None

    monday = report_date - dt.timedelta(days=report_date.weekday())
    sunday = monday + dt.timedelta(days=6)
    next_monday = monday + dt.timedelta(days=7)
    next_sunday = next_monday + dt.timedelta(days=6)

    # Friday: WK1ST would only span Fri-Sun (3 days) - too short to stand alone, so
    # it merges with WK2ND into one combined backlog-through-next-Sunday window, and
    # the second STOCK BALANCE END WK slot shifts forward to use WK3RD's booking
    # window instead (the week after next) rather than sitting blank - per sirichai
    # 2026-09-18 then 2026-09-19.
    merge_weeks = report_date.weekday() == 4   # Monday=0 .. Friday=4
    wk1_hi = next_sunday if merge_weeks else sunday
    if merge_weeks:
        wk2_lo = next_monday + dt.timedelta(days=7)    # WK3RD monday
        wk2_hi = wk2_lo + dt.timedelta(days=6)          # WK3RD sunday
    else:
        wk2_lo, wk2_hi = next_monday, next_sunday       # WK2ND, as usual

    fi = full_inbound(actual)
    cs = current_stock(stay)
    bk1 = booking(bkg, None, wk1_hi)                       # TODAY+WK1ST: no lower bound
    bk2 = booking(bkg, wk2_lo, wk2_hi)                     # WK2ND, or WK3RD on Fridays
    rf = rf_seasonal_counts(stay)
    repo = repo_ep_totals(ep2)                              # REPO (E/P), row 14 - optional EP2 file

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

    # REPO (E/P) row 14 - not tied to a location group, so no colour_row() call;
    # the template already carries its own preset font colour on this row.
    write_row(ws, 14, BKK_COLS_STOCK, repo["BKK"])
    write_row(ws, 14, LCH_COLS_STOCK, repo["LCH"])

    # hide a bare 0 on every STOCK END WK formula cell, full block height
    for row in list(range(31, 41)) + list(range(45, 55)):
        hide_zero_format(ws, row, BKK_COLS_STOCK)
        hide_zero_format(ws, row, LCH_COLS_STOCK)

    # report date + d/m/yyyy display
    ws["AL1"] = dt.datetime(report_date.year, report_date.month, report_date.day)
    ws["R2"].number_format = "d/m/yyyy"
    ws["AM2"].number_format = "d/m/yyyy"

    # bare date-range labels (no "TODAY+WK 1ST" / "WK 2ND" text)
    wk1_lbl = fmt_range(report_date, wk1_hi)               # spans both weeks when merged
    wk2_lbl = fmt_range(wk2_lo, wk2_hi)                     # WK2ND, or WK3RD on Fridays
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
    # Balance = CURRENT STOCK total (incl. REPO E/P row 14, like Daily!L15/AG15) +
    #           FULL INBOUND total - BOOKING wk1 total - BOOKING wk2 total
    bal = {
        "BKK": {t: sum(cs[g][t] for g in BKK_GROUPS) + repo["BKK"][t] + sum(fi[g][t] for g in BKK_GROUPS)
                - sum(bk1[g][t] for g in BKK_GROUPS) - sum(bk2[g][t] for g in BKK_GROUPS)
                for t in TEMPLATE_TYPES},
        "LCH": {t: sum(cs[g][t] for g in LCH_GROUPS) + repo["LCH"][t] + sum(fi[g][t] for g in LCH_GROUPS)
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
                    fi, cs, bk1, bk2, sew1, sew2, bal, rf, repo,
                    os.path.basename(out_path))
    print(f"Dashboard   : {os.path.basename(dash)}")

    # ---- console report -------------------------------------------------- #
    print(f"\nReport date : {report_date:%Y-%m-%d} ({report_date:%A})")
    print(f"WK 1ST      : {wk1_lbl}   (TRAN DT <= {wk1_hi:%d/%m}"
          f"{'  -- merged with WK2ND (Friday)' if merge_weeks else ''})")
    print(f"WK 2ND      : {wk2_lbl}   ({wk2_lo:%d/%m} .. {wk2_hi:%d/%m}"
          f"{'  -- this is WK3RD (Friday)' if merge_weeks else ''})")
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
# NOTE: this must stay byte-for-byte in sync with what's live on the public
# dashboard repo (hal-stock-hq-dashboard) - this is Numtarn's merged-table /
# filter-buttons / live-clock redesign (commits 9f9d7c8..a074448, 2026-09-14),
# reused verbatim (CSS + JS) with our own data plugged in. If it gets
# redesigned again, re-sync from the live site rather than silently reverting.
DASH_CSS = """
:root{--bg-a:#eaf1fb;--bg-b:#f8f5f2;--card:#fff;--ink:#1c2530;--mut:#66738a;--line:#e3e8f0;
--depot:#1a56db;--neg:#c0192b;--negbg:#fff1f0;--head:#0d2b46;--accent:#0f6fff;--accent2:#d0202f;--chip:#eaf2ff;
--grid:#c7ccd6;--headbg:#eef0f4}
*{box-sizing:border-box}body{margin:0;color:var(--ink);
background:linear-gradient(160deg,var(--bg-a) 0%,#f7f9fc 40%,var(--bg-b) 100%) fixed;
font:14px/1.5 -apple-system,Segoe UI,Roboto,"Noto Sans Thai",sans-serif}
.accent-bar{height:5px;background:linear-gradient(90deg,var(--accent),var(--accent2))}
.wrap{max-width:1180px;margin:0 auto;padding:24px 20px 60px}
.topbar{display:flex;justify-content:space-between;align-items:center;gap:16px;
background:var(--card);border:1px solid var(--line);border-radius:14px;
padding:18px 24px;margin:0 0 18px;box-shadow:0 3px 10px rgba(16,24,40,.06)}
.topbar-right{display:flex;flex-direction:column;align-items:flex-end;gap:8px}
.topbar img{height:46px;width:auto;object-fit:contain}
h1{font-size:21px;margin:0 0 2px;color:var(--head)}.sub{color:var(--mut);margin:0}
.filter-bar{display:flex;gap:8px;margin:0 0 20px}
.filter-btn{cursor:pointer;border:1px solid var(--line);background:var(--card);color:var(--mut);
font-size:12.5px;font-weight:700;letter-spacing:.03em;padding:7px 18px;border-radius:999px;transition:all .15s}
.filter-btn:hover{border-color:var(--accent);color:var(--head)}
.filter-btn.active{color:#fff;border-color:transparent}
.filter-btn[data-filter="all"].active{background:var(--head)}
.filter-btn[data-filter="bkk"].active{background:var(--accent)}
.filter-btn[data-filter="lch"].active{background:var(--accent2)}
.hide{display:none!important}
.alert{background:linear-gradient(180deg,#fff6f5,#fff);border:1px solid var(--line);border-left:4px solid var(--neg);
border-radius:12px;padding:14px 16px;margin:30px 0 0;box-shadow:0 2px 8px rgba(16,24,40,.05)}
.alert h2{font-size:14px;margin:0 0 8px;color:var(--neg)}
.alert ul{margin:0;padding-left:18px}.alert li{margin:3px 0}
.alert .ok{color:var(--mut)}
.section{margin:0 0 28px}
.section h2{display:inline-block;font-size:12.5px;font-weight:700;letter-spacing:.03em;
margin:0 0 12px;color:var(--head);background:var(--chip);padding:5px 14px;border-radius:999px}
.grid{display:grid;gap:22px;grid-template-columns:1fr}
.card{background:var(--card);border:1px solid var(--line);border-top:3px solid var(--line);
border-radius:12px;padding:14px 16px 16px;overflow-x:auto;box-shadow:0 2px 8px rgba(16,24,40,.05)}
.card.bkk{border-top-color:var(--accent)}
.card.lch{border-top-color:var(--accent2)}
.card h3{font-size:13px;letter-spacing:.04em;text-transform:uppercase;color:var(--head);margin:0 0 10px}
.side-tag{display:inline-block;font-size:11px;font-weight:700;letter-spacing:.05em;
margin:0 0 10px;padding:3px 10px;border-radius:6px}
.side-tag.bkk{background:var(--chip);color:var(--accent)}
.side-tag.lch{background:#fdecec;color:var(--accent2)}
table{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums}
th,td{padding:5px 7px;text-align:right;border:1px solid var(--grid);white-space:nowrap}
th:first-child,td:first-child,th:nth-child(2),td:nth-child(2){text-align:left}
thead th{font-size:11px;color:var(--ink);font-weight:700;background:var(--headbg)}
thead tr:first-child th[colspan]{text-align:center;background:var(--headbg);color:var(--ink);
font-size:11px;text-transform:uppercase;letter-spacing:.03em;padding:6px 4px}
thead tr:first-child th.blk2{color:var(--depot)}
tr.total td{font-weight:700;background:var(--headbg);border-top:2px solid var(--grid)}
td.depot,th.depot{color:var(--depot)}
td.type{color:var(--ink);font-size:11.5px}
td.neg{color:var(--neg);font-weight:700;background:var(--negbg)}
.zero{color:var(--line)}
th.divider,td.divider{border-left:2px solid var(--grid)}
footer{margin-top:30px;color:var(--mut);font-size:12px;text-align:center}
"""


def _mcell(v, neg=False, divider=False, depot=False):
    cls = []
    if v == 0:
        cls.append("zero")
    elif neg and v < 0:
        cls.append("neg")
    if divider:
        cls.append("divider")
    if depot:
        cls.append("depot")
    cls_str = f' class="{" ".join(cls)}"' if cls else ""
    text = "·" if v == 0 else str(v)
    return f"<td{cls_str}>{text}</td>"


def _mrow(label, type_label, left_vals, right_vals, is_depot, right_neg_ok):
    lcls = ' class="depot"' if is_depot else ""
    left_cells = "".join(_mcell(int(left_vals[t]), depot=is_depot) for t in TEMPLATE_TYPES)
    right_cells = "".join(
        _mcell(int(right_vals[t]), neg=right_neg_ok, divider=(i == 0), depot=is_depot)
        for i, t in enumerate(TEMPLATE_TYPES)
    )
    return f"<tr><td{lcls}>{label}</td><td class='type'>{type_label}</td>{left_cells}{right_cells}</tr>"


def _mtotal(left_tot, right_tot):
    left_cells = "".join(f"<td>{left_tot[t] or ''}</td>" for t in TEMPLATE_TYPES)
    right_cells = "".join(
        f'<td class="divider">{right_tot[t] or ""}</td>' if i == 0 else f"<td>{right_tot[t] or ''}</td>"
        for i, t in enumerate(TEMPLATE_TYPES)
    )
    return f"<tr class='total'><td>TOTAL</td><td></td>{left_cells}{right_cells}</tr>"


def _merged_table(side_label, side_cls, groups, left, right, left_title, right_title,
                  right_neg_ok=False, extra_rows=None):
    """One BKK/LCH card: Location + Type + a 9-col left block + a 9-col right block
    (divider rule on the right block's first column) - e.g. FULL INBOUND | CURRENT
    STOCK, or BOOKING | STOCK BALANCE END WK. extra_rows: optional
    [(label, left_vals, right_vals, is_depot), ...] appended before TOTAL, e.g.
    REPO (E/P) which isn't tied to a location group."""
    left_head = "".join(f"<th>{t}</th>" for t in TEMPLATE_TYPES)
    right_head = (f'<th class="divider">{TEMPLATE_TYPES[0]}</th>'
                  + "".join(f"<th>{t}</th>" for t in TEMPLATE_TYPES[1:]))
    rows = []
    tot_l = {t: 0 for t in TEMPLATE_TYPES}
    tot_r = {t: 0 for t in TEMPLATE_TYPES}
    for g in groups:
        is_depot = GROUP_TYPE[g] == "DEPOT"
        lv, rv = left[g], right[g]
        for t in TEMPLATE_TYPES:
            tot_l[t] += int(lv[t]); tot_r[t] += int(rv[t])
        rows.append(_mrow(g, GROUP_TYPE[g], lv, rv, is_depot, right_neg_ok))
    for label, lv, rv, is_depot in (extra_rows or []):
        for t in TEMPLATE_TYPES:
            tot_l[t] += int(lv[t]); tot_r[t] += int(rv[t])
        rows.append(_mrow(label, "", lv, rv, is_depot, right_neg_ok))
    rows.append(_mtotal(tot_l, tot_r))
    return (f'<div class="card {side_cls}"><div class="side-tag {side_cls}">{side_label}</div>'
            f"<table><thead><tr><th rowspan='2'>Location</th><th rowspan='2'>Type</th>"
            f"<th colspan='9'>{left_title}</th><th colspan='9' class=\"divider blk2\">{right_title}</th></tr>"
            f"<tr>{left_head}{right_head}</tr></thead><tbody>{''.join(rows)}</tbody></table></div>")


def _msection(title, bkk_card, lch_card):
    return f"<div class='section'><h2>{title}</h2><div class='grid'>{bkk_card}{lch_card}</div></div>"


def write_dashboard(path, report_date, wk1_lbl, wk2_lbl,
                    fi, cs, bk1, bk2, sew1, sew2, bal, rf, repo, xlsx_name):
    repo_extra = {
        "BKK": [("REPO (E/P)", _empty_counts(), repo["BKK"], True)],
        "LCH": [("REPO (E/P)", _empty_counts(), repo["LCH"], True)],
    }
    sec1 = _msection(
        "FULL INBOUND &amp; CURRENT STOCK",
        _merged_table("BKK", "bkk", BKK_GROUPS, fi, cs, "FULL INBOUND", "CURRENT STOCK",
                      extra_rows=repo_extra["BKK"]),
        _merged_table("LCH", "lch", LCH_GROUPS, fi, cs, "FULL INBOUND", "CURRENT STOCK",
                      extra_rows=repo_extra["LCH"]),
    )
    sec2 = _msection(
        f"BOOKING &amp; STOCK BALANCE END WK &nbsp;{wk1_lbl}",
        _merged_table("BKK", "bkk", BKK_GROUPS, bk1, sew1,
                      f"BOOKING {wk1_lbl}", f"STOCK BALANCE END WK {wk1_lbl}", right_neg_ok=True),
        _merged_table("LCH", "lch", LCH_GROUPS, bk1, sew1,
                      f"BOOKING {wk1_lbl}", f"STOCK BALANCE END WK {wk1_lbl}", right_neg_ok=True),
    )
    sec3 = _msection(
        f"BOOKING &amp; STOCK BALANCE END WK &nbsp;{wk2_lbl}",
        _merged_table("BKK", "bkk", BKK_GROUPS, bk2, sew2,
                      f"BOOKING {wk2_lbl}", f"STOCK BALANCE END WK {wk2_lbl}", right_neg_ok=True),
        _merged_table("LCH", "lch", LCH_GROUPS, bk2, sew2,
                      f"BOOKING {wk2_lbl}", f"STOCK BALANCE END WK {wk2_lbl}", right_neg_ok=True),
    )
    sections_html = sec1 + sec2 + sec3

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
<style>{DASH_CSS}</style></head><body>
<div class="accent-bar"></div>
<div class="wrap">
<div class="topbar">
<div><h1>HAL Stock HQ &ndash; Daily Container Stock</h1>
<p class="sub">รายงานประจำวันที่ {report_date:%d/%m/%Y} ({report_date:%A})</p></div>
<div class="topbar-right">
<img src="logo.png" alt="logo">
</div>
</div>
<div class="filter-bar">
<button type="button" class="filter-btn active" data-filter="all">ALL</button>
<button type="button" class="filter-btn" data-filter="bkk">THBKK</button>
<button type="button" class="filter-btn" data-filter="lch">THLCH</button>
</div>
{sections_html}
{rf_html}
<footer>สร้างอัตโนมัติจาก build_stock_hq.py &middot; ตัวเลขเป็นยอดรวมรายกลุ่มสถานที่ (ไม่มีชื่อลูกค้า/เลขบุ๊คกิ้ง)</footer>
</div>
<script>
(function(){{
  var buttons = Array.prototype.slice.call(document.querySelectorAll('.filter-btn'));
  var bkkCards = document.querySelectorAll('.card.bkk');
  var lchCards = document.querySelectorAll('.card.lch');
  function apply(filter){{
    for (var i = 0; i < bkkCards.length; i++) {{
      bkkCards[i].classList.toggle('hide', filter === 'lch');
    }}
    for (var j = 0; j < lchCards.length; j++) {{
      lchCards[j].classList.toggle('hide', filter === 'bkk');
    }}
    buttons.forEach(function(b){{ b.classList.toggle('active', b.dataset.filter === filter); }});
    try {{ localStorage.setItem('halStockFilter', filter); }} catch (e) {{}}
  }}
  buttons.forEach(function(b){{
    b.addEventListener('click', function(){{ apply(b.dataset.filter); }});
  }});
  var saved = 'all';
  try {{ saved = localStorage.getItem('halStockFilter') || 'all'; }} catch (e) {{}}
  apply(saved);
}})();
</script>
</body></html>"""
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
def _sniff_role(path):
    """Fallback for a file whose name matches none of the keywords below (or was
    misspelled, e.g. 9-16-ATCUAL.xls) - peek at its columns instead. ACTUAL and
    STAYING share an identical column set, so they're told apart by Full/Empty:
    ACTUAL is always a snapshot of FULL containers only (every row 'F'), while
    STAYING is the empty-stock snapshot (every row 'E')."""
    try:
        for header in (0, 1):
            df = load_data_sheet(path, header=header)
            cols = set(df.columns)
            if {"BK No", "Pickup", "TRAN DT"} <= cols:
                return "bkg"
            if "P.O.D" in cols and "BK No" not in cols:
                return "ep2"
            if {"Move Code", "Size/Type", "Full/Empty", "Location"} <= cols:
                fe = df["Full/Empty"].dropna().astype(str).str.upper()
                if len(fe) and (fe == "F").all():
                    return "actual"
                if len(fe):
                    return "staying"
    except Exception:
        pass
    return None


def _autodetect():
    """bkg is a LIST - the booking/pickup export sometimes comes split across more
    than one file (backlog view + forward view, see load_booking_all); every
    matching candidate is collected and later combined, not just the newest.

    That makes it essential to first narrow the candidate pool down to today's
    upload batch (same calendar day as the most-recently-modified file in the
    folder) - otherwise a stale BKG+PD.xls left over from a previous day would
    get swept in alongside the current one and double-count old bookings."""
    found = {"actual": None, "bkg": [], "staying": None, "ep2": None}
    cand = glob.glob(os.path.join(HERE, "*.xls")) + glob.glob(os.path.join(HERE, "*.xlsx"))
    cand = [f for f in cand
            if "STOCK HQ" not in os.path.basename(f).upper()
            and "STOCK DAILY" not in os.path.basename(f).upper()
            and not os.path.basename(f).upper().startswith("(HAL)")]
    if not cand:
        return found
    cand.sort(key=os.path.getmtime, reverse=True)   # newest upload wins (for the single-file roles)
    latest_day = dt.date.fromtimestamp(os.path.getmtime(cand[0]))
    cand = [f for f in cand if dt.date.fromtimestamp(os.path.getmtime(f)) == latest_day]
    unclaimed = []
    for f in cand:
        name = os.path.basename(f).upper()
        if "ACTUAL" in name or "ATCUAL" in name:  # ATCUAL: seen typo'd 2026-09-16
            found["actual"] = found["actual"] or f
        elif "STAY" in name:
            found["staying"] = found["staying"] or f
        elif "EP2" in name or "EMPTY" in name:
            found["ep2"] = found["ep2"] or f
        elif "BKG" in name or "BKGPD" in name or "PD" in name or "PENDING" in name:
            found["bkg"].append(f)
        else:
            unclaimed.append(f)
    for f in unclaimed:                              # content sniff, name-based pass found nothing
        if found["bkg"] and found["ep2"] and found["actual"] and found["staying"]:
            break
        role = _sniff_role(f)
        if role == "bkg":
            found["bkg"].append(f)
        elif role == "ep2" and not found["ep2"]:
            found["ep2"] = f
        elif role == "actual" and not found["actual"]:
            found["actual"] = f
        elif role == "staying" and not found["staying"]:
            found["staying"] = f
    return found


def main(argv=None):
    ap = argparse.ArgumentParser(description="Build the HAL Stock HQ daily report.")
    ap.add_argument("date", nargs="?", help="report date YYYY-MM-DD (default: today)")
    ap.add_argument("--actual"); ap.add_argument("--staying")
    ap.add_argument("--bkg", action="append",
                    help="booking/pickup file - repeat if the export is split "
                         "(e.g. --bkg backlog.xlsx --bkg forward.xls)")
    ap.add_argument("--ep2", help="optional P.O.D file for the REPO (E/P) row (row 14)")
    ap.add_argument("--date", dest="date_opt")
    ap.add_argument("--out")
    ap.add_argument("--dashboard-dir", default=os.path.join(HERE, "dashboard-public"),
                    help="local clone of the public dashboard repo (default: ./dashboard-public)")
    ap.add_argument("--publish", action="store_true",
                    help="copy index.html into --dashboard-dir and git commit + push it")
    a = ap.parse_args(argv)

    det = _autodetect()
    actual = a.actual or det["actual"]
    bkg = a.bkg or det["bkg"]          # list - one or more booking/pickup files
    staying = a.staying or det["staying"]
    ep2 = a.ep2 or det["ep2"]          # optional - a day with no EP2 file just skips row 14
    missing = [n for n, v in (("ACTUAL", actual), ("BKG+PD", bkg), ("STAYING", staying)) if not v]
    if missing:
        raise SystemExit("Missing source file(s): " + ", ".join(missing) +
                         "  - pass them with --actual / --bkg / --staying")

    date_str = a.date_opt or a.date
    report_date = (dt.datetime.strptime(date_str, "%Y-%m-%d").date()
                   if date_str else dt.date.today())

    out = a.out or os.path.join(HERE, f"Stock Daily {report_date:%Y-%m-%d}.xlsx")

    print("Sources:")
    print("  ACTUAL :", os.path.basename(actual))
    print("  BKG+PD :", ", ".join(os.path.basename(p) for p in bkg))
    print("  STAYING:", os.path.basename(staying))
    print("  EP2    :", os.path.basename(ep2) if ep2 else "(none - REPO (E/P) row left blank)")
    build(actual, bkg, staying, report_date, out, ep2_path=ep2)

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

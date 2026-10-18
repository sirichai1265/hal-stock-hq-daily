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
    "PAT": "TERMINAL", "UNITHAI": "TERMINAL", "ESCO #2": "ICD", "TIFFA": "ICD",
    "SMART": "DEPOT", "BC2": "DEPOT", "BMTP": "ICD", "SCT": "ICD",
    "LCMT": "TERMINAL", "ESCO": "TERMINAL", "LCIT": "TERMINAL", "HPT": "TERMINAL",
    "HAST": "DEPOT", "CELLO": "DEPOT", "PW": "DEPOT",
}
GROUP_CODES = {
    "PAT": ["BKK01", "BKK04"], "UNITHAI": ["BKK02"], "ESCO #2": ["LCH55"],
    "TIFFA": ["LCH54"], "SMART": ["BKK25"], "BC2": ["BKK27"],
    "BMTP": ["BKKY4"], "SCT": ["LCHY4"],
    "LCMT": ["LCH04", "LCH07"], "ESCO": ["LCH01"], "LCIT": ["LCH05", "LCH10"],
    "HPT": ["LCH02", "LCH06", "LCH08", "LCH09"],
    "HAST": ["LCH27"], "CELLO": ["LCH28"], "PW": ["LCHY5"],
}

# template row numbers per group
# BKK order (per sirichai 2026-10-12): TIFFA (LCH54) takes BMTP's old slot; BMTP and
# the new SCT (LCHY4) sit below BC2 in the template's spare rows 10-11.
BKK_STOCK_ROW = {"PAT": 4, "UNITHAI": 5, "ESCO #2": 6, "TIFFA": 7, "SMART": 8, "BC2": 9,
                 "BMTP": 10, "SCT": 11}
LCH_STOCK_ROW = {"LCMT": 4, "ESCO": 5, "LCIT": 6, "HPT": 7, "HAST": 8, "CELLO": 9, "PW": 10}
BKK_BKG_ROW_WK1 = {"PAT": 31, "UNITHAI": 32, "ESCO #2": 33, "TIFFA": 34, "SMART": 35, "BC2": 36,
                   "BMTP": 37, "SCT": 38}
LCH_BKG_ROW_WK1 = {"LCMT": 31, "ESCO": 32, "LCIT": 33, "HPT": 34, "HAST": 35, "CELLO": 36, "PW": 37}
BKK_BKG_ROW_WK2 = {g: r + 14 for g, r in BKK_BKG_ROW_WK1.items()}   # 45-52
LCH_BKG_ROW_WK2 = {g: r + 14 for g, r in LCH_BKG_ROW_WK1.items()}   # 45-51

BKK_GROUPS = ["PAT", "UNITHAI", "ESCO #2", "TIFFA", "SMART", "BC2", "BMTP", "SCT"]
LCH_GROUPS = ["LCMT", "ESCO", "LCIT", "HPT", "HAST", "CELLO", "PW"]
# LCH terminal groups whose empty-container STOCK BALANCE END WK cells are left
# blank in the Excel (formulas removed, so the block TOTAL excludes them too) -
# per sirichai 2026-10-05.
NO_BALANCE_GROUPS = {"LCMT", "ESCO", "LCIT", "HPT"}

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
    """REPO (E/P) row: EP2/EMPTY rows summed by P.O.D (THBKK/THLCH) x container type.
    ep2=None (no file that day) -> all zeros, row 14 just stays blank.

    Two schema variants seen so far (same "match by content" rule as load_booking):
    - the original: one row per vessel/voyage, with container-type COUNT columns
      (22GP, 42GP, ...) already matching TEMPLATE_TYPES 1:1.
    - since 2026-09-18: one row per CONTAINER, with a 'Type Size' column holding
      the type as text (e.g. '22GP') - count rows instead of summing columns."""
    out = {"BKK": _empty_counts(), "LCH": _empty_counts()}
    if ep2 is None:
        return out
    pod_side = {"THBKK": "BKK", "THLCH": "LCH"}
    per_container = "Type Size" in ep2.columns and not any(t in ep2.columns for t in TEMPLATE_TYPES)
    for pod, side in pod_side.items():
        sub = ep2[ep2["P.O.D"] == pod]
        if per_container:
            for t, n in sub["Type Size"].value_counts().items():
                if t in out[side]:
                    out[side][t] = int(n)
        else:
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


def clear_row(ws, row, col_letters):
    for col in col_letters:
        ws[f"{col}{row}"] = None


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
            if g in NO_BALANCE_GROUPS:
                clear_row(ws, r1[g], cstock)
            else:
                style_negative(ws, r1[g], cstock, sew1[g])
            # BOOKING wk2 (rows 45-50/45-51) + its STOCK END WK formula cells
            write_row(ws, r2[g], cin, bk2[g])
            colour_row(ws, r2[g], cin, g)
            colour_row(ws, r2[g], cstock, g)
            if g in NO_BALANCE_GROUPS:
                clear_row(ws, r2[g], cstock)
            else:
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
    # block headers carry their week slot as a suffix (per sirichai 2026-10-05):
    # row 29 block = WK1ST, row 43 block = WK2ND
    for row, wk in ((29, "WK1ST"), (43, "WK2ND")):
        ws[f"C{row}"] = f"BOOKING (BKK) {wk}"
        ws[f"L{row}"] = f"STOCK BALANCE END (BKK) {wk}"
        ws[f"X{row}"] = f"BOOKING (LCH) {wk}"
        ws[f"AG{row}"] = f"STOCK BALANCE END (LCH) {wk}"
    ws["I29"] = ws["AD29"] = wk1_lbl
    ws["I43"] = ws["AD43"] = wk2_lbl

    # --- RF SEASONAL rows 4-7 (row 8 TOTAL formula untouched) ---
    # Blank single 0s here too, except TOTAL - reverted back per sirichai 2026-09-22
    # (years 2022/2024/2026 dropped from the sheet entirely, see RF_SEASONAL_YEARS).
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

    # NEW FORMAT Balance colour: red when zero or negative (<= 0, per sirichai 2026-10-06), else blue.
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
            cell.font = Font(color=RED if bal[side][t] <= 0 else BLUE,
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
# Redesigned 2026-10-18 ("ทำให้สวยงาม เข้าใจง่ายกว่าเดิม"): KPI summary per zone,
# colour-coded blocks (FULL INBOUND green / CURRENT STOCK blue / BOOKING amber /
# STOCK BALANCE END navy), Thai one-line explanations, sticky toolbar with section
# links, sticky location column. Kept from Numtarn's version: logo + header photo,
# ALL/THBKK/THLCH filter, merged BKK/LCH tables with a Type column, REPO (E/P) row,
# RF SEASONAL table, title "Stock Report HQ". Still removed on purpose (earlier
# requests): live clock, 'created at' chip, WK1ST/WK2ND/Excel meta chips and the
# negative-alert box - negatives are shown in red inside the tables.
DASH_CSS = """
:root{--navy:#0d2b46;--ink:#1b2735;--mut:#64748b;--line:#e2e8f0;--grid:#e3e9f2;--card:#fff;--bg:#eef3f9;
--bkk:#1d6fe8;--bkk-soft:#e7f0ff;--lch:#d72638;--lch-soft:#fdeaec;
--in:#0f8a5f;--in-soft:#e1f5ec;--in-soft2:#f1faf6;
--stock:#1d5fd1;--stock-soft:#e4edff;--stock-soft2:#f1f5ff;
--bkg:#b25e00;--bkg-soft:#fff1d6;--bkg-soft2:#fff9ec;
--bal:#0d2b46;--bal-soft:#e6ebf3;--bal-soft2:#f3f6fa;
--neg:#c0192b;--negbg:#ffe1e5;--depot:#1a56db;--zero:#c1cad8;--divider:#b9c5d6}
*{box-sizing:border-box}html{scroll-behavior:smooth}
body{margin:0;color:var(--ink);
background:radial-gradient(1100px 480px at 8% -8%,#dbe8fb 0,rgba(219,232,251,0) 62%),var(--bg);
font:14px/1.55 "Segoe UI","Leelawadee UI","Noto Sans Thai",Roboto,sans-serif}
.accent-bar{height:5px;background:linear-gradient(90deg,var(--bkk),var(--lch))}
.wrap{max-width:1240px;margin:0 auto;padding:22px 20px 56px}
.topbar{display:flex;justify-content:space-between;align-items:center;gap:16px;
background:linear-gradient(rgba(255,255,255,.88),rgba(255,255,255,.88)),url('header.jpg') center/cover;
border:1px solid var(--line);border-radius:18px;padding:22px 28px;margin:0 0 14px;
box-shadow:0 6px 20px rgba(15,35,60,.08)}
.topbar img{height:50px;width:auto;object-fit:contain}
h1{font-size:26px;margin:0 0 2px;color:var(--navy);font-weight:800;letter-spacing:-.01em}
.sub{color:var(--mut);margin:0;font-size:14px}
.toolbar{position:sticky;top:0;z-index:30;display:flex;flex-wrap:wrap;gap:10px 18px;align-items:center;
justify-content:space-between;margin:0 -10px 6px;padding:10px 10px;
background:rgba(238,243,249,.88);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);
border-bottom:1px solid var(--line)}
.filter-bar{display:flex;gap:6px}
.filter-btn{cursor:pointer;border:1px solid var(--line);background:#fff;color:var(--mut);
font-family:inherit;font-weight:700;font-size:12.5px;line-height:1;letter-spacing:.03em;
padding:9px 18px;border-radius:999px;transition:all .15s}
.filter-btn:hover{border-color:var(--bkk);color:var(--navy)}
.filter-btn.active{color:#fff;border-color:transparent}
.filter-btn[data-filter="all"].active{background:var(--navy)}
.filter-btn[data-filter="bkk"].active{background:var(--bkk)}
.filter-btn[data-filter="lch"].active{background:var(--lch)}
.nav{display:flex;flex-wrap:wrap;gap:4px}
.nav a{font-size:12.5px;font-weight:600;color:var(--mut);text-decoration:none;padding:7px 12px;border-radius:8px}
.nav a:hover{background:#fff;color:var(--navy)}
.nav a.on{background:var(--navy);color:#fff}
.hide{display:none!important}
.zone-title{display:flex;align-items:center;gap:10px;margin:16px 0 8px;color:var(--navy);font-weight:800;font-size:15px}
.side-tag{display:inline-block;font-weight:800;font-size:12.5px;letter-spacing:.04em;padding:4px 14px;border-radius:999px}
.side-tag.bkk{background:var(--bkk-soft);color:var(--bkk)}
.side-tag.lch{background:var(--lch-soft);color:var(--lch)}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
.kpi{background:var(--card);border:1px solid var(--line);border-top:4px solid var(--c);border-radius:14px;
padding:13px 16px;box-shadow:0 3px 10px rgba(15,35,60,.05)}
.kpi .lbl{font-size:12.5px;color:var(--mut);font-weight:600}
.kpi .num{font-size:30px;font-weight:800;color:var(--c);line-height:1.15;font-variant-numeric:tabular-nums;margin:2px 0}
.kpi .sub2{font-size:12px;color:var(--mut)}
.kpi.in{--c:var(--in)}.kpi.stock{--c:var(--stock)}.kpi.bkg{--c:var(--bkg)}
.legend{display:flex;flex-wrap:wrap;gap:6px 18px;font-size:12.5px;color:var(--mut);margin:18px 0 0;
padding:10px 14px;background:rgba(255,255,255,.7);border:1px dashed var(--divider);border-radius:12px}
.dot{display:inline-block;width:11px;height:11px;border-radius:3px;vertical-align:-1px;margin-right:6px}
.section{margin:34px 0 0;scroll-margin-top:72px}
.sec-head{display:flex;gap:12px;align-items:flex-start;margin:0 0 14px}
.step{flex:none;width:32px;height:32px;border-radius:50%;background:var(--navy);color:#fff;font-weight:800;
display:grid;place-items:center;font-size:15px}
.sec-head h2{margin:0;font-size:19px;color:var(--navy);display:flex;flex-wrap:wrap;gap:4px 10px;align-items:center}
.range{font-size:12.5px;font-weight:700;color:var(--bkg);background:var(--bkg-soft);padding:3px 11px;border-radius:999px}
.sec-head p{margin:3px 0 0;color:var(--mut);font-size:13px;max-width:900px}
.grid{display:grid;gap:18px;grid-template-columns:1fr}
.card{background:var(--card);border:1px solid var(--line);border-radius:16px;box-shadow:0 4px 14px rgba(15,35,60,.06);overflow:hidden}
.card.bkk{border-top:4px solid var(--bkk)}
.card.lch{border-top:4px solid var(--lch)}
.card-head{padding:12px 16px 2px}
.tbl-wrap{overflow-x:auto;padding:6px 10px 14px}
table{border-collapse:separate;border-spacing:0;width:100%;font-variant-numeric:tabular-nums;font-size:13px}
th,td{padding:7px 9px;text-align:right;white-space:nowrap;border-bottom:1px solid var(--grid)}
th.loc,td.loc{text-align:left}
td.loc{font-weight:700;position:sticky;left:0;z-index:1;background:#fff}
td.type{text-align:left}
thead th{font-size:11.5px;font-weight:700;color:#475569;background:#f4f7fb}
thead tr.top th{text-align:center;font-size:12px;letter-spacing:.05em;text-transform:uppercase;padding:9px 6px}
th.b-in{background:var(--in-soft);color:var(--in)}
th.b-stock{background:var(--stock-soft);color:var(--stock)}
th.b-bkg{background:var(--bkg-soft);color:var(--bkg)}
th.b-bal{background:var(--bal-soft);color:var(--bal)}
thead tr.sub th.b-in{background:var(--in-soft2)}
thead tr.sub th.b-stock{background:var(--stock-soft2)}
thead tr.sub th.b-bkg{background:var(--bkg-soft2)}
thead tr.sub th.b-bal{background:var(--bal-soft2)}
th.divider,td.divider{border-left:2px solid var(--divider)}
tbody tr:nth-child(even) td{background:#fafcfe}
tbody tr:hover td{background:#eaf2ff}
td.depot{color:var(--depot)}
td.neg{color:var(--neg)!important;font-weight:800;background:var(--negbg)!important}
.zero{color:var(--zero)}
.pill{display:inline-block;font-size:10.5px;font-weight:700;letter-spacing:.03em;padding:2px 9px;border-radius:999px}
.t-terminal{background:#e8eefc;color:#2a4a99}
.t-icd{background:#efe8fb;color:#5b3aa6}
.t-depot{background:#e1f0ff;color:#0b5ed7}
tr.total td{font-weight:800;color:var(--navy);background:#eaeff6!important;border-top:2px solid #b8c4d6;border-bottom:none}
.rf{max-width:760px}
.rf tbody td:first-child{text-align:left;font-weight:700}
.rf th.zone-bkk{background:var(--bkk-soft);color:var(--bkk)}
.rf th.zone-lch{background:var(--lch-soft);color:var(--lch)}
footer{margin-top:34px;color:var(--mut);font-size:12px;text-align:center}
@media(max-width:900px){.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}
.topbar{padding:16px 18px}h1{font-size:21px}.topbar img{height:40px}}
@media(max-width:600px){.topbar{flex-direction:column;align-items:flex-start}.topbar img{height:34px}
.toolbar{gap:6px;padding:8px 10px}.nav{flex-wrap:nowrap;overflow-x:auto;width:100%}
.nav a{white-space:nowrap;padding:6px 10px}.kpi .num{font-size:26px}.wrap{padding:14px 12px 40px}}
"""

DASH_JS = """
(function(){
  var buttons = Array.prototype.slice.call(document.querySelectorAll('.filter-btn'));
  var zoned = document.querySelectorAll('[data-zone]');
  function apply(filter){
    for (var i = 0; i < zoned.length; i++) {
      var z = zoned[i].getAttribute('data-zone');
      zoned[i].classList.toggle('hide', filter !== 'all' && z !== filter);
    }
    buttons.forEach(function(b){ b.classList.toggle('active', b.dataset.filter === filter); });
    try { localStorage.setItem('halStockFilter', filter); } catch (e) {}
  }
  buttons.forEach(function(b){
    b.addEventListener('click', function(){ apply(b.dataset.filter); });
  });
  var saved = 'all';
  try { saved = localStorage.getItem('halStockFilter') || 'all'; } catch (e) {}
  apply(saved);

  var links = Array.prototype.slice.call(document.querySelectorAll('.nav a'));
  var secs = links.map(function(a){ return document.querySelector(a.getAttribute('href')); });
  if ('IntersectionObserver' in window) {
    var io = new IntersectionObserver(function(entries){
      entries.forEach(function(e){
        if (e.isIntersecting) {
          links.forEach(function(a){ a.classList.toggle('on', a.getAttribute('href') === '#' + e.target.id); });
        }
      });
    }, {rootMargin: '-30% 0px -60% 0px'});
    secs.forEach(function(s){ if (s) io.observe(s); });
  }
})();
"""

ZONE_NAME = {"bkk": "THBKK · กรุงเทพ", "lch": "THLCH · แหลมฉบัง"}
TH_DAYS = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]
GP_TYPES = ("22GP", "42GP", "45GP")
RE_TYPES = ("22RE", "45RE")


def _num(v):
    return f"{int(v):,}"


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
    text = "·" if v == 0 else _num(v)
    return f"<td{cls_str}>{text}</td>"


def _type_pill(type_label):
    if not type_label:
        return ""
    return f'<span class="pill t-{type_label.lower()}">{type_label}</span>'


def _mrow(label, type_label, left_vals, right_vals, is_depot, right_neg_ok):
    lcls = "loc depot" if is_depot else "loc"
    left_cells = "".join(_mcell(int(left_vals[t]), depot=is_depot) for t in TEMPLATE_TYPES)
    right_cells = "".join(
        _mcell(int(right_vals[t]), neg=right_neg_ok, divider=(i == 0), depot=is_depot)
        for i, t in enumerate(TEMPLATE_TYPES)
    )
    return (f"<tr><td class='{lcls}'>{label}</td>"
            f"<td class='type'>{_type_pill(type_label)}</td>{left_cells}{right_cells}</tr>")


def _mtotal(left_tot, right_tot):
    left_cells = "".join(f"<td>{_num(left_tot[t]) if left_tot[t] else ''}</td>" for t in TEMPLATE_TYPES)
    right_cells = "".join(
        (f'<td class="divider">{_num(right_tot[t]) if right_tot[t] else ""}</td>' if i == 0
         else f"<td>{_num(right_tot[t]) if right_tot[t] else ''}</td>")
        for i, t in enumerate(TEMPLATE_TYPES)
    )
    return f"<tr class='total'><td class='loc'>TOTAL</td><td></td>{left_cells}{right_cells}</tr>"


def _merged_table(side_label, side_cls, groups, left, right, left_title, right_title,
                  right_neg_ok=False, extra_rows=None, left_cls="b-in", right_cls="b-stock"):
    """One BKK/LCH card: Location + Type + a 9-col left block + a 9-col right block
    (divider rule on the right block's first column), blocks colour-coded by
    left_cls / right_cls. extra_rows: optional [(label, left_vals, right_vals,
    is_depot), ...] appended before TOTAL, e.g. REPO (E/P)."""
    left_head = "".join(f'<th class="{left_cls}">{t}</th>' for t in TEMPLATE_TYPES)
    right_head = (f'<th class="divider {right_cls}">{TEMPLATE_TYPES[0]}</th>'
                  + "".join(f'<th class="{right_cls}">{t}</th>' for t in TEMPLATE_TYPES[1:]))
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
    return (f'<div class="card {side_cls}" data-zone="{side_cls}">'
            f'<div class="card-head"><span class="side-tag {side_cls}">{ZONE_NAME[side_cls]}</span></div>'
            f'<div class="tbl-wrap"><table><thead>'
            f"<tr class='top'><th rowspan='2' class='loc'>Location</th><th rowspan='2'>Type</th>"
            f"<th colspan='9' class='{left_cls}'>{left_title}</th>"
            f"<th colspan='9' class='divider {right_cls}'>{right_title}</th></tr>"
            f"<tr class='sub'>{left_head}{right_head}</tr></thead>"
            f"<tbody>{''.join(rows)}</tbody></table></div></div>")


def _msection(sec_id, step, title, desc, bkk_card, lch_card, rng=None):
    chip = f'<span class="range">{rng}</span>' if rng else ""
    return (f"<section class='section' id='{sec_id}'><div class='sec-head'><div class='step'>{step}</div>"
            f"<div><h2>{title}{chip}</h2><p>{desc}</p></div></div>"
            f"<div class='grid'>{bkk_card}{lch_card}</div></section>")


def _kpi_sub(data, groups, extra=None):
    """(grand total, 'GP x · RE y · อื่นๆ z') for `data` summed over `groups`."""
    tot = {t: sum(int(data[g][t]) for g in groups) + (int(extra[t]) if extra else 0)
           for t in TEMPLATE_TYPES}
    gp = sum(tot[t] for t in GP_TYPES)
    re_ = sum(tot[t] for t in RE_TYPES)
    other = sum(tot.values()) - gp - re_
    sub = f"GP {gp:,} · RE {re_:,}" + (f" · อื่นๆ {other:,}" if other else "")
    return sum(tot.values()), sub


def _kpi_card(kind, label, total, sub):
    return (f"<div class='kpi {kind}'><div class='lbl'>{label}</div>"
            f"<div class='num'>{total:,}</div><div class='sub2'>{sub}</div></div>")


def _kpi_zone(side_cls, groups, fi, cs, bk1, bk2, repo_side, wk1_lbl, wk2_lbl):
    stock_n, stock_sub = _kpi_sub(cs, groups, repo_side)
    if any(repo_side.values()):
        stock_sub += f" · รวม REPO (E/P) {sum(int(v) for v in repo_side.values()):,}"
    in_n, in_sub = _kpi_sub(fi, groups)
    b1_n, b1_sub = _kpi_sub(bk1, groups)
    b2_n, b2_sub = _kpi_sub(bk2, groups)
    cards = (_kpi_card("stock", "ตู้เปล่าในสต๊อก (CURRENT STOCK)", stock_n, stock_sub)
             + _kpi_card("in", "ตู้ FULL ขาเข้า รอ devanning (FULL INBOUND)", in_n, in_sub)
             + _kpi_card("bkg", f"บุ๊คกิ้ง WK1ST · {wk1_lbl}", b1_n, b1_sub)
             + _kpi_card("bkg", f"บุ๊คกิ้ง WK2ND · {wk2_lbl}", b2_n, b2_sub))
    return (f"<div data-zone='{side_cls}'><div class='zone-title'>"
            f"<span class='side-tag {side_cls}'>{ZONE_NAME[side_cls]}</span>ภาพรวมวันนี้</div>"
            f"<div class='kpis'>{cards}</div></div>")


def write_dashboard(path, report_date, wk1_lbl, wk2_lbl,
                    fi, cs, bk1, bk2, sew1, sew2, bal, rf, repo, xlsx_name):
    repo_extra = {
        "BKK": [("REPO (E/P)", _empty_counts(), repo["BKK"], True)],
        "LCH": [("REPO (E/P)", _empty_counts(), repo["LCH"], True)],
    }
    kpis_html = (_kpi_zone("bkk", BKK_GROUPS, fi, cs, bk1, bk2, repo["BKK"], wk1_lbl, wk2_lbl)
                 + _kpi_zone("lch", LCH_GROUPS, fi, cs, bk1, bk2, repo["LCH"], wk1_lbl, wk2_lbl))

    sec1 = _msection(
        "s-stock", 1, "สต๊อก &amp; ตู้ขาเข้า",
        "<b>FULL INBOUND</b> = ตู้ FULL ที่ถึงแล้ว รอ devanning &nbsp;·&nbsp; "
        "<b>CURRENT STOCK</b> = ตู้เปล่าคงเหลือในลาน (<b>REPO (E/P)</b> = ตู้เปล่าที่ย้ายเข้ามา)",
        _merged_table("BKK", "bkk", BKK_GROUPS, fi, cs, "FULL INBOUND", "CURRENT STOCK",
                      extra_rows=repo_extra["BKK"], left_cls="b-in", right_cls="b-stock"),
        _merged_table("LCH", "lch", LCH_GROUPS, fi, cs, "FULL INBOUND", "CURRENT STOCK",
                      extra_rows=repo_extra["LCH"], left_cls="b-in", right_cls="b-stock"),
    )
    sec2 = _msection(
        "s-wk1", 2, "บุ๊คกิ้ง &amp; สต๊อกคงเหลือ · WK1ST",
        "<b>BOOKING</b> = ตู้ที่ลูกค้าจองรับ (รวมบุ๊คที่ค้างรับ) &nbsp;·&nbsp; "
        "<b>STOCK BALANCE END</b> = ตู้เปล่าคงเหลือหลังหักบุ๊คกิ้ง "
        "<span style='color:var(--neg);font-weight:700'>ตัวเลขแดง = ไม่พอ (ติดลบ)</span>",
        _merged_table("BKK", "bkk", BKK_GROUPS, bk1, sew1, "BOOKING", "STOCK BALANCE END",
                      right_neg_ok=True, left_cls="b-bkg", right_cls="b-bal"),
        _merged_table("LCH", "lch", LCH_GROUPS, bk1, sew1, "BOOKING", "STOCK BALANCE END",
                      right_neg_ok=True, left_cls="b-bkg", right_cls="b-bal"),
        rng=wk1_lbl,
    )
    sec3 = _msection(
        "s-wk2", 3, "บุ๊คกิ้ง &amp; สต๊อกคงเหลือ · WK2ND",
        "หักบุ๊คกิ้งสัปดาห์ถัดไป <b>ต่อจาก WK1ST</b> &nbsp;·&nbsp; "
        "<b>STOCK BALANCE END</b> = สต๊อกคงเหลือของ WK1ST หักบุ๊ค WK2ND",
        _merged_table("BKK", "bkk", BKK_GROUPS, bk2, sew2, "BOOKING", "STOCK BALANCE END",
                      right_neg_ok=True, left_cls="b-bkg", right_cls="b-bal"),
        _merged_table("LCH", "lch", LCH_GROUPS, bk2, sew2, "BOOKING", "STOCK BALANCE END",
                      right_neg_ok=True, left_cls="b-bkg", right_cls="b-bal"),
        rng=wk2_lbl,
    )
    sections_html = sec1 + sec2 + sec3

    # RF SEASONAL - 45RE reefers by build year x brand
    brands = ("CARRIER", "DAIKIN", "TRMK")
    rf_rows = []
    tot = {(a, c): 0 for a in ("BKK", "LCH") for c in brands}
    for y in RF_SEASONAL_YEARS:
        cells = ""
        for a in ("BKK", "LCH"):
            for c in brands:
                v = rf[(a, y, c)]
                tot[(a, c)] += v
                cls = ' class="divider"' if (a == "LCH" and c == "CARRIER") else ""
                cells += f"<td{cls}>{v or ''}</td>"
        rf_rows.append(f"<tr><td>{y}</td>{cells}</tr>")
    tcells = "".join(
        (f'<td class="divider">{tot[(a, c)]}</td>' if (a == "LCH" and c == "CARRIER")
         else f"<td>{tot[(a, c)]}</td>")
        for a in ("BKK", "LCH") for c in brands)
    rf_rows.append(f"<tr class='total'><td>TOTAL</td>{tcells}</tr>")
    rf_card = (
        "<div class='card rf'><div class='tbl-wrap'><table><thead>"
        "<tr class='top'><th rowspan='2' class='loc'>ปีผลิต</th>"
        "<th colspan='3' class='zone-bkk'>THBKK</th><th colspan='3' class='divider zone-lch'>THLCH</th></tr>"
        "<tr class='sub'><th>CARRIER</th><th>DAIKIN</th><th>TRMK</th>"
        "<th class='divider'>CARRIER</th><th>DAIKIN</th><th>TRMK</th></tr></thead>"
        f"<tbody>{''.join(rf_rows)}</tbody></table></div></div>")
    rf_html = (
        "<section class='section' id='s-rf'><div class='sec-head'><div class='step'>4</div><div>"
        "<h2>RF SEASONAL <span class='range' style='color:var(--stock);background:var(--stock-soft)'>45RE</span></h2>"
        "<p>จำนวนตู้ reefer 45RE ในสต๊อก แยกตามปีผลิตและยี่ห้อเครื่องทำความเย็น "
        "(CARRIER · DAIKIN · TRMK = Thermo King)</p></div></div>"
        f"{rf_card}</section>")

    weekday = TH_DAYS[report_date.weekday()]
    html = f"""<!doctype html><html lang="th"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Stock Report HQ &ndash; {report_date:%Y-%m-%d}</title>
<style>{DASH_CSS}</style></head><body>
<div class="accent-bar"></div>
<div class="wrap">
<header class="topbar">
<div><h1>Stock Report HQ</h1>
<p class="sub">รายงานสต๊อกตู้ประจำวันที่ {report_date:%d/%m/%Y} (วัน{weekday})</p></div>
<img src="logo.png" alt="logo">
</header>
<div class="toolbar">
<div class="filter-bar">
<button type="button" class="filter-btn active" data-filter="all">ALL</button>
<button type="button" class="filter-btn" data-filter="bkk">THBKK</button>
<button type="button" class="filter-btn" data-filter="lch">THLCH</button>
</div>
<nav class="nav">
<a href="#s-stock">สต๊อก &amp; ตู้ขาเข้า</a><a href="#s-wk1">WK1ST</a><a href="#s-wk2">WK2ND</a><a href="#s-rf">RF SEASONAL</a>
</nav>
</div>
{kpis_html}
<div class="legend">
<span><i class="dot" style="background:var(--in)"></i>FULL INBOUND (ตู้ FULL)</span>
<span><i class="dot" style="background:var(--stock)"></i>CURRENT STOCK (ตู้เปล่า)</span>
<span><i class="dot" style="background:var(--bkg)"></i>BOOKING (ลูกค้าจอง)</span>
<span><i class="dot" style="background:var(--bal)"></i>STOCK BALANCE END (คงเหลือ)</span>
<span style="color:var(--depot)"><b>น้ำเงิน</b> = DEPOT</span>
<span style="color:var(--neg)"><b>แดง</b> = ติดลบ</span>
<span><b>·</b> = 0</span>
</div>
{sections_html}
{rf_html}
<footer>สร้างอัตโนมัติจาก build_stock_hq.py &middot; ตัวเลขเป็นยอดรวมรายกลุ่มสถานที่ (ไม่มีชื่อลูกค้า/เลขบุ๊คกิ้ง)</footer>
</div>
<script>{DASH_JS}</script>
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

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


# STAYING rows with these Move Codes are not counted (user rule 2026-10-19:
# "STAYING ไม่นับ OEP, OEV, OEW") - applies to CURRENT STOCK and RF SEASONAL.
STAYING_EXCLUDE_MOVES = {"OEP", "OEV", "OEW"}


def load_staying(path):
    df = load_data_sheet(path, header=0)
    df = df[df["Location"].notna()].copy()
    skip = df["Move Code"].astype(str).str.strip().str.upper().isin(STAYING_EXCLUDE_MOVES)
    if skip.any():
        counts = df.loc[skip, "Move Code"].astype(str).str.upper().value_counts().to_dict()
        print(f"  STAYING: excluded {int(skip.sum())} rows by Move Code {counts}")
    return df[~skip].copy()


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


ACTUAL_MAX_DAYS = 30


def full_inbound(actual):
    """FULL INBOUND: ACTUAL rows by group + type. User rules (2026-10-21):
    - Move Code OFD is not counted;
    - containers older than 30 days (Days > 30) are not counted;
    - IFV whose BKNO(Cur) starts with HASLS22 is not counted, but an IFV within
      30 days is counted (as IFI) - so in effect an IFV counts only while Days <= 30."""
    mc = actual["Move Code"].astype(str).str.strip().str.upper()
    days = pd.to_numeric(actual["Days"], errors="coerce").fillna(0) if "Days" in actual.columns else 0
    within = days <= ACTUAL_MAX_DAYS
    drop = (mc == "OFD") | ~within
    inb = actual[~drop].copy()
    n_old = int((~within & (mc != "OFD")).sum())
    n_ifv = int(((mc == "IFV") & ~drop).sum())
    if n_old or n_ifv:
        print(f"  ACTUAL: excluded {n_old} rows older than {ACTUAL_MAX_DAYS} days; "
              f"{n_ifv} IFV within {ACTUAL_MAX_DAYS} days counted as IFI")
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
    # Balance mirrors the template formula exactly (so colours and the dashboard show
    # what Excel shows): NEW FORMAT 'Current Stock' = Daily!L15/AG15, which already
    # includes REPO (E/P) row 14, and the 'E/P In' row adds Daily!L14/AG14 again - so
    # REPO is counted twice in the template's Balance (hence 2 * repo below).
    #   Balance = CURRENT STOCK total + E/P In + FULL INBOUND - BOOKING wk1 - BOOKING wk2
    bal = {
        "BKK": {t: sum(cs[g][t] for g in BKK_GROUPS) + 2 * repo["BKK"][t] + sum(fi[g][t] for g in BKK_GROUPS)
                - sum(bk1[g][t] for g in BKK_GROUPS) - sum(bk2[g][t] for g in BKK_GROUPS)
                for t in TEMPLATE_TYPES},
        "LCH": {t: sum(cs[g][t] for g in LCH_GROUPS) + 2 * repo["LCH"][t] + sum(fi[g][t] for g in LCH_GROUPS)
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
    write_overview(os.path.join(os.path.dirname(dash), "overview.html"), report_date,
                   wk1_lbl, wk2_lbl, fi, cs, bk1, bk2, bal, repo)
    print(f"Dashboard   : {os.path.basename(dash)} + overview.html")

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
# 2026-10-18: the dashboard now mirrors the Excel "Daily" sheet cell for cell
# (sirichai: "dashboard โชว์แบบชีท daily ใน excel"): same title row + company banner
# + DATE box, BKK block on the left and LCH on the right, three bands (FULL INBOUND |
# CURRENT STOCK, BOOKING WK1ST | STOCK BALANCE END, same for WK2ND), grey N/A cells,
# yellow LEASE IN / REPO (E/P) rows, dark-grey BMTP/SCT labels, blue DEPOT rows, pink
# negatives, blank single 0s except TOTAL rows, centred numbers. Page chrome kept from
# earlier requests: header photo + logo + title "Stock Report HQ" and the
# ALL/THBKK/THLCH filter. Still removed on purpose: live clock, 'created at' chip,
# WK/Excel meta chips, negative-alert box.
DASH_CSS = """
:root{--navy:#0d2b46;--ink:#1b2735;--mut:#64748b;--line:#e2e8f0;--bkk:#1d6fe8;--lch:#d72638}
*{box-sizing:border-box}
body{margin:0;color:var(--ink);background:#e9eef5;font:14px/1.5 "Segoe UI","Leelawadee UI","Noto Sans Thai",Roboto,sans-serif}
.accent-bar{height:5px;background:linear-gradient(90deg,var(--bkk),var(--lch))}
.wrap{max-width:1800px;margin:0 auto;padding:18px 18px 48px}
.topbar{display:flex;justify-content:space-between;align-items:center;gap:16px;
background:linear-gradient(rgba(255,255,255,.88),rgba(255,255,255,.88)),url('header.jpg') center/cover;
border:1px solid var(--line);border-radius:16px;padding:18px 26px;margin:0 0 12px;box-shadow:0 5px 16px rgba(15,35,60,.08)}
.topbar img{height:46px;width:auto;object-fit:contain}
h1{font-size:24px;margin:0 0 2px;color:var(--navy);font-weight:800}
.sub{color:var(--mut);margin:0;font-size:14px}
.toolbar{display:flex;gap:6px;margin:0 0 12px}
.filter-btn{cursor:pointer;border:1px solid var(--line);background:#fff;color:var(--mut);
font-family:inherit;font-weight:700;font-size:12.5px;line-height:1;letter-spacing:.03em;
padding:9px 18px;border-radius:999px;transition:all .15s}
.filter-btn:hover{border-color:var(--bkk);color:var(--navy)}
.filter-btn.active{color:#fff;border-color:transparent}
.filter-btn[data-filter="all"].active{background:var(--navy)}
.filter-btn[data-filter="bkk"].active{background:var(--bkk)}
.filter-btn[data-filter="lch"].active{background:var(--lch)}
.hide{display:none!important}
.xl-wrap{overflow-x:auto;background:#fff;border:1px solid #cbd5e1;border-radius:10px;padding:10px}
.xl{width:max-content;font-family:Arial,"Segoe UI","Noto Sans Thai",sans-serif;font-size:13px;color:#000}
.xl-title{display:flex;align-items:stretch;border:2px solid #000;height:56px;margin:0 0 6px}
.xl-title .t-name{display:flex;align-items:center;padding:0 14px;font-weight:700;font-size:20px;color:#1f4e79;
min-width:250px;border-right:2px solid #000}
.xl-title .t-logo{flex:1;display:flex;align-items:center;padding:0 10px}
.xl-title .t-logo img{height:44px;width:auto}
.xl-title .t-date{display:flex;align-items:center;gap:22px;padding:0 18px;border-left:2px solid #000;min-width:210px;font-weight:700}
.band{display:flex;align-items:flex-start}
.band + .band{margin-top:8px}
.zone{flex:none;border:2px solid #000}
.xl.stacked .band{flex-direction:column;gap:8px}
.xl table{border-collapse:collapse;table-layout:fixed}
.xl col.c-loc,.xl col.c-typ{width:84px}
.xl col.c-n{width:37px}
.xl td,.xl th{border:1px solid #bfbfbf;height:21px;padding:0 1px;text-align:center;font-weight:400;
white-space:nowrap;overflow:hidden;background:#fff}
.xl th.hd{background:#d9d9d9;font-weight:700}
.xl th.hd2{background:#d9d9d9;font-size:12.5px}
.xl th.dt{background:#fff;font-weight:700}
.xl .blue{color:#0000ff}
.xl .sepl{border-left:2px solid #000}
.xl thead tr:last-child th{border-bottom:2px solid #000}
.xl .dep{color:#0000ff}
.xl .na{background:#d9d9d9}
.xl .ylw{background:#ffff99}
.xl .mute{background:#b0b0b0}
.xl td.loc,.xl td.typ{text-align:center}
.xl tr.tot td{border-top:2px solid #000}
.xl td.totl{background:#d9d9d9;font-weight:700}
.xl tr.tot td.blue{color:#0000ff}
.xl td.neg{background:#ffc7ce!important;color:#9c0006!important;font-weight:700}
.nf-wrap{display:flex;flex-wrap:wrap;gap:22px;overflow-x:auto}
.nf{font-family:Calibri,"Segoe UI","Noto Sans Thai",Arial,sans-serif;font-size:14px;color:#000}
.nf table{border-collapse:collapse}
.nf td,.nf th{border:1.5px solid #808080;height:27px;min-width:56px;padding:0 6px;text-align:center;font-weight:400;background:#fff}
.nf th{background:#9dc3e6;font-weight:400}
.nf thead tr:first-child th{font-weight:700;font-size:15px}
.nf td.it{min-width:150px;text-align:center}
.nf tr.gry td{background:#bfbfbf}
.nf tr.org td{background:#fff2cc}
.nf tr.bal td{font-weight:700}
.nf tr.bal td.it{color:#000}
.nf td.bl{color:#0000ff}
.nf td.rd{color:#ff0000}
.nf td.rng{border:none;background:none;text-align:left;padding-left:12px;white-space:nowrap;font-weight:400}
.rf-wrap{display:flex;flex-wrap:wrap;gap:14px;margin-top:16px}
.rf{border:2px solid #000;font-family:Arial,"Segoe UI",sans-serif;font-size:13px}
.rf table{border-collapse:collapse}
.rf td,.rf th{border:1px solid #bfbfbf;height:21px;min-width:78px;padding:0 6px;text-align:center;font-weight:400;background:#fff}
.rf th{background:#d9d9d9;font-weight:700}
.rf tr.tot td{border-top:3px double #000;font-weight:700;background:#d9d9d9}
.rf-title{font-weight:700;color:#1f4e79;margin:16px 0 0;font-size:15px}
footer{margin-top:26px;color:var(--mut);font-size:12px;text-align:center}
@media(max-width:600px){.topbar{flex-direction:column;align-items:flex-start;padding:14px 16px}
.topbar img{height:34px}h1{font-size:20px}.wrap{padding:12px 10px 36px}}
"""

DASH_JS = """
(function(){
  var buttons = Array.prototype.slice.call(document.querySelectorAll('.filter-btn'));
  var zoned = document.querySelectorAll('[data-zone]');
  var wrap = document.getElementById('xlWrap');
  var xl = document.getElementById('xl');
  function fit(){
    if (!wrap || !xl) return;
    xl.style.zoom = 1;
    xl.classList.remove('stacked');
    var avail = wrap.clientWidth - 22;
    var natural = xl.getBoundingClientRect().width;
    var zonesShown = xl.querySelectorAll('.band')[0].querySelectorAll('.zone:not(.hide)').length;
    var s = Math.min(1, avail / natural);
    if (zonesShown > 1 && s < 0.72) {
      xl.classList.add('stacked');
      natural = xl.getBoundingClientRect().width;
      s = Math.max(0.7, Math.min(1, avail / natural));
    } else {
      s = Math.max(0.5, s);
    }
    xl.style.zoom = s;
  }
  function apply(filter){
    for (var i = 0; i < zoned.length; i++) {
      var z = zoned[i].getAttribute('data-zone');
      zoned[i].classList.toggle('hide', filter !== 'all' && z !== filter);
    }
    buttons.forEach(function(b){ b.classList.toggle('active', b.dataset.filter === filter); });
    try { localStorage.setItem('halStockFilter', filter); } catch (e) {}
    fit();
  }
  buttons.forEach(function(b){
    b.addEventListener('click', function(){ apply(b.dataset.filter); });
  });
  var saved = 'all';
  try { saved = localStorage.getItem('halStockFilter') || 'all'; } catch (e) {}
  apply(saved);
  window.addEventListener('resize', fit);
  window.addEventListener('load', fit);
})();
"""


def _sheet_logo_uri():
    """The company banner picture embedded in the template's Daily sheet, as a data
    URI (so the page stays one self-contained file)."""
    import base64
    import zipfile
    try:
        with zipfile.ZipFile(TEMPLATE) as z:
            name = next(n for n in z.namelist() if n.startswith("xl/media/"))
            data = z.read(name)
        mime = "image/png" if name.lower().endswith(".png") else "image/jpeg"
        return f"data:{mime};base64,{base64.b64encode(data).decode()}"
    except Exception:
        return ""


TH_DAYS = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]
# per-zone layout facts, mirroring the template
XL_ZONES = {
    "bkk": {"code": "BKK", "groups": BKK_GROUPS, "muted": {"BMTP", "SCT"}, "no_book": {"BMTP", "SCT"},
            "no_bal": set(), "blanks": 1},
    "lch": {"code": "LCH", "groups": LCH_GROUPS, "muted": set(), "no_book": set(NO_BALANCE_GROUPS),
            "no_bal": set(NO_BALANCE_GROUPS), "blanks": 2},
}


def _xl_block(vals, cls="", neg=False, zero_blank=True, blank=False, na=False):
    """9 numeric cells (one per container type); first cell carries the thick left rule."""
    parts = []
    for i, t in enumerate(TEMPLATE_TYPES):
        v = int(vals[t]) if vals is not None else 0
        c = [x for x in (cls, "na" if na else "", "sepl" if i == 0 else "",
                         "neg" if (neg and v < 0) else "") if x]
        txt = "" if (blank or na or (zero_blank and v == 0)) else str(v)
        parts.append(f"<td class='{' '.join(c)}'>{txt}</td>")
    return "".join(parts)


def _xl_types():
    return "".join(f"<th class='hd2{' sepl' if i == 0 else ''}'>{t}</th>"
                   for i, t in enumerate(TEMPLATE_TYPES))


def _xl_colgroup():
    return ("<colgroup><col class='c-loc'><col class='c-typ'>"
            + "<col class='c-n'>" * 18 + "</colgroup>")


def _xl_sum(data, groups):
    return {t: sum(int(data[g][t]) for g in groups) for t in TEMPLATE_TYPES}


def _xl_label_cells(g, cfg):
    dep = " dep" if GROUP_TYPE[g] == "DEPOT" else ""
    mute = " mute" if g in cfg["muted"] else ""
    return (f"<td class='loc{dep}{mute}'>{g}</td><td class='typ{dep}{mute}'>{GROUP_TYPE[g]}</td>")


def _xl_blank_row():
    return ("<tr><td class='loc'></td><td class='typ'></td>"
            + _xl_block(None, blank=True) + _xl_block(None, blank=True) + "</tr>")


def _xl_top_zone(zone, fi, cs, repo_side, date_txt):
    cfg = XL_ZONES[zone]
    code, groups = cfg["code"], cfg["groups"]
    head = ("<thead><tr><th class='hd' rowspan='2' colspan='2'>LOCATION</th>"
            f"<th class='hd sepl' colspan='9'>FULL INBOUND ({code})</th>"
            f"<th class='hd blue sepl' colspan='6'>CURRENT STOCK ({code})</th>"
            f"<th class='dt blue' colspan='3'>{date_txt}</th></tr>"
            f"<tr>{_xl_types()}{_xl_types()}</tr></thead>")
    rows = []
    for g in groups:
        depot = GROUP_TYPE[g] == "DEPOT"
        rows.append("<tr>" + _xl_label_cells(g, cfg)
                    + _xl_block(fi[g], na=depot)
                    + _xl_block(cs[g], cls="dep" if depot else "") + "</tr>")
    rows += [_xl_blank_row() for _ in range(cfg["blanks"])]
    rows.append("<tr><td class='loc ylw'>LEASE IN</td><td class='typ ylw'></td>"
                + _xl_block(None, na=True) + _xl_block(None, cls="ylw", blank=True) + "</tr>")
    rows.append("<tr><td class='loc ylw'>REPO (E/P)</td><td class='typ ylw'></td>"
                + _xl_block(None, na=True) + _xl_block(repo_side, cls="ylw dep") + "</tr>")
    cs_tot = {t: _xl_sum(cs, groups)[t] + int(repo_side[t]) for t in TEMPLATE_TYPES}
    rows.append(f"<tr class='tot'><td class='totl' colspan='2'>TOTAL AT {code}</td>"
                + _xl_block(_xl_sum(fi, groups), zero_blank=False)
                + _xl_block(cs_tot, cls="blue", zero_blank=False) + "</tr>")
    return (f"<div class='zone' data-zone='{zone}'><table>{_xl_colgroup()}{head}"
            f"<tbody>{''.join(rows)}</tbody></table></div>")


def _xl_book_zone(zone, bk, sew, wk, rng):
    cfg = XL_ZONES[zone]
    code, groups = cfg["code"], cfg["groups"]
    head = ("<thead><tr><th class='hd' rowspan='2' colspan='2'>LOCATION</th>"
            f"<th class='hd sepl' colspan='6'>BOOKING ({code}) {wk}</th>"
            f"<th class='dt' colspan='3'>{rng}</th>"
            f"<th class='hd blue sepl' colspan='9'>STOCK BALANCE END ({code}) {wk}</th></tr>"
            f"<tr>{_xl_types()}{_xl_types()}</tr></thead>")
    rows = []
    for g in groups:
        dep = "dep" if GROUP_TYPE[g] == "DEPOT" else ""
        rows.append("<tr>" + _xl_label_cells(g, cfg)
                    + _xl_block(bk[g], cls=dep, na=(g in cfg["no_book"]))
                    + _xl_block(sew[g], cls=dep, neg=True, blank=(g in cfg["no_bal"])) + "</tr>")
    rows += [_xl_blank_row() for _ in range(cfg["blanks"])]
    rows.append("<tr><td class='loc ylw'>LEASE IN</td><td class='typ ylw'></td>"
                + _xl_block(None, cls="ylw", blank=True) + _xl_block(None, cls="ylw", blank=True) + "</tr>")
    bal_groups = [g for g in groups if g not in cfg["no_bal"]]
    rows.append(f"<tr class='tot'><td class='totl' colspan='2'>TOTAL AT {code}</td>"
                + _xl_block(_xl_sum(bk, groups), zero_blank=False)
                + _xl_block(_xl_sum(sew, bal_groups), cls="blue", neg=True, zero_blank=False) + "</tr>")
    return (f"<div class='zone' data-zone='{zone}'><table>{_xl_colgroup()}{head}"
            f"<tbody>{''.join(rows)}</tbody></table></div>")


def _xl_rf_zone(zone, rf):
    area = XL_ZONES[zone]["code"]
    brands = ("CARRIER", "DAIKIN", "TRMK")
    body, tot = [], {c: 0 for c in brands}
    for y in RF_SEASONAL_YEARS:
        cells = ""
        for c in brands:
            v = rf[(area, y, c)]
            tot[c] += v
            cells += f"<td>{v or ''}</td>"
        body.append(f"<tr><td>{y}</td>{cells}</tr>")
    body.append("<tr class='tot'><td>TOTAL</td>" + "".join(f"<td>{tot[c]}</td>" for c in brands) + "</tr>")
    return (f"<div class='rf' data-zone='{zone}'><table><thead>"
            f"<tr><th colspan='4'>HAL (TH{area})</th></tr>"
            "<tr><th rowspan='2'>YEAR BUILT</th><th>CARRIER</th><th>DAIKIN</th><th>TRMK</th></tr>"
            "<tr><th>45RE</th><th>45RE</th><th>45RE</th></tr></thead>"
            f"<tbody>{''.join(body)}</tbody></table></div>")


def _nf_zone(zone, cs, fi, bk1, bk2, repo_side, bal_side, wk1_lbl, wk2_lbl):
    """The NEW FORMAT sheet's table for one zone (same rows/labels/colours as Excel).
    Current Stock = Daily total incl. REPO; Balance comes from build()'s `bal` (the
    template formula). Single 0s are blank except Current Stock and Balance."""
    cfg = XL_ZONES[zone]
    groups, code = cfg["groups"], cfg["code"]
    cs_tot = {t: _xl_sum(cs, groups)[t] + int(repo_side[t]) for t in TEMPLATE_TYPES}
    fi_tot = _xl_sum(fi, groups)
    b1_tot = _xl_sum(bk1, groups)
    b2_tot = _xl_sum(bk2, groups)

    def cells(vals, zero_blank=True, cls_fn=None):
        out = []
        for t in TEMPLATE_TYPES:
            v = int(vals[t])
            c = cls_fn(v) if cls_fn else ""
            out.append(f"<td class='{c}'>{'' if (zero_blank and v == 0) else v}</td>")
        return "".join(out)

    def row(label, vals, tr_cls="", zero_blank=True, rng="", cls_fn=None):
        rng_td = f"<td class='rng'>{rng}</td>" if rng else ""
        return (f"<tr class='{tr_cls}'><td class='it'>{label}</td>"
                f"{cells(vals, zero_blank, cls_fn)}{rng_td}</tr>")

    empty = {t: 0 for t in TEMPLATE_TYPES}
    types = "".join(f"<th>{t}</th>" for t in TEMPLATE_TYPES)
    body = (row("Current Stock", cs_tot, zero_blank=False)
            + row("E/P In (this week)", repo_side)
            + row("", empty, "gry") + row("", empty, "gry")
            + row("INBOUND MT RETURN", fi_tot, "org")
            + row("Booking(This week)", b1_tot, rng=wk1_lbl)
            + row("Booking(Next week)", b2_tot, rng=wk2_lbl)
            + row("Balance", bal_side, "bal", zero_blank=False,
                  cls_fn=lambda v: "rd" if v <= 0 else "bl"))
    return (f"<div class='nf' data-zone='{zone}'><table><thead>"
            f"<tr><th colspan='10'>HAL (TH{code})</th></tr>"
            f"<tr><th>Item</th>{types}</tr></thead><tbody>{body}</tbody></table></div>")


OV_CSS = """
body{background:#eaf0f7}
.ov{max-width:1240px;margin:0 auto;padding:22px 20px 48px}
.ov-banner{position:relative;height:158px;border-radius:18px;overflow:hidden;
background:#32414f url('header.jpg') center 38%/cover;box-shadow:0 8px 22px rgba(15,35,60,.18)}
.ov-banner:after{content:"";position:absolute;inset:0;background:linear-gradient(180deg,rgba(15,30,45,0) 25%,rgba(15,30,45,.62) 100%)}
.ov-banner .bt{position:absolute;z-index:1;left:20px;bottom:14px;color:#fff}
.ov-banner h1{margin:0;color:#fff;font-size:36px;line-height:1.15;font-weight:800;letter-spacing:.2px}
.ov-banner .bs{font-size:13px;letter-spacing:.06em;opacity:.9;margin-top:2px}
.ov-banner .bd{position:absolute;z-index:1;right:20px;bottom:14px;color:#fff;font:700 22px ui-monospace,Consolas,Menlo,monospace}
.ov-bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:14px 0 4px}
.nav{display:flex;gap:8px;margin:0 12px 0 0}
.nav a{text-decoration:none;font-weight:700;font-size:13px;padding:8px 16px;border-radius:999px;
border:1px solid var(--line);background:#fff;color:var(--mut)}
.nav a:hover{border-color:var(--bkk);color:var(--navy)}
.nav a.on{background:var(--navy);border-color:var(--navy);color:#fff}
.ov-bar .toolbar{margin:0}
.zone-h{display:flex;align-items:baseline;gap:10px;margin:26px 0 12px}
.zone-h h2{margin:0;font-size:36px;line-height:1.1;color:#1f6aa5;font-weight:800}
.zone-h small{color:#8a97a6;font-size:13px;letter-spacing:.04em}
.zone-h:after{content:"";flex:1;border-top:1px solid #cfd8e3;transform:translateY(-6px)}
.ccard{background:#fff;border-radius:16px;border-left:5px solid #1e5bc6;box-shadow:0 6px 18px rgba(15,35,60,.09);
padding:16px 20px 12px;overflow-x:auto}
.cc-h{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:10px}
.cc-n{display:flex;align-items:center;gap:10px}
.cc-n i{width:11px;height:11px;border-radius:50%;background:#1e5bc6;display:inline-block}
.cc-n b{display:block;color:#1e5bc6;font-size:17px;font-weight:800;letter-spacing:.02em;line-height:1.1}
.cc-n small{display:block;color:#8a97a6;font-size:12.5px}
.cc-u{text-align:right}
.cc-u b{display:block;font:700 26px ui-monospace,Consolas,Menlo,monospace;color:#16222f;line-height:1.1}
.cc-u small{color:#8a97a6;font-size:11px;letter-spacing:.08em}
table.ct{width:100%;border-collapse:collapse;min-width:640px;font-family:ui-monospace,Consolas,Menlo,monospace;font-size:14px}
.ct th{background:#f2f6fa;color:#8a97a6;font:600 11.5px "Segoe UI",sans-serif;padding:7px 8px;text-align:right}
.ct th:first-child,.ct td:first-child{text-align:left}
.ct td{padding:8px 8px;text-align:right;border-bottom:1px solid #e8eef4;color:#2b3a4a}
.ct td:first-child{font:600 13px "Segoe UI","Leelawadee UI",sans-serif;color:#44566a;white-space:nowrap}
.ct tr.cur td{font-weight:800;font-size:15.5px;color:#0f1b27;padding:11px 8px}
.ct tr.cur td:first-child{font-size:15px;color:#0f1b27}
.ct td.z{color:#aab6c3}
.ct tr.bal td{border-bottom:0;padding-top:12px}
.ct tr.bal td:first-child{font-weight:800;color:#0f1b27}
.pill{display:inline-block;min-width:44px;text-align:center;padding:3px 9px;border-radius:9px;font-weight:700;font-size:13.5px}
.p-sur{background:#dbeafe;color:#1d5fa8}.p-nor{background:#dcf3dc;color:#15803d}
.p-tig{background:#fdf1c9;color:#946400}.p-sho{background:#fbd9d9;color:#b42318}
.ov-det{margin:10px 0 0}
.ov-det summary{cursor:pointer;color:#1f6aa5;font-size:13px;font-weight:700;padding:4px 0}
table.gt{width:100%;border-collapse:collapse;font-size:13.5px;margin-top:8px}
.gt th{white-space:nowrap;text-align:right;color:var(--mut);font-size:12px;font-weight:600;padding:6px;border-bottom:2px solid var(--line)}
.gt th:first-child,.gt td:first-child{text-align:left}
.gt td{padding:7px 6px;border-bottom:1px solid var(--line);text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.gt td.gname b{color:var(--navy)}
.badge{display:inline-block;margin-left:6px;font-size:10px;font-weight:700;padding:1px 7px;border-radius:999px;
background:#e2e8f0;color:#475569;vertical-align:1px}
.badge.dep{background:#dbeafe;color:#1d4ed8}
.gt td.dash{color:#cbd5e1}
.gt tr.tot td{border-top:2px solid var(--navy);font-weight:800;color:var(--navy);border-bottom:0}
.ov-leg{display:flex;flex-wrap:wrap;gap:18px;margin:26px 0 0;padding-top:14px;border-top:1px solid #cfd8e3;color:#5b6b7c;font-size:12.5px}
.ov-leg span:before{content:"";display:inline-block;width:11px;height:11px;border-radius:3px;margin-right:6px;vertical-align:-1px;background:var(--c)}
.ov-leg .sur{--c:#1d5fa8}.ov-leg .nor{--c:#15803d}.ov-leg .tig{--c:#946400}.ov-leg .sho{--c:#b42318}
.ov-note{margin:8px 0 0;color:#8a97a6;font-size:12px}
@media(max-width:600px){.ov-banner h1{font-size:26px}.zone-h h2{font-size:28px}.ov-banner .bd{font-size:16px}}
"""

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_ZONE_NAME = {"bkk": ("Bangkok", "THBKK"), "lch": ("Laem Chabang", "THLCH")}


def _nav(active):
    items = (("index", "index.html", "📋 ตารางรายงาน"), ("overview", "overview.html", "📊 ภาพรวม"))
    links = []
    for key, href, label in items:
        on = " class='on'" if key == active else ""
        links.append(f"<a href='{href}'{on}>{label}</a>")
    return "<nav class='nav'>" + "".join(links) + "</nav>"


def _bal_class(bal, demand):
    """Colour of a Balance pill. Shortage = balance <= 0 (same red rule as the NEW FORMAT
    sheet); otherwise by cover = balance / booking of the two weeks: Tight < 1x,
    Normal 1-3x, Surplus > 3x (no booking at all: Tight under 5 boxes, else Normal)."""
    if bal <= 0:
        return "p-sho"
    if demand <= 0:
        return "p-tig" if bal < 5 else "p-nor"
    cover = bal / demand
    return "p-tig" if cover < 1 else ("p-nor" if cover <= 3 else "p-sur")


def _ov_zone(zone, cs, fi, bk1, bk2, repo_side, bal_side, wk1_lbl, wk2_lbl):
    cfg = XL_ZONES[zone]
    groups, code = cfg["groups"], cfg["code"]
    name, th = _ZONE_NAME[zone]
    stock = {t: _xl_sum(cs, groups)[t] + int(repo_side[t]) for t in TEMPLATE_TYPES}
    fin, b1, b2 = _xl_sum(fi, groups), _xl_sum(bk1, groups), _xl_sum(bk2, groups)
    bal = {t: int(bal_side[t]) for t in TEMPLATE_TYPES}

    def plain(vals, cls=""):
        out = ""
        for t in TEMPLATE_TYPES:
            v = int(vals[t])
            out += f"<td class='z'>–</td>" if v == 0 else f"<td>{v:,}</td>"
        return f"<tr class='{cls}'>" + "{label}" + out + "</tr>"

    def row(label, vals, cls=""):
        return plain(vals, cls).replace("{label}", f"<td>{label}</td>")

    pills = "".join(
        f"<td><span class='pill {_bal_class(bal[t], b1[t] + b2[t])}'>{bal[t]:,}</span></td>"
        for t in TEMPLATE_TYPES)
    head = "".join(f"<th>{t}</th>" for t in TEMPLATE_TYPES)
    table = (f"<table class='ct'><thead><tr><th>TYPE</th>{head}</tr></thead><tbody>"
             + row("Current stock", stock, "cur")
             + row("MT return (inbound)", fin)
             + row("Booking, this wk", b1)
             + row("Booking, next wk", b2)
             + f"<tr class='bal'><td>Balance</td>{pills}</tr></tbody></table>")

    grows = ""
    for g in groups:
        st = sum(int(cs[g][t]) for t in TEMPLATE_TYPES)
        badge = ("<span class='badge dep'>DEPOT</span>" if GROUP_TYPE[g] == "DEPOT"
                 else f"<span class='badge'>{GROUP_TYPE[g]}</span>")
        if g in cfg["no_book"]:
            c1 = c2 = "<td class='dash'>–</td>"
        else:
            c1 = f"<td>{sum(int(bk1[g][t]) for t in TEMPLATE_TYPES):,}</td>"
            c2 = f"<td>{sum(int(bk2[g][t]) for t in TEMPLATE_TYPES):,}</td>"
        fcell = ("<td class='dash'>–</td>" if GROUP_TYPE[g] == "DEPOT"
                 else f"<td>{sum(int(fi[g][t]) for t in TEMPLATE_TYPES):,}</td>")
        grows += (f"<tr><td class='gname'><b>{g}</b>{badge}</td><td>{st:,}</td>{fcell}{c1}{c2}</tr>")
    grows += (f"<tr class='tot'><td>รวม {code}</td><td>{sum(stock.values()):,}</td><td>{sum(fin.values()):,}</td>"
              f"<td>{sum(b1.values()):,}</td><td>{sum(b2.values()):,}</td></tr>")
    detail = ("<details class='ov-det'><summary>ดูแยกตามสถานที่</summary>"
              "<table class='gt'><thead><tr><th>สถานที่</th><th>ตู้ว่างในสต็อก</th><th>FULL ขาเข้า</th>"
              f"<th title='{wk1_lbl}'>Booking WK1ST</th><th title='{wk2_lbl}'>Booking WK2ND</th></tr></thead>"
              f"<tbody>{grows}</tbody></table></details>")

    return f"""<section data-zone="{zone}">
<div class="zone-h"><h2>{name}</h2><small>{th}</small></div>
<div class="ccard"><div class="cc-h"><div class="cc-n"><i></i><div><b>HEUNG A</b><small>{name}</small></div></div>
<div class="cc-u"><b>{sum(stock.values()):,}</b><small>UNITS ON HAND</small></div></div>
{table}{detail}</div></section>"""


def write_overview(path, report_date, wk1_lbl, wk2_lbl, fi, cs, bk1, bk2, bal, repo):
    date_txt = f"{report_date.day} {_MONTHS[report_date.month - 1]} {report_date.year}"
    zones = (_ov_zone("bkk", cs, fi, bk1, bk2, repo["BKK"], bal["BKK"], wk1_lbl, wk2_lbl)
             + _ov_zone("lch", cs, fi, bk1, bk2, repo["LCH"], bal["LCH"], wk1_lbl, wk2_lbl))
    html = f"""<!doctype html><html lang="th"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Stock Report HQ &ndash; ภาพรวม {report_date:%Y-%m-%d}</title>
<style>{DASH_CSS}{OV_CSS}</style></head><body>
<div class="ov">
<header class="ov-banner"><div class="bt"><h1>Container Stock Condition</h1>
<div class="bs">HEUNG A LINE (THAILAND)</div></div><div class="bd">{date_txt}</div></header>
<div class="ov-bar">{_nav("overview")}
<div class="toolbar">
<button type="button" class="filter-btn active" data-filter="all">ALL</button>
<button type="button" class="filter-btn" data-filter="bkk">THBKK</button>
<button type="button" class="filter-btn" data-filter="lch">THLCH</button>
</div></div>
{zones}
<div class="ov-leg"><span class="sur">Surplus</span><span class="nor">Normal</span>
<span class="tig">Tight</span><span class="sho">Shortage</span></div>
<p class="ov-note">Balance = ตู้ในสต็อก + FULL ขาเข้า &minus; Booking 2 สัปดาห์ (WK1ST {wk1_lbl}, WK2ND {wk2_lbl}) ·
Shortage = คงเหลือ ≤ 0 · Tight = คงเหลือน้อยกว่า booking 2 สัปดาห์ · Normal = 1&ndash;3 เท่า · Surplus = มากกว่า 3 เท่า</p>
</div>
<script>{DASH_JS}</script>
</body></html>"""
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)


def write_dashboard(path, report_date, wk1_lbl, wk2_lbl,
                    fi, cs, bk1, bk2, sew1, sew2, bal, rf, repo, xlsx_name):
    date_txt = f"{report_date.day}/{report_date.month}/{report_date.year}"
    top = ("<div class='band'>"
           + _xl_top_zone("bkk", fi, cs, repo["BKK"], date_txt)
           + _xl_top_zone("lch", fi, cs, repo["LCH"], date_txt) + "</div>")
    wk1 = ("<div class='band'>"
           + _xl_book_zone("bkk", bk1, sew1, "WK1ST", wk1_lbl)
           + _xl_book_zone("lch", bk1, sew1, "WK1ST", wk1_lbl) + "</div>")
    wk2 = ("<div class='band'>"
           + _xl_book_zone("bkk", bk2, sew2, "WK2ND", wk2_lbl)
           + _xl_book_zone("lch", bk2, sew2, "WK2ND", wk2_lbl) + "</div>")
    nf_html = (_nf_zone("bkk", cs, fi, bk1, bk2, repo["BKK"], bal["BKK"], wk1_lbl, wk2_lbl)
               + _nf_zone("lch", cs, fi, bk1, bk2, repo["LCH"], bal["LCH"], wk1_lbl, wk2_lbl))
    logo = _sheet_logo_uri()
    logo_html = f"<img src='{logo}' alt='Heung A Line'>" if logo else ""
    weekday = TH_DAYS[report_date.weekday()]
    html = f"""<!doctype html><html lang="th"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Stock Report HQ &ndash; {report_date:%Y-%m-%d}</title>
<style>{DASH_CSS}</style></head><body>
<div class="accent-bar"></div>
<div class="wrap">
<header class="topbar">
<div><h1>Stock Report HQ</h1>
<p class="sub">รายงานประจำวันที่ {report_date:%d/%m/%Y} (วัน{weekday})</p></div>
<img src="logo.png" alt="logo">
</header>
{_nav("index")}
<div class="toolbar">
<button type="button" class="filter-btn active" data-filter="all">ALL</button>
<button type="button" class="filter-btn" data-filter="bkk">THBKK</button>
<button type="button" class="filter-btn" data-filter="lch">THLCH</button>
</div>
<div class="xl-wrap" id="xlWrap"><div class="xl" id="xl">
<div class="xl-title"><div class="t-name">DAILY CONTAINER STOCK</div>
<div class="t-logo">{logo_html}</div>
<div class="t-date"><span>DATE:</span><span>{report_date:%d-%b-%Y}</span></div></div>
{top}
{wk1}
{wk2}
</div></div>
<div class="rf-title">RF SEASONAL</div>
<div class="rf-wrap">{_xl_rf_zone("bkk", rf)}{_xl_rf_zone("lch", rf)}</div>
<div class="rf-title">NEW FORMAT</div>
<div class="nf-wrap">{nf_html}</div>
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
    """Copy the freshly built index.html + overview.html into the public dashboard repo and push."""
    if not os.path.isdir(os.path.join(dash_dir, ".git")):
        print(f"  publish skipped: {dash_dir} is not a git repo "
              f"(clone https://github.com/sirichai1265/hal-stock-hq-dashboard there first)")
        return
    pages = ["index.html"]
    ov = os.path.join(os.path.dirname(src_html), "overview.html")
    if os.path.exists(ov):
        pages.append("overview.html")
    for name in pages:
        shutil.copyfile(src_html if name == "index.html" else ov, os.path.join(dash_dir, name))
    subprocess.run(["git", "-C", dash_dir, "add"] + pages, check=True)
    r = subprocess.run(["git", "-C", dash_dir, "commit", "-m",
                        f"dashboard {report_date:%Y-%m-%d}"], capture_output=True, text=True)
    if r.returncode and "nothing to commit" in (r.stdout + r.stderr):
        print("  publish: no dashboard change")
        return
    subprocess.run(["git", "-C", dash_dir, "push", "-q"], check=True)
    print("  published -> https://sirichai1265.github.io/hal-stock-hq-dashboard/")


if __name__ == "__main__":
    main()

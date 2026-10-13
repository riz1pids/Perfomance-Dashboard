#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shwapno Overall Performance Dashboard - PART 1: DATA BUILDER
=============================================================
Reads raw POS exports + master lookups, applies Same-Store logic and the
SPLY (Same Period Last Year) same-day mapping, and emits browser-loadable
JavaScript data files consumed by the Part 2 HTML dashboard.

Design rules (per project conventions):
  * Single CONFIG block at the top - no code edits needed for routine runs.
  * Preflight hard-stops instead of silent fallbacks.
  * Raw values only in the output; every ratio (Growth %, GP %, ACH %,
    Basket Size) is derived in the browser so figures stay internally
    consistent and file size stays small.

Output layout (under OUTPUT_DIR), one folder per month:
  manifest.js / .json   the months available; the dashboard reads this first
  2026-09/              this month's payloads, each as .js and .json, split
                        into .partN files when larger than MAX_FILE_MB
Rebuilding a month replaces that month's folder; other months are untouched.
Payloads:
  data_meta.js        - month, day list, SPLY map, filter vocabularies, run log
  data_daily_cat.js    - day x Division x Cat1 x Cat3 x Brand x CRP, actuals
  data_daily_target.js - day x Division x Cat1 x Cat3 x CRP, targets only
  data_sku.js          - SKU / Article rows, period totals, no targets

Targets are published at Category 3, NOT by brand, so they live in their own
payload. Joining them onto a cat3 x brand grain would duplicate one target
across every brand in the category.

Usage:  python build_dashboard_data.py
"""

import base64
import gzip
import json
import os
import re
import sys
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd

# =============================================================================
# CONFIG
# =============================================================================
HERE = os.path.dirname(os.path.abspath(__file__))

CONFIG = {
    # ---- Input files -------------------------------------------------------
    # input/ and data/ sit beside this script. SHWAPNO_INPUT_DIR and
    # SHWAPNO_OUTPUT_DIR override them (the GitHub workflow uses both).
    "INPUT_DIR": os.environ.get("SHWAPNO_INPUT_DIR") or os.path.join(HERE, "input"),
    "FILE_SALES_TY": "sept 2026.csv",
    "FILE_SALES_LY": "sept 2025.csv",
    "FILE_SAME_STORE": "SEPT'26 Same Store List.xlsx",
    "FILE_SAME_DAY": "SAME DAY Sep'26 Proposed.xlsx",
    "FILE_TARGET": "SEPT'26 Cat3 wise Target.xlsx",
    "FILE_PERSON": "Personwise Category Distribution.xlsx",
    "FILE_KEY_ARTICLE": "Key_Article.xlsx",
    "FILE_KEY_STORE": "Key_Store.xlsx",
    "FILE_ZONE": "Zone Distribution Sep 2026 w location type.xlsx",

    # ---- Output ------------------------------------------------------------
    "OUTPUT_DIR": os.environ.get("SHWAPNO_OUTPUT_DIR") or os.path.join(HERE, "data"),

    # ---- Period ------------------------------------------------------------
    "PERIOD_LABEL": "September 2026",
    "TY_YEAR": 2026,
    "LY_YEAR": 2025,
    "SAME_DAY_SHEET": "Sep'26",
    "SAME_DAY_TY_ROW": 3,    # 0-indexed row holding TY (2026) dates
    "SAME_DAY_LY_ROW": 9,    # 0-indexed row holding SPLY (2025) dates
    "SAME_DAY_FIRST_COL": 1,  # first data column

    # ---- Same-store rule ---------------------------------------------------
    # True  -> same-store universe = codes in the Same Store List that traded
    #          in BOTH TY and LY within the loaded window.
    # False -> use the Same Store List strictly as published (business rule).
    "SAME_STORE_REQUIRE_BOTH_YEARS": False,

    # ---- Store-wise target -------------------------------------------------
    # No store-level target is published, so the chain day-wise target is
    # allocated across outlets. Modes:
    #   "allocate_sply" -> split each day's chain target by that outlet's share
    #                      of SPLY sales for the same day (new outlets with no
    #                      SPLY base get no target and show "-")
    #   "none"          -> store page carries no Target / ACH columns
    "STORE_TARGET_MODE": "none",

    "ZONE_SHEET": "Final_Zone Dis",
    "ZONE_COLS": {
        "outlet": "CODE",
        "leader": "Regional Head HR Name",
        "zonal": "Zonal HR Name",
        "area": "Area",
    },

    # ---- Target scope ------------------------------------------------------
    # Targets are published chain-wide (all-store). Options:
    #   "ALLSTORE_ONLY" -> targets shown on All Store scope only
    #   "BOTH"          -> same chain target repeated on both scopes
    "TARGET_SCOPE": "ALLSTORE_ONLY",

    # ---- Master-data sheets / columns -------------------------------------
    "TARGET_HEADER_ROW": 1,
    "PERSON_SHEET": "raw",
    "KEY_ARTICLE_SHEET": "Key_Article",
    "ARTICLE_LEVEL_COLS": {
        "code": "Article_Code",
        "name": "Article_Name",
        "division": "Division",
        "cat1": "Category_Level_1",
        "cat3": "Category_Level_3",
        "brand": "Brand_Name_Description",
    },
    "UNASSIGNED_LABEL": "UNASSIGNED",

    # ---- Output tuning -----------------------------------------------------
    # Emit a plain .json beside every .js payload. The dashboard loads the
    # .js (script tags work from file://); the .json is for anything else
    # that wants the same numbers.
    "EMIT_JSON": False,
    # Delete any other file left in OUTPUT_DIR from an earlier build, so the
    # folder only ever holds the current .js / .json set.
    "PRUNE_OUTPUT_DIR": True,

    # GitHub rejects files over 25 MB. Any payload larger than this is split
    # on row boundaries into parts, each of which is still valid JSON.
    "MAX_FILE_MB": 24,


    # How many months to keep in the data folder. None keeps every month.
    # Set 2 to hold, say, September and October, and have September's
    # folder deleted automatically when November is built.
    "KEEP_MONTHS": 3,

    "ROUND_MONEY": 3,
    "ROUND_QTY": 3,
    "JS_VAR_PREFIX": "SHWAPNO_",
    "PRETTY_JSON": False,
}

METRIC_COLS = ["sales", "gpv", "ff", "qty"]
LOG = []


CR_FIGURE = re.compile(r"-?[\d,]*\.?\d+\s*cr\b")


def log(msg):
    LOG.append(msg)
    # Actions logs on a public repository are public. SHWAPNO_REDACT_LOG=1
    # masks the crore figures on screen; the build log kept in meta.js is not
    # touched.
    shown = CR_FIGURE.sub("[x cr]", msg) if os.environ.get("SHWAPNO_REDACT_LOG") else msg
    print(f"[{datetime.now():%H:%M:%S}] {shown}", flush=True)


def halt(msg):
    print(f"\n!!! PREFLIGHT FAILURE: {msg}\n", file=sys.stderr)
    sys.exit(1)


def norm(s):
    """Uppercase, trim, collapse whitespace - the canonical join key form."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    return " ".join(s.strip().upper().split())


INPUT_EXTS = (".csv", ".xlsx", ".xlsm", ".xls")


def _input_files():
    out = []
    for root, _dirs, files in os.walk(CONFIG["INPUT_DIR"]):
        for f in files:
            if not f.startswith(("~$", ".")):
                out.append(os.path.join(root, f))
    return sorted(out)


def ipath(key):
    """Find CONFIG[key] under INPUT_DIR, searching subfolders too (a Drive
    download lands in one). An exact name wins; otherwise the same name in any
    letter case, as .csv or as an Excel workbook. Two candidates is an error,
    never a guess."""
    want = CONFIG[key]
    exact = os.path.join(CONFIG["INPUT_DIR"], want)
    if os.path.exists(exact):
        return exact
    files = _input_files()
    hits = [f for f in files if os.path.basename(f) == want]
    if not hits:
        stem = os.path.splitext(want)[0].strip().lower()
        hits = [f for f in files
                if os.path.splitext(os.path.basename(f))[0].strip().lower() == stem
                and os.path.splitext(f)[1].lower() in INPUT_EXTS]
    if len(hits) > 1:
        halt(f"{key}: more than one file matches '{want}': "
             + ", ".join(os.path.relpath(h, CONFIG['INPUT_DIR']) for h in hits))
    return hits[0] if hits else exact


# =============================================================================
# 1. DAY MAPPING (TY date <-> SPLY date)
# =============================================================================
def load_day_mapping():
    """Authoritative TY->SPLY pairing straight from the Same Day workbook.

    The retail calendar is NOT chronological - the Friday/Saturday weekend
    means last year's comparable date is pinned weekday-to-weekday. Never
    derive this by subtracting 365 days.
    """
    raw = pd.read_excel(ipath("FILE_SAME_DAY"),
                        sheet_name=CONFIG["SAME_DAY_SHEET"], header=None)
    ty_row = raw.iloc[CONFIG["SAME_DAY_TY_ROW"]]
    ly_row = raw.iloc[CONFIG["SAME_DAY_LY_ROW"]]

    pairs = []
    for c in range(CONFIG["SAME_DAY_FIRST_COL"], raw.shape[1]):
        t, l = ty_row.iloc[c], ly_row.iloc[c]
        if pd.isna(t):
            continue
        t = pd.to_datetime(t, errors="coerce")
        l = pd.to_datetime(l, errors="coerce")
        if pd.isna(t) or t.year != CONFIG["TY_YEAR"]:
            continue
        if pd.isna(l):
            halt(f"Same-day sheet has TY date {t:%Y-%m-%d} with no SPLY pair.")
        if l.year != CONFIG["LY_YEAR"]:
            halt(f"Same-day sheet: SAME_DAY_LY_ROW={CONFIG['SAME_DAY_LY_ROW']} "
                 f"pairs TY {t:%Y-%m-%d} with {l:%Y-%m-%d}, which is not "
                 f"{CONFIG['LY_YEAR']}. That row is probably the SPLM "
                 f"(last month) block, not SPLY - check the row index.")
        pairs.append((t.normalize(), l.normalize()))

    if not pairs:
        halt("No TY/SPLY date pairs parsed from the same-day sheet.")

    pairs.sort(key=lambda x: x[0])
    dedup, seen = [], set()
    for t, l in pairs:
        if t in seen:
            continue
        seen.add(t)
        dedup.append((t, l))

    log(f"Day mapping: {len(dedup)} calendar days "
        f"({dedup[0][0]:%d %b %Y} -> {dedup[-1][0]:%d %b %Y}); "
        f"SPLY {dedup[0][1]:%d %b %Y} -> {dedup[-1][1]:%d %b %Y}")
    return dedup


# =============================================================================
# 2. MASTER DATA
# =============================================================================
def load_key_article():
    c = CONFIG["ARTICLE_LEVEL_COLS"]
    ka = pd.read_excel(ipath("FILE_KEY_ARTICLE"), engine="calamine",
                       sheet_name=CONFIG["KEY_ARTICLE_SHEET"],
                       usecols=list(c.values()))
    ka = ka.rename(columns={v: k for k, v in c.items()})
    ka["code"] = pd.to_numeric(ka["code"], errors="coerce")
    ka = ka.dropna(subset=["code"])
    ka["code"] = ka["code"].astype("int64")

    dups = int(ka["code"].duplicated().sum())
    if dups:
        ka = ka.drop_duplicates(subset=["code"], keep="first")
        log(f"Key_Article: dropped {dups} duplicate article code(s), kept first.")

    ka["name"] = ka["name"].fillna("").astype(str).str.strip()
    for lvl in ("division", "cat1", "cat3", "brand"):
        ka[lvl] = ka[lvl].map(norm)
        ka.loc[ka[lvl].isin(["", "NAN", "0", "-", "N/A", "NA", "NONE"]), lvl] = \
            CONFIG["UNASSIGNED_LABEL"]

    log(f"Key_Article: {len(ka):,} articles | "
        f"{ka.division.nunique()} divisions, {ka.cat1.nunique()} cat1, "
        f"{ka.cat3.nunique()} cat3, {ka.brand.nunique():,} brands")
    return ka


def load_key_store():
    ks = pd.read_excel(ipath("FILE_KEY_STORE"))
    ks.columns = [str(c).strip() for c in ks.columns]
    need = ["Outlet_Code", "Outlet_Name", "Outlet_Format",
            "Outlet_Division", "Outlet_Zone", "Outlet_Cluster",
            "Outlet_Status", "Outlet_Size"]
    missing = [c for c in need if c not in ks.columns]
    if missing:
        halt(f"Key_Store is missing column(s): {missing}")

    ks["outlet"] = ks["Outlet_Code"].fillna("").astype(str).str.strip().str.upper()
    ks = ks[ks["outlet"] != ""].drop_duplicates(subset=["outlet"], keep="first")

    out = pd.DataFrame({
        "outlet": ks["outlet"],
        "store_name": ks["Outlet_Name"].fillna("").astype(str).str.strip(),
        "format": ks["Outlet_Format"].fillna("").astype(str).str.strip()
                  .replace("", CONFIG["UNASSIGNED_LABEL"]),
        "geo": ks["Outlet_Division"].fillna("").astype(str).str.strip()
               .str.title().str.replace(r"\s+", " ", regex=True)
               .replace("", CONFIG["UNASSIGNED_LABEL"]),
        "zone": ks["Outlet_Zone"].fillna("").astype(str).str.strip()
                .replace("", CONFIG["UNASSIGNED_LABEL"]),
        "cluster": ks["Outlet_Cluster"].fillna("").astype(str).str.strip()
                   .replace("", CONFIG["UNASSIGNED_LABEL"]),
        "status": ks["Outlet_Status"].fillna("").astype(str).str.strip().str.upper(),
    })
    log(f"Key_Store: {len(out):,} outlets "
        f"({(out.status == 'ACTIVE').sum():,} active)")
    return out


def load_zone(stores):
    """Attach Leader (regional head) and Zonal manager to the outlet master."""
    c = CONFIG["ZONE_COLS"]
    z = pd.read_excel(ipath("FILE_ZONE"), sheet_name=CONFIG["ZONE_SHEET"])
    z.columns = [str(x).strip() for x in z.columns]
    missing = [v for v in c.values() if v not in z.columns]
    if missing:
        halt(f"Zone file is missing column(s): {missing}")

    z = z[list(c.values())].rename(columns={v: k for k, v in c.items()})
    z["outlet"] = z["outlet"].fillna("").astype(str).str.strip().str.upper()
    z = z[z["outlet"] != ""].drop_duplicates(subset=["outlet"], keep="first")
    for k in ("leader", "zonal", "area"):
        z[k] = (z[k].fillna("").astype(str).str.strip()
                .replace("", CONFIG["UNASSIGNED_LABEL"]))

    out = stores.merge(z, on="outlet", how="left")
    for k in ("leader", "zonal", "area"):
        gap = int(out[k].isna().sum())
        out[k] = out[k].fillna(CONFIG["UNASSIGNED_LABEL"])
        if gap:
            log(f"  ! {gap:,} outlet(s) absent from the zone file -> "
                f"{k} = {CONFIG['UNASSIGNED_LABEL']}.")
    log(f"Zone file: {len(z):,} outlets | {z['leader'].nunique()} leaders, "
        f"{z['zonal'].nunique()} zonal managers, {z['area'].nunique()} areas")
    return out


def load_same_store():
    df = pd.read_excel(ipath("FILE_SAME_STORE"), header=1)
    df.columns = [str(c).strip().upper() for c in df.columns]
    if "CODE" not in df.columns:
        halt("Same Store List: no CODE column found on header row 2.")
    codes = (df["CODE"].fillna("").astype(str).str.strip().str.upper())
    codes = set(c for c in codes if c and c != "NAN")
    log(f"Same Store List: {len(codes):,} outlet codes")
    return codes


def load_person_map():
    p = pd.read_excel(ipath("FILE_PERSON"), sheet_name=CONFIG["PERSON_SHEET"])
    p.columns = [str(c).strip() for c in p.columns]
    cmap = {"Division[Article]": "division",
            "Category Level 1[Article]": "cat1",
            "Category Level 3": "cat3",
            "CRP": "crp", "Head": "head"}
    missing = [c for c in cmap if c not in p.columns]
    if missing:
        halt(f"Person-wise file is missing column(s): {missing}")
    p = p.rename(columns=cmap)[["division", "cat1", "cat3", "crp", "head"]]
    for lvl in ("division", "cat1", "cat3"):
        p[lvl] = p[lvl].map(norm)
    p["crp"] = p["crp"].map(norm).replace("", CONFIG["UNASSIGNED_LABEL"])
    p["head"] = p["head"].map(norm).replace("", CONFIG["UNASSIGNED_LABEL"])

    before = len(p)
    p = p.drop_duplicates(subset=["division", "cat1", "cat3"], keep="first")
    if before != len(p):
        log(f"Person map: collapsed {before - len(p)} duplicate cat3 key(s) "
            f"(case variants, identical CRP).")
    log(f"Person map: {len(p):,} cat3 keys | {p['crp'].nunique()} CRPs, "
        f"{p['head'].nunique()} heads")
    return p


def load_targets(day_pairs):
    """Day-wise Sales and GPV targets at Division x Cat1 x Cat3 x CRP grain."""
    t = pd.read_excel(ipath("FILE_TARGET"), header=CONFIG["TARGET_HEADER_ROW"])
    t.columns = [str(c).strip() for c in t.columns]
    base = {"DIVISION": "division",
            "Category Level 1[Article]": "cat1",
            "Category Level 3": "cat3",
            "CRP": "crp"}
    missing = [c for c in base if c not in t.columns]
    if missing:
        halt(f"Target file is missing column(s): {missing}")

    # Date-like columns appear twice: first block = Sales target, second = GPV.
    # A bare-time header such as "00:00:00" (the blank total column at the end
    # of each block) would otherwise be parsed by pandas as *today's* date and
    # silently shadow a real day's target - so require an explicit ISO date.
    iso = re.compile(r"^\d{4}-\d{2}-\d{2}")
    date_cols = []
    for i, c in enumerate(t.columns):
        if not iso.match(str(c).strip()):
            continue
        d = pd.to_datetime(str(c).strip()[:10], errors="coerce")
        if pd.notna(d) and d.year == CONFIG["TY_YEAR"]:
            date_cols.append((i, d.normalize(), c))
    if not date_cols:
        halt("Target file: no TY day-wise columns detected.")

    n_days = len(day_pairs)
    by_date = {}
    for i, d, c in date_cols:
        by_date.setdefault(d, []).append((i, c))

    bad = {d: [c for _, c in v] for d, v in by_date.items() if len(v) != 2}
    if bad:
        halt("Target file: every TY date must map to exactly two columns "
             f"(Sales then GPV). Offenders: "
             f"{ {d.strftime('%Y-%m-%d'): v for d, v in bad.items()} }")

    sales_cols, gpv_cols = {}, {}
    for d, lst in by_date.items():
        lst.sort()
        sales_cols[d] = lst[0][1]
        gpv_cols[d] = lst[1][1]

    miss = [d for d, _ in day_pairs if d not in sales_cols]
    if miss:
        log(f"  ! {len(miss)} calendar day(s) have no target column "
            f"(first: {miss[0]:%Y-%m-%d}) - treated as zero target.")

    t["division"] = t["DIVISION"].map(norm)
    t["cat1"] = t["Category Level 1[Article]"].map(norm)
    t["cat3"] = t["Category Level 3"].map(norm)
    t["crp"] = t["CRP"].map(norm).replace("", CONFIG["UNASSIGNED_LABEL"])
    t = t[t["division"] != ""].copy()

    recs = []
    for day_idx, (ty_date, _ly) in enumerate(day_pairs, start=1):
        sc = sales_cols.get(ty_date)
        gc = gpv_cols.get(ty_date)
        if sc is None:
            continue
        blk = pd.DataFrame({
            "day": day_idx,
            "division": t["division"], "cat1": t["cat1"],
            "cat3": t["cat3"], "crp": t["crp"],
            "tgt_sales": pd.to_numeric(t[sc], errors="coerce").fillna(0.0),
            "tgt_gpv": pd.to_numeric(t[gc], errors="coerce").fillna(0.0),
        })
        recs.append(blk)

    tgt = pd.concat(recs, ignore_index=True)
    tgt = tgt.groupby(["day", "division", "cat1", "cat3", "crp"],
                      as_index=False)[["tgt_sales", "tgt_gpv"]].sum()
    log(f"Targets: {t.shape[0]:,} cat3 rows x {n_days} days | "
        f"month total Sales {tgt.tgt_sales.sum()/1e7:,.2f} cr, "
        f"GPV {tgt.tgt_gpv.sum()/1e7:,.2f} cr")
    return tgt


# =============================================================================
# 3. SALES FACTS
# =============================================================================
# Accepted header names for each field, first match wins. The current export
# uses "Outlet Code" / "Product Code" / "POS Sales"; the older names are kept
# so a file from the previous export format still loads. Matching ignores
# case and surrounding spaces. Columns not listed here (Profit %, Basket Size,
# POS Discount) are ignored - those ratios are recomputed from the sums.
SALES_COLS = {
    "date":   ["Date"],
    "outlet": ["Outlet Code", "Outlet"],
    "code":   ["Product Code", "Product", "Article Code"],
    "sales":  ["POS NSI"],
    "gpv":    ["POS GPV"],
    "ff":     ["Footfall"],
    "qty":    ["POS Sales", "POS Sales Qty", "POS Qty"],
}


def resolve_sales_columns(columns, label):
    """Map each field to the header actually present in the file."""
    lookup = {str(c).strip().lower(): c for c in columns}
    chosen, missing = {}, []
    for field, names in SALES_COLS.items():
        hit = next((lookup[n.lower()] for n in names if n.lower() in lookup), None)
        if hit is None:
            missing.append(f"{field} (expected one of: {', '.join(names)})")
        else:
            chosen[hit] = field
    if missing:
        halt(f"{label} sales file is missing column(s): " + "; ".join(missing) +
             f". Headers found: {list(columns)}")
    return chosen


def load_sales(path_key, label):
    path = ipath(path_key)
    if os.path.splitext(path)[1].lower() in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(path, engine="calamine")
    else:
        df = pd.read_csv(path, low_memory=False)
    df.columns = [str(c).strip() for c in df.columns]
    mapping = resolve_sales_columns(df.columns, label)
    df = df[list(mapping)].rename(columns=mapping)
    log(f"{label}: columns " + ", ".join(f"{src} -> {dst}" for src, dst in mapping.items()))

    raw_date = df["date"].copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce",
                                format="mixed", dayfirst=False).dt.normalize()
    bad_mask = df["date"].isna()
    if bad_mask.any():
        n_bad = int(bad_mask.sum())
        bad_vals = raw_date[bad_mask]
        # Tolerate a small number of trailing junk rows (e.g. a "Total" /
        # "Grand Total" footer row some exports append) - drop them instead
        # of hard-stopping. Anything else still halts.
        # .str accessors propagate NA, so a genuinely blank date cell would
        # fall through isin() and halt. Flatten it to "" first.
        txt = (bad_vals.astype(str).fillna("")
               .str.strip().str.lower().fillna(""))
        looks_like_footer = txt.isin(
            {"", "nan", "nat", "none", "total", "totals", "grand total",
             "grand total:", "sum", "sum total", "subtotal"})
        if n_bad <= 5 and looks_like_footer.all():
            log(f"{label}: dropped {n_bad} footer/summary row(s) with "
                f"non-date value(s): {bad_vals.tolist()}")
            df = df[~bad_mask].copy()
        else:
            preview = bad_vals.head(10).to_string()
            halt(f"{label}: {n_bad} unparseable date value(s). Offending "
                 f"row(s) (up to 10 shown):\n{preview}")

    df["outlet"] = df["outlet"].fillna("").astype(str).str.strip().str.upper()
    df["code"] = pd.to_numeric(df["code"], errors="coerce")
    if df["code"].isna().any():
        halt(f"{label}: {int(df['code'].isna().sum())} non-numeric article code(s).")
    df["code"] = df["code"].astype("int64")

    for m in METRIC_COLS:
        df[m] = pd.to_numeric(df[m], errors="coerce").fillna(0.0)

    log(f"{label} sales: {len(df):,} rows | {df.outlet.nunique():,} outlets | "
        f"{df.code.nunique():,} SKUs | {df.date.nunique()} day(s) | "
        f"NSI {df.sales.sum()/1e7:,.2f} cr")
    return df


# =============================================================================
# 4. ASSEMBLY
# =============================================================================
def build_fact(ty, ly, day_pairs, ka, person):
    """One long fact table with a year tag and a shared day index."""
    ty_idx = {d: i for i, (d, _) in enumerate(day_pairs, start=1)}
    ly_idx = {d: i for i, (_, d) in enumerate(day_pairs, start=1)}

    ty = ty.copy(); ly = ly.copy()
    ty["day"] = ty["date"].map(ty_idx)
    ly["day"] = ly["date"].map(ly_idx)

    drop_ty = int(ty["day"].isna().sum())
    drop_ly = int(ly["day"].isna().sum())
    if drop_ty:
        log(f"  ! TY: {drop_ty:,} rows outside the mapped calendar - excluded.")
    if drop_ly:
        log(f"  ! LY: {drop_ly:,} rows outside the SPLY calendar - excluded.")
    ty = ty.dropna(subset=["day"]); ly = ly.dropna(subset=["day"])
    ty["day"] = ty["day"].astype(int); ly["day"] = ly["day"].astype(int)

    ty["yr"] = "ty"; ly["yr"] = "ly"
    fact = pd.concat([ty, ly], ignore_index=True)

    fact = fact.merge(ka, on="code", how="left")
    orphan = fact["division"].isna()
    if orphan.any():
        share = fact.loc[orphan, "sales"].sum() / max(fact["sales"].sum(), 1)
        log(f"  ! {int(orphan.sum()):,} rows ({share:.3%} of NSI) have no "
            f"Key_Article match -> bucketed as {CONFIG['UNASSIGNED_LABEL']}.")
        for c in ("division", "cat1", "cat3", "brand"):
            fact[c] = fact[c].fillna(CONFIG["UNASSIGNED_LABEL"])
        fact["name"] = fact["name"].fillna("")

    fact = fact.merge(person, on=["division", "cat1", "cat3"], how="left")
    u = fact["crp"].isna()
    if u.any():
        share = fact.loc[u, "sales"].sum() / max(fact["sales"].sum(), 1)
        log(f"  ! {int(u.sum()):,} rows ({share:.3%} of NSI) have no CRP "
            f"mapping -> {CONFIG['UNASSIGNED_LABEL']}.")
    fact["crp"] = fact["crp"].fillna(CONFIG["UNASSIGNED_LABEL"])
    fact["head"] = fact["head"].fillna(CONFIG["UNASSIGNED_LABEL"])
    return fact


def resolve_same_store(fact, ss_codes):
    ty_out = set(fact.loc[fact.yr == "ty", "outlet"].unique())
    ly_out = set(fact.loc[fact.yr == "ly", "outlet"].unique())
    if CONFIG["SAME_STORE_REQUIRE_BOTH_YEARS"]:
        ss = ss_codes & ty_out & ly_out
    else:
        ss = set(ss_codes)
    log(f"Same-store universe: {len(ss):,} outlets "
        f"(list {len(ss_codes):,}; traded TY {len(ss_codes & ty_out):,}, "
        f"LY {len(ss_codes & ly_out):,}) | all-store TY {len(ty_out):,}, "
        f"LY {len(ly_out):,}")
    return ss


# ---------------------------------------------------------------------------
def pivot_years(df, keys):
    """Group by keys x year, then widen into *_ty / *_ly columns."""
    g = df.groupby(keys + ["yr"], as_index=False)[METRIC_COLS].sum()
    wide = g.pivot_table(index=keys, columns="yr", values=METRIC_COLS,
                         fill_value=0.0)
    wide.columns = [f"{m}_{y}" for m, y in wide.columns]
    for m in METRIC_COLS:
        for y in ("ty", "ly"):
            if f"{m}_{y}" not in wide.columns:
                wide[f"{m}_{y}"] = 0.0
    return wide.reset_index()


def attach_targets(df, tgt, keys, enabled):
    cols = ["tgt_sales", "tgt_gpv"]
    if not enabled or tgt is None or not keys:
        for c in cols:
            df[c] = np.nan
        return df
    agg = tgt.groupby(keys, as_index=False)[cols].sum()
    out = df.merge(agg, on=keys, how="left")
    for c in cols:
        out[c] = out[c].fillna(0.0)
    return out


def rnd(v, nd):
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return None
    return round(float(v), nd)


def pack(df, dim_cols, with_target):
    """Compact array-of-arrays payload: [*dims, sales_ty, sales_ly, gpv_ty,
    gpv_ly, ff_ty, ff_ly, qty_ty, qty_ly, tgt_sales, tgt_gpv]"""
    m, q = CONFIG["ROUND_MONEY"], CONFIG["ROUND_QTY"]
    rows = []
    for r in df.itertuples(index=False):
        d = r._asdict()
        row = [d[c] for c in dim_cols]
        row += [rnd(d["sales_ty"], m), rnd(d["sales_ly"], m),
                rnd(d["gpv_ty"], m), rnd(d["gpv_ly"], m),
                int(round(d["ff_ty"])), int(round(d["ff_ly"])),
                rnd(d["qty_ty"], q), rnd(d["qty_ly"], q)]
        if with_target:
            row += [rnd(d.get("tgt_sales"), m), rnd(d.get("tgt_gpv"), m)]
        else:
            row += [None, None]
        rows.append(row)
    return rows


SCHEMA_TAIL = ["sales_ty", "sales_ly", "gpv_ty", "gpv_ly",
               "ff_ty", "ff_ly", "qty_ty", "qty_ly",
               "tgt_sales", "tgt_gpv"]


# =============================================================================
# 5. BUILD EACH PAYLOAD
# =============================================================================
def daily_wide(f, keys):
    """Group by day x keys x year, widened into *_ty / *_ly columns."""
    g = f.groupby(["day"] + keys + ["yr"], as_index=False)[METRIC_COLS].sum()
    w = g.pivot_table(index=["day"] + keys, columns="yr", values=METRIC_COLS,
                      fill_value=0.0)
    w.columns = [f"{m}_{y}" for m, y in w.columns]
    for m in METRIC_COLS:
        for y in ("ty", "ly"):
            if f"{m}_{y}" not in w.columns:
                w[f"{m}_{y}"] = 0.0
    return w.reset_index()


def pack_daily(df, key_cols, key_index, with_target):
    """[key_id, day, sales_ty, sales_ly, gpv_ty, gpv_ly, ff_ty, ff_ly,
        qty_ty, qty_ly, tgt_sales, tgt_gpv] - money kept to 2 dp so browser
    totals reconcile exactly with the source workbooks."""
    q = CONFIG["ROUND_QTY"]
    rows = []
    for r in df.itertuples(index=False):
        d = r._asdict()
        kid = key_index[tuple(str(d[c]) for c in key_cols)]
        m = CONFIG["ROUND_MONEY"]
        row = [kid, int(d["day"]),
               rnd(num0(d["sales_ty"]), m), rnd(num0(d["sales_ly"]), m),
               rnd(num0(d["gpv_ty"]), m), rnd(num0(d["gpv_ly"]), m),
               int(round(num0(d["ff_ty"]))), int(round(num0(d["ff_ly"]))),
               rnd(num0(d["qty_ty"]), q), rnd(num0(d["qty_ly"]), q)]
        if with_target:
            row += [rnd(num0(d.get("tgt_sales")), m), rnd(num0(d.get("tgt_gpv")), m)]
        else:
            row += [None, None]
        rows.append(row)
    return rows


def num0(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return 0.0
    return 0.0 if (np.isnan(f) or np.isinf(f)) else f


DAILY_TAIL = ["sales_ty", "sales_ly", "gpv_ty", "gpv_ly", "ff_ty", "ff_ly",
              "qty_ty", "qty_ly", "tgt_sales", "tgt_gpv"]


def build_daily_cat(fact, scopes):
    """Day x Division x Cat1 x Cat3 x Brand x CRP - actuals only.

    Targets are deliberately absent: they are published at Category 3, and
    attaching them here would repeat one category target on every brand.
    """
    keys = ["division", "cat1", "cat3", "brand", "crp", "head"]
    frames = {scope: daily_wide(fact[mask], keys) for scope, mask in scopes.items()}

    key_rows, key_index = [], {}
    for scope in frames:
        for r in frames[scope][keys].drop_duplicates().itertuples(index=False):
            t = tuple(str(x) for x in r)
            if t not in key_index:
                key_index[t] = len(key_rows)
                key_rows.append(list(t))

    out = {}
    for scope, w in frames.items():
        out[scope] = pack_daily(w, keys, key_index, False)
        log(f"  daily_cat[{scope}]: {len(out[scope]):,} category-brand-day rows")
    log(f"  daily_cat: {len(key_rows):,} distinct keys")
    return {"keys": keys, "key_rows": key_rows,
            "schema": ["key", "day"] + DAILY_TAIL, "data": out}


def build_daily_target(tgt, scopes, tgt_on, head_by_crp, fact):
    """Day x Division x Cat1 x Cat3 x CRP - targets only, at their own grain."""
    keys = ["division", "cat1", "cat3", "crp", "head"]
    head_of = dict(head_by_crp)
    agg = tgt.groupby(["day", "division", "cat1", "cat3", "crp"],
                      as_index=False)[["tgt_sales", "tgt_gpv"]].sum()
    agg["head"] = agg["crp"].map(head_of).fillna(CONFIG["UNASSIGNED_LABEL"])

    key_rows, key_index = [], {}
    for r in agg[keys].drop_duplicates().itertuples(index=False):
        t = tuple(str(x) for x in r)
        if t not in key_index:
            key_index[t] = len(key_rows)
            key_rows.append(list(t))

    rows = []
    m = CONFIG["ROUND_MONEY"]
    for r in agg.itertuples(index=False):
        d = r._asdict()
        rows.append([key_index[tuple(str(d[c]) for c in keys)], int(d["day"]),
                     rnd(num0(d["tgt_sales"]), m), rnd(num0(d["tgt_gpv"]), m)])

    out = {}
    for scope in scopes:
        out[scope] = rows if tgt_on[scope] else []
        log(f"  daily_target[{scope}]: {len(out[scope]):,} target-days")
    return {"keys": keys, "key_rows": key_rows,
            "schema": ["key", "day", "tgt_sales", "tgt_gpv"], "data": out}


OC_COLS = ["sales_ty", "sales_ly", "gpv_ty", "gpv_ly",
           "ff_ty", "ff_ly", "qty_ty", "qty_ly"]
OC_SCALE = [1000, 1000, 1000, 1000, 1, 1, 1000, 1000]   # 3 dp for money and qty
INT32_MAX = 2_147_483_647


def build_outlet_cat3(fact, stores, ss_codes):
    """Day x Outlet x Category 3 - the single source for everything that is
    split by outlet: the outlet table, the outlet-filtered view, the
    Outlet division / Area / Leader / Zonal tabs and the degrowth card.

    Being day-wise, it follows any date range. It is stored once, for all
    outlets; same-store is a per-outlet flag, because an outlet's figures are
    identical in both universes.

    Encoding: plain JSON would be ~190 MB for a month, so the rows are packed
    as column arrays of 32-bit integers (money and quantity in thousandths,
    exact to the taka when summed), gzip-compressed and base64-encoded. The
    page decompresses them with the browser's built-in DecompressionStream.
    """
    import numpy as np
    keys = ["outlet", "division", "cat1", "cat3"]
    w = daily_wide(fact, keys)
    smap = stores.set_index("outlet")

    def attr(o, col, dflt=CONFIG["UNASSIGNED_LABEL"]):
        v = smap[col].get(o) if o in smap.index else None
        return dflt if v is None or (isinstance(v, float) and np.isnan(v)) else str(v)

    oc, outs = pd.factorize(w["outlet"].astype(str))
    path = (w["division"].astype(str) + "\u0001" + w["cat1"].astype(str) +
            "\u0001" + w["cat3"].astype(str))
    cc, cats = pd.factorize(path)
    if len(outs) > 65535 or len(cats) > 65535:
        halt("outlet_cat3: more than 65,535 outlets or categories - widen the index type.")

    vals = []
    for col, sc in zip(OC_COLS, OC_SCALE):
        v = np.rint(w[col].astype(float).to_numpy() * sc)
        if np.abs(v).max(initial=0) > INT32_MAX:
            halt(f"outlet_cat3: {col} exceeds the 32-bit range at scale {sc}.")
        vals.append(v.astype("<i4"))
    outlets = [[o, attr(o, "store_name", ""), attr(o, "format"), attr(o, "geo"),
                attr(o, "zone"), attr(o, "cluster"), attr(o, "leader"),
                attr(o, "zonal"), attr(o, "area"), 1 if o in ss_codes else 0]
               for o in outs]
    log(f"  outlet_cat3: {len(w):,} day-outlet-category rows | "
        f"{len(outs):,} outlets x {len(cats):,} category paths")
    return {"outlets": outlets, "cats": list(cats),
            "rows": {"oi": oc.astype("<u2"), "ci": cc.astype("<u2"),
                     "day": w["day"].astype("<u1").to_numpy(), "vals": vals}}


def _pack_chunk(rows, lo, hi):
    """Column-major block: 8 int32 value columns, then outlet (u16),
    category (u16) and day (u8) - int32 first so every view is aligned."""
    parts = [v[lo:hi].tobytes() for v in rows["vals"]]
    parts += [rows["oi"][lo:hi].tobytes(), rows["ci"][lo:hi].tobytes(),
              rows["day"][lo:hi].tobytes()]
    return base64.b64encode(gzip.compress(b"".join(parts), 6, mtime=0)).decode("ascii")


def write_outlet_cat3(month_dir, month_id, oc):
    """Write the binary payload as a header part plus as many compressed
    chunks as it takes to keep each file under MAX_FILE_MB."""
    limit = int(CONFIG["MAX_FILE_MB"] * 1_000_000) - 4096
    n = len(oc["rows"]["oi"])
    probe = min(n, 200_000)
    bpr = max(1.0, len(_pack_chunk(oc["rows"], 0, probe)) / max(probe, 1)) if n else 1.0
    step = max(1, int(limit * 0.85 / bpr))
    chunks, lo = [], 0
    while lo < n:
        hi = min(n, lo + step)
        b64 = _pack_chunk(oc["rows"], lo, hi)
        while len(b64) > limit and hi - lo > 1:        # rare: shrink and retry
            hi = lo + (hi - lo) // 2
            b64 = _pack_chunk(oc["rows"], lo, hi)
        chunks.append({"rows": hi - lo, "b64": b64})
        lo = hi

    head = {"binary": True, "outlets": oc["outlets"], "cats": oc["cats"],
            "cols": OC_COLS, "scale": OC_SCALE,
            "outlet_cols": ["outlet", "store_name", "format", "geo", "zone",
                            "cluster", "leader", "zonal", "area", "is_ss"],
            "layout": "int32 x 8 cols, then u16 outlet, u16 category, u8 day; little-endian"}
    of = 1 + len(chunks)
    names = []
    for i, v in enumerate([head] + [{"chunk": c} for c in chunks]):
        stem = "outlet_cat3" if of == 1 else f"outlet_cat3.part{i + 1}"
        body = _dumps({"m": month_id, "n": "outlet_cat3", "p": i, "of": of, "v": v})
        path = os.path.join(month_dir, stem + ".js")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("// Auto-generated by build_dashboard_data.py - do not edit.\n")
            fh.write("(window.SHWAPNO_REG=window.SHWAPNO_REG||[]).push(" + body + ");\n")
        WRITTEN.append(stem + ".js")
        names.append(stem + ".js")
        log(f"  wrote {month_id}/{stem}.js ({os.path.getsize(path) / 1e6:,.1f} MB)")
    log(f"  outlet_cat3: {n:,} rows in {len(chunks)} compressed chunk(s); "
        "no .json copy (binary payload)")
    return names


def build_sku(fact, scopes):
    out = {}
    names = (fact[["code", "name"]].drop_duplicates("code")
             .set_index("code")["name"].to_dict())
    for scope, mask in scopes.items():
        f = fact[mask]
        keys = ["division", "cat1", "cat3", "brand", "code"]
        w = pivot_years(f, keys)
        w["name"] = w["code"].map(names).fillna("")
        dims = ["division", "cat1", "cat3", "brand", "code", "name"]
        w = w[dims + SCHEMA_TAIL[:-2]]
        for c in ("tgt_sales", "tgt_gpv"):
            w[c] = np.nan
        rows = pack(w, dims, False)
        out[scope] = rows
        log(f"  sku[{scope}]: {len(rows):,} rows")
    return {"schema": dims + SCHEMA_TAIL, "data": out}


# =============================================================================
# 6. EMIT
# =============================================================================
WRITTEN = []          # file names written this run, relative to the month folder


def _dumps(obj):
    return json.dumps(obj, ensure_ascii=False, allow_nan=False,
                      separators=(",", ":"),
                      indent=2 if CONFIG["PRETTY_JSON"] else None)


def _split(payload, limit):
    """Split a payload into parts no larger than `limit` bytes.

    Part 0 carries every header field plus empty row lists; later parts carry
    only row slices under "data". The browser concatenates them back in order.
    Payloads without a row block ("data") must fit in one file.
    """
    whole = _dumps(payload)
    if len(whole.encode("utf-8")) <= limit:
        return [payload]
    data = payload.get("data")
    if not isinstance(data, dict):
        halt(f"Payload of {len(whole.encode()) / 1e6:.1f} MB has no row block "
             "to split on and exceeds MAX_FILE_MB.")

    head = {k: v for k, v in payload.items() if k != "data"}
    head["data"] = {sc: [] for sc in data}
    if len(_dumps(head).encode("utf-8")) > limit:
        halt("Payload header alone exceeds MAX_FILE_MB - raise the limit.")

    parts, base = [head], 200          # 200 bytes of envelope headroom
    cur, size = {sc: [] for sc in data}, base
    for sc, rows in data.items():
        for r in rows:
            rs = len(_dumps(r).encode("utf-8")) + 1
            if size + rs > limit and any(cur.values()):
                parts.append({"data": cur})
                cur, size = {sc2: [] for sc2 in data}, base
            cur[sc].append(r)
            size += rs
    if any(cur.values()):
        parts.append({"data": cur})
    return parts


def write_payload(month_dir, month_id, name, payload, sub=""):
    """Write one payload as <name>.js (+ .json), splitting into
    <name>.partN.js when it would exceed MAX_FILE_MB.

    Every file is self-describing - {"m": month, "n": name, "p": part,
    "of": total, "v": value} - so the browser can reassemble without guessing,
    and a JSON consumer can tell which month and part it is holding.
    Returns the .js file names in load order."""
    # Decimal megabytes: 24 means 24,000,000 bytes, safely under GitHub's
    # 25 MB (and 25 MiB) web-upload ceiling.
    limit = int(CONFIG["MAX_FILE_MB"] * 1_000_000) - 4096     # JS wrapper room
    parts = _split(payload, limit)
    out_dir = os.path.join(month_dir, sub) if sub else month_dir
    os.makedirs(out_dir, exist_ok=True)
    rel = (sub + "/") if sub else ""
    n = len(parts)
    js_names = []
    for i, v in enumerate(parts):
        stem = name if n == 1 else f"{name}.part{i + 1}"
        env = {"m": month_id, "n": name, "p": i, "of": n, "v": v}
        body = _dumps(env)

        js = os.path.join(out_dir, stem + ".js")
        with open(js, "w", encoding="utf-8") as fh:
            fh.write("// Auto-generated by build_dashboard_data.py - do not edit.\n")
            fh.write("(window.SHWAPNO_REG=window.SHWAPNO_REG||[]).push(" + body + ");\n")
        WRITTEN.append(rel + stem + ".js")
        js_names.append(rel + stem + ".js")
        mb = os.path.getsize(js) / 1e6
        if mb > CONFIG["MAX_FILE_MB"] + 0.5:
            halt(f"{stem}.js is {mb:.1f} MB after splitting - lower MAX_FILE_MB.")
        msg = f"  wrote {month_id}/{rel}{stem}.js ({mb:,.1f} MB)"

        if CONFIG["EMIT_JSON"]:
            jp = os.path.join(out_dir, stem + ".json")
            with open(jp, "w", encoding="utf-8") as fh:
                fh.write(body)
            WRITTEN.append(rel + stem + ".json")
            msg += f" + .json ({os.path.getsize(jp) / 1e6:,.1f} MB)"
        log(msg)
    if n > 1:
        log(f"  {name}: split into {n} parts to stay under {CONFIG['MAX_FILE_MB']} MB")
    return js_names


def prune_month_dir(month_dir):
    """Leave only this run's files in the month folder. Each month's folder is
    replaced wholesale on every rebuild; other months are never touched."""
    if not CONFIG["PRUNE_OUTPUT_DIR"]:
        return
    keep, removed = set(WRITTEN), []
    listing = [n for n in os.listdir(month_dir)]
    sub = os.path.join(month_dir, "sku_outlet")          # retired payload
    if os.path.isdir(sub):
        listing += ["sku_outlet/" + n for n in os.listdir(sub)]
    for name in sorted(listing):
        full = os.path.join(month_dir, name)
        if not os.path.isfile(full) or name in keep:
            continue
        if os.path.splitext(name)[1].lower() not in (".js", ".json"):
            continue
        os.remove(full)
        removed.append(name)
    if os.path.isdir(sub) and not os.listdir(sub):
        os.rmdir(sub)
    if removed:
        shown = ", ".join(removed[:6]) + (f" and {len(removed) - 6} more" if len(removed) > 6 else "")
        log(f"  pruned {len(removed)} stale file(s) from this month: {shown}")


def prune_legacy_root(root):
    """Remove single-month payloads left in the data root by older builds."""
    if not CONFIG["PRUNE_OUTPUT_DIR"]:
        return
    removed = []
    for name in sorted(os.listdir(root)):
        full = os.path.join(root, name)
        if not os.path.isfile(full) or name.startswith("manifest."):
            continue
        if name.startswith("data_") and name.endswith((".js", ".json")):
            os.remove(full)
            removed.append(name)
    if removed:
        log(f"  removed {len(removed)} pre-monthly file(s) from the data root")


def update_manifest(root, month_id, label, js_files, lazy=None):
    """manifest.json is the list of months the dashboard can show. This month's
    entry is replaced; every other month's entry is kept as it was."""
    path = os.path.join(root, "manifest.json")
    man = {"months": []}
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as fh:
                man = json.load(fh)
        except (OSError, ValueError):
            log("  ! manifest.json unreadable - rebuilding it from this month only.")
            man = {"months": []}

    months = [m for m in man.get("months", []) if m.get("id") != month_id]
    # Drop entries whose folder has been deleted by hand.
    months = [m for m in months if os.path.isdir(os.path.join(root, m["id"]))]
    entry = {"id": month_id, "label": label,
             "files": [f"{month_id}/{f}" for f in js_files],
             "built": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    if lazy:
        entry["lazy"] = lazy
    months.append(entry)
    months.sort(key=lambda m: m["id"])

    keep_n = CONFIG.get("KEEP_MONTHS")
    if keep_n and len(months) > keep_n:
        for old in months[:-keep_n]:
            d = os.path.join(root, old["id"])
            # Only a YYYY-MM folder, and only generated files inside it.
            if re.fullmatch(r"\d{4}-\d{2}", old["id"]) and os.path.isdir(d):
                for f in os.listdir(d):
                    if f.endswith((".js", ".json")):
                        os.remove(os.path.join(d, f))
                try:
                    os.rmdir(d)
                except OSError:
                    log(f"  ! {old['id']}/ kept: it holds files this script did not write.")
            log(f"  KEEP_MONTHS={keep_n}: removed {old['label']}")
        months = months[-keep_n:]
    man = {"months": months, "latest": months[-1]["id"]}

    with open(path, "w", encoding="utf-8") as fh:
        fh.write(_dumps(man))
    with open(os.path.join(root, "manifest.js"), "w", encoding="utf-8") as fh:
        fh.write("// Auto-generated - lists the months this dashboard can show.\n")
        fh.write("window.SHWAPNO_MANIFEST = " + _dumps(man) + ";\n")
    log(f"  manifest: {len(months)} month(s) -> "
        + ", ".join(m["label"] for m in months))


# =============================================================================
# MAIN
# =============================================================================
def main():
    t0 = datetime.now()
    log("=" * 68)
    log(f"SHWAPNO PERFORMANCE DASHBOARD - DATA BUILD ({CONFIG['PERIOD_LABEL']})")
    log("=" * 68)

    if not os.path.isdir(CONFIG["INPUT_DIR"]):
        halt(f"Input folder not found: {CONFIG['INPUT_DIR']}")
    missing = [f"{k} = '{CONFIG[k]}'" for k in CONFIG
               if k.startswith("FILE_") and not os.path.exists(ipath(k))]
    if missing:
        found = [os.path.relpath(f, CONFIG["INPUT_DIR"]) for f in _input_files()]
        halt("Input file(s) not found:\n  " + "\n  ".join(missing)
             + f"\nFiles in {CONFIG['INPUT_DIR']}:\n  "
             + ("\n  ".join(found) if found else "(none)")
             + "\nRename the file or change the name in CONFIG.")
    for k in [k for k in CONFIG if k.startswith("FILE_")]:
        log(f"  {k:<17} -> {os.path.relpath(ipath(k), CONFIG['INPUT_DIR'])}")

    log("-- Step 1/6  day mapping")
    day_pairs = load_day_mapping()

    log("-- Step 2/6  master data")
    ka = load_key_article()
    stores = load_zone(load_key_store())
    ss_codes = load_same_store()
    person = load_person_map()
    tgt = load_targets(day_pairs)

    log("-- Step 3/6  sales facts")
    ty = load_sales("FILE_SALES_TY", "TY")
    ly = load_sales("FILE_SALES_LY", "LY")
    fact = build_fact(ty, ly, day_pairs, ka, person)

    log("-- Step 4/6  scope resolution")

    # The SPLY file may carry the whole month while this year is only part-way
    # through. Trim BOTH sides to the days that actually have this-year data,
    # so a full last-year file can never inflate the comparison.
    loaded = sorted(fact.loc[fact.yr == "ty", "day"].unique())
    max_day = int(max(loaded)) if loaded else 0
    if max_day == 0:
        halt("No this-year sales rows fell inside the mapped calendar.")
    before = len(fact)
    extra = fact[(fact.yr == "ly") & (fact["day"] > max_day)]
    if len(extra):
        log(f"  SPLY file runs past this year: dropped {len(extra):,} rows "
            f"({extra['sales'].sum()/1e7:,.2f} cr) beyond day {max_day}.")
    fact = fact[fact["day"] <= max_day].copy()
    if len(fact) != before - len(extra):
        log(f"  Trimmed {before - len(fact) - len(extra):,} further row(s) "
            f"beyond day {max_day}.")

    day_pairs = day_pairs[:max_day]
    tgt = tgt[tgt["day"] <= max_day]

    ss = resolve_same_store(fact, ss_codes)
    scopes = {"all": pd.Series(True, index=fact.index),
              "same": fact["outlet"].isin(ss)}
    tgt_on = {"all": True,
              "same": CONFIG["TARGET_SCOPE"] == "BOTH"}
    log(f"Loaded window: day 1..{max_day} "
        f"({day_pairs[0][0]:%d %b} - {day_pairs[-1][0]:%d %b %Y}; "
        f"SPLY {day_pairs[0][1]:%d %b} - {day_pairs[-1][1]:%d %b %Y})")

    log("-- Step 5/6  aggregation")
    head_by_crp = (fact[["crp", "head"]].drop_duplicates("crp")
                   .set_index("crp")["head"].to_dict())
    daily_cat = build_daily_cat(fact, scopes)
    daily_target = build_daily_target(tgt, scopes, tgt_on, head_by_crp, fact)
    outlet_cat3 = build_outlet_cat3(fact, stores, ss)
    sku = build_sku(fact, scopes)

    # Cascading filter vocabulary: division -> cat1 -> cat3
    tree = {}
    for d, c1, c3 in (fact[["division", "cat1", "cat3"]]
                      .drop_duplicates().itertuples(index=False)):
        tree.setdefault(d, {}).setdefault(c1, set()).add(c3)
    tree = {d: {c1: sorted(c3) for c1, c3 in sorted(v.items())}
            for d, v in sorted(tree.items())}

    outlets = {}
    for scope, mask in scopes.items():
        f = fact[mask]
        outlets[scope] = {
            "ty": int(f.loc[f.yr == "ty", "outlet"].nunique()),
            "ly": int(f.loc[f.yr == "ly", "outlet"].nunique()),
        }

    meta = {
        "period_label": CONFIG["PERIOD_LABEL"],
        "ty_year": CONFIG["TY_YEAR"], "ly_year": CONFIG["LY_YEAR"],
        "days": [{"day": i,
                  "ty": t.strftime("%Y-%m-%d"), "ly": l.strftime("%Y-%m-%d"),
                  "label": t.strftime("%d %b"), "dow": t.strftime("%a"),
                  "sply_label": l.strftime("%d %b")}
                 for i, (t, l) in enumerate(day_pairs, start=1)],
        "same_store_count": len(ss),
        "same_store_rule": ("Same Store List ∩ traded in both TY and SPLY"
                            if CONFIG["SAME_STORE_REQUIRE_BOTH_YEARS"]
                            else "Same Store List as published"),
        "target_scope": CONFIG["TARGET_SCOPE"],
        "store_target_mode": CONFIG["STORE_TARGET_MODE"],
        "outlets": outlets,
        "filters": {"tree": tree},
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "build_log": LOG.copy(),
    }

    log("-- Step 6/6  emit JS + JSON")
    root = CONFIG["OUTPUT_DIR"]
    month_id = day_pairs[0][0].strftime("%Y-%m")
    month_dir = os.path.join(root, month_id)
    os.makedirs(month_dir, exist_ok=True)
    log(f"  month folder: {month_id}/  (replaces any earlier build of this month)")

    files = []
    for name, payload in [("meta", meta), ("daily_cat", daily_cat),
                          ("daily_target", daily_target), ("sku", sku)]:
        files += write_payload(month_dir, month_id, name, payload)
    files += write_outlet_cat3(month_dir, month_id, outlet_cat3)
    prune_month_dir(month_dir)
    prune_legacy_root(root)
    update_manifest(root, month_id, CONFIG["PERIOD_LABEL"], files)

    log("=" * 68)
    log(f"DONE in {(datetime.now() - t0).total_seconds():.1f}s -> "
        f"{CONFIG['OUTPUT_DIR']}")
    log("=" * 68)


if __name__ == "__main__":
    main()
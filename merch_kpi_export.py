"""
merch_kpi_export.py
===================

JSON layer for the Merchandising KPI dashboard.

It does NOT touch any Merchandising SQL / KPI calculation. It takes the
DataFrames merch_source.py already produces for A-Class, B-Class, MOH and
Inventory (MRP / CP) and:

  1. adds spoc_name from spoc_name.xlsx          (server side only)
  2. adds target_moh to the MOH report           (from Moh_targets.xlsx)
  3. appends "<SPoC> - TOTAL" rows and one "MINISO" row, using the same
     row labels OPS uses, so ops_kpi_json.py classifies them unchanged
  4. gives every period column a calendar label
  5. serialises through ops_kpi_json.py (same schema as the OPS dashboard)

Both Excel files are read here, in Python. The deployed dashboard only ever
sees the generated JSON -- it never needs access to the G:\\ drive.
"""

from __future__ import annotations

import os
from datetime import date, timedelta

import pandas as pd

import merch_kpi_json as jl
 
SPOC_FILE = os.environ.get(
    "MERCH_SPOC_FILE",
    r"G:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt"
    r"\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard"
    r"\spoc_name\spoc_name.xlsx",
)
MOH_TARGET_FILE = os.environ.get(
    "MERCH_MOH_TARGET_FILE",
    r"G:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt"
    r"\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard"
    r"\mer_kpi\Moh_targets.xlsx",
)

NO_SPOC = "(No SPoC)"

# Store code renamed in the database. The MOH targets file may still list the
# old code, so its target is also used for the new one (unless the new code has
# its own row). Deliberately NOT applied to SPoC mapping: INES has no SPoC.
STORE_CODE_ALIASES = {"INFK": "INES"}

# SPoC values that mean "no SPoC"
_NO_SPOC_VALUES = {"", "NA", "N/A", "NAN", "NONE", "NULL", "-"}

# Colour rules are data, evaluated by the dashboard (app.js statusFor).
TIER_A = {"type": "tier", "scale": 100, "green_above": 94.99, "yellow_from": 87.99}
TIER_B = {"type": "tier", "scale": 100, "green_above": 87.99, "yellow_from": 7.99}
MOH_RULE = {"type": "moh_variance"}
INV_RULE = {"type": "moh_variance", "source_report": "moh"}   # same colour as MOH

# prefix of the Excel file name in merch_source.py -> dashboard report
KPIS = [
    {"key": "a_class", "prefix": "a_class", "name": "A-Class",
     "description": "A-class SKU availability. Green above 94.99%, yellow 87.99-94.99%, orange below.",
     "rename": "availability", "rollup": "mean", "labels": "iso",
     "display": {"scale": 100, "decimals": 0, "suffix": "%"},
     "value_format": "percent0", "color_rule": TIER_A},
    {"key": "b_class", "prefix": "b_class", "name": "B-Class",
     "description": "B-class SKU availability. Green above 87.99%, yellow 79.99-87.99%, orange below.",
     "rename": "availability", "rollup": "mean", "labels": "iso",
     "display": {"scale": 100, "decimals": 0, "suffix": "%"},
     "value_format": "percent0", "color_rule": TIER_B},
    {"key": "new_sku", "prefix": "new_sku_availability", "name": "New SKU availability",
     "description": "Availability of new SKUs by store.",
     "rename": "availability", "rollup": "mean", "labels": "iso",
     "display": {"scale": 100, "decimals": 0, "suffix": "%"},
     "value_format": "percent0", "color_rule": None},
    {"key": "moh", "prefix": "moh", "name": "MOH",
     "description": "Months of hold: stock / last 30 days sales, against the store's target MOH.",
     "rename": None, "rollup": "mean", "labels": "snapshot",
     "display": {"decimals": 1}, "value_format": "decimal1",
     "color_rule": MOH_RULE, "target_column": "target_moh"},
    {"key": "inv_mrp", "prefix": "inventory_value_mrp", "name": "Inventory (MRP)",
     "description": "Inventory value at MRP, in lakhs. Coloured by the store's MOH status.",
     "rename": None, "rollup": "sum", "labels": "iso",
     "display": {"divisor": 100000, "decimals": 2}, "value_format": "lakh2",
     "color_rule": INV_RULE},
    {"key": "inv_cp", "prefix": "inventory_value_cp", "name": "Inventory (CP)",
     "description": "Inventory value at CP, in lakhs. Coloured by the store's MOH status.",
     "rename": None, "rollup": "sum", "labels": "iso",
     "display": {"divisor": 100000, "decimals": 2}, "value_format": "lakh2",
     "color_rule": INV_RULE},
]

_SHORT = {"d": ("day", 7), "w": ("week", 4), "m": ("month", 3)}


# ----------------------------------------------------------------------
# Excel mappings (read here, server side)
# ----------------------------------------------------------------------

def _read_excel(path: str, **kwargs) -> pd.DataFrame:
    df = pd.read_excel(path, **kwargs)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _find_col(df: pd.DataFrame, *needles: str):
    for col in df.columns:
        low = col.lower().replace(" ", "_")
        if all(n in low for n in needles):
            return col
    return None


def _code(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.upper()


def load_spoc_map(path: str = SPOC_FILE) -> dict:
    # keep_default_na=False: a literal "NA" must stay text, not become blank
    df = _read_excel(path, keep_default_na=False)
    code_col = _find_col(df, "store", "code") or _find_col(df, "store") or df.columns[0]
    spoc_col = _find_col(df, "spoc")
    if spoc_col is None:
        raise ValueError(f"No SPoC column found in {path}. Columns: {list(df.columns)}")
    df = df[[code_col, spoc_col]].dropna(subset=[code_col])
    df[spoc_col] = df[spoc_col].astype(str).str.strip()
    df = df[~df[spoc_col].str.upper().isin(_NO_SPOC_VALUES)]
    return dict(zip(_code(df[code_col]), df[spoc_col]))


def target_moh_from_base(base: float):
    """Target MOH rule: <1 -> +15% capped at 1.5; 1..1.5 -> unchanged; >1.5 -> -15% floored at 1.5."""
    if pd.isna(base):
        return None
    if base < 1:
        return round(min(base * 1.15, 1.5), 4)
    if base <= 1.5:
        return round(base, 4)
    return round(max(1.5, base * 0.85), 4)


def load_moh_targets(path: str = MOH_TARGET_FILE) -> dict:
    """store_code -> target MOH.

    The file's base MOH column (name contains 'base', or the first numeric
    column) is run through the target rule. If the file only has a column
    named 'target', it is taken as already final.
    """
    df = _read_excel(path)
    code_col = _find_col(df, "store", "code") or _find_col(df, "store") or df.columns[0]
    base_col = _find_col(df, "base")
    target_col = _find_col(df, "target")
    if base_col is None and target_col is None:
        numeric = [c for c in df.columns if c != code_col
                   and pd.api.types.is_numeric_dtype(df[c])]
        if not numeric:
            raise ValueError(f"No MOH column found in {path}. Columns: {list(df.columns)}")
        base_col = numeric[0]

    values = pd.to_numeric(df[base_col or target_col], errors="coerce")
    if base_col is not None:
        values = values.map(target_moh_from_base)
    targets = dict(zip(_code(df[code_col]), values))
    for old, new in STORE_CODE_ALIASES.items():
        if old in targets and new not in targets:
            targets[new] = targets[old]
    return targets


# ----------------------------------------------------------------------
# Rows: SPoC + totals
# ----------------------------------------------------------------------

def _aggregate(frame: pd.DataFrame, cols: list, how: str) -> dict:
    out = {}
    for col in cols:
        s = pd.to_numeric(frame[col], errors="coerce")
        out[col] = s.sum(min_count=1) if how == "sum" else s.mean()
    return out


def add_spoc_and_totals(df: pd.DataFrame, spoc_map: dict, how: str) -> pd.DataFrame:
    df = df.copy()
    df["spoc_name"] = _code(df["store_code"]).map(spoc_map).fillna(NO_SPOC)

    metric_cols = [c for c in df.columns
                   if jl.parse_metric_key(c)[0] in ("day", "week", "month", "target")]
    rows = []
    ordered = df.sort_values(["spoc_name", "store_code"])
    for spoc, grp in ordered[ordered["spoc_name"] != NO_SPOC].groupby("spoc_name", sort=True):
        rows.append(grp)
        total = {"store_code": f"{spoc} - TOTAL", "store_name": f"{spoc} - TOTAL",
                 "spoc_name": None}
        total.update(_aggregate(grp, metric_cols, how))
        rows.append(pd.DataFrame([total]))
    # stores without a SPoC: after every SPoC group, no SPoC total, still in MINISO
    rows.append(ordered[ordered["spoc_name"] == NO_SPOC])
    grand = {"store_code": "MINISO", "store_name": "MINISO", "spoc_name": None}
    grand.update(_aggregate(df, metric_cols, how))
    rows.append(pd.DataFrame([grand]))
    return pd.concat(rows, ignore_index=True)


# ----------------------------------------------------------------------
# Labels
# ----------------------------------------------------------------------

def _month_start(today: date, back: int) -> date:
    total = today.year * 12 + today.month - 1 - back
    return date(total // 12, total % 12 + 1, 1)


def build_labels(columns, style: str, today: date) -> dict:
    labels = {}
    dow = (today.weekday() + 1) % 7          # Postgres DOW: Sunday = 0
    for col in columns:
        gran, idx, _metric = jl.parse_metric_key(col)
        if gran == "day":
            labels[col] = (today - timedelta(days=idx)).strftime("%d-%b-%Y")
        elif gran == "week" and style == "snapshot":
            snap = today - timedelta(days=dow) - timedelta(days=7 * (idx - 1))
            labels[col] = f"As on {snap:%d-%b-%Y}"
        elif gran == "week":                 # ISO week of (today - idx weeks)
            ref = today - timedelta(weeks=idx)
            mon = ref - timedelta(days=ref.weekday())
            sun = mon + timedelta(days=6)
            labels[col] = f"{mon:%d-%b} to {sun:%d-%b-%Y}"
        elif gran == "month":
            labels[col] = _month_start(today, idx).strftime("%b-%Y")
    return labels


# ----------------------------------------------------------------------
# Export
# ----------------------------------------------------------------------

def _pick(collected: dict, prefix: str):
    for file_name, item in collected.items():
        if file_name.lower().startswith(prefix):
            return item
    raise KeyError(f"report '{prefix}' was not produced in this run")


def _short_to_long(df: pd.DataFrame, metric: str) -> pd.DataFrame:
    """d1 -> day_1_<metric>, w1 -> week_1_<metric>, m1 -> month_1_<metric>."""
    rename = {}
    for col in df.columns:
        if len(col) == 2 and col[0] in _SHORT and col[1].isdigit():
            rename[col] = f"{_SHORT[col[0]][0]}_{col[1]}_{metric}"
    return df.rename(columns=rename)


def build_blocks(collected: dict, today: date | None = None):
    """collected: {excel_file_name: (DataFrame, excel_path)} -> (blocks, notes)."""
    today = today or date.today()
    spoc_map = load_spoc_map()
    moh_targets = load_moh_targets()
    notes = []

    names = {}
    for item in collected.values():
        frame = item[0]
        if "store_name" in frame.columns:
            names.update(dict(zip(_code(frame["store_code"]), frame["store_name"])))

    blocks = {}
    for kpi in KPIS:
        raw, excel_path = _pick(collected, kpi["prefix"])
        df = raw.drop(columns=["update_date"], errors="ignore").copy()
        if kpi["rename"]:
            df = _short_to_long(df, kpi["rename"])
        if "store_name" not in df.columns:
            df.insert(1, "store_name", _code(df["store_code"]).map(names))
        df["store_code"] = df["store_code"].astype(str).str.strip()

        # renamed store: drop the old code's empty row when the new code exists
        present = set(_code(df["store_code"]))
        stale = [old for old, new in STORE_CODE_ALIASES.items() if new in present]
        if stale:
            df = df[~_code(df["store_code"]).isin(stale)].copy()

        if kpi["key"] == "moh":
            df["target_moh"] = _code(df["store_code"]).map(moh_targets)
            missing = df.loc[df["target_moh"].isna(), "store_code"].tolist()
            if missing:
                notes.append(f"MOH: no target for {len(missing)} store(s): {', '.join(missing[:10])}")

        df = add_spoc_and_totals(df, spoc_map, kpi["rollup"])
        ordered = ["store_code", "store_name", "spoc_name"] + [
            c for c in df.columns if c not in ("store_code", "store_name", "spoc_name")]
        df = df[ordered]

        labels = build_labels(df.columns, kpi["labels"], today)
        config = {"display_name": kpi["name"], "description": kpi["description"],
                  "value_format": kpi["value_format"],
                  "target_column": kpi.get("target_column")}
        block = jl.build_report_block(kpi["key"], config, df, labels, excel_path)
        if kpi["color_rule"]:
            block["color_rule"] = kpi["color_rule"]
        block["display"] = kpi["display"]
        blocks[kpi["key"]] = block
    return blocks, notes


def _reporting_date(blocks: dict, today: date) -> date:
    """Freshest day that EVERY report has data for (MINISO total row)."""
    stalest = None
    for block in blocks.values():
        mini = next(r for r in block["data"] if r["_row_type"] == "miniso_total")
        idxs = [p["index"] for p in block["periods"].get("day", [])
                if mini.get(p["key"]) is not None]
        if idxs:
            stalest = min(idxs) if stalest is None else max(stalest, min(idxs))
    return today - timedelta(days=stalest or 1)


def export(collected: dict, json_path: str, today: date | None = None) -> list:
    """Build, validate and atomically write the dashboard JSON. Returns warnings."""
    today = today or date.today()
    blocks, notes = build_blocks(collected, today)
    payload = jl.build_payload(blocks, _reporting_date(blocks, today), [])
    warnings = jl.validate_payload(payload, list(blocks))
    jl.write_atomic(json_path, payload)
    return notes + warnings
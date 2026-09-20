"""
Production Analytics — data layer
================================================
All Excel reading, cleaning, normalization, quality flagging and metric
calculation lives here. app.py only draws.

Keeping this separate means the numbers can be tested without launching
Streamlit:  python validate.py
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.json"
EXCEL_PATH = BASE_DIR / "Production_Report.xlsx"
SHEET_NAME = "Production_Data"

# Columns the dashboard depends on. Missing ones are reported, not guessed.
REQUIRED_COLUMNS = [
    "SR.NO", "COIL NO", "GRADE", "P. CARD NO", "OUTWARD DATE", "RM/SIZE",
    "SLITTING OUT WT (TONS)", "SLITTING IN WT (TONS)", "SHOT WT (TONS)",
    "EXCESS WT (TONS)", "TRIMING (TONS)", "SCRAP (TONS)", "TOTAL SCRAP (TONS)",
    "INWARD DATE", "OPRETAR NAME", "MC NO", "MATERIAL TYPE",
]

WEIGHT_COLUMNS = {
    "inputWt": "SLITTING IN WT (TONS)",
    "outputWt": "SLITTING OUT WT (TONS)",
    "shotWt": "SHOT WT (TONS)",
    "excessWt": "EXCESS WT (TONS)",
    "trimWt": "TRIMING (TONS)",
    "scrapWt": "SCRAP (TONS)",
    "totalScrapWt": "TOTAL SCRAP (TONS)",
}


# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
def load_config() -> dict:
    """Read config.json. Keys starting with '_comment' are documentation."""
    with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
        return json.load(fh)


# ---------------------------------------------------------------------------
# NORMALIZERS
# ---------------------------------------------------------------------------
# Ordered, anchored patterns. Anchoring on word boundaries means a new grade
# string can no longer be swallowed by an accidental substring match
# (the previous `"GI" in s` test would have caught anything containing "gi").
GRADE_RULES: list[tuple[str, str]] = [
    (r"\bBRASS\b",                      "Brass"),
    (r"\bBER[IY]",                      "Beryllium Copper"),
    (r"\bCOPPER\b",                     "Copper"),
    (r"\bNIFE\b",                       "NIFE"),
    (r"\bCRC\b|\bSIST\b",               "CRC"),
    (r"\bAL+U?M[IY]",                   "Aluminium"),
    (r"\bGI\b",                         "GI"),
    (r"\bSIGMA\b",                      "Sigma"),
    (r"^PB\b|\bPB\s+(GR|IV)",           "PB"),
    (r"\bSS\s?\d*\b|^SS$",              "SS"),
    (r"\bTIN\s+PLATED\b",               "Tin Plated Steel"),
]


def normalize_grade(raw) -> tuple[str, bool]:
    """Collapse messy grade strings into one metal group.

    Returns (group, matched). matched=False means no rule fired and the value
    fell through to a title-cased passthrough — surfaced on the Data Health tab
    so new grades get a rule instead of quietly becoming their own category.
    """
    s = str(raw).strip().upper()
    if not s or s == "NAN":
        return "Unknown", False
    for pattern, group in GRADE_RULES:
        if re.search(pattern, s):
            return group, True
    return str(raw).strip().title(), False


def normalize_machine(raw) -> str:
    s = str(raw).strip().upper()
    if s in ("OUTSIDE", "OUT"):
        return "Outside"
    try:
        return f"Machine {int(float(s))}"
    except (ValueError, TypeError):
        return str(raw).strip().title()


def machine_sort_key(m: str) -> int:
    if m == "Outside":
        return 99
    try:
        return int(str(m).replace("Machine ", ""))
    except ValueError:
        return 50


_SIZE_RE = re.compile(r"^\s*([\d.]+)\s*[xX*]\s*([\d.]+)")


def parse_size(raw) -> tuple[float | None, float | None]:
    """'0.80X1250' -> (0.80 thickness mm, 1250 width mm).

    Tolerates trailing text ('0.80X48 BAL') and stray spaces. Returns
    (None, None) when the string cannot be read, which the Data Health tab
    counts rather than silently defaulting to zero.
    """
    m = _SIZE_RE.match(str(raw).replace(" ", ""))
    if not m:
        return None, None
    try:
        thickness, width = float(m.group(1)), float(m.group(2))
    except ValueError:
        return None, None
    if thickness <= 0 or width <= 0 or thickness > 50 or width > 3000:
        return None, None  # implausible — treat as unparseable
    return thickness, width


def band_labels(edges: list[float], unit: str) -> list[str]:
    labels = [f"0–{edges[0]:g} {unit}"]
    for lo, hi in zip(edges, edges[1:]):
        labels.append(f"{lo:g}–{hi:g} {unit}")
    labels.append(f">{edges[-1]:g} {unit}")
    return labels


def cut_bands(series: pd.Series, edges: list[float], unit: str) -> pd.Categorical:
    bins = [0] + list(edges) + [float("inf")]
    return pd.cut(series, bins=bins, labels=band_labels(edges, unit), right=True)


# ---------------------------------------------------------------------------
# LOAD
# ---------------------------------------------------------------------------
def load_data(token: int, path_str: str, sheet: str, cfg: dict):
    """Read and clean the production sheet.

    `token` is a plain (NOT underscore-prefixed) argument on purpose: Streamlit
    excludes underscore-prefixed parameters from the cache key, so the previous
    `_token` never actually invalidated the cache — refresh only worked because
    of an accompanying cache.clear(). This name makes the invalidation real.
    """
    path = Path(path_str)
    if not path.exists():
        raise FileNotFoundError(
            f"Could not find '{path.name}' next to app.py. "
            "Run 'python make_sample_data.py' to write a synthetic one, or put "
            "the real register there under that exact filename."
        )

    raw = pd.read_excel(path, sheet_name=sheet)
    raw.columns = [" ".join(str(c).split()) for c in raw.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in raw.columns]
    if missing:
        raise ValueError(
            "These columns are missing from the sheet: " + ", ".join(missing)
        )

    df = pd.DataFrame(index=raw.index)
    df["sr"] = pd.to_numeric(raw["SR.NO"], errors="coerce").fillna(0).astype(int)
    df["coil"] = raw["COIL NO"].astype(str).str.strip()
    df["gradeRaw"] = raw["GRADE"].astype(str).str.strip()

    grade_pairs = df["gradeRaw"].apply(normalize_grade)
    df["gradeGroup"] = [g for g, _ in grade_pairs]
    df["gradeMatched"] = [ok for _, ok in grade_pairs]

    df["pcard"] = raw["P. CARD NO"].astype(str).str.strip()
    df["sizeRaw"] = raw["RM/SIZE"].astype(str).str.strip()
    size_pairs = df["sizeRaw"].apply(parse_size)
    df["thicknessMm"] = [t for t, _ in size_pairs]
    df["widthMm"] = [w for _, w in size_pairs]
    df["widthBand"] = cut_bands(df["widthMm"], cfg["width_band_edges_mm"], "mm")
    df["thicknessBand"] = cut_bands(
        df["thicknessMm"], cfg["thickness_band_edges_mm"], "mm"
    )

    df["finishSize"] = raw.get("FINISH/SIZE", "").astype(str).str.strip()
    df["hardness"] = raw.get("HARDNESS", "").astype(str).str.strip()
    df["machine"] = raw["MC NO"].apply(normalize_machine)

    mtype = raw["MATERIAL TYPE"].astype(str).str.strip().str.upper()
    df["materialTypeValid"] = mtype.isin(["FG", "RM"])
    df["materialType"] = mtype.where(df["materialTypeValid"], "Unknown")

    # ---- dates -------------------------------------------------------------
    # OUTWARD DATE = material sent out to the machine for production (job start)
    # INWARD  DATE = material returned from production (job complete)
    # So the production month is the INWARD (completion) date, and
    # turnaround = INWARD - OUTWARD. Outward AFTER inward is a data-entry error.
    outward = pd.to_datetime(raw["OUTWARD DATE"], errors="coerce")
    inward = pd.to_datetime(raw["INWARD DATE"], errors="coerce")
    df["jobStart"] = outward
    df["jobEnd"] = inward
    df["tatDays"] = (inward - outward).dt.days

    df["monthKey"] = inward.dt.strftime("%Y-%m").fillna("9999-99")
    df["month"] = inward.dt.strftime("%b %Y").fillna("Unknown")
    df["dateStart"] = outward.dt.strftime("%d %b %Y").fillna("")
    df["dateEnd"] = inward.dt.strftime("%d %b %Y").fillna("")
    df["dayKey"] = inward.dt.date

    # ---- operators ---------------------------------------------------------
    aliases = {k.upper(): v.upper() for k, v in cfg["operator_aliases"].items()}
    vendors = {v.upper() for v in cfg["vendor_operators"]}
    op_raw = raw["OPRETAR NAME"].fillna("").astype(str).str.strip().str.upper()
    df["operatorRaw"] = op_raw
    df["operator"] = op_raw.replace(aliases)
    df["operatorType"] = df["operator"].apply(
        lambda o: "Vendor" if o in vendors else ("Unknown" if not o else "In-house")
    )
    # Shared jobs ("GANESH/AJAY") stay as typed but are marked so they don't
    # silently inflate or deflate one person's numbers.
    df["operatorShared"] = df["operator"].str.contains("/", na=False)

    # ---- weights -----------------------------------------------------------
    for out_col, src_col in WEIGHT_COLUMNS.items():
        df[out_col] = pd.to_numeric(raw[src_col], errors="coerce").fillna(0.0)

    # Loss % is the real efficiency metric. Output/Input ("yield") is NOT used
    # as a KPI: scrap is not deducted from SLITTING OUT WT, so that ratio sits
    # at 99.8-100.0% for every machine and cannot discriminate anything.
    df["lossPct"] = (df["totalScrapWt"] / df["inputWt"].replace(0, pd.NA) * 100).astype(float)
    df["trimPct"] = (df["trimWt"] / df["inputWt"].replace(0, pd.NA) * 100).astype(float)

    df = _add_quality_flags(df, raw, cfg)

    meta = _build_meta(df, cfg)
    return df, meta


def _add_quality_flags(df: pd.DataFrame, raw: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    tol = cfg["mass_balance_tolerance_tons"]
    max_tat = cfg["max_reasonable_tat_days"]
    today = pd.Timestamp(datetime.now().date())

    residual = df["inputWt"] - (
        df["outputWt"] + df["trimWt"] + df["scrapWt"] + df["excessWt"] + df["shotWt"]
    )
    df["massBalanceResidual"] = residual

    df["flagOutputOverInput"] = df["outputWt"] > df["inputWt"]
    df["flagDateOrder"] = df["tatDays"] < 0
    df["flagLongTat"] = df["tatDays"] > max_tat
    df["flagFutureDate"] = (df["jobStart"] > today) | (df["jobEnd"] > today)
    df["flagMassBalance"] = residual.abs() > tol
    df["flagScrapMismatch"] = (
        (df["totalScrapWt"] - (df["trimWt"] + df["scrapWt"])).abs() > 1e-6
    )
    df["flagDuplicateCoil"] = df["coil"].duplicated(keep=False) & (df["coil"] != "")
    df["flagUnparsedSize"] = df["widthMm"].isna()
    df["flagBadMaterialType"] = ~df["materialTypeValid"]
    df["flagUnmappedGrade"] = ~df["gradeMatched"]
    df["flagZeroInput"] = df["inputWt"] <= 0

    flag_cols = [c for c in df.columns if c.startswith("flag")]
    df["issueCount"] = df[flag_cols].sum(axis=1)
    return df


def _build_meta(df: pd.DataFrame, cfg: dict) -> dict:
    machines = sorted(df["machine"].unique().tolist(), key=machine_sort_key)
    grades = sorted(df["gradeGroup"].unique().tolist())
    months_lookup = (
        df[df["monthKey"] != "9999-99"][["monthKey", "month"]]
        .drop_duplicates()
        .sort_values("monthKey")
    )
    return {
        "machines": machines,
        "grades": grades,
        "months": months_lookup["month"].tolist(),
        "monthKeys": months_lookup["monthKey"].tolist(),
        "loadedAt": datetime.now(),
        "rowCount": len(df),
        "dailyTargets": cfg["machine_daily_target_tons"],
        "workingDays": cfg["working_days_per_month"],
    }


# ---------------------------------------------------------------------------
# METRICS
# ---------------------------------------------------------------------------
def machine_target(machine: str, n_months: int, cfg: dict) -> float | None:
    """Committed target in tons. None = no target agreed for this machine."""
    daily = cfg["machine_daily_target_tons"].get(machine)
    if daily is None:
        return None
    return daily * cfg["working_days_per_month"] * n_months


def loss_rate(frame: pd.DataFrame) -> float:
    """Volume-weighted loss %, not the mean of per-job percentages.

    Averaging per-job percentages lets a 0.04-ton coil with 27% loss count
    the same as a 200-ton coil with 1% loss. Every loss figure in this
    dashboard is weighted by input tonnage.
    """
    inp = frame["inputWt"].sum()
    return (frame["totalScrapWt"].sum() / inp * 100) if inp else 0.0


def machine_summary(fdf: pd.DataFrame, n_months: int, cfg: dict) -> pd.DataFrame:
    """One row per machine, with size-mix-adjusted loss.

    `Expected Loss %` is what this machine's loss would be if it performed at
    the plant-wide rate for each width band, given the width mix it actually
    ran. `Loss Gap` = observed - expected, so a machine that only looks bad
    because it runs narrow coil shows a gap near zero.
    """
    band_rates = _band_loss_rates(fdf)

    rows = []
    for machine, sub in fdf.groupby("machine", observed=True):
        inp = sub["inputWt"].sum()
        out = sub["outputWt"].sum()
        scrap = sub["totalScrapWt"].sum()
        target = machine_target(machine, n_months, cfg)
        active_days = sub["dayKey"].nunique()

        expected_scrap = sum(
            band_rates.get(band, 0.0) / 100 * grp["inputWt"].sum()
            for band, grp in sub.groupby("widthBand", observed=True)
        )
        expected_pct = (expected_scrap / inp * 100) if inp else 0.0
        observed_pct = (scrap / inp * 100) if inp else 0.0

        rows.append({
            "Machine": machine,
            "Output (T)": round(out, 2),
            "Target (T)": round(target, 2) if target is not None else None,
            "Achieved %": round(out / target * 100, 1) if target else None,
            "Loss %": round(observed_pct, 2),
            "Expected Loss %": round(expected_pct, 2),
            "Loss Gap": round(observed_pct - expected_pct, 2),
            "Jobs": len(sub),
            "T / Job": round(out / len(sub), 2) if len(sub) else 0,
            "Active Days": active_days,
            "T / Active Day": round(out / active_days, 2) if active_days else 0,
            "Median TAT (d)": sub.loc[sub["tatDays"] >= 0, "tatDays"].median(),
        })

    out_df = pd.DataFrame(rows)
    if out_df.empty:
        return out_df
    # Machines with no committed target carry NaN, not None, so the UI renders
    # an empty cell rather than the literal text "None".
    for col in ("Target (T)", "Achieved %"):
        out_df[col] = pd.to_numeric(out_df[col], errors="coerce")
    return out_df.sort_values("Machine", key=lambda s: s.map(machine_sort_key))


def _band_loss_rates(fdf: pd.DataFrame) -> dict:
    grp = fdf.dropna(subset=["widthBand"]).groupby("widthBand", observed=True)
    agg = grp.agg(inp=("inputWt", "sum"), scrap=("totalScrapWt", "sum"))
    return {
        band: (row.scrap / row.inp * 100) if row.inp else 0.0
        for band, row in agg.iterrows()
    }


def width_band_summary(fdf: pd.DataFrame) -> pd.DataFrame:
    grp = fdf.dropna(subset=["widthBand"]).groupby("widthBand", observed=True)
    out = grp.agg(
        Jobs=("sr", "count"),
        Input=("inputWt", "sum"),
        Scrap=("totalScrapWt", "sum"),
    ).reset_index()
    out["Loss %"] = (out["Scrap"] / out["Input"] * 100).round(2)
    out["Input (T)"] = out["Input"].round(1)
    out["Scrap (T)"] = out["Scrap"].round(2)
    out = out.rename(columns={"widthBand": "RM Width Band"})
    return out[["RM Width Band", "Jobs", "Input (T)", "Scrap (T)", "Loss %"]]


def machine_band_matrix(fdf: pd.DataFrame, min_tons: float):
    """Loss % per machine x width band, plus the tonnage behind each cell.

    Cells below `min_tons` are blanked: a 3-ton cell showing 16% loss is noise,
    and blanking it stops the heatmap from inventing a hotspot.
    """
    sub = fdf.dropna(subset=["widthBand"])
    inp = sub.pivot_table(index="machine", columns="widthBand",
                          values="inputWt", aggfunc="sum", observed=True)
    scrap = sub.pivot_table(index="machine", columns="widthBand",
                            values="totalScrapWt", aggfunc="sum", observed=True)
    pct = (scrap / inp * 100).where(inp >= min_tons)
    order = sorted(pct.index.tolist(), key=machine_sort_key)
    return pct.reindex(order), inp.reindex(order)


def monthly_summary(fdf: pd.DataFrame, months: list[str], n_machines_target: float) -> pd.DataFrame:
    """Month-by-month output, target, loss and month-over-month deltas."""
    grp = fdf.groupby("month", observed=True).agg(
        Output=("outputWt", "sum"),
        Input=("inputWt", "sum"),
        Scrap=("totalScrapWt", "sum"),
        Jobs=("sr", "count"),
    )
    grp = grp.reindex(months).fillna(0.0)
    grp["Loss %"] = (grp["Scrap"] / grp["Input"].replace(0, pd.NA) * 100).round(2)
    grp["Target (T)"] = round(n_machines_target, 2)
    grp["Achieved %"] = (grp["Output"] / n_machines_target * 100).round(1) if n_machines_target else None
    grp["MoM Output %"] = (grp["Output"].pct_change() * 100).round(1)
    grp["MoM Loss (pp)"] = grp["Loss %"].diff().round(2)
    grp["Output (T)"] = grp["Output"].round(2)
    return grp.reset_index().rename(columns={"index": "Month", "month": "Month"})


def narrow_slit_opportunity(fdf: pd.DataFrame, cfg: dict) -> dict:
    """What the narrow-coil loss gap is worth in tons per month.

    Compares the loss rate on coil below the second band edge against the
    plant's best-performing wide band, and prices the difference in tons.
    """
    edges = cfg["width_band_edges_mm"]
    narrow_cut = edges[1]
    sub = fdf.dropna(subset=["widthMm"])
    narrow = sub[sub["widthMm"] <= narrow_cut]
    wide = sub[sub["widthMm"] > edges[-1]]

    if narrow.empty or wide.empty:
        return {}

    narrow_rate = loss_rate(narrow)
    wide_rate = loss_rate(wide)
    mid = sub[(sub["widthMm"] > edges[1]) & (sub["widthMm"] <= edges[2])]
    benchmark_rate = loss_rate(mid) if not mid.empty else wide_rate

    n_months = fdf.loc[fdf["monthKey"] != "9999-99", "monthKey"].nunique() or 1
    narrow_input = narrow["inputWt"].sum()
    recoverable = narrow_input * (narrow_rate - benchmark_rate) / 100

    return {
        "narrow_cut_mm": narrow_cut,
        "narrow_rate": narrow_rate,
        "benchmark_rate": benchmark_rate,
        "wide_rate": wide_rate,
        "narrow_input_tons": narrow_input,
        "narrow_jobs": len(narrow),
        "recoverable_tons_total": recoverable,
        "recoverable_tons_per_month": recoverable / n_months,
        "months": n_months,
    }


def data_health_report(df: pd.DataFrame) -> pd.DataFrame:
    """One row per check, so the tab is data-driven rather than hand-written."""
    checks = [
        ("flagDateOrder",        "Outward date after Inward date",
         "Job cannot finish before it starts — check the entry."),
        ("flagFutureDate",       "Date in the future",
         "Outward or Inward date is later than today."),
        ("flagLongTat",          "Turnaround longer than the configured limit",
         "Possibly a typo in the year or month."),
        ("flagOutputOverInput",  "Output weight greater than input weight",
         "Physically impossible — weighbridge or entry error."),
        ("flagMassBalance",      "Weights do not balance",
         "IN should equal OUT + TRIM + SCRAP + EXCESS + SHOT."),
        ("flagScrapMismatch",    "TOTAL SCRAP does not equal TRIM + SCRAP",
         "The three scrap columns disagree."),
        ("flagZeroInput",        "Zero or missing input weight",
         "Loss % cannot be computed for these rows."),
        ("flagDuplicateCoil",    "Duplicate COIL NO",
         "May be a genuine second pass, or a double entry."),
        ("flagUnparsedSize",     "RM/SIZE could not be read",
         "Excluded from all width and thickness analysis."),
        ("flagBadMaterialType",  "MATERIAL TYPE is neither FG nor RM",
         "Previously these were silently relabelled as FG."),
        ("flagUnmappedGrade",    "GRADE did not match any grouping rule",
         "Add a rule in datalayer.py GRADE_RULES to group it."),
    ]
    rows = []
    for col, name, why in checks:
        count = int(df[col].sum())
        rows.append({
            "Check": name,
            "Rows": count,
            "% of Data": round(count / len(df) * 100, 1) if len(df) else 0,
            "Why it matters": why,
            "_col": col,
        })
    return pd.DataFrame(rows)

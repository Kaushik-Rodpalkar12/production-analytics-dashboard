"""
Production Analytics — data layer checks
=======================================================
Runs the numbers without launching Streamlit, so a change to datalayer.py can
be verified in seconds.

    python validate.py

Every check either prints PASS or raises. If this script passes, the dashboard
is reading the sheet correctly.
"""

from __future__ import annotations

import sys

import pandas as pd

import datalayer as dl


def check(name: str, condition: bool, detail: str = ""):
    status = "PASS" if condition else "FAIL"
    print(f"  [{status}] {name}" + (f"  — {detail}" if detail else ""))
    if not condition:
        check.failures += 1


check.failures = 0


def main() -> int:
    cfg = dl.load_config()
    df, meta = dl.load_data(0, str(dl.EXCEL_PATH), dl.SHEET_NAME, cfg)

    print(f"\nLoaded {len(df):,} rows · {meta['months'][0]} – {meta['months'][-1]}\n")

    print("Structure")
    check("every row has a machine", df["machine"].notna().all())
    check("every row has a month", (df["monthKey"] != "9999-99").all())
    check("no negative weights",
          all((df[c] >= 0).all() for c in dl.WEIGHT_COLUMNS))
    check("grade rules cover the file",
          df["gradeMatched"].all(),
          f"{(~df['gradeMatched']).sum()} unmapped: "
          f"{sorted(df.loc[~df['gradeMatched'], 'gradeRaw'].unique())[:5]}")
    check("RM/SIZE parses",
          df["widthMm"].notna().mean() > 0.99,
          f"{df['widthMm'].isna().sum()} unparsed")

    print("\nNormalisation")
    ops_raw = df["operatorRaw"].nunique()
    ops = df["operator"].nunique()
    check("operator aliases collapse duplicates", ops < ops_raw,
          f"{ops_raw} raw names -> {ops} canonical")
    check("vendors are separated",
          set(df["operatorType"].unique()) <= {"In-house", "Vendor", "Unknown"})
    vendor_outside = df[df["operatorType"] == "Vendor"]["machine"].eq("Outside").mean()
    check("vendors work almost entirely on Outside", vendor_outside > 0.9,
          f"{vendor_outside:.0%} of vendor jobs are Outside")

    print("\nMetrics")
    plant = dl.loss_rate(df)
    check("plant loss % is in a plausible range", 0.5 < plant < 10,
          f"{plant:.2f}%")
    weighted = df["totalScrapWt"].sum() / df["inputWt"].sum() * 100
    check("loss_rate is tonnage-weighted, not a mean of percentages",
          abs(plant - weighted) < 1e-9 and abs(plant - df["lossPct"].mean()) > 0.1,
          f"weighted {plant:.2f}% vs unweighted mean {df['lossPct'].mean():.2f}%")

    n_months = df["monthKey"].nunique()
    t1 = dl.machine_target("Machine 1", n_months, cfg)
    expect = cfg["machine_daily_target_tons"]["Machine 1"] * cfg["working_days_per_month"] * n_months
    check("machine target matches config", abs(t1 - expect) < 1e-9,
          f"Machine 1 = {t1:,.0f} T over {n_months} months")
    check("machines without a config entry have no target",
          dl.machine_target("Outside", n_months, cfg) is None)

    print("\nSize-mix adjustment")
    summ = dl.machine_summary(df, n_months, cfg)
    check("summary has one row per machine", len(summ) == df["machine"].nunique())
    check("loss gap = observed - expected",
          ((summ["Loss %"] - summ["Expected Loss %"] - summ["Loss Gap"]).abs() < 0.011).all())
    # The adjustment must actually do something: the spread between machines
    # should shrink once coil width is controlled for.
    spread_raw = summ["Loss %"].max() - summ["Loss %"].min()
    spread_adj = summ["Loss Gap"].max() - summ["Loss Gap"].min()
    check("adjusting for coil width narrows the spread between machines",
          spread_adj < spread_raw,
          f"raw spread {spread_raw:.2f} pp -> adjusted {spread_adj:.2f} pp")

    print("\nWidth bands")
    wb = dl.width_band_summary(df)
    check("bands are ordered narrow to wide", len(wb) == len(cfg["width_band_edges_mm"]) + 1)
    check("loss falls as coil gets wider",
          wb["Loss %"].iloc[0] > wb["Loss %"].iloc[-1],
          f"{wb['Loss %'].iloc[0]:.2f}% narrowest vs {wb['Loss %'].iloc[-1]:.2f}% widest")
    check("every band carries volume", (wb["Input (T)"] > 0).all())

    pct, inp = dl.machine_band_matrix(df, cfg["min_tons_for_insight"])
    thin = (inp < cfg["min_tons_for_insight"]) & inp.notna()
    check("thin cells are blanked in the matrix",
          pct.where(thin).isna().all().all(),
          f"{int(thin.sum().sum())} cells below "
          f"{cfg['min_tons_for_insight']:g} T suppressed")

    print("\nMonthly")
    monthly = dl.monthly_summary(df, meta["months"], 100.0)
    check("one row per month", len(monthly) == len(meta["months"]))
    check("output reconciles with the raw total",
          abs(monthly["Output"].sum() - df["outputWt"].sum()) < 0.01)

    print("\nData health")
    health = dl.data_health_report(df)
    check("every check reports a count", health["Rows"].notna().all())
    for _, r in health.iterrows():
        if r["Rows"]:
            print(f"        · {r['Check']}: {r['Rows']:,} rows ({r['% of Data']}%)")

    print("\nReconciliation")
    scope = cfg["scopes"]["In-house slitting (M1-M6)"]
    inhouse = df[df["machine"].isin(scope)]
    check("scope filter is a strict subset", len(inhouse) < len(df),
          f"{len(inhouse):,} in-house of {len(df):,} total")
    check("in-house + M7 + Outside = everything",
          len(inhouse) + len(df[df["machine"] == "Machine 7"])
          + len(df[df["machine"] == "Outside"]) == len(df))

    print()
    if check.failures:
        print(f"{check.failures} check(s) FAILED\n")
        return 1
    print("All checks passed.\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())

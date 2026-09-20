"""Generate a synthetic Production_Report.xlsx with the schema the dashboard expects.

The real sheet is plant data and is not published. This writes a stand-in with
the same columns and the same structural relationships, so the app runs and the
width-adjusted loss comparison has something to show.

    python make_sample_data.py

Relationships built into the output, because the dashboard exists to measure
them:
  - loss falls as coil width rises (edge trim is a roughly fixed width)
  - machines run different width mixes, so raw loss differs between machines
    while true machine-to-machine difference stays small
  - Machine 7 is the wide, high-tonnage line; Machines 1-3 run narrow work
  - a small number of rows carry the data-entry faults the Data Health tab
    reports: unparseable sizes, reversed dates, duplicate coil numbers,
    misspelled operator names, blank material type
"""

from __future__ import annotations

import random
from datetime import date, timedelta

import pandas as pd

SEED = 20250416
ROWS = 1800
START = date(2025, 4, 1)
MONTHS = 12

GRADES = [
    ("CRC IS 513", 0.30), ("CRC AISI 1008", 0.10), ("PB GR III", 0.12),
    ("BRASS", 0.10), ("ALLUMINIUM", 0.08), ("COPPER", 0.07),
    ("SS 304", 0.09), ("GI", 0.06), ("BERYLIUM COPPER", 0.04),
    ("TIN PLATED STEEL", 0.04),
]

# machine -> (width range mm, thickness range mm, tons per job, share of jobs)
MACHINES = {
    "1": ((8, 60), (0.15, 0.60), (0.20, 1.20), 0.13),
    "2": ((10, 80), (0.15, 0.70), (0.25, 1.40), 0.13),
    "3": ((12, 90), (0.20, 0.80), (0.30, 1.60), 0.11),
    "4": ((40, 260), (0.30, 1.20), (1.00, 4.50), 0.12),
    "5": ((50, 300), (0.30, 1.40), (1.20, 5.00), 0.12),
    "6": ((60, 340), (0.35, 1.50), (1.40, 5.50), 0.11),
    "7": ((300, 1250), (0.50, 2.50), (6.00, 22.00), 0.19),
    "OUTSIDE": ((20, 900), (0.20, 2.00), (0.50, 12.00), 0.09),
}

# Plant-wide loss rate by width band. Narrow strip loses a far larger share to
# edge trim than wide strip does, which is the confounder the dashboard adjusts
# for before comparing machines.
BAND_LOSS = [(50, 0.072), (150, 0.046), (400, 0.026), (800, 0.017), (1e9, 0.013)]

# Small true machine effect, in multiples of the band rate. The point of the
# dashboard is that this is much smaller than the raw spread suggests.
MACHINE_EFFECT = {
    "1": 1.04, "2": 0.98, "3": 1.06, "4": 1.01,
    "5": 0.97, "6": 1.02, "7": 1.00, "OUTSIDE": 1.09,
}

OPERATORS = ["OPERATOR A", "OPERATOR B", "OPERATOR C", "OPERATOR D",
             "OPERATOR E", "OPERATOR F", "OPERATOR G", "OPERATOR H"]
VENDORS = ["VENDOR ONE", "VENDOR TWO", "VENDOR THREE", "VENDOR FOUR"]

rng = random.Random(SEED)


def band_rate(width: float) -> float:
    for edge, rate in BAND_LOSS:
        if width <= edge:
            return rate
    return BAND_LOSS[-1][1]


def pick(weighted):
    total = sum(w for _, w in weighted)
    r = rng.uniform(0, total)
    acc = 0.0
    for value, weight in weighted:
        acc += weight
        if r <= acc:
            return value
    return weighted[-1][0]


def build_row(sr: int) -> dict:
    machine = pick([(m, cfg[3]) for m, cfg in MACHINES.items()])
    (w_lo, w_hi), (t_lo, t_hi), (ton_lo, ton_hi), _ = MACHINES[machine]

    width = round(rng.uniform(w_lo, w_hi))
    thickness = round(rng.uniform(t_lo, t_hi), 2)
    input_wt = round(rng.uniform(ton_lo, ton_hi), 3)

    rate = band_rate(width) * MACHINE_EFFECT[machine] * rng.uniform(0.75, 1.30)
    total_scrap = round(input_wt * rate, 4)
    trim = round(total_scrap * rng.uniform(0.55, 0.85), 4)
    scrap = round(total_scrap - trim, 4)
    shot = round(input_wt * rng.uniform(0, 0.004), 4) if rng.random() < 0.15 else 0.0
    excess = round(input_wt * rng.uniform(0, 0.006), 4) if rng.random() < 0.20 else 0.0
    # IN = OUT + TRIM + SCRAP + EXCESS + SHOT, which is what the Data Health
    # mass-balance check tests.
    output_wt = round(input_wt - trim - scrap - shot - excess, 4)

    start = START + timedelta(days=rng.randint(0, MONTHS * 30))
    tat = rng.choice([1, 1, 2, 2, 3, 3, 4, 5, 7, 9, 12, 18])
    end = start + timedelta(days=tat)

    is_vendor = machine == "OUTSIDE"
    operator = rng.choice(VENDORS if is_vendor else OPERATORS)
    if not is_vendor and rng.random() < 0.04:
        operator = f"{rng.choice(OPERATORS)}/{rng.choice(OPERATORS)}"

    return {
        "SR.NO": sr,
        "COIL NO": f"SMP/{rng.randint(1000, 9999)}/{rng.randint(1, 99):02d}",
        "GRADE": pick(GRADES),
        "P. CARD NO": rng.randint(7000, 9999),
        "OUTWARD      DATE": pd.Timestamp(start),
        "RM/SIZE": f"{thickness:.2f}X{width}",
        "HARDNESS": rng.choice(["HARD", "SOFT", "D", "H", "1/2H"]),
        "HARDNESS/OBSERVED": None,
        "SLITTING OUT WT (TONS)": output_wt,
        "SLITTING IN WT (TONS)": input_wt,
        "SHOT WT (TONS)": shot,
        "EXCESS WT (TONS)": excess,
        "FINISH/SIZE": f"{thickness:.2f}X{max(4, round(width / rng.randint(2, 8)))}",
        "TRIMING (TONS)": trim,
        "SCRAP (TONS)": scrap,
        "TOTAL SCRAP (TONS)": total_scrap,
        "TRIMING(%)": round(100 * trim / input_wt, 2),
        "SCRAP(%)": round(100 * scrap / input_wt, 2),
        "PERCENTAGE(%)": round(100 * total_scrap / input_wt, 2),
        "INWARD DATE": pd.Timestamp(end),
        "OPRETAR NAME": operator,
        "MC NO": machine,
        "MATERIAL TYPE": rng.choice(["RM", "RM", "RM", "FG"]),
    }


def inject_faults(rows: list[dict]) -> None:
    """Add the data-entry faults the Data Health tab is built to catch."""
    for i in rng.sample(range(len(rows)), 12):
        rows[i]["RM/SIZE"] = rng.choice(["AS PER SAMPLE", "0.5 X", "-", "COIL"])
    for i in rng.sample(range(len(rows)), 9):
        rows[i]["OPRETAR NAME"] = rng.choice(["OPERATOR A.", "OPRATOR B"])
    for i in rng.sample(range(len(rows)), 5):
        rows[i]["OUTWARD      DATE"], rows[i]["INWARD DATE"] = (
            rows[i]["INWARD DATE"], rows[i]["OUTWARD      DATE"])
    for i in rng.sample(range(len(rows)), 6):
        rows[i]["TOTAL SCRAP (TONS)"] = round(
            rows[i]["TOTAL SCRAP (TONS)"] * rng.uniform(1.5, 3.0), 4)
    for i in rng.sample(range(len(rows)), 4):
        rows[i]["MATERIAL TYPE"] = ""
    for i in rng.sample(range(len(rows)), 3):
        rows[i]["OPRETAR NAME"] = ""


def main() -> None:
    rows = [build_row(i + 1) for i in range(ROWS)]
    inject_faults(rows)
    df = pd.DataFrame(rows).sort_values("INWARD DATE").reset_index(drop=True)
    df["SR.NO"] = range(1, len(df) + 1)
    df.to_excel("Production_Report.xlsx", sheet_name="Production_Data", index=False)

    tons = df["SLITTING IN WT (TONS)"].sum()
    loss = 100 * df["TOTAL SCRAP (TONS)"].sum() / tons
    print(f"{len(df):,} rows written to Production_Report.xlsx")
    print(f"{tons:,.0f} tons in, plant loss {loss:.2f}%")


if __name__ == "__main__":
    main()

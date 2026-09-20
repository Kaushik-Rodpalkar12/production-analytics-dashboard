"""
Production Analytics
===================================
Streamlit dashboard over Production_Report.xlsx (sheet "Production_Data").

Run:  streamlit run app.py
Data does not auto-refresh — press "Refresh data" in the sidebar after saving
the Excel file.

Targets, scope groups, width bands and operator/vendor names all live in
config.json. Edit that file, not this one.
"""

from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import datalayer as dl

# =============================================================================
# PAGE CONFIG
# =============================================================================
st.set_page_config(
    page_title="Production Analytics",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# PALETTE
# -----------------------------------------------------------------------------
# Categorical order validated for colour-vision deficiency and normal-vision
# separation (worst adjacent pair: CVD ΔE 9.2, normal ΔE 19.6, both above the
# required floors). Do not reorder or extend past 8 without re-validating —
# the ORDER is what makes it CVD-safe, not the individual hues.
# Status colours are reserved for good/bad meaning and are never used as series.
# =============================================================================
SERIES = ["#1baf7a", "#eb6834", "#2a78d6", "#eda100",
          "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
S1 = SERIES[0]

ST_GOOD, ST_WARN, ST_SERIOUS, ST_CRIT = "#0ca30c", "#fab219", "#ec835a", "#d03b3b"

BLUE_RAMP = ["#cde2fb", "#b7d3f6", "#9ec5f4", "#86b6ef", "#6da7ec",
             "#5598e7", "#3987e5", "#2a78d6", "#256abf", "#1c5cab",
             "#184f95", "#104281", "#0d366b"]

SURFACE = "#FFFFFF"
INK_1, INK_2, INK_MUTED = "#0b0b0b", "#52514e", "#898781"
GRID, AXIS = "#e1e0d9", "#c3c2b7"
FONT = 'system-ui, -apple-system, "Segoe UI", sans-serif'

st.markdown(f"""
<style>
    .block-container{{padding-top:1.4rem; padding-bottom:3rem;}}
    div[data-testid="stMetric"]{{
        background:{SURFACE}; border:1px solid #E6E5DE; border-radius:12px;
        padding:12px 14px 8px;
    }}
    div[data-testid="stMetricLabel"]{{font-size:11px; text-transform:uppercase;
        letter-spacing:.05em; color:{INK_MUTED}; font-weight:600;}}
    div[data-testid="stMetricValue"]{{font-size:23px; color:{INK_1}; font-weight:650;}}
    .note{{padding:10px 13px; border-radius:9px; font-size:13.5px; margin-bottom:7px;
        line-height:1.45; border-left:3px solid;}}
    .note-good{{background:#EAF7EC; border-color:{ST_GOOD}; color:#12420f;}}
    .note-warn{{background:#FDF4E0; border-color:{ST_WARN}; color:#5c4406;}}
    .note-crit{{background:#FBEAEA; border-color:{ST_CRIT}; color:#6b1f1f;}}
    .note-info{{background:#F2F5F4; border-color:{INK_MUTED}; color:{INK_2};}}
    .scope-badge{{display:inline-block; background:#EAF7F1; color:#0f5b43;
        border:1px solid #BDE5D5; border-radius:999px; padding:3px 11px;
        font-size:12px; font-weight:650;}}
    section[data-testid="stSidebar"] .stButton button{{width:100%;}}
</style>
""", unsafe_allow_html=True)


# =============================================================================
# HELPERS
# =============================================================================
@st.cache_data(show_spinner="Reading Production_Report.xlsx…")
def cached_load(token: int, path_str: str, sheet: str, cfg: dict):
    return dl.load_data(token, path_str, sheet, cfg)


def fmt(n, dp: int = 2) -> str:
    try:
        if pd.isna(n):
            return "—"
        return f"{n:,.{dp}f}"
    except (TypeError, ValueError):
        return "—"


def note(kind: str, text: str):
    st.markdown(f'<div class="note note-{kind}">{text}</div>', unsafe_allow_html=True)


def style_fig(fig, height=330, ylab=None, xlab=None, legend=True, hover="closest"):
    fig.update_layout(
        height=height,
        margin=dict(t=14, b=8, l=8, r=8),
        plot_bgcolor=SURFACE, paper_bgcolor=SURFACE,
        font=dict(family=FONT, size=12, color=INK_2),
        showlegend=legend,
        legend=dict(orientation="h", y=-0.18, x=0, bgcolor="rgba(0,0,0,0)",
                    font=dict(color=INK_2)),
        hovermode=hover,
        hoverlabel=dict(bgcolor=SURFACE, bordercolor=AXIS,
                        font=dict(family=FONT, size=12, color=INK_1)),
    )
    fig.update_yaxes(title=ylab, gridcolor=GRID, zeroline=True, zerolinecolor=AXIS,
                     linecolor=AXIS, tickfont=dict(color=INK_MUTED, size=11),
                     title_font=dict(color=INK_MUTED, size=11))
    fig.update_xaxes(title=xlab, gridcolor=GRID, zeroline=False,
                     linecolor=AXIS, tickfont=dict(color=INK_MUTED, size=11),
                     title_font=dict(color=INK_MUTED, size=11))
    return fig


def bar(x, y, color=S1, name=None, horizontal=False, text=None, colors=None):
    """Thin bar with 4px rounded data-end and a 2px surface gap between bars."""
    kw = dict(
        marker=dict(color=colors if colors is not None else color,
                    line=dict(width=2, color=SURFACE)),
        name=name, text=text,
        textposition="outside", textfont=dict(size=11, color=INK_2),
        cliponaxis=False,
    )
    try:
        kw["marker"]["cornerradius"] = 4
    except Exception:  # older plotly
        pass
    if horizontal:
        return go.Bar(x=y, y=x, orientation="h", **kw)
    return go.Bar(x=x, y=y, **kw)


def auto_labels(values, suffix: str = ""):
    """Pick decimal places from the magnitude of the series, so a 0.36-ton bar
    is not labelled '0' next to a 1,055-ton one."""
    vals = [abs(v) for v in values if pd.notna(v)]
    top = max(vals) if vals else 0
    dp = 0 if top >= 100 else (1 if top >= 10 else 2)
    return [f"{v:,.{dp}f}{suffix}" if pd.notna(v) else "" for v in values]


def pad_axis(fig, values, axis: str = "x", factor: float = 1.20):
    """Outside data labels get clipped at the plot edge — give the value axis
    headroom so the last label always fits."""
    vals = [v for v in values if pd.notna(v)]
    if not vals:
        return fig
    hi, lo = max(vals), min(vals)
    hi = hi * factor if hi > 0 else 0
    lo = lo * factor if lo < 0 else 0
    if axis == "x":
        fig.update_xaxes(range=[lo, hi])
    else:
        fig.update_yaxes(range=[lo, hi])
    return fig


def dl_button(df: pd.DataFrame, label: str, filename: str, key: str):
    st.download_button(
        label, df.to_csv(index=False).encode("utf-8"),
        file_name=filename, mime="text/csv", key=key, width='content',
    )


def sync_options(state_key: str, known_key: str, options: list):
    """Keep a multi-select valid as the underlying option list changes:
    drop values that vanished, auto-select values that are genuinely new."""
    known = st.session_state.get(known_key, [])
    if state_key not in st.session_state:
        st.session_state[state_key] = list(options)
    else:
        current = [v for v in st.session_state[state_key] if v in options]
        for o in options:
            if o not in known and o not in current:
                current.append(o)
        st.session_state[state_key] = current
    st.session_state[known_key] = list(options)


def loss_colors(values, threshold: float):
    """Status colouring — the colour means over/under target, not identity."""
    return [ST_CRIT if (pd.notna(v) and v > threshold * 1.5)
            else ST_WARN if (pd.notna(v) and v > threshold)
            else ST_GOOD for v in values]


# =============================================================================
# LOAD CONFIG + DATA
# =============================================================================
try:
    CFG = dl.load_config()
    cfg_error = None
except Exception as e:  # noqa: BLE001
    CFG, cfg_error = {}, str(e)

if cfg_error:
    st.title("Production Dashboard")
    st.error(f"Could not read config.json — {cfg_error}")
    st.stop()

if "refresh_token" not in st.session_state:
    st.session_state.refresh_token = 0
if "reset_requested" not in st.session_state:
    st.session_state.reset_requested = False
if st.session_state.reset_requested:
    for k in ["sel_material", "sel_machine", "sel_grade", "sel_month"]:
        st.session_state.pop(k, None)
    st.session_state.reset_requested = False

with st.sidebar:
    st.markdown("### 🏭 PRODUCTION ANALYTICS")
    st.caption("PRODUCTION ANALYTICS")
    st.divider()

    if st.button("🔄  Refresh data", type="primary", width='stretch'):
        st.session_state.refresh_token += 1

    try:
        df_all, meta = cached_load(
            st.session_state.refresh_token, str(dl.EXCEL_PATH), dl.SHEET_NAME, CFG
        )
        load_error = None
    except Exception as e:  # noqa: BLE001
        df_all, meta, load_error = pd.DataFrame(), {}, str(e)

if load_error:
    st.title("Production Dashboard")
    st.error(f"Could not load data — {load_error}")
    st.info("**Production_Report.xlsx** must sit next to app.py with a sheet named "
            "**Production_Data**. Then press Refresh data.")
    st.stop()

# =============================================================================
# SIDEBAR FILTERS
# =============================================================================
with st.sidebar:
    st.caption(f"Loaded {meta['loadedAt'].strftime('%d %b %Y, %I:%M %p')} · "
               f"{meta['rowCount']:,} records")
    st.divider()

    st.markdown("**Scope**")
    scope_names = [k for k in CFG["scopes"]]
    scope = st.radio("Scope", scope_names, key="scope", label_visibility="collapsed")
    scope_machines = CFG["scopes"][scope] or meta["machines"]
    scope_machines = [m for m in meta["machines"] if m in scope_machines]
    st.caption("Applies to every number on every tab.")

    df = df_all[df_all["machine"].isin(scope_machines)].copy()
    st.divider()

    st.markdown("**Material type**")
    sync_options("sel_material", "known_material", ["FG", "RM", "Unknown"])
    avail_types = [t for t in ["FG", "RM", "Unknown"] if t in df["materialType"].unique()]
    sel_material = st.pills("Material type", avail_types, selection_mode="multi",
                            key="sel_material", label_visibility="collapsed")

    st.markdown("**Machine**")
    sync_options("sel_machine", "known_machine", scope_machines)
    sel_machine = st.pills("Machine", scope_machines, selection_mode="multi",
                           key="sel_machine", label_visibility="collapsed")

    st.markdown("**Grade / metal**")
    grades_in_scope = sorted(df["gradeGroup"].unique().tolist())
    sync_options("sel_grade", "known_grade", grades_in_scope)
    sel_grade = st.multiselect("Grade / metal", grades_in_scope,
                               key="sel_grade", label_visibility="collapsed")

    st.markdown("**Month**")
    months_in_scope = [m for m in meta["months"] if m in df["month"].unique()]
    sync_options("sel_month", "known_month", months_in_scope)
    sel_month = st.multiselect("Month", months_in_scope,
                               key="sel_month", label_visibility="collapsed")

    st.divider()
    if st.button("↺  Reset all filters", width='stretch'):
        st.session_state.reset_requested = True
        st.rerun()

    st.divider()
    st.caption("Targets and scope groups are set in **config.json**.")

# =============================================================================
# APPLY FILTERS
# =============================================================================
fdf = df[
    df["materialType"].isin(sel_material)
    & df["machine"].isin(sel_machine)
    & df["gradeGroup"].isin(sel_grade)
    & df["month"].isin(sel_month)
].copy()

MONTHS = [m for m in meta["months"] if m in sel_month]
N_MONTHS = len(MONTHS)
LOSS_TARGET = CFG["plant_loss_target_pct"]
MIN_TONS = CFG["min_tons_for_insight"]

targeted_machines = [m for m in sel_machine
                     if CFG["machine_daily_target_tons"].get(m) is not None]
total_target = sum(dl.machine_target(m, N_MONTHS, CFG) or 0 for m in targeted_machines)
output_targeted = fdf[fdf["machine"].isin(targeted_machines)]["outputWt"].sum()

# =============================================================================
# HEADER
# =============================================================================
st.title("Production Dashboard")
h1, h2 = st.columns([2.2, 1])
with h1:
    st.markdown(f'<span class="scope-badge">{scope}</span>', unsafe_allow_html=True)
    st.caption(
        f"{len(fdf):,} of {len(df_all):,} records · "
        f"{MONTHS[0] if MONTHS else '—'} – {MONTHS[-1] if MONTHS else '—'} · "
        f"production month = job completion (Inward) date"
    )
with h2:
    st.caption(f"Showing **{len(sel_machine)}/{len(scope_machines)}** machines · "
               f"**{N_MONTHS}/{len(months_in_scope)}** months")

if fdf.empty:
    note("warn", "<b>No records match the current filters.</b> Widen the selection "
                 "in the sidebar, or press <i>Reset all filters</i>.")
    st.stop()

tabs = st.tabs([
    "📊 Overview", "⚙️ Machine Performance", "📐 Size & Loss", "🧱 Material",
    "📅 Monthly Trend", "👷 Operators", "🔎 Job Explorer", "🩺 Data Health",
])

# =============================================================================
# TAB 1 — OVERVIEW
# =============================================================================
with tabs[0]:
    total_out = fdf["outputWt"].sum()
    total_in = fdf["inputWt"].sum()
    total_scrap = fdf["totalScrapWt"].sum()
    total_trim = fdf["trimWt"].sum()
    plant_loss = dl.loss_rate(fdf)
    med_tat = fdf.loc[fdf["tatDays"] >= 0, "tatDays"].median()
    achieved = (output_targeted / total_target * 100) if total_target else None

    k = st.columns(6)
    k[0].metric("Output", f"{fmt(total_out)} T")
    k[1].metric("Input", f"{fmt(total_in)} T")
    k[2].metric("Loss %", f"{plant_loss:.2f}%",
                delta=f"{plant_loss - LOSS_TARGET:+.2f} pp vs target",
                delta_color="inverse")
    k[3].metric("Coils processed", f"{len(fdf):,}")
    k[4].metric("Target achieved",
                f"{achieved:.1f}%" if achieved is not None else "—")
    k[5].metric("Median turnaround", f"{fmt(med_tat, 0)} days")

    st.caption(
        f"Loss % = TOTAL SCRAP ÷ SLITTING IN, weighted by tonnage. "
        f"Target achievement covers only machines with a committed target "
        f"({', '.join(targeted_machines) if targeted_machines else 'none selected'}) "
        f"at {CFG['working_days_per_month']} working days/month over {N_MONTHS} month(s)."
    )
    st.write("")

    summary = dl.machine_summary(fdf, N_MONTHS, CFG)

    cA, cB = st.columns([1.25, 1])
    with cA:
        st.subheader("Output vs committed target")
        tgt_rows = summary[summary["Target (T)"].notna()]
        if tgt_rows.empty:
            note("info", "No machine in the current selection has a committed "
                         "target in config.json.")
        else:
            fig = go.Figure()
            fig.add_trace(bar(tgt_rows["Machine"].tolist(),
                              tgt_rows["Output (T)"].tolist(),
                              color=S1, name="Actual output",
                              text=[f"{v:,.0f}" for v in tgt_rows["Output (T)"]]))
            fig.add_trace(go.Scatter(
                x=tgt_rows["Machine"], y=tgt_rows["Target (T)"],
                mode="markers", name="Committed target",
                marker=dict(symbol="line-ew", size=26, line=dict(width=2.5, color=INK_1)),
                hovertemplate="Target %{y:,.0f} T<extra></extra>",
            ))
            pad_axis(fig, list(tgt_rows["Output (T)"]) + list(tgt_rows["Target (T)"]),
                     axis="y", factor=1.18)
            st.plotly_chart(style_fig(fig, 300, ylab="Tons"), width='stretch')

    with cB:
        st.subheader("Loss % by machine")
        s = summary.sort_values("Loss %", ascending=True)
        if len(s) < 2:
            # A one-bar bar chart is just a number wearing a costume.
            only = s.iloc[0]
            st.metric(f"{only['Machine']} loss", f"{only['Loss %']:.2f}%",
                      f"{only['Loss %'] - LOSS_TARGET:+.2f} pp vs {LOSS_TARGET}% target",
                      delta_color="inverse")
            st.caption("Only one machine is selected — a single bar would say no "
                       "more than this number does.")
        else:
            fig = go.Figure()
            fig.add_trace(bar(s["Machine"].tolist(), s["Loss %"].tolist(),
                              horizontal=True,
                              colors=loss_colors(s["Loss %"], LOSS_TARGET),
                              text=[f"{v:.2f}%" for v in s["Loss %"]]))
            fig.add_vline(x=LOSS_TARGET, line=dict(color=INK_1, width=1.5, dash="solid"),
                          annotation_text=f"target {LOSS_TARGET}%",
                          annotation_position="top",
                          annotation_font=dict(size=11, color=INK_2))
            pad_axis(fig, s["Loss %"])
            st.plotly_chart(style_fig(fig, 300, xlab="Loss %", legend=False),
                            width='stretch')
            st.caption("🟢 at or under target · 🟡 over target · 🔴 more than 1.5× "
                       "target. A machine can look bad here purely because it runs "
                       "narrow coil — the Machine Performance tab strips that out.")

    st.write("")
    st.subheader("What the data says")
    st.caption(f"Generated from the current filters. Any machine or band with "
               f"under {MIN_TONS:g} T of input is excluded from these statements.")

    big = summary[summary["Output (T)"] >= MIN_TONS]
    i1, i2 = st.columns(2)

    with i1:
        if not big.empty and big["Achieved %"].notna().any():
            ach = big[big["Achieved %"].notna()]
            best = ach.loc[ach["Achieved %"].idxmax()]
            worst = ach.loc[ach["Achieved %"].idxmin()]
            note("good", f"<b>{best['Machine']}</b> is the strongest against target at "
                         f"<b>{best['Achieved %']:.0f}%</b> "
                         f"({fmt(best['Output (T)'],0)} T against "
                         f"{fmt(best['Target (T)'],0)} T).")
            kind = "crit" if worst["Achieved %"] < 70 else "warn"
            note(kind, f"<b>{worst['Machine']}</b> is furthest behind at "
                       f"<b>{worst['Achieved %']:.0f}%</b> of target "
                       f"({fmt(worst['Output (T)'],0)} T against "
                       f"{fmt(worst['Target (T)'],0)} T).")
        else:
            note("info", "Not enough volume in the current selection to compare "
                         "machines against target.")

    with i2:
        kind = "good" if plant_loss <= LOSS_TARGET else (
            "warn" if plant_loss <= LOSS_TARGET * 1.5 else "crit")
        note(kind, f"Plant loss is <b>{plant_loss:.2f}%</b> of input "
                   f"({fmt(total_scrap)} T) against a {LOSS_TARGET:.1f}% target. "
                   f"Trimming is {fmt(total_trim)} T of that.")
        if not big.empty and big["Loss Gap"].notna().any():
            gap = big.loc[big["Loss Gap"].idxmax()]
            if gap["Loss Gap"] > 0.3:
                note("warn", f"After adjusting for the coil widths it actually runs, "
                             f"<b>{gap['Machine']}</b> still loses "
                             f"<b>{gap['Loss Gap']:+.2f} pp</b> more than expected "
                             f"({gap['Loss %']:.2f}% actual vs "
                             f"{gap['Expected Loss %']:.2f}% expected).")
            else:
                note("good", "No machine loses materially more than its coil-width "
                             "mix predicts — differences between machines are "
                             "explained by what they run, not how they run it. "
                             "See the Size &amp; Loss tab.")

    st.write("")
    st.subheader("Where the material goes")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Trimming", f"{fmt(total_trim)} T",
              f"{total_trim / total_in * 100:.2f}% of input" if total_in else None,
              delta_color="off")
    m2.metric("Scrap", f"{fmt(fdf['scrapWt'].sum())} T",
              f"{fdf['scrapWt'].sum() / total_in * 100:.2f}% of input" if total_in else None,
              delta_color="off")
    m3.metric("Excess", f"{fmt(fdf['excessWt'].sum())} T")
    fg = fdf.loc[fdf["materialType"] == "FG", "outputWt"].sum()
    m4.metric("Finished goods share",
              f"{fg / total_out * 100:.1f}%" if total_out else "—",
              f"{fmt(fg)} T of {fmt(total_out)} T", delta_color="off")

# =============================================================================
# TAB 2 — MACHINE PERFORMANCE
# =============================================================================
with tabs[1]:
    st.subheader("Machine summary")
    st.caption(
        "**Expected Loss %** is what this machine's loss would be if it ran at the "
        "plant-wide rate for every coil-width band, given the width mix it actually "
        "processed. **Loss Gap** = actual − expected: a positive gap is a genuine "
        "machine problem, a gap near zero means the headline loss is explained by "
        "the material mix."
    )
    show_cols = ["Machine", "Output (T)", "Target (T)", "Achieved %", "Loss %",
                 "Expected Loss %", "Loss Gap", "Jobs", "T / Job",
                 "Active Days", "T / Active Day", "Median TAT (d)"]
    num_fmt = {
        "Output (T)": "%.2f", "Target (T)": "%.0f", "Achieved %": "%.1f%%",
        "Loss %": "%.2f%%", "Expected Loss %": "%.2f%%", "Loss Gap": "%+.2f",
        "T / Job": "%.2f", "T / Active Day": "%.2f", "Median TAT (d)": "%.0f",
    }
    st.dataframe(
        summary[show_cols], hide_index=True, width='stretch',
        column_config={c: st.column_config.NumberColumn(c, format=f)
                       for c, f in num_fmt.items()},
    )
    dl_button(summary, "⬇ Download machine summary (CSV)",
              "machine_summary.csv", "dl_mach")

    st.write("")
    cA, cB = st.columns(2)
    with cA:
        st.subheader("Actual vs expected loss")
        st.caption("Bars to the right of zero lose more than their coil-width mix "
                   "predicts — that part is the machine. Bars to the left do better "
                   "than their mix.")
        s = summary.sort_values("Loss Gap", ascending=True)
        fig = go.Figure()
        fig.add_trace(bar(
            s["Machine"].tolist(), s["Loss Gap"].tolist(), horizontal=True,
            colors=[ST_CRIT if v > 0.3 else ST_WARN if v > 0 else ST_GOOD
                    for v in s["Loss Gap"]],
            text=[f"{v:+.2f}" for v in s["Loss Gap"]]))
        fig.add_vline(x=0, line=dict(color=INK_1, width=1.5))
        pad_axis(fig, list(s["Loss Gap"]) + [-max(abs(v) for v in s["Loss Gap"])],
                 factor=1.35)
        st.plotly_chart(style_fig(fig, 300, xlab="Percentage points vs expected",
                                  legend=False), width='stretch')

    with cB:
        st.subheader("Throughput per active day")
        st.caption("Total tonnage is not comparable across machines that run very "
                   "different coil sizes — this is.")
        s = summary.sort_values("T / Active Day", ascending=True)
        fig = go.Figure()
        fig.add_trace(bar(s["Machine"].tolist(), s["T / Active Day"].tolist(),
                          horizontal=True, color=S1,
                          text=[f"{v:,.1f}" for v in s["T / Active Day"]]))
        pad_axis(fig, s["T / Active Day"])
        st.plotly_chart(style_fig(fig, 300, xlab="Tons per active day", legend=False),
                        width='stretch')

    st.write("")
    st.subheader("Monthly output vs target")
    focus = st.selectbox("Machine", sel_machine, key="focus_machine")
    f_target = dl.machine_target(focus, 1, CFG)
    by_month = (fdf[fdf["machine"] == focus]
                .groupby("month", observed=True)["outputWt"].sum()
                .reindex(MONTHS).fillna(0))

    fig = go.Figure()
    fig.add_trace(bar(MONTHS, by_month.tolist(), color=S1, name=f"{focus} output",
                      text=[f"{v:,.0f}" for v in by_month]))
    if f_target:
        fig.add_trace(go.Scatter(
            x=MONTHS, y=[f_target] * len(MONTHS), mode="lines",
            name=f"Target ({f_target:,.0f} T/month)",
            line=dict(color=INK_1, width=2)))
    pad_axis(fig, list(by_month) + ([f_target] if f_target else []), axis="y", factor=1.20)
    if not f_target:
        st.caption(f"{focus} has no committed target in config.json — job work.")
    st.plotly_chart(style_fig(fig, 320, ylab="Tons", hover="x unified"),
                    width='stretch')

# =============================================================================
# TAB 3 — SIZE & LOSS
# =============================================================================
with tabs[2]:
    st.subheader("Loss is driven by coil width")
    st.caption("The single strongest pattern in the data. Read this tab before "
               "concluding that any one machine is weak.")

    wb = dl.width_band_summary(fdf)
    wb["Enough volume"] = wb["Input (T)"] >= MIN_TONS
    cA, cB = st.columns([1.15, 1])
    with cA:
        fig = go.Figure()
        fig.add_trace(bar(
            wb["RM Width Band"].astype(str).tolist(), wb["Loss %"].tolist(),
            colors=[S1 if ok else "#D8D6CF" for ok in wb["Enough volume"]],
            text=[f"{v:.2f}%" for v in wb["Loss %"]]))
        fig.add_hline(y=LOSS_TARGET, line=dict(color=INK_1, width=1.5),
                      annotation_text=f"plant target {LOSS_TARGET}%",
                      annotation_position="bottom right",
                      annotation_font=dict(size=11, color=INK_2))
        pad_axis(fig, wb["Loss %"], axis="y", factor=1.22)
        st.plotly_chart(style_fig(fig, 320, ylab="Loss %", xlab="RM width band",
                                  legend=False), width='stretch')
        st.caption(f"Grey bars carry under {MIN_TONS:g} T of input — shown for "
                   f"completeness, but too little volume to conclude anything from.")
    with cB:
        st.dataframe(wb.drop(columns=["Enough volume"]), hide_index=True,
                     width='stretch')

    solid = wb[wb["Enough volume"]]
    if len(solid) >= 2:
        worst = solid.loc[solid["Loss %"].idxmax()]
        best = solid.loc[solid["Loss %"].idxmin()]
        note("warn", f"Coil in the <b>{worst['RM Width Band']}</b> band loses "
                     f"<b>{worst['Loss %']:.2f}%</b> against "
                     f"<b>{best['Loss %']:.2f}%</b> in the "
                     f"<b>{best['RM Width Band']}</b> band — a "
                     f"{worst['Loss %'] / best['Loss %']:.1f}× difference. Edge trim "
                     f"is a roughly fixed width, so on a narrow strip it eats a far "
                     f"larger share of the coil.")

    st.write("")
    st.subheader("Loss % by machine and coil width")
    st.caption(f"Blank cells have under {MIN_TONS:g} T of input — too little to "
               f"read anything into. Compare <b>down a column</b> (machines running "
               f"the same width), never across a row.", unsafe_allow_html=True)

    pct, inp = dl.machine_band_matrix(fdf, MIN_TONS)
    if pct.empty or pct.isna().all().all():
        note("info", "Not enough volume per machine and width band to build the matrix.")
    else:
        cols = [str(c) for c in pct.columns]
        fig = go.Figure(go.Heatmap(
            z=pct.values, x=cols, y=pct.index.tolist(),
            colorscale=[[i / (len(BLUE_RAMP) - 1), c] for i, c in enumerate(BLUE_RAMP)],
            colorbar=dict(title=dict(text="Loss %", font=dict(size=11, color=INK_MUTED)),
                          tickfont=dict(size=11, color=INK_MUTED), thickness=12),
            xgap=2, ygap=2,
            text=[[f"{v:.2f}%" if pd.notna(v) else "" for v in row] for row in pct.values],
            texttemplate="%{text}", textfont=dict(size=11),
            hovertemplate="%{y} · %{x}<br>Loss %{z:.2f}%<extra></extra>",
        ))
        fig.update_yaxes(autorange="reversed")  # Machine 1 at the top
        st.plotly_chart(style_fig(fig, 90 + 42 * len(pct), legend=False),
                        width='stretch')

        with st.expander("Table view — loss % and the tonnage behind each cell"):
            tv = pct.round(2).astype(object)
            for r in tv.index:
                for c in tv.columns:
                    tons = inp.loc[r, c] if c in inp.columns else None
                    if pd.notna(tv.loc[r, c]):
                        tv.loc[r, c] = f"{tv.loc[r, c]:.2f}%  ({tons:,.0f} T)"
                    else:
                        tv.loc[r, c] = f"— ({tons:,.0f} T)" if pd.notna(tons) else "—"
            tv.columns = [str(c) for c in tv.columns]
            st.dataframe(tv, width='stretch')
            dl_button(pct.reset_index(), "⬇ Download matrix (CSV)",
                      "machine_width_loss.csv", "dl_matrix")

    st.write("")
    cA, cB = st.columns(2)
    with cA:
        st.subheader("Loss % by coil thickness")
        tb = (fdf.dropna(subset=["thicknessBand"])
              .groupby("thicknessBand", observed=True)
              .agg(inp=("inputWt", "sum"), scrap=("totalScrapWt", "sum"),
                   jobs=("sr", "count")).reset_index())
        tb["Loss %"] = (tb["scrap"] / tb["inp"] * 100).round(2)
        fig = go.Figure()
        fig.add_trace(bar(tb["thicknessBand"].astype(str).tolist(),
                          tb["Loss %"].tolist(), color=S1,
                          text=[f"{v:.2f}%" for v in tb["Loss %"]]))
        pad_axis(fig, tb["Loss %"], axis="y", factor=1.22)
        st.plotly_chart(style_fig(fig, 300, ylab="Loss %", xlab="RM thickness band",
                                  legend=False), width='stretch')

    with cB:
        st.subheader("The narrow-slit opportunity")
        opp = dl.narrow_slit_opportunity(fdf, CFG)
        if not opp:
            note("info", "Not enough narrow and wide coil in the current selection "
                         "to size this.")
        else:
            o1, o2 = st.columns(2)
            o1.metric(f"Coil ≤ {opp['narrow_cut_mm']:g} mm",
                      f"{fmt(opp['narrow_input_tons'], 1)} T",
                      f"{opp['narrow_jobs']:,} jobs", delta_color="off")
            o2.metric("Recoverable", f"{fmt(opp['recoverable_tons_per_month'], 2)} T/mo",
                      f"{fmt(opp['recoverable_tons_total'], 1)} T over "
                      f"{opp['months']} months", delta_color="off")
            note("warn",
                 f"Narrow coil loses <b>{opp['narrow_rate']:.2f}%</b> against "
                 f"<b>{opp['benchmark_rate']:.2f}%</b> on mid-width work. Closing "
                 f"that gap on the narrow volume alone recovers roughly "
                 f"<b>{opp['recoverable_tons_per_month']:.2f} tons a month</b>. "
                 f"Multiply by your realised price per ton for the rupee figure.")
            st.caption("Assumes narrow work can be brought to the mid-width loss "
                       "rate — the benchmark a dedicated narrow-slit setup with "
                       "tighter edge control would target. It is an upper bound, "
                       "not a committed saving.")

# =============================================================================
# TAB 4 — MATERIAL
# =============================================================================
with tabs[3]:
    mat = (fdf.groupby("gradeGroup", observed=True)
           .agg(Output=("outputWt", "sum"), Input=("inputWt", "sum"),
                Scrap=("totalScrapWt", "sum"), Jobs=("sr", "count"))
           .reset_index())
    mat["Loss %"] = (mat["Scrap"] / mat["Input"] * 100).round(2)
    mat["Output (T)"] = mat["Output"].round(2)
    mat["Share %"] = (mat["Output"] / mat["Output"].sum() * 100).round(1)
    mat = mat.rename(columns={"gradeGroup": "Metal"})

    cA, cB = st.columns(2)
    with cA:
        st.subheader("Output by metal")
        s = mat.sort_values("Output (T)", ascending=True)
        fig = go.Figure()
        fig.add_trace(bar(s["Metal"].tolist(), s["Output (T)"].tolist(),
                          horizontal=True, color=S1,
                          text=auto_labels(s["Output (T)"])))
        pad_axis(fig, s["Output (T)"])
        st.plotly_chart(style_fig(fig, 60 + 34 * len(s), xlab="Tons", legend=False),
                        width='stretch')

    with cB:
        st.subheader("Loss % by metal")
        st.caption(f"Metals with under {MIN_TONS:g} T of input are greyed out — "
                   "too small to draw conclusions from.")
        s = mat.sort_values("Loss %", ascending=True)
        colors = [INK_MUTED if i < MIN_TONS else c
                  for i, c in zip(s["Input"], loss_colors(s["Loss %"], LOSS_TARGET))]
        fig = go.Figure()
        fig.add_trace(bar(s["Metal"].tolist(), s["Loss %"].tolist(),
                          horizontal=True, colors=colors,
                          text=[f"{v:.2f}%" for v in s["Loss %"]]))
        fig.add_vline(x=LOSS_TARGET, line=dict(color=INK_1, width=1.5))
        pad_axis(fig, s["Loss %"])
        st.plotly_chart(style_fig(fig, 40 + 34 * len(s), xlab="Loss %", legend=False),
                        width='stretch')

    st.write("")
    st.subheader("Material summary")
    st.dataframe(
        mat[["Metal", "Output (T)", "Share %", "Loss %", "Jobs"]]
        .sort_values("Output (T)", ascending=False),
        hide_index=True, width='stretch',
        column_config={
            "Output (T)": st.column_config.NumberColumn(format="%.2f"),
            "Share %": st.column_config.NumberColumn(format="%.1f%%"),
            "Loss %": st.column_config.NumberColumn(format="%.2f%%"),
        })
    dl_button(mat, "⬇ Download material summary (CSV)", "material_summary.csv", "dl_mat")

    st.write("")
    st.subheader("What each machine mostly runs")
    rows = []
    for m in sorted(fdf["machine"].unique(), key=dl.machine_sort_key):
        sub = fdf[fdf["machine"] == m]
        top = sub.groupby("gradeGroup", observed=True)["inputWt"].sum().sort_values(ascending=False)
        if top.empty:
            continue
        rows.append({
            "Machine": m,
            "Main metal": top.index[0],
            "Its share of the machine's input": f"{top.iloc[0] / top.sum() * 100:.0f}%",
            "Median coil width (mm)": round(sub["widthMm"].median(), 1)
            if sub["widthMm"].notna().any() else None,
            "Jobs": len(sub),
        })
    st.dataframe(pd.DataFrame(rows), hide_index=True, width='stretch')

# =============================================================================
# TAB 5 — MONTHLY TREND
# =============================================================================
with tabs[4]:
    monthly = dl.monthly_summary(fdf, MONTHS, total_target / N_MONTHS if N_MONTHS else 0)

    cA, cB = st.columns(2)
    with cA:
        st.subheader("Output vs target by month")
        fig = go.Figure()
        fig.add_trace(bar(monthly["Month"].tolist(), monthly["Output (T)"].tolist(),
                          color=S1, name="Output",
                          text=[f"{v:,.0f}" for v in monthly["Output (T)"]]))
        if total_target:
            fig.add_trace(go.Scatter(
                x=monthly["Month"], y=monthly["Target (T)"], mode="lines",
                name="Target", line=dict(color=INK_1, width=2)))
        pad_axis(fig, list(monthly["Output (T)"]) + list(monthly["Target (T)"]),
                 axis="y", factor=1.18)
        st.plotly_chart(style_fig(fig, 320, ylab="Tons", hover="x unified"),
                        width='stretch')

    with cB:
        st.subheader("Loss % trend with control limits")
        mean = monthly["Loss %"].mean()
        sd = monthly["Loss %"].std(ddof=0)
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=monthly["Month"], y=monthly["Loss %"], mode="lines+markers",
            name="Loss %", line=dict(color=S1, width=2),
            marker=dict(size=8, line=dict(width=2, color=SURFACE))))
        for val, label, dash in [(mean, "mean", "solid"),
                                 (mean + 2 * sd, "+2σ", "dot"),
                                 (max(0, mean - 2 * sd), "−2σ", "dot")]:
            fig.add_hline(y=val, line=dict(color=INK_MUTED, width=1, dash=dash),
                          annotation_text=f"{label} {val:.2f}%",
                          annotation_font=dict(size=10, color=INK_MUTED))
        fig.add_hline(y=LOSS_TARGET, line=dict(color=ST_CRIT, width=1.5),
                      annotation_text=f"target {LOSS_TARGET}%",
                      annotation_font=dict(size=10, color=ST_CRIT))
        st.plotly_chart(style_fig(fig, 320, ylab="Loss %", legend=False),
                        width='stretch')
        st.caption("Points inside ±2σ are normal variation, not a story. Only a "
                   "point outside the band is a real signal worth investigating.")

    st.write("")
    st.subheader("Month by month")
    disp = monthly[["Month", "Output (T)", "Target (T)", "Achieved %",
                    "MoM Output %", "Loss %", "MoM Loss (pp)", "Jobs"]]
    st.dataframe(
        disp, hide_index=True, width='stretch',
        column_config={
            "Output (T)": st.column_config.NumberColumn(format="%.2f"),
            "Target (T)": st.column_config.NumberColumn(format="%.0f"),
            "Achieved %": st.column_config.NumberColumn(format="%.1f%%"),
            "MoM Output %": st.column_config.NumberColumn(format="%+.1f%%"),
            "Loss %": st.column_config.NumberColumn(format="%.2f%%"),
            "MoM Loss (pp)": st.column_config.NumberColumn(format="%+.2f"),
        })
    dl_button(disp, "⬇ Download monthly summary (CSV)", "monthly_summary.csv", "dl_mo")

    st.write("")
    st.subheader("Output trend by machine")
    pivot = (fdf.pivot_table(index="month", columns="machine", values="outputWt",
                             aggfunc="sum", observed=True)
             .reindex(MONTHS).fillna(0))
    ordered = sorted(pivot.columns, key=dl.machine_sort_key)
    # Colour follows the machine, not its rank — filtering never repaints a series.
    color_of = {m: SERIES[i % len(SERIES)]
                for i, m in enumerate(sorted(meta["machines"], key=dl.machine_sort_key))}
    fig = go.Figure()
    for m in ordered:
        fig.add_trace(go.Scatter(
            x=MONTHS, y=pivot[m], mode="lines+markers", name=m,
            line=dict(color=color_of[m], width=2),
            marker=dict(size=7, line=dict(width=2, color=SURFACE))))
    st.plotly_chart(style_fig(fig, 360, ylab="Tons", hover="x unified"),
                    width='stretch')
    with st.expander("Table view"):
        st.dataframe(pivot.round(2), width='stretch')

# =============================================================================
# TAB 6 — OPERATORS
# =============================================================================
with tabs[5]:
    st.caption("Spelling variants are merged using the alias list in config.json. "
               "Outside job-work firms are reported separately from in-house "
               "operators — mixing them makes both comparisons meaningless.")

    for label, kind in [("In-house operators", "In-house"), ("Outside job work", "Vendor")]:
        sub = fdf[fdf["operatorType"] == kind]
        if sub.empty:
            continue
        st.subheader(label)
        ops = (sub.groupby("operator", observed=True)
               .agg(Output=("outputWt", "sum"), Input=("inputWt", "sum"),
                    Scrap=("totalScrapWt", "sum"), Jobs=("sr", "count"),
                    Shared=("operatorShared", "max"))
               .reset_index())
        ops["Loss %"] = (ops["Scrap"] / ops["Input"] * 100).round(2)
        ops["Output (T)"] = ops["Output"].round(2)
        ops["T / Job"] = (ops["Output"] / ops["Jobs"]).round(2)
        ops["Operator"] = ops.apply(
            lambda r: f"{r['operator']} (shared job)" if r["Shared"] else r["operator"], axis=1)

        cA, cB = st.columns([1.1, 1])
        with cA:
            s = ops.sort_values("Output (T)", ascending=True)
            fig = go.Figure()
            fig.add_trace(bar(s["Operator"].tolist(), s["Output (T)"].tolist(),
                              horizontal=True, color=S1,
                              text=auto_labels(s["Output (T)"])))
            pad_axis(fig, s["Output (T)"])
            st.plotly_chart(style_fig(fig, 80 + 32 * len(s), xlab="Tons", legend=False),
                            width='stretch')
        with cB:
            st.dataframe(
                ops[["Operator", "Output (T)", "Loss %", "Jobs", "T / Job"]]
                .sort_values("Output (T)", ascending=False),
                hide_index=True, width='stretch', height=80 + 35 * len(s),
                column_config={
                    "Output (T)": st.column_config.NumberColumn(format="%.2f"),
                    "Loss %": st.column_config.NumberColumn(format="%.2f%%"),
                    "T / Job": st.column_config.NumberColumn(format="%.2f"),
                })
        st.caption("Loss % here reflects the coil the operator was given as much as "
                   "how they ran it — read it alongside the Size & Loss tab.")
        st.write("")

# =============================================================================
# TAB 7 — JOB EXPLORER
# =============================================================================
with tabs[6]:
    st.subheader("Highest-loss jobs")
    st.caption("The individual jobs worth pulling the P-card for. Only jobs above "
               "0.05 T of input are listed, so a 40-kg offcut does not top the table.")
    worst = (fdf[fdf["inputWt"] >= 0.05]
             .nlargest(20, "lossPct")
             [["dateEnd", "coil", "gradeRaw", "sizeRaw", "machine", "operator",
               "inputWt", "totalScrapWt", "lossPct"]]
             .rename(columns={"dateEnd": "Completed", "coil": "Coil No",
                              "gradeRaw": "Grade", "sizeRaw": "RM/Size",
                              "machine": "Machine", "operator": "Operator",
                              "inputWt": "Input (T)", "totalScrapWt": "Scrap (T)",
                              "lossPct": "Loss %"}))
    st.dataframe(
        worst.round(3), hide_index=True, width='stretch',
        column_config={
            "Input (T)": st.column_config.NumberColumn(format="%.3f"),
            "Scrap (T)": st.column_config.NumberColumn(format="%.3f"),
            "Loss %": st.column_config.NumberColumn(format="%.2f%%"),
        })
    dl_button(worst, "⬇ Download highest-loss jobs (CSV)", "worst_jobs.csv", "dl_worst")

    st.write("")
    st.subheader("Search all jobs")
    q = st.text_input("Coil no, grade, P-card no, operator, date or size", "")
    show = fdf
    if q.strip():
        ql = q.strip().lower()
        cols = ["coil", "gradeRaw", "gradeGroup", "pcard", "operator", "dateEnd",
                "dateStart", "sizeRaw", "machine"]
        hit = pd.Series(False, index=fdf.index)
        for c in cols:
            hit |= fdf[c].astype(str).str.lower().str.contains(ql, na=False, regex=False)
        show = fdf[hit]

    display_cols = {
        "sr": "Sr.No", "coil": "Coil No", "gradeRaw": "Grade", "gradeGroup": "Metal",
        "pcard": "P.Card No", "machine": "Machine", "materialType": "Type",
        "sizeRaw": "RM/Size", "widthMm": "Width (mm)", "dateStart": "Sent Out",
        "dateEnd": "Returned", "tatDays": "TAT (d)", "operator": "Operator",
        "inputWt": "Input (T)", "outputWt": "Output (T)",
        "totalScrapWt": "Scrap (T)", "lossPct": "Loss %",
    }
    table = show[list(display_cols)].rename(columns=display_cols).round(3)

    PAGE = 200
    pages = max(1, -(-len(table) // PAGE))
    c1, c2 = st.columns([1, 3])
    page = c1.number_input(f"Page (of {pages})", 1, pages, 1, key="job_page")
    c2.caption(f"**{len(table):,}** matching records · showing rows "
               f"{(page - 1) * PAGE + 1:,}–{min(page * PAGE, len(table)):,}. "
               f"The CSV below contains all {len(table):,}.")
    st.dataframe(table.iloc[(page - 1) * PAGE: page * PAGE],
                 hide_index=True, width='stretch', height=460)
    dl_button(table, f"⬇ Download all {len(table):,} matching records (CSV)",
              "jobs.csv", "dl_jobs")

# =============================================================================
# TAB 8 — DATA HEALTH
# =============================================================================
with tabs[7]:
    st.subheader("Data quality checks")
    st.caption("Run against **all** records in the file, not the current filters. "
               "Nothing here is corrected automatically — the dashboard reports "
               "what is in the sheet so it can be fixed at source.")

    health = dl.data_health_report(df_all)
    clean = int((df_all["issueCount"] == 0).sum())

    h = st.columns(4)
    h[0].metric("Records in file", f"{len(df_all):,}")
    h[1].metric("Clean records", f"{clean:,}",
                f"{clean / len(df_all) * 100:.1f}%", delta_color="off")
    h[2].metric("Records with an issue", f"{len(df_all) - clean:,}")
    h[3].metric("Checks failing", f"{int((health['Rows'] > 0).sum())} of {len(health)}")

    st.write("")
    st.dataframe(health[["Check", "Rows", "% of Data", "Why it matters"]],
                 hide_index=True, width='stretch')

    st.write("")
    st.subheader("Inspect a check")
    failing = health[health["Rows"] > 0]
    if failing.empty:
        note("good", "Every check passes. Nothing to fix.")
    else:
        pick = st.selectbox("Check", failing["Check"].tolist(), key="health_pick")
        col = failing.loc[failing["Check"] == pick, "_col"].iloc[0]
        bad = df_all[df_all[col]]
        cols = {"sr": "Sr.No", "coil": "Coil No", "gradeRaw": "Grade",
                "sizeRaw": "RM/Size", "machine": "Machine", "operatorRaw": "Operator",
                "dateStart": "Sent Out", "dateEnd": "Returned", "tatDays": "TAT (d)",
                "inputWt": "Input (T)", "outputWt": "Output (T)",
                "totalScrapWt": "Scrap (T)", "massBalanceResidual": "Balance Gap (T)",
                "materialType": "Type"}
        out = bad[list(cols)].rename(columns=cols).round(4)
        st.caption(f"{len(out):,} record(s).")
        st.dataframe(out.head(300), hide_index=True, width='stretch', height=400)
        dl_button(out, f"⬇ Download these {len(out):,} records (CSV)",
                  "data_issues.csv", "dl_health")

    st.write("")
    with st.expander("Why 'Yield %' was removed as a KPI"):
        y = df_all["outputWt"].sum() / df_all["inputWt"].sum() * 100
        st.markdown(
            f"Output ÷ Input across the whole file is **{y:.2f}%**, and it sits "
            f"between 99.8% and 100.0% for every single machine. Scrap and trimming "
            f"are *not* deducted from `SLITTING OUT WT`, so that ratio cannot "
            f"discriminate between a good month and a bad one — it was a flat line "
            f"by construction.\n\n"
            f"**Loss % (TOTAL SCRAP ÷ SLITTING IN)** is used everywhere instead. It "
            f"ranges from about 1% to 10% across the data and matches the "
            f"`PERCENTAGE(%)` column already in the sheet, so the dashboard and the "
            f"shop-floor Excel tell the same story."
        )

st.caption("Production Analytics · reads Production_Report.xlsx on "
           "demand · press **Refresh data** in the sidebar after saving the Excel file.")

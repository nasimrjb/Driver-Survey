"""
Driver Survey Routine Analysis — HTML Dashboard
================================================
Reads the same processed CSV files as survey_routine_analysis.py and renders
color-coded matrix tables (city × metric) directly as an HTML dashboard.
Cutoff values (SHEET_MIN_N) live in survey_routine_analysis.py and are applied
via sra.apply_min_n_cutoff — exactly as in the Excel version.

Usage:
    python build_routine_dashboard.py              # auto-detect latest week
    python build_routine_dashboard.py 52           # specific week number
"""

import sys, os, re
import numpy as np
import pandas as pd

# ── Locate and import the analysis module ───────────────────────────────────
SRA_DIR = r"D:\OneDrive\Documents\Snapp\Driver Survey"
if SRA_DIR not in sys.path:
    sys.path.insert(0, SRA_DIR)

import survey_routine_analysis as sra   # noqa: E402

# ── Output path ─────────────────────────────────────────────────────────────
OUTPUT = r"D:\Work\Driver Survey\PowerBI\RoutineAnalysis_Dashboard.html"

# ── Brand / UI colors ────────────────────────────────────────────────────────
SNAPP_GREEN      = "#00C853"
HEADER_BG        = "#1e3a5f"
HEADER_BG_LIGHT  = "#D9E1F2"
GROUP_HEADER_BG  = "#4472C4"
BODY_BG          = "#f5f6fa"

# Excel-matching color scale RGB tuples
WHITE   = (255, 255, 255)
GREEN   = (99,  190, 123)   # #63BE7B
RED     = (248, 105, 107)   # #F8696B
YELLOW  = (255, 235, 132)   # #FFEB84
LT_WHT  = (252, 252, 255)   # #FCFCFF — near-white baseline

# ── Tab group definitions (prefix → tab name) ────────────────────────────────
TAB_GROUPS = [
    ("Incentive",          ["#1_", "#2_", "#4_", "#5_6_", "#8_", "#9_"]),
    ("Satisfaction",       ["#3_Sat_", "#12_"]),
    ("Market Share",       ["#13_", "#15_Persona_PartTime"]),
    ("Personas",           ["#14_", "#15_", "#16_", "#17_"]),
    ("Commission & Wheel", ["#18_", "#19_", "#20_"]),
    ("Support & NPS",      ["#CS_", "#Reco_", "#NavReco_"]),
    ("Referral & Reg",     ["#Refer_", "#Reg_", "#Income_", "#Decline_"]),
    ("Services",           ["#Carfix_", "#Garage_"]),
    ("Operations",         ["#Demand_", "#Speed_", "#DistOrigin_", "#GPS_", "#Unpaid_"]),
]

# Satisfaction sheets with fixed 1-5 red/yellow/green scale
_SAT_FIXED = {"#3_Sat_", "#CS_Sat_", "#Carfix_Sat_", "#Garage_Sat_"}

# Dissatisfaction sheets: higher = worse (white → red)
_DISSAT = {"#8_Dissat", "#9_Dissat"}


# ══════════════════════════════════════════════════════════════════════════════
#  COLOR HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def lerp(c1, c2, t):
    t = max(0.0, min(1.0, t))
    return tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))


def css_rgb(c):
    return f"rgb({c[0]},{c[1]},{c[2]})"


def lum(c):
    return 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]


def cell_color(val, col, sheet_name, lo, hi):
    """Return (bg_rgb, text_color_str) for a data cell, or (None, None)."""
    if pd.isna(val) or col in sra.NON_PCT_COLS or lo == hi:
        return None, None

    is_pct_col = "Part%" in col or "GotMsg%" in col
    is_fixed_sat = any(sheet_name.startswith(p) for p in _SAT_FIXED)
    is_dissat = any(sheet_name.startswith(p) for p in _DISSAT)

    t = max(0.0, min(1.0, (val - lo) / (hi - lo))) if hi != lo else 0.0

    if is_fixed_sat and not is_pct_col:
        # Fixed 1-5 scale: red→yellow→green
        if val <= 3:
            t2 = (val - 1) / 2 if val >= 1 else 0.0
            bg = lerp(RED, YELLOW, t2)
        else:
            t2 = (val - 3) / 2
            bg = lerp(YELLOW, GREEN, t2)
    elif is_dissat:
        bg = lerp(WHITE, RED, t)
    elif sheet_name.startswith("#12_Cities_Overview") and "Dual SU" in col:
        bg = lerp(WHITE, RED, t)
    elif sheet_name.startswith("#13_RideShare") and "Tapsi" in col and "@" in col:
        bg = lerp(WHITE, RED, t)
    else:
        bg = lerp(LT_WHT, GREEN, t)

    text = "#000" if lum(bg) > 140 else "#fff"
    return bg, text


# ══════════════════════════════════════════════════════════════════════════════
#  VALUE FORMATTING
# ══════════════════════════════════════════════════════════════════════════════

def fmt(val, col, sheet_name):
    """Return a display string for a cell value."""
    if pd.isna(val):
        return ""
    if col in sra.NON_PCT_COLS:
        try:
            return f"{int(round(float(val))):,}"
        except (ValueError, TypeError):
            return str(val)

    is_pct_col   = "Part%" in col or "GotMsg%" in col
    is_fixed_sat = any(sheet_name.startswith(p) for p in _SAT_FIXED)

    if is_fixed_sat and not is_pct_col:
        return f"{val:.2f}"
    if sheet_name.startswith("#NavReco_"):
        return f"{val:.2f}"
    if sra.is_pct_sheet(sheet_name):
        return f"{val:.1f}%"
    return f"{val:.2f}"


# ══════════════════════════════════════════════════════════════════════════════
#  PER-COLUMN MIN/MAX (group-aware for shared color scales)
# ══════════════════════════════════════════════════════════════════════════════

def col_ranges(df, grp_headers=None):
    """Return {col: (lo, hi)} for color-scalable numeric columns."""
    data = df[df.index != "Total"] if "Total" in df.index else df
    rng = {}
    for col in df.columns:
        if col in sra.NON_PCT_COLS:
            continue
        if not pd.api.types.is_numeric_dtype(df[col]):
            continue
        vals = pd.to_numeric(data[col], errors="coerce").dropna()
        if len(vals):
            rng[col] = (float(vals.min()), float(vals.max()))

    # Apply group-wide range for grouped sheets
    if grp_headers:
        for prefix, _ in grp_headers:
            grp = [c for c in df.columns if c.startswith(prefix) and c in rng]
            if len(grp) >= 2:
                all_v = pd.concat([pd.to_numeric(data[c], errors="coerce") for c in grp]).dropna()
                if len(all_v):
                    g_lo, g_hi = float(all_v.min()), float(all_v.max())
                    for c in grp:
                        rng[c] = (g_lo, g_hi)
    return rng


# ══════════════════════════════════════════════════════════════════════════════
#  HTML TABLE RENDERER
# ══════════════════════════════════════════════════════════════════════════════

_COL_RENAME = {"E_n": "#Resp", "F_n": "#Joint", "G_n": "#Cmpt"}


def render_table(df, sheet_name, title):
    """Render a single DataFrame as a color-coded HTML table block."""
    if df is None or df.empty:
        return f'<div class="table-wrapper"><h3 class="sheet-title">{title}</h3><p class="no-data">No data</p></div>'

    grp_headers = None
    for prefix, headers in sra.SHEET_GROUP_HEADERS.items():
        if sheet_name.startswith(prefix) and headers:
            grp_headers = headers
            break

    ranges = col_ranges(df, grp_headers)

    def _disp_col(col):
        if grp_headers:
            for pfx, _ in grp_headers:
                if col.startswith(pfx):
                    s = col[len(pfx):]
                    if s.endswith("_WoW"):
                        return "WoW"
                    if s.endswith("_Last"):
                        return "Last"
                    return _COL_RENAME.get(s, s)
        if col.endswith("_WoW"):
            return "WoW"
        if col.endswith("_Last"):
            return "Last"
        return _COL_RENAME.get(col, col)

    out = [f'<div class="table-wrapper">'
           f'<h3 class="sheet-title">{title}</h3>'
           f'<div class="table-scroll"><table class="heatmap-table">']
    out.append("<thead>")

    # Group header row
    if grp_headers:
        out.append('<tr class="group-header-row"><th class="group-header empty-group" rowspan="2">City</th>')
        covered = set()
        for pfx, label in grp_headers:
            gcols = [c for c in df.columns if c.startswith(pfx)]
            if gcols:
                out.append(f'<th class="group-header" colspan="{len(gcols)}">{label}</th>')
                covered.update(gcols)
        for c in df.columns:
            if c not in covered:
                out.append('<th class="group-header empty-group"></th>')
        out.append("</tr>")
        # Sub-column headers
        out.append('<tr class="col-header-row">')
        for col in df.columns:
            d = _disp_col(col)
            cls = "col-header count-col" if col in sra.NON_PCT_COLS else "col-header"
            out.append(f'<th class="{cls}">{d}</th>')
        out.append("</tr>")
    else:
        out.append('<tr class="col-header-row"><th class="col-header">City</th>')
        for col in df.columns:
            d = _disp_col(col)
            cls = "col-header count-col" if col in sra.NON_PCT_COLS else "col-header"
            out.append(f'<th class="{cls}">{d}</th>')
        out.append("</tr>")

    out.append("</thead><tbody>")

    for city, row in df.iterrows():
        is_total = city == "Total"
        out.append(f'<tr class="{"total-row" if is_total else "data-row"}">')
        out.append(f'<td class="city-cell">{city}</td>')
        for col in df.columns:
            val = row[col]
            if pd.isna(val):
                out.append('<td class="na-cell">-</td>')
                continue
            lo, hi = ranges.get(col, (0, 1))
            bg, tc = cell_color(val, col, sheet_name, lo, hi) if not is_total else (None, None)
            text = fmt(val, col, sheet_name)
            if bg:
                out.append(f'<td style="background:{css_rgb(bg)};color:{tc}">{text}</td>')
            else:
                out.append(f'<td>{text}</td>')
        out.append("</tr>")

    out.append("</tbody></table></div></div>")
    return "\n".join(out)


# ══════════════════════════════════════════════════════════════════════════════
#  DATA COLLECTION (same logic as survey_routine_analysis.run_all, no Excel export)
# ══════════════════════════════════════════════════════════════════════════════

def collect_sheets(week=None):
    """Load CSVs → run analysis → return cleaned dict of DataFrames."""
    data = sra.load_data()
    if week is None:
        week = sra.get_latest_week(data["short_main"])
    print(f"\nRunning analyses for Week {week}")
    print("=" * 55)

    sh = {}

    def _s(tag, fn, *a):
        try:
            return fn(*a)
        except Exception as e:
            print(f"  [skip] {tag}: {e}")
            return pd.DataFrame()

    def _m(tag, fn, *a):
        try:
            return fn(*a)
        except Exception as e:
            print(f"  [skip] {tag}: {e}")
            return {}

    sh["#1_Snapp_Incentive_Amt"] = _s("#1",  sra.analysis_incentive_amounts_snapp, data, week)
    sh["#2_Tapsi_Incentive_Amt"] = _s("#2",  sra.analysis_incentive_amounts_tapsi, data, week)
    for seg, df in _m("#3", sra.analysis_satisfaction_review, data, week).items():
        sh[f"#3_Sat_{seg[:20]}"] = df
    sh["#4_Incentive_Duration"] = _s("#4",   sra.analysis_incentive_time_limitation, data, week)
    sh["#5_6_IncType"]          = _s("#5_6", sra.analysis_received_incentive_types, data, week)
    try:
        d = sra.analysis_incentive_dissatisfaction(data, week)
        sh["#8_Dissat"] = d["combined"]
        if isinstance(d.get("summary"), pd.DataFrame) and not d["summary"].empty:
            sh["#9_Dissat_Sum"] = d["summary"]
        if isinstance(d.get("summary_last_week"), pd.DataFrame) and not d["summary_last_week"].empty:
            sh["#9_Dissat_Sum_LastWk"] = d["summary_last_week"]
    except Exception as e:
        print(f"  [skip] #8/#9 dissatisfaction: {e}")

    sh["#12_Cities_Overview"]    = _s("#12",    sra.analysis_all_cities_overview, data, week)
    sh["#13_RideShare"]          = _s("#13",    sra.analysis_ride_share, data, week)
    for label, df in _m("#14", sra.analysis_navigation_usage, data, week).items():
        sh[f"#14_Nav_{label[:20]}"] = df
    for label, df in _m("#15_Persona", sra.analysis_driver_persona, data, week).items():
        sh[f"#15_Persona_{label[:16]}"] = df
    sh["#15_Persona_PartTime"]   = _s("#15_PT", sra.analysis_driver_persona_parttime_rides, data, week)
    for label, df in _m("#16", sra.analysis_referral_plan, data, week).items():
        sh[f"#16_Ref_{label[:20]}"] = df
    sh["#17_Inactivity"]         = _s("#17",    sra.analysis_inactivity_before_incentive, data, week)
    for label, df in _m("#18", sra.analysis_commission_free, data, week).items():
        sh[f"#18_CommFree_{label}"] = df
    sh["#19_LuckyWheel"]         = _s("#19",    sra.analysis_lucky_wheel, data, week)
    for label, df in _m("#20", sra.analysis_request_refusal, data, week).items():
        sh[f"#20_Refusal_{label.replace(' ','_')[:20]}"] = df
    for plat, df in _m("#CS_Sat",    sra.analysis_cs_satisfaction,     data, week).items():
        sh[f"#CS_Sat_{plat}"] = df
    for label, df in _m("#CS_Cat",   sra.analysis_cs_categories,       data, week).items():
        sh[f"#CS_Cat_{label.replace(' ','_')[:18]}"] = df
    for plat, df in _m("#CS_Reason", sra.analysis_cs_important_reason, data, week).items():
        sh[f"#CS_Reason_{plat}"] = df
    reco = _s("#Reco", sra.analysis_recommend, data, week)
    if isinstance(reco, pd.DataFrame) and not reco.empty:
        sh["#Reco_NPS"] = reco
    for label, df in _m("#Refer", sra.analysis_refer_others, data, week).items():
        sh[f"#Refer_{label.replace(' ','_')[:18]}"] = df
    nav_reco = _s("#NavReco", sra.analysis_navigation_recommendations, data, week)
    if isinstance(nav_reco, pd.DataFrame) and not nav_reco.empty:
        sh["#NavReco_Scores"] = nav_reco
    for label, df in _m("#Reg",    sra.analysis_registration,  data, week).items():
        sh[f"#Reg_{label.replace(' ','_')[:18]}"] = df
    for label, df in _m("#Income", sra.analysis_better_income, data, week).items():
        sh[f"#Income_{label.replace(' ','_')[:18]}"] = df
    decline = _s("#Decline", sra.analysis_decline_reasons, data, week)
    if isinstance(decline, pd.DataFrame) and not decline.empty:
        sh["#Decline_Reasons"] = decline
    for label, df in _m("#Carfix", sra.analysis_snappcarfix_satisfaction, data, week).items():
        sh[f"#Carfix_{label[:20]}"] = df
    for label, df in _m("#Garage", sra.analysis_tapsigarage_satisfaction, data, week).items():
        sh[f"#Garage_{label[:20]}"] = df
    for tag, fn, key in [
        ("#Demand",     sra.analysis_demand,             "#Demand_Perception"),
        ("#Speed",      sra.analysis_speed_satisfaction, "#Speed_Satisfaction"),
        ("#DistOrigin", sra.analysis_distance_to_origin, "#DistOrigin_Sat"),
    ]:
        v = _s(tag, fn, data, week)
        if isinstance(v, pd.DataFrame) and not v.empty:
            sh[key] = v
    for label, df in _m("#GPS",    sra.analysis_gps,                 data, week).items():
        sh[f"#GPS_{label.replace(' ','_')[:18]}"] = df
    for label, df in _m("#Unpaid", sra.analysis_unpaid_by_passenger, data, week).items():
        sh[f"#Unpaid_{label.replace(' ','_')[:18]}"] = df

    # Drop all-NaN-column sheets / empty sheets (mirrors run_all cleanup)
    meta = {"n", "n_joint", "n_dissatisfied", "n_contacted"}
    cleaned = {}
    for name, df in sh.items():
        if df is None or (isinstance(df, pd.DataFrame) and df.empty):
            continue
        drop = {c for c in df.columns if c not in meta and df[c].isna().all()}
        df = df[[c for c in df.columns if c not in drop]]
        if not [c for c in df.columns if c not in meta]:
            continue
        cleaned[name] = df

    print(f"\n  {len(cleaned)} sheets ready for rendering.")
    return cleaned, week


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET PREPARATION (same transforms as Excel export)
# ══════════════════════════════════════════════════════════════════════════════

def prepare(df, sheet_name):
    df = sra.sort_cities(df)
    df = sra.reorder_columns(df, sheet_name)
    df = sra.apply_min_n_cutoff(df, sheet_name)
    return df


# ══════════════════════════════════════════════════════════════════════════════
#  HTML ASSEMBLY
# ══════════════════════════════════════════════════════════════════════════════

def build_html(sheets, week):
    tab_panels   = []
    tab_buttons  = []
    all_assigned = set()

    for i, (tab_name, prefixes) in enumerate(TAB_GROUPS):
        # Exact name match OR prefix match
        matching = {}
        for k, v in sheets.items():
            for p in prefixes:
                if k == p or k.startswith(p):
                    matching[k] = v
                    break
        if not matching:
            continue
        all_assigned.update(matching.keys())

        cards = []
        for sheet_name, df in matching.items():
            df    = prepare(df, sheet_name)
            title = sheet_name.lstrip("#").replace("_", " ")
            cards.append(render_table(df, sheet_name, title))

        active = "active" if i == 0 else ""
        disp   = "block"  if i == 0 else "none"
        tab_buttons.append(
            f'<button class="tab-btn {active}" onclick="switchTab(event,\'tab-{i}\')">'
            f'{tab_name}</button>'
        )
        tab_panels.append(
            f'<div id="tab-{i}" class="tab-panel" style="display:{disp}">'
            + "\n".join(cards) + "</div>"
        )

    # Unassigned sheets → Other tab
    unassigned = {k: v for k, v in sheets.items() if k not in all_assigned}
    if unassigned:
        idx = len(tab_buttons)
        cards = []
        for sheet_name, df in unassigned.items():
            df    = prepare(df, sheet_name)
            title = sheet_name.lstrip("#").replace("_", " ")
            cards.append(render_table(df, sheet_name, title))
        tab_buttons.append(
            f'<button class="tab-btn" onclick="switchTab(event,\'tab-{idx}\')">'
            f'Other</button>'
        )
        tab_panels.append(
            f'<div id="tab-{idx}" class="tab-panel" style="display:none">'
            + "\n".join(cards) + "</div>"
        )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Driver Survey Routine Analysis — Week {week}</title>
<style>
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background: {BODY_BG};
    color: #333;
    line-height: 1.4;
}}
.dashboard-header {{
    background: linear-gradient(135deg, {HEADER_BG} 0%, #2c5282 100%);
    color: white;
    padding: 16px 30px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 100;
    box-shadow: 0 2px 8px rgba(0,0,0,0.18);
}}
.dashboard-header h1 {{ font-size: 20px; font-weight: 600; }}
.week-badge {{
    background: {SNAPP_GREEN};
    color: white;
    padding: 3px 14px;
    border-radius: 20px;
    font-size: 13px;
    font-weight: bold;
}}
.legend-bar {{
    background: white;
    border-bottom: 1px solid #e2e8f0;
    padding: 7px 24px;
    display: flex;
    gap: 20px;
    align-items: center;
    flex-wrap: wrap;
    font-size: 11px;
    color: #555;
}}
.legend-item {{ display: flex; align-items: center; gap: 5px; }}
.swatch {{
    width: 56px; height: 12px;
    border-radius: 3px;
    border: 1px solid #ccc;
}}
.sw-green  {{ background: linear-gradient(90deg, {css_rgb(LT_WHT)}, {css_rgb(GREEN)}); }}
.sw-red    {{ background: linear-gradient(90deg, {css_rgb(WHITE)}, {css_rgb(RED)}); }}
.sw-3color {{ background: linear-gradient(90deg, {css_rgb(RED)}, {css_rgb(YELLOW)}, {css_rgb(GREEN)}); }}
.tab-bar {{
    background: #f0f4f8;
    border-bottom: 2px solid #e2e8f0;
    padding: 0 16px;
    display: flex;
    gap: 2px;
    overflow-x: auto;
    position: sticky;
    top: 56px;
    z-index: 99;
    scrollbar-width: thin;
}}
.tab-btn {{
    padding: 9px 16px;
    border: none;
    background: transparent;
    font-size: 12px;
    font-weight: 500;
    color: #666;
    cursor: pointer;
    border-bottom: 3px solid transparent;
    white-space: nowrap;
    transition: all 0.15s;
}}
.tab-btn:hover {{ color: #333; background: rgba(0,0,0,0.04); }}
.tab-btn.active {{
    color: {HEADER_BG};
    border-bottom-color: {HEADER_BG};
    font-weight: 600;
}}
.tab-panel {{ padding: 16px 20px; }}
.table-wrapper {{
    background: white;
    border-radius: 7px;
    box-shadow: 0 1px 5px rgba(0,0,0,0.08);
    margin-bottom: 20px;
    overflow: hidden;
}}
.sheet-title {{
    font-size: 14px;
    font-weight: 600;
    padding: 10px 14px 7px;
    color: {HEADER_BG};
    border-bottom: 1px solid #eee;
    background: #fafbfc;
}}
.table-scroll {{
    overflow-x: auto;
    overflow-y: auto;
    max-height: 680px;
}}
.no-data {{ padding: 10px 14px; color: #999; font-style: italic; font-size: 12px; }}
.heatmap-table {{
    border-collapse: collapse;
    font-size: 11px;
    white-space: nowrap;
    min-width: max-content;
}}
.heatmap-table th,
.heatmap-table td {{
    padding: 3px 8px;
    border: 1px solid #d4d8dd;
    text-align: center;
}}
.group-header-row {{ background: {GROUP_HEADER_BG}; }}
.group-header {{
    color: white;
    font-weight: 700;
    font-size: 11px;
    padding: 5px 9px;
    background: {GROUP_HEADER_BG};
    border-color: #3a6bb5;
}}
.empty-group {{ background: #e0e6f0; }}
.col-header-row {{ background: {HEADER_BG_LIGHT}; }}
.col-header {{
    font-weight: 600;
    font-size: 10px;
    color: #333;
    padding: 5px 7px;
    background: {HEADER_BG_LIGHT};
    position: sticky;
    top: 0;
    z-index: 2;
    max-width: 120px;
    min-width: 50px;
    white-space: normal;
    word-wrap: break-word;
}}
.count-col {{ background: #ebebeb; color: #777; font-weight: 500; }}
/* Sticky group header + sub-header */
thead tr:nth-child(1) th {{ position: sticky; top: 0; z-index: 4; }}
thead tr:nth-child(2) th {{ position: sticky; top: 25px; z-index: 3; }}
.city-cell {{
    text-align: left !important;
    font-weight: 500;
    background: white;
    position: sticky;
    left: 0;
    z-index: 2;
    border-right: 2px solid #b0b8c4;
    min-width: 105px;
    max-width: 145px;
}}
.total-row {{ font-weight: 700; }}
.total-row td {{ background: #f0f3f8 !important; border-top: 2px solid #8899aa; }}
.total-row .city-cell {{ background: #e4eaf4 !important; }}
.na-cell {{ color: #ccc; font-style: italic; font-size: 10px; }}
.data-row:hover td {{ filter: brightness(0.93); }}
@media print {{
    .dashboard-header {{ position: static; }}
    .tab-bar {{ display: none; }}
    .tab-panel {{ display: block !important; }}
    .table-scroll {{ max-height: none; overflow: visible; }}
}}
</style>
</head>
<body>

<div class="dashboard-header">
    <h1>Driver Survey — Routine Analysis</h1>
    <span class="week-badge">Week {week}</span>
</div>

<div class="legend-bar">
    <span style="font-weight:600;color:#444">Color scales:</span>
    <div class="legend-item">
        <div class="swatch sw-green"></div>
        <span>% (white → green)</span>
    </div>
    <div class="legend-item">
        <div class="swatch sw-red"></div>
        <span>Dissatisfaction (white → red)</span>
    </div>
    <div class="legend-item">
        <div class="swatch sw-3color"></div>
        <span>Satisfaction 1-5 (red → yellow → green)</span>
    </div>
    <span style="margin-left:auto;color:#aaa;font-size:10px">Gray columns = sample count (no color scale)</span>
</div>

<div class="tab-bar">
    {"".join(tab_buttons)}
</div>

{"".join(tab_panels)}

<script>
function switchTab(evt, tabId) {{
    document.querySelectorAll('.tab-panel').forEach(p => p.style.display = 'none');
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById(tabId).style.display = 'block';
    evt.currentTarget.classList.add('active');
}}
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    week_arg = None
    if len(sys.argv) > 1:
        m = re.search(r"(\d+)", sys.argv[1])
        if m:
            week_arg = int(m.group(1))

    print("=" * 55)
    print("  Driver Survey Routine Dashboard (HTML)")
    print("=" * 55)

    sheets, week = collect_sheets(week_arg)

    print(f"\nRendering {len(sheets)} tables...")
    html = build_html(sheets, week)

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(OUTPUT) / 1024
    print(f"\nSaved: {OUTPUT}  ({size_kb:.0f} KB)")
    print("Opening in browser...")
    os.startfile(OUTPUT)


if __name__ == "__main__":
    main()

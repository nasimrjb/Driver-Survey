"""
Build an interactive HTML dashboard for the Driver Survey Routine Analysis.
Reads the weekly Excel output and renders it as colored heatmap tables
that match the Excel conditional-formatting style.

Usage:
    python build_routine_dashboard.py                     # auto-detect latest
    python build_routine_dashboard.py week_52              # specific week
    python build_routine_dashboard.py path/to/file.xlsx    # specific file
"""

import os, sys, json, re
import numpy as np
import pandas as pd
from pathlib import Path

# ─── Config ─────────────────────────────────────────────────────────────────
BASE_DIR = r"D:\Work\Driver Survey"
PROCESSED_DIR = os.path.join(BASE_DIR, "processed")
OUTPUT_DIR = os.path.join(BASE_DIR, "PowerBI")

SNAPP_GREEN = "#00C853"
TAPSI_ORANGE = "#FF6D00"
HEADER_BG = "#1e3a5f"
HEADER_BG_LIGHT = "#D9E1F2"
GROUP_HEADER_BG = "#4472C4"
TAB_BG = "#f0f4f8"
BODY_BG = "#f5f6fa"

# Color scale endpoints (matching Excel)
WHITE = (255, 255, 255)
GREEN = (99, 190, 123)      # 63BE7B
RED = (248, 105, 107)       # F8696B
YELLOW = (255, 235, 132)    # FFEB84
LIGHT_WHITE = (252, 252, 255)  # FCFCFF

# ─── Column sets that are NOT percentages ────────────────────────────────────
NON_PCT_COLS = {
    "n", "n_joint", "Snapp_n", "Tapsi_n", "ExSnapp_n", "JntSnapp_n",
    "n_dissatisfied", "n_contacted", "respondent_count", "joint_count",
    "avg_snapp_LOC", "avg_tapsi_LOC", "avg_snapp_ride", "avg_tapsi_ride",
    "avg_magical_window_income", "count", "tapsi_carpooling_count",
    "E_n", "F_n", "G_n",
    "AvgLOC_All Snapp", "AvgLOC_Joint Snapp",
    "AvgLOC_Joint Cmpt", "AvgLOC_Joint Cmpt SU",
    "total Res", "Joint Res", "Ex drivers",
    "Total Ride", "Total Ride Snapp", "Ex drivers Ride in Snapp",
    "Snapp Ride", "Tapsi Tide",
    "RidePerBoarded_Snapp", "RidePerBoarded_Tapsi", "AvgAllRides",
    "Who Got Message", "Free Comm Drivers",
    "GotMsg_Money", "GotMsg_Free-Commission",
    "GotMsg_Money & Free-commission",
    "Total Rides", "RidesAmong_Total Rides", "RidesAmong_Free Comm Rides",
    "# Respondents", "# Joint", "# Cmpt Signup",
    "wheel_usage", "max_demand",
}

# Sheets where values are satisfaction scores (1-5), not percentages
SATISFACTION_PREFIXES = {"#3_Sat_", "#CS_Sat_", "#Carfix_Sat_", "#Garage_Sat_", "#NavReco_"}

# Sheets where dissatisfaction → use red scale (higher = worse)
DISSAT_PREFIXES = {"#8_Dissat", "#9_Dissat"}

# ─── Tab Organization ───────────────────────────────────────────────────────
TAB_GROUPS = [
    ("Incentive", [
        "#1_Snapp_Incentive_Amt",
        "#2_Tapsi_Incentive_Amt",
        "#4_Incentive_Duration",
        "#5_6_IncType",
        "#8_Dissat",
        "#9_Dissat_Sum",
        "#9_Dissat_Sum_LastWk",
    ]),
    ("Satisfaction", [
        "#3_Sat_All Drivers",
        "#3_Sat_Part-Time",
        "#3_Sat_Full-Time",
        "#12_Cities_Overview",
    ]),
    ("Market Share", [
        "#13_RideShare",
        "#15_Persona_PartTime",
    ]),
    ("Operations", [
        "#14_Nav_",
        "#15_Persona_",
        "#17_Inactivity",
        "#19_LuckyWheel",
    ]),
    ("Commission Free", [
        "#18_CommFree_",
    ]),
    ("Support & NPS", [
        "#CS_Sat_",
        "#CS_Cat_",
        "#CS_Reason_",
        "#Reco_NPS",
    ]),
    ("Registration", [
        "#Refer_",
        "#Reg_",
        "#16_Ref_",
        "#Income_",
    ]),
    ("Decline & Demand", [
        "#Decline_",
        "#Demand_",
        "#20_Refusal_",
    ]),
]


# ═══════════════════════════════════════════════════════════════════════════
#  COLOR INTERPOLATION
# ═══════════════════════════════════════════════════════════════════════════

def lerp_color(c1, c2, t):
    """Linear interpolation between two RGB tuples, t in [0,1]."""
    t = max(0, min(1, t))
    return tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))


def color_to_css(rgb):
    return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"


def text_color_for_bg(rgb):
    """Return black or white text depending on background luminance."""
    lum = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
    return "#000" if lum > 140 else "#fff"


def get_color_scale(sheet_name, col_name):
    """Determine the color scale type (no group context)."""
    return get_color_scale_with_group(sheet_name, col_name, "")


def get_color_scale_with_group(sheet_name, col_name, group_name=""):
    """Determine the color scale type for a given sheet+column+group.

    Returns: 'green' | 'red' | 'sat_3color' | 'green_abs' | None
    """
    col_name = str(col_name)
    if col_name in NON_PCT_COLS:
        return None
    if "WoW" in col_name or "Last" in col_name:
        return None

    # Dissatisfaction sheets → red scale
    for prefix in DISSAT_PREFIXES:
        if sheet_name.startswith(prefix):
            return "red"

    # Satisfaction sheets → context-aware coloring
    for prefix in SATISFACTION_PREFIXES:
        if sheet_name.startswith(prefix):
            # % columns within sat sheets use green
            if col_name.endswith("_%") or "Part%" in col_name or "GotMsg%" in col_name:
                return "green"
            if "solved" in col_name.lower() or "contacted" in col_name.lower():
                return "green"
            # Use group header to determine: % groups → green, Avg/Sat groups → 3-color
            grp_lower = group_name.lower()
            if "%" in group_name or "participation" in grp_lower:
                return "green"
            if "satisfaction" in grp_lower or "avg" in grp_lower:
                return "sat_3color"
            # Fallback: if value range is 0-1, it's probably a percentage
            return "sat_3color"

    # RideShare: @Tapsi → red, others → green
    if sheet_name.startswith("#13_RideShare"):
        if "Tapsi" in col_name and "@" in col_name:
            return "red"

    # Cities Overview: % Dual SU → red
    if sheet_name.startswith("#12_Cities"):
        if "Dual SU" in col_name:
            return "red"

    return "green"


def compute_cell_color(value, col_min, col_max, scale_type):
    """Compute the background RGB for a cell value given min/max and scale type."""
    if pd.isna(value) or col_min == col_max:
        return None

    t = (value - col_min) / (col_max - col_min) if col_max != col_min else 0

    if scale_type == "green":
        return lerp_color(WHITE, GREEN, t)
    elif scale_type == "red":
        return lerp_color(WHITE, RED, t)
    elif scale_type == "sat_3color":
        # Map 1-5 scale: 1=red, 3=yellow, 5=green
        if value <= 3:
            t2 = (value - 1) / 2 if value >= 1 else 0
            return lerp_color(RED, YELLOW, t2)
        else:
            t2 = (value - 3) / 2
            return lerp_color(YELLOW, GREEN, t2)
    elif scale_type == "green_abs":
        return lerp_color(LIGHT_WHITE, GREEN, t)

    return None


# ═══════════════════════════════════════════════════════════════════════════
#  READ & PARSE EXCEL
# ═══════════════════════════════════════════════════════════════════════════

def find_latest_excel():
    """Find the most recent routine_analysis_week_*.xlsx in PROCESSED_DIR."""
    files = list(Path(PROCESSED_DIR).glob("routine_analysis_week_*.xlsx"))
    if not files:
        raise FileNotFoundError(f"No routine_analysis_week_*.xlsx found in {PROCESSED_DIR}")
    # Sort by week number
    def week_num(f):
        m = re.search(r"week_(\d+)", f.name)
        return int(m.group(1)) if m else 0
    files.sort(key=week_num)
    return str(files[-1])


def read_excel_smart(path, sheet_name):
    """Read a sheet, handling merged/multi-row headers intelligently."""
    # First pass: read raw to detect header structure
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)

    if raw.empty or raw.shape[0] < 2:
        return pd.DataFrame(), None, None

    # Detect if there's a group header row (row 0 has merged cells with labels)
    first_row = raw.iloc[0].tolist()
    second_row = raw.iloc[1].tolist()

    # Check if first column of row 0 is "City" → single header row
    if str(first_row[0]).strip() == "City":
        df = pd.read_excel(path, sheet_name=sheet_name, header=0)
        df = df.rename(columns={"Unnamed: 0": "City"})
        return df, None, None

    # Check if second row starts with "City" → has group header
    if str(second_row[0]).strip() == "City":
        # Build group headers from row 0
        group_headers = {}
        current_group = None
        group_start = None
        for i, val in enumerate(first_row):
            if pd.notna(val) and str(val).strip():
                if current_group is not None:
                    group_headers[current_group] = (group_start, i - 1)
                current_group = str(val).strip()
                group_start = i
            # Handle unnamed columns that are part of merged group
        if current_group is not None:
            group_headers[current_group] = (group_start, len(first_row) - 1)

        # Build column names from row 1 (sub-headers)
        # Determine which group each sub-column belongs to for prefixing
        col_to_group = {}
        for gname, (gstart, gend) in group_headers.items():
            for ci in range(gstart, gend + 1):
                col_to_group[ci] = gname

        # Build unique internal names AND clean display names
        col_names = []       # unique internal names for DataFrame
        display_names = []   # clean names for HTML rendering
        seen = {}
        for i, val in enumerate(second_row):
            display = str(val).strip() if pd.notna(val) else None
            if display is not None and display.startswith("Unnamed"):
                display = None
            # If sub-header is empty, inherit name from group header above
            if display is None:
                display = col_to_group.get(i, f"col_{i}")
            internal = display
            # Make internal name unique by appending index if duplicate
            if internal in seen:
                internal = f"{internal}__{i}"
            seen[internal] = True
            col_names.append(internal)
            display_names.append(display)

        # Data starts at row 2
        df = raw.iloc[2:].copy()
        df.columns = col_names
        df = df.rename(columns={col_names[0]: "City"})
        display_names[0] = "City"
        df = df.reset_index(drop=True)

        # Convert numeric columns
        for col in df.columns:
            if col != "City":
                try:
                    df[col] = pd.to_numeric(df[col], errors="coerce")
                except (TypeError, ValueError):
                    pass

        # Build internal→display mapping
        display_map = dict(zip(col_names, display_names))

        return df, group_headers, display_map

    # Fallback: just read normally
    df = pd.read_excel(path, sheet_name=sheet_name, header=0)
    first_col = df.columns[0]
    if "Unnamed" in str(first_col):
        df = df.rename(columns={first_col: "City"})
    return df, None, None


# ═══════════════════════════════════════════════════════════════════════════
#  HTML TABLE RENDERING
# ═══════════════════════════════════════════════════════════════════════════

def _col_to_group(df_columns, group_headers):
    """Map each column name → group name (from group_headers).

    group_headers uses raw Excel column indices (0-based).
    df_columns[0] = 'City' maps to raw index 0.
    So df_columns[i] maps to raw index i.
    """
    if not group_headers:
        return {}
    mapping = {}
    for gname, (gstart, gend) in group_headers.items():
        for i in range(gstart, gend + 1):
            if 0 <= i < len(df_columns):
                mapping[str(df_columns[i])] = gname
    return mapping


def render_table_html(df, sheet_name, group_headers=None, display_map=None):
    """Render a DataFrame as an HTML table with conditional formatting colors."""
    if df.empty:
        return "<p>No data available</p>"

    # display_map: internal col name → clean display name
    if display_map is None:
        display_map = {str(c): str(c) for c in df.columns}

    # Build column→group mapping for context-aware coloring
    col_group = _col_to_group(df.columns, group_headers)

    # Compute column min/max for color scaling
    col_stats = {}
    for col in df.columns:
        col_str = str(col)
        disp = display_map.get(col_str, col_str)
        if disp == "City" or disp.startswith("col_"):
            continue
        # Use display name for NON_PCT_COLS matching, group for context
        grp = col_group.get(col_str, "")
        scale = get_color_scale_with_group(sheet_name, disp, grp)
        if scale and pd.api.types.is_numeric_dtype(df[col]):
            valid = df[col].dropna()
            if len(valid) > 0:
                if scale == "sat_3color":
                    col_stats[col_str] = (1, 5, scale)
                else:
                    col_stats[col_str] = (valid.min(), valid.max(), scale)

    lines = []
    lines.append('<div class="table-wrapper">')
    lines.append(f'<h3 class="sheet-title">{sheet_name}</h3>')
    lines.append('<div class="table-scroll">')
    lines.append('<table class="heatmap-table">')

    # ── Group header row (if present) ──
    if group_headers:
        lines.append('<tr class="group-header-row">')
        # Track which columns are covered by groups
        covered = set()
        # First: empty cell for City column
        lines.append(f'<th class="group-header empty-group"></th>')

        # Build a mapping: col_index → group_name
        col_group_map = {}
        for gname, (gstart, gend) in group_headers.items():
            for i in range(gstart, gend + 1):
                col_group_map[i] = gname

        # Emit group header cells
        i = 1  # skip City column (index 0)
        while i < len(df.columns) + 1:  # +1 because City is col 0 in original
            actual_i = i  # column index in the original data
            if actual_i in col_group_map:
                gname = col_group_map[actual_i]
                gstart, gend = group_headers[gname]
                span = gend - gstart + 1
                # Clamp span to actual columns
                span = min(span, len(df.columns) + 1 - i)
                lines.append(f'<th class="group-header" colspan="{span}">{gname}</th>')
                i += span
            else:
                lines.append(f'<th class="group-header empty-group"></th>')
                i += 1
        lines.append('</tr>')

    # ── Column header row ──
    lines.append('<tr class="col-header-row">')
    for col in df.columns:
        col_s = str(col)
        disp = display_map.get(col_s, col_s)
        if disp in NON_PCT_COLS or col_s in NON_PCT_COLS:
            css_class = "col-header count-col"
        else:
            css_class = "col-header"
        lines.append(f'<th class="{css_class}">{disp}</th>')
    lines.append('</tr>')

    # ── Data rows ──
    for idx in range(len(df)):
        row = df.iloc[idx]
        city = str(row.iloc[0]) if len(row) > 0 else ""
        is_total = city == "Total"
        row_class = "total-row" if is_total else "data-row"
        lines.append(f'<tr class="{row_class}">')

        for col_idx, col in enumerate(df.columns):
            col_s = str(col)
            disp_name = display_map.get(col_s, col_s)
            val = row.iloc[col_idx]

            if disp_name == "City":
                lines.append(f'<td class="city-cell">{val}</td>')
                continue

            # Format value
            try:
                is_na = bool(pd.isna(val)) if np.isscalar(val) else False
            except (ValueError, TypeError):
                is_na = False
            if is_na:
                lines.append('<td class="na-cell">-</td>')
                continue

            # Determine display format
            display_val = ""
            style = ""

            if col_s in col_stats:
                cmin, cmax, scale = col_stats[col_s]
                bg = compute_cell_color(val, cmin, cmax, scale)
                if bg:
                    text_col = text_color_for_bg(bg)
                    style = f'background-color:{color_to_css(bg)};color:{text_col}'

            # Format the number based on scale type
            is_count = disp_name in NON_PCT_COLS
            cell_scale = col_stats.get(col_s, (None, None, None))[2] if col_s in col_stats else None

            if is_count:
                if isinstance(val, float) and val == int(val):
                    display_val = f"{int(val)}"
                else:
                    display_val = f"{val:,.1f}" if isinstance(val, float) else str(val)
            elif cell_scale == "sat_3color":
                display_val = f"{val:.2f}"
            elif cell_scale in ("green", "red"):
                if isinstance(val, (int, float)):
                    if abs(val) <= 1.5:
                        display_val = f"{val*100:.1f}%"
                    else:
                        display_val = f"{val:.1f}%"
                else:
                    display_val = str(val)
            elif "WoW" in disp_name:
                if isinstance(val, (int, float)):
                    sign = "+" if val > 0 else ""
                    if abs(val) <= 1.5:
                        display_val = f"{sign}{val*100:.1f}%"
                    else:
                        display_val = f"{sign}{val:.2f}"
                else:
                    display_val = str(val)
            else:
                is_pct_sheet = not any(sheet_name.startswith(p) for p in {"#Demand_"})
                if is_pct_sheet and isinstance(val, (int, float)) and disp_name not in NON_PCT_COLS:
                    if abs(val) <= 1.5:
                        display_val = f"{val*100:.1f}%"
                    else:
                        display_val = f"{val:.1f}%"
                elif isinstance(val, float):
                    display_val = f"{val:.2f}"
                else:
                    display_val = str(val)

            if style:
                lines.append(f'<td style="{style}">{display_val}</td>')
            else:
                lines.append(f'<td>{display_val}</td>')

        lines.append('</tr>')

    lines.append('</table>')
    lines.append('</div>')  # table-scroll
    lines.append('</div>')  # table-wrapper

    return '\n'.join(lines)


# ═══════════════════════════════════════════════════════════════════════════
#  ASSEMBLE FULL DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════

def build_dashboard(excel_path):
    """Build the full HTML dashboard from an Excel file."""
    print(f"Reading: {excel_path}")

    # Get week number from filename
    week_match = re.search(r"week_(\d+)", excel_path)
    week_num = week_match.group(1) if week_match else "?"

    # Read all sheets
    xl = pd.ExcelFile(excel_path)
    all_sheets = xl.sheet_names
    print(f"Found {len(all_sheets)} sheets: {all_sheets[:5]}...")

    # Assign sheets to tabs
    tab_contents = {}
    used_sheets = set()

    for tab_name, prefixes in TAB_GROUPS:
        tab_sheets = []
        for prefix in prefixes:
            for sname in all_sheets:
                if sname.startswith(prefix) and sname not in used_sheets:
                    # Exclude PartTime from Operations (it's in Market Share)
                    if tab_name == "Operations" and "PartTime" in sname:
                        continue
                    tab_sheets.append(sname)
                    used_sheets.add(sname)
        if tab_sheets:
            tab_contents[tab_name] = tab_sheets

    # Any remaining sheets go to "Other"
    remaining = [s for s in all_sheets if s not in used_sheets]
    if remaining:
        tab_contents["Other"] = remaining

    # Build HTML for each tab
    tab_html = {}
    for tab_name, sheet_names in tab_contents.items():
        print(f"\n  Building tab: {tab_name} ({len(sheet_names)} sheets)")
        html_parts = []
        for sname in sheet_names:
            print(f"    Sheet: {sname}")
            df, grp_headers, col_names = read_excel_smart(excel_path, sname)
            if df.empty:
                continue
            table_html = render_table_html(df, sname, grp_headers, col_names)
            html_parts.append(table_html)
        tab_html[tab_name] = '\n'.join(html_parts)

    # Count total sheets and rows
    total_rows = sum(
        pd.read_excel(excel_path, sheet_name=s).shape[0]
        for s in all_sheets
    )

    # Assemble final HTML
    html = build_full_html(tab_html, week_num, len(all_sheets), total_rows)

    output_path = os.path.join(OUTPUT_DIR, "RoutineAnalysis_Dashboard.html")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\nDashboard saved to: {output_path}")
    print(f"  Size: {os.path.getsize(output_path) / 1024:.0f} KB")
    return output_path


def build_full_html(tab_html, week_num, n_sheets, n_rows):
    """Assemble the full HTML document."""

    # Build tab buttons
    tab_buttons = []
    tab_panels = []
    for i, (tab_name, content) in enumerate(tab_html.items()):
        active = "active" if i == 0 else ""
        tab_buttons.append(
            f'<button class="tab-btn {active}" onclick="switchTab(event, \'tab-{i}\')">'
            f'{tab_name}</button>'
        )
        display = "block" if i == 0 else "none"
        tab_panels.append(
            f'<div id="tab-{i}" class="tab-panel" style="display:{display}">'
            f'{content}</div>'
        )

    tab_buttons_html = '\n'.join(tab_buttons)
    tab_panels_html = '\n'.join(tab_panels)

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Driver Survey Routine Analysis - Week {week_num}</title>
<style>
/* ─── Reset & Base ─────────────────────────────────────────────── */
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background: {BODY_BG};
    color: #333;
    line-height: 1.4;
}}

/* ─── Header ───────────────────────────────────────────────────── */
.dashboard-header {{
    background: linear-gradient(135deg, {HEADER_BG} 0%, #2c5282 100%);
    color: white;
    padding: 18px 30px;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    position: sticky;
    top: 0;
    z-index: 100;
}}
.dashboard-header h1 {{
    font-size: 22px;
    font-weight: 600;
    margin-bottom: 4px;
}}
.dashboard-header .subtitle {{
    font-size: 13px;
    opacity: 0.85;
}}

/* ─── Legend ────────────────────────────────────────────────────── */
.legend-bar {{
    background: white;
    border-bottom: 1px solid #e2e8f0;
    padding: 8px 30px;
    display: flex;
    gap: 24px;
    align-items: center;
    flex-wrap: wrap;
    font-size: 12px;
    color: #555;
}}
.legend-item {{
    display: flex;
    align-items: center;
    gap: 6px;
}}
.legend-swatch {{
    width: 60px;
    height: 14px;
    border-radius: 3px;
    border: 1px solid #ccc;
}}
.swatch-green {{
    background: linear-gradient(90deg, #fff, {color_to_css(GREEN)});
}}
.swatch-red {{
    background: linear-gradient(90deg, #fff, {color_to_css(RED)});
}}
.swatch-3color {{
    background: linear-gradient(90deg, {color_to_css(RED)}, {color_to_css(YELLOW)}, {color_to_css(GREEN)});
}}

/* ─── Tabs ─────────────────────────────────────────────────────── */
.tab-bar {{
    background: {TAB_BG};
    border-bottom: 2px solid #e2e8f0;
    padding: 0 20px;
    display: flex;
    gap: 2px;
    overflow-x: auto;
    position: sticky;
    top: 70px;
    z-index: 99;
}}
.tab-btn {{
    padding: 10px 18px;
    border: none;
    background: transparent;
    font-size: 13px;
    font-weight: 500;
    color: #666;
    cursor: pointer;
    border-bottom: 3px solid transparent;
    white-space: nowrap;
    transition: all 0.2s;
}}
.tab-btn:hover {{
    color: #333;
    background: rgba(0,0,0,0.04);
}}
.tab-btn.active {{
    color: {HEADER_BG};
    border-bottom-color: {HEADER_BG};
    font-weight: 600;
}}

/* ─── Content ──────────────────────────────────────────────────── */
.tab-panel {{
    padding: 20px 24px;
    max-width: 100%;
}}

/* ─── Table Wrappers ───────────────────────────────────────────── */
.table-wrapper {{
    background: white;
    border-radius: 8px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    margin-bottom: 24px;
    overflow: hidden;
}}
.sheet-title {{
    font-size: 15px;
    font-weight: 600;
    padding: 12px 16px 8px;
    color: {HEADER_BG};
    border-bottom: 1px solid #eee;
    background: #fafbfc;
}}
.table-scroll {{
    overflow-x: auto;
    overflow-y: auto;
    max-height: 700px;
}}

/* ─── Heatmap Table ────────────────────────────────────────────── */
.heatmap-table {{
    border-collapse: collapse;
    font-size: 12px;
    width: 100%;
    min-width: max-content;
}}
.heatmap-table th,
.heatmap-table td {{
    padding: 4px 8px;
    border: 1px solid #d4d8dd;
    text-align: center;
    white-space: nowrap;
}}

/* Group header row (blue band) */
.group-header-row {{
    background: {GROUP_HEADER_BG};
}}
.group-header {{
    color: white;
    font-weight: 700;
    font-size: 12px;
    padding: 6px 10px;
    background: {GROUP_HEADER_BG};
    border-color: #3a6bb5;
}}
.empty-group {{
    background: #e8ecf2;
}}

/* Column headers */
.col-header-row {{
    background: {HEADER_BG_LIGHT};
}}
.col-header {{
    font-weight: 600;
    font-size: 11px;
    color: #333;
    padding: 6px 8px;
    background: {HEADER_BG_LIGHT};
    position: sticky;
    top: 0;
    z-index: 2;
    max-width: 120px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: normal;
    word-wrap: break-word;
    min-width: 55px;
}}
.count-col {{
    background: #f0f0f0;
    color: #888;
    font-weight: 500;
}}

/* City column (frozen) */
.city-cell {{
    text-align: left;
    font-weight: 500;
    background: white;
    position: sticky;
    left: 0;
    z-index: 1;
    border-right: 2px solid #b0b8c4;
    min-width: 100px;
    max-width: 140px;
}}

/* Total row */
.total-row {{
    font-weight: 700;
    background: #f7f8fa;
}}
.total-row .city-cell {{
    background: #f0f2f5;
    font-weight: 700;
}}

/* N/A cells */
.na-cell {{
    color: #ccc;
    font-style: italic;
}}

/* Data row hover */
.data-row:hover td {{
    filter: brightness(0.95);
}}

/* ─── Responsive ───────────────────────────────────────────────── */
@media (max-width: 768px) {{
    .tab-btn {{ padding: 8px 12px; font-size: 12px; }}
    .tab-panel {{ padding: 12px; }}
    .heatmap-table {{ font-size: 11px; }}
}}

/* ─── Search / Filter ──────────────────────────────────────────── */
.filter-bar {{
    padding: 8px 16px;
    background: #fafbfc;
    border-bottom: 1px solid #eee;
    display: flex;
    align-items: center;
    gap: 12px;
}}
.filter-input {{
    padding: 5px 10px;
    border: 1px solid #ccc;
    border-radius: 4px;
    font-size: 12px;
    width: 200px;
}}
.filter-label {{
    font-size: 12px;
    color: #666;
}}

/* ─── Print ────────────────────────────────────────────────────── */
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
    <h1>Driver Survey Routine Analysis</h1>
    <div class="subtitle">
        Week {week_num} | {n_sheets} sheets | {n_rows:,} data rows |
        Generated from SQL Views
    </div>
</div>

<div class="legend-bar">
    <span style="font-weight:600;">Color Scales:</span>
    <div class="legend-item">
        <div class="legend-swatch swatch-green"></div>
        <span>Percentage (white &rarr; green = higher)</span>
    </div>
    <div class="legend-item">
        <div class="legend-swatch swatch-red"></div>
        <span>Dissatisfaction (white &rarr; red = higher)</span>
    </div>
    <div class="legend-item">
        <div class="legend-swatch swatch-3color"></div>
        <span>Satisfaction (1-5: red &rarr; yellow &rarr; green)</span>
    </div>
    <div class="legend-item" style="margin-left:auto;">
        <span style="color:#999">Gray n columns = sample size</span>
    </div>
</div>

<div class="tab-bar">
    {tab_buttons_html}
</div>

{tab_panels_html}

<script>
function switchTab(evt, tabId) {{
    // Hide all panels
    document.querySelectorAll('.tab-panel').forEach(p => p.style.display = 'none');
    // Deactivate all buttons
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    // Show target & activate button
    document.getElementById(tabId).style.display = 'block';
    evt.currentTarget.classList.add('active');
}}

// City filter
function filterCities(input) {{
    const query = input.value.toLowerCase();
    document.querySelectorAll('.heatmap-table').forEach(table => {{
        table.querySelectorAll('.data-row, .total-row').forEach(row => {{
            const city = row.querySelector('.city-cell');
            if (city) {{
                const match = city.textContent.toLowerCase().includes(query) ||
                              row.classList.contains('total-row');
                row.style.display = match ? '' : 'none';
            }}
        }});
    }});
}}
</script>

</body>
</html>'''


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if os.path.isfile(arg):
            excel_path = arg
        else:
            # Try to find file matching the argument
            m = re.search(r"(\d+)", arg)
            if m:
                week = m.group(1)
                excel_path = os.path.join(PROCESSED_DIR, f"routine_analysis_week_{week}.xlsx")
            else:
                excel_path = find_latest_excel()
    else:
        excel_path = find_latest_excel()

    if not os.path.exists(excel_path):
        print(f"ERROR: File not found: {excel_path}")
        sys.exit(1)

    build_dashboard(excel_path)

"""
Driver Survey Interactive Dashboard Builder
============================================
Connects to SQL Server views and generates a standalone HTML dashboard.
Open the output file in any browser — no Power BI needed.

Usage:
    python build_dashboard.py
"""

import pyodbc
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import jdatetime
from datetime import date, timedelta
import json, os

# ============================================================
# CONFIG
# ============================================================
SERVER = "192.168.18.37"
DATABASE = "Cab_Studies"
SQL_USER = "nasim.rajabi"
SQL_PASS = "ISLRv2_corrected_June_2023"
OUTPUT = r"D:\Work\Driver Survey\PowerBI\DriverSurvey_Dashboard.html"

SNAPP_COLOR = "#00C853"   # Snapp green
TAPSI_COLOR = "#FF6D00"   # Tapsi orange
BG_COLOR = "#f8f9fa"
CARD_BG = "#ffffff"

CITY_ORDER = [
    "Tehran(city)", "Karaj", "Isfahan", "Shiraz", "Mashhad",
    "Qom", "Tabriz", "Ahwaz", "Sari", "Rasht",
    "Urumieh", "Yazd", "Kerman", "Gorgan", "Ghazvin",
    "Arak", "Kermanshah", "Hamedan", "Ardebil", "Zanjan", "Kish",
]

PERSIAN_MONTHS = [
    'Far', 'Ord', 'Kho', 'Tir', 'Mor', 'Sha',
    'Meh', 'Aba', 'Aza', 'Dey', 'Bah', 'Esf',
]

def _week_start_monday(year, weeknum):
    """Return the Monday that starts ISO week `weeknum` of `year`."""
    # ISO week 1 contains the year's first Thursday.
    # Jan 4 is always in ISO week 1.
    jan4 = date(year, 1, 4)
    # Monday of that week
    iso_w1_monday = jan4 - timedelta(days=jan4.weekday())
    return iso_w1_monday + timedelta(weeks=weeknum - 1)


def yearweek_to_persian(yw):
    """Convert 'YY-WW' yearweek string to a Persian week label like '25-W4: 13-19 Dey'."""
    try:
        yy, ww = yw.split('-')
        yr = 2000 + int(yy)
        wk = int(ww)
        mon = _week_start_monday(yr, wk)
        sun = mon + timedelta(days=6)
        j_mon = jdatetime.date.fromgregorian(date=mon)
        j_sun = jdatetime.date.fromgregorian(date=sun)
        m1 = PERSIAN_MONTHS[j_mon.month - 1]
        m2 = PERSIAN_MONTHS[j_sun.month - 1]
        if m1 == m2:
            return f"{yy}-W{wk}: {j_mon.day}-{j_sun.day} {m1}"
        return f"{yy}-W{wk}: {j_mon.day} {m1}-{j_sun.day} {m2}"
    except Exception:
        return yw


def add_persian_week_labels(df, yearweek_col="yearweek"):
    """Add a 'week_label' column with Persian week descriptions, ordered chronologically."""
    df = df.copy()
    df["week_label"] = df[yearweek_col].astype(str).map(yearweek_to_persian)
    # Build chronological order from the sorted yearweek values
    ordered_labels = df.sort_values(yearweek_col)["week_label"].unique().tolist()
    df["week_label"] = pd.Categorical(df["week_label"], categories=ordered_labels, ordered=True)
    return df


def query(sql):
    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={SERVER};DATABASE={DATABASE};"
        f"UID={SQL_USER};PWD={SQL_PASS};TrustServerCertificate=yes;"
    )
    df = pd.read_sql(sql, conn)
    conn.close()
    return df


def load_all_views():
    """Load all views into a dict of DataFrames."""
    views = {}
    names = [
        "vw_KPISummary", "vw_WeeklySatisfaction", "vw_WeeklyNPS",
        "vw_SatisfactionByCity", "vw_SatisfactionByCityWeek",
        "vw_SatisfactionByDemographics", "vw_HoneymoonEffect",
        "vw_IncentiveByWeek", "vw_IncentiveByCity",
        "vw_IncentiveAmountByCity", "vw_WideIncentiveTypes",
        "vw_WideUnsatisfactionReasons",
        "vw_RideShareByCityWeek", "vw_NavigationUsage",
        "vw_NavigationByWeek", "vw_Demographics",
        "vw_PersonaByCity",
        "vw_LongSurveyAnswers", "vw_LongRareSurveyAnswers",
        "vw_LongSurveyByCity",
    ]
    for name in names:
        print(f"  Loading {name}...")
        views[name] = query(f"SELECT * FROM [Cab].[{name}]")
    return views


# ============================================================
# PAGE 1: EXECUTIVE OVERVIEW
# ============================================================
def page1_executive(views):
    kpi = views["vw_KPISummary"].iloc[0]
    ws = add_persian_week_labels(views["vw_WeeklySatisfaction"].sort_values("yearweek"))
    nps = add_persian_week_labels(views["vw_WeeklyNPS"].sort_values("yearweek"))

    figs = []

    # --- KPI Cards as a table-like figure ---
    kpi_data = [
        ("Total Responses", f"{int(kpi['total_responses']):,}"),
        ("Survey Weeks", f"{int(kpi['survey_weeks'])}"),
        ("Cities", f"{int(kpi['cities'])}"),
        ("Joint Driver %", f"{kpi['joint_driver_pct']:.1f}%"),
        ("Snapp Fare Sat", f"{kpi['snapp_fare_sat']:.2f}"),
        ("Tapsi Fare Sat", f"{kpi['tapsi_fare_sat']:.2f}"),
        ("Snapp NPS", f"{kpi['snapp_nps']:.1f}"),
        ("Tapsi NPS", f"{kpi['tapsi_nps']:.1f}" if pd.notna(kpi.get('tapsi_nps')) else "N/A"),
    ]

    fig_kpi = go.Figure(data=[go.Table(
        header=dict(
            values=[d[0] for d in kpi_data],
            fill_color="#2c3e50",
            font=dict(color="white", size=13),
            align="center", height=35
        ),
        cells=dict(
            values=[[d[1]] for d in kpi_data],
            fill_color=CARD_BG,
            font=dict(size=18, color="#2c3e50"),
            align="center", height=40
        )
    )])
    fig_kpi.update_layout(
        title="Key Performance Indicators",
        height=160, margin=dict(t=40, b=10, l=20, r=20)
    )
    figs.append(fig_kpi)

    # --- Weekly Satisfaction Trends ---
    fig_sat = make_subplots(specs=[[{"secondary_y": True}]])
    for col, name, color, dash in [
        ("snapp_fare_sat_avg", "Snapp Fare", SNAPP_COLOR, "solid"),
        ("snapp_income_sat_avg", "Snapp Income", SNAPP_COLOR, "dash"),
        ("snapp_req_sat_avg", "Snapp Request", SNAPP_COLOR, "dot"),
        ("tapsi_fare_sat_avg", "Tapsi Fare", TAPSI_COLOR, "solid"),
        ("tapsi_income_sat_avg", "Tapsi Income", TAPSI_COLOR, "dash"),
    ]:
        if col in ws.columns:
            fig_sat.add_trace(go.Scatter(
                x=ws["week_label"], y=ws[col], name=name,
                line=dict(color=color, dash=dash, width=2),
                hovertemplate=f"{name}: %{{y:.2f}}<extra></extra>"
            ), secondary_y=False)

    # Joint driver % on secondary axis
    fig_sat.add_trace(go.Scatter(
        x=ws["week_label"], y=ws["joint_driver_pct"], name="Joint %",
        line=dict(color="#7f8c8d", dash="dot", width=1.5),
        opacity=0.6
    ), secondary_y=True)

    ws_labels = ws["week_label"].tolist()
    fig_sat.update_layout(
        title="Weekly Satisfaction Trends (1-5 scale)",
        height=400, hovermode="x unified",
        legend=dict(orientation="h", y=-0.15),
        margin=dict(t=40, b=60),
        xaxis=dict(categoryorder="array", categoryarray=ws_labels),
    )
    fig_sat.update_yaxes(title_text="Satisfaction (1-5)", range=[1, 5], secondary_y=False)
    fig_sat.update_yaxes(title_text="Joint %", secondary_y=True)
    figs.append(fig_sat)

    # --- Weekly NPS (filter to weeks with data, connect gaps) ---
    nps_valid = nps.dropna(subset=["snapp_nps"])
    fig_nps = go.Figure()
    fig_nps.add_trace(go.Scatter(
        x=nps_valid["week_label"], y=nps_valid["snapp_nps"], name="Snapp NPS",
        line=dict(color=SNAPP_COLOR, width=3), mode="lines+markers",
        marker=dict(size=5),
        fill="tozeroy", fillcolor="rgba(0,200,83,0.1)",
        connectgaps=True
    ))
    nps_tapsi = nps.dropna(subset=["tapsi_nps"])
    if len(nps_tapsi) >= 3:
        fig_nps.add_trace(go.Scatter(
            x=nps_tapsi["week_label"], y=nps_tapsi["tapsi_nps"], name="Tapsi NPS",
            line=dict(color=TAPSI_COLOR, width=3), mode="lines+markers",
            marker=dict(size=5),
            fill="tozeroy", fillcolor="rgba(255,109,0,0.1)",
            connectgaps=True
        ))
    fig_nps.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
    nps_labels = nps.sort_values("yearweek")["week_label"].tolist()
    fig_nps.update_layout(
        title=f"Weekly NPS Score (Snapp: {len(nps_valid)} weeks, Tapsi: {len(nps_tapsi)} weeks — rare question)",
        height=350, hovermode="x unified",
        yaxis_title="NPS",
        legend=dict(orientation="h", y=-0.15),
        xaxis=dict(categoryorder="array", categoryarray=nps_labels),
    )
    figs.append(fig_nps)

    # --- NPS Decomposition (Stacked Area, filter to valid weeks) ---
    nps_decomp = nps.dropna(subset=["snapp_detractor_pct"])
    fig_nps_decomp = go.Figure()
    fig_nps_decomp.add_trace(go.Scatter(
        x=nps_decomp["week_label"], y=nps_decomp["snapp_detractor_pct"], name="Detractors",
        stackgroup="one", line=dict(color="#e74c3c"), fillcolor="rgba(231,76,60,0.4)"
    ))
    fig_nps_decomp.add_trace(go.Scatter(
        x=nps_decomp["week_label"], y=nps_decomp["snapp_passive_pct"], name="Passives",
        stackgroup="one", line=dict(color="#f39c12"), fillcolor="rgba(243,156,18,0.4)"
    ))
    fig_nps_decomp.add_trace(go.Scatter(
        x=nps_decomp["week_label"], y=nps_decomp["snapp_promoter_pct"], name="Promoters",
        stackgroup="one", line=dict(color="#27ae60"), fillcolor="rgba(39,174,96,0.4)"
    ))
    fig_nps_decomp.update_layout(
        title="Snapp NPS Decomposition (Detractors / Passives / Promoters %)",
        height=350, yaxis_title="%", hovermode="x unified",
        legend=dict(orientation="h", y=-0.15),
        xaxis=dict(categoryorder="array", categoryarray=nps_labels),
    )
    figs.append(fig_nps_decomp)

    # --- Response Count ---
    fig_resp = go.Figure(go.Bar(
        x=ws["week_label"], y=ws["response_count"],
        marker_color="#3498db", opacity=0.7, name="Responses"
    ))
    fig_resp.update_layout(
        title="Weekly Response Count",
        height=250, yaxis_title="Responses",
        xaxis=dict(categoryorder="array", categoryarray=ws_labels),
        margin=dict(t=40, b=30)
    )
    figs.append(fig_resp)

    return figs


# ============================================================
# PAGE 2: SATISFACTION DEEP-DIVE
# ============================================================
def page2_satisfaction(views):
    figs = []
    sc_all = views["vw_SatisfactionByCity"]
    sc = sc_all[sc_all["city"].isin(CITY_ORDER)].copy()
    sc["city"] = pd.Categorical(sc["city"], categories=list(reversed(CITY_ORDER)), ordered=True)
    sc = sc.sort_values("city")
    sd = views["vw_SatisfactionByDemographics"]
    hm = views["vw_HoneymoonEffect"]
    scw = add_persian_week_labels(views["vw_SatisfactionByCityWeek"])

    # --- Top Cities Satisfaction (dot plot) ---
    fig_city = go.Figure()
    dot_metrics = [
        ("snapp_fare_sat", "Snapp Fare", SNAPP_COLOR, "circle"),
        ("snapp_income_sat", "Snapp Income", "#2196F3", "diamond"),
        ("snapp_req_sat", "Snapp Request", "#9C27B0", "square"),
        ("tapsi_fare_sat", "Tapsi Fare", TAPSI_COLOR, "triangle-up"),
    ]
    for col, name, color, symbol in dot_metrics:
        if col in sc.columns:
            fig_city.add_trace(go.Scatter(
                y=sc["city"], x=sc[col], name=name,
                mode="markers",
                marker=dict(color=color, size=12, symbol=symbol,
                            line=dict(width=1, color="white")),
                hovertemplate=f"{name}: %{{x:.2f}}<extra>%{{y}}</extra>"
            ))
    fig_city.update_layout(
        title="Satisfaction by City",
        height=700, xaxis_title="Satisfaction (1-5)", xaxis_range=[1, 5],
        legend=dict(orientation="h", y=-0.08),
        margin=dict(l=120),
        yaxis=dict(dtick=1),
        xaxis=dict(gridcolor="#e0e0e0", gridwidth=1),
        plot_bgcolor="white",
    )
    figs.append(fig_city)

    # --- Satisfaction by Demographics (dumbbell / dot plot) ---
    demo_metrics = [
        ("snapp_fare_sat", "Snapp Fare", SNAPP_COLOR, "circle"),
        ("snapp_income_sat", "Snapp Income", "#2196F3", "diamond"),
        ("tapsi_fare_sat", "Tapsi Fare", TAPSI_COLOR, "triangle-up"),
    ]
    for dim_name in ["cooperation_type", "driver_type", "age_group", "gender"]:
        sub = sd[sd["dimension"] == dim_name].copy()
        if len(sub) == 0:
            continue
        # Draw connecting lines between min/max per category (dumbbell)
        fig_dim = go.Figure()
        metric_cols = [c for c, _, _, _ in demo_metrics if c in sub.columns]
        for _, row in sub.iterrows():
            vals = [row[c] for c in metric_cols if pd.notna(row[c])]
            if len(vals) >= 2:
                fig_dim.add_trace(go.Scatter(
                    x=[min(vals), max(vals)], y=[row["category"]] * 2,
                    mode="lines", line=dict(color="#ddd", width=6),
                    showlegend=False, hoverinfo="skip"
                ))
        for col, name, color, symbol in demo_metrics:
            if col in sub.columns:
                fig_dim.add_trace(go.Scatter(
                    x=sub[col], y=sub["category"], name=name,
                    mode="markers+text",
                    marker=dict(color=color, size=14, symbol=symbol,
                                line=dict(width=1, color="white")),
                    text=[f"{v:.2f}" for v in sub[col]],
                    textposition="top center", textfont=dict(size=10),
                    hovertemplate=f"{name}: %{{x:.2f}} (n=%{{customdata:,}})<extra>%{{y}}</extra>",
                    customdata=sub["n"],
                ))
        fig_dim.update_layout(
            title=f"Satisfaction by {dim_name.replace('_', ' ').title()}",
            height=max(280, len(sub) * 70 + 120),
            xaxis_title="Satisfaction (1-5)", xaxis_range=[1, 4],
            xaxis=dict(gridcolor="#e0e0e0", gridwidth=1),
            legend=dict(orientation="h", y=-0.2),
            margin=dict(l=120),
            plot_bgcolor="white",
        )
        figs.append(fig_dim)

    # --- Honeymoon Effect ---
    tenure_order = [
        "less_than_1_month", "1_to_3_months", "less_than_3_months",
        "3_to_6_months", "6_months_to_1_year", "6_to_12_months",
        "1_to_2_years", "1_to_3_years", "2_to_3_years",
        "3_to_4_years", "3_to_5_years", "more_than_4_years",
        "5_to_7_years", "more_than_7_years",
    ]
    tenure_labels = {
        "less_than_1_month": "< 1 Month",
        "1_to_3_months": "1-3 Months",
        "less_than_3_months": "< 3 Months",
        "3_to_6_months": "3-6 Months",
        "6_months_to_1_year": "6-12 Months",
        "6_to_12_months": "6-12 Months",
        "1_to_2_years": "1-2 Years",
        "1_to_3_years": "1-3 Years",
        "2_to_3_years": "2-3 Years",
        "3_to_4_years": "3-4 Years",
        "3_to_5_years": "3-5 Years",
        "more_than_4_years": "4+ Years",
        "5_to_7_years": "5-7 Years",
        "more_than_7_years": "7+ Years",
    }
    hm_sorted = hm.copy()
    hm_sorted["sort_key"] = hm_sorted["tenure"].apply(
        lambda x: tenure_order.index(x.strip()) if x.strip() in tenure_order else 99
    )
    hm_sorted = hm_sorted.sort_values("sort_key")
    hm_sorted["tenure_label"] = hm_sorted["tenure"].str.strip().map(tenure_labels).fillna(hm_sorted["tenure"])

    # Add sample size as a subtle bar in the background
    fig_hon = make_subplots(specs=[[{"secondary_y": True}]])
    fig_hon.add_trace(go.Bar(
        x=hm_sorted["tenure_label"], y=hm_sorted["n"], name="Sample Size",
        marker_color="#e8e8e8", showlegend=True,
        hovertemplate="n=%{y:,}<extra></extra>",
    ), secondary_y=True)
    for col, name, color, dash in [
        ("snapp_fare_sat", "Fare Sat", SNAPP_COLOR, "solid"),
        ("snapp_income_sat", "Income Sat", "#2196F3", "dash"),
        ("snapp_overall_sat", "Overall Sat", "#9C27B0", "dot"),
    ]:
        if col in hm_sorted.columns:
            series = hm_sorted[col]
            if series.notna().sum() < len(series) * 0.5:
                continue
            fig_hon.add_trace(go.Scatter(
                x=hm_sorted["tenure_label"], y=hm_sorted[col], name=name,
                mode="lines+markers", line=dict(color=color, width=2.5, dash=dash),
                marker=dict(size=8, line=dict(width=1, color="white")),
                connectgaps=True,
                hovertemplate=f"{name}: %{{y:.2f}}<extra></extra>",
            ), secondary_y=False)
    fig_hon.update_layout(
        title="Honeymoon Effect: Satisfaction by Snapp Tenure",
        height=420, legend=dict(orientation="h", y=-0.22),
        xaxis_tickangle=-30, plot_bgcolor="white",
        xaxis=dict(title="Tenure"),
    )
    fig_hon.update_yaxes(title_text="Satisfaction (1-5)", range=[1, 5],
                         gridcolor="#e0e0e0", secondary_y=False)
    fig_hon.update_yaxes(title_text="Sample Size", showgrid=False, secondary_y=True)
    figs.append(fig_hon)

    # --- City x Week Heatmap ---
    label_order = scw.sort_values("yearweek")["week_label"].unique()
    hm_data = scw[scw["city"].isin(CITY_ORDER)].pivot_table(
        index="city", columns="week_label", values="snapp_fare_sat"
    )
    hm_data = hm_data[[c for c in label_order if c in hm_data.columns]]
    hm_data = hm_data.reindex([c for c in CITY_ORDER if c in hm_data.index])
    fig_heatmap = go.Figure(go.Heatmap(
        z=hm_data.values,
        x=[str(w) for w in hm_data.columns],
        y=hm_data.index,
        colorscale="RdYlGn",
        zmin=1.5, zmax=3.0,
        text=np.round(hm_data.values, 2),
        texttemplate="%{text}",
        textfont=dict(size=9),
        hovertemplate="City: %{y}<br>Week: %{x}<br>Fare Sat: %{z:.2f}<extra></extra>"
    ))
    fig_heatmap.update_layout(
        title="Snapp Fare Satisfaction Heatmap: City x Week",
        height=550, margin=dict(l=120),
        xaxis_title="Week"
    )
    figs.append(fig_heatmap)

    return figs


# ============================================================
# PAGE 3: INCENTIVE ANALYSIS
# ============================================================
def page3_incentive(views):
    figs = []
    iw = add_persian_week_labels(views["vw_IncentiveByWeek"].sort_values("yearweek"))
    it = views["vw_WideIncentiveTypes"]
    ur = views["vw_WideUnsatisfactionReasons"]
    ic_all = views["vw_IncentiveByCity"]
    ic = ic_all[ic_all["city"].isin(CITY_ORDER)].copy()
    ic["city"] = pd.Categorical(ic["city"], categories=list(reversed(CITY_ORDER)), ordered=True)
    ic = ic.sort_values("city")

    # --- Incentive Amount + Satisfaction over time ---
    fig_inc = make_subplots(specs=[[{"secondary_y": True}]])
    fig_inc.add_trace(go.Bar(
        x=iw["week_label"], y=iw["snapp_incentive_avg_mrial"],
        name="Avg Incentive (M Rials)",
        marker_color="rgba(0,200,83,0.35)", marker_line=dict(color=SNAPP_COLOR, width=1),
        hovertemplate="Incentive: %{y:.2f} M Rials<extra></extra>",
    ), secondary_y=False)
    fig_inc.add_trace(go.Scatter(
        x=iw["week_label"], y=iw["snapp_inc_sat_avg"],
        name="Incentive Satisfaction",
        line=dict(color="#e74c3c", width=3), mode="lines+markers",
        marker=dict(size=7, line=dict(width=1, color="white")),
        hovertemplate="Satisfaction: %{y:.2f}<extra></extra>",
    ), secondary_y=True)
    if "snapp_commfree_pct" in iw.columns:
        fig_inc.add_trace(go.Scatter(
            x=iw["week_label"], y=iw["snapp_commfree_pct"],
            name="Comm-Free %",
            line=dict(color="#9b59b6", dash="dash", width=2),
            hovertemplate="Comm-Free: %{y:.1f}%<extra></extra>",
        ), secondary_y=True)
    iw_labels = iw["week_label"].tolist()
    fig_inc.update_layout(
        title="Incentive ROI: Spend vs Satisfaction",
        height=420, hovermode="x unified", plot_bgcolor="white",
        legend=dict(orientation="h", y=-0.18),
        xaxis=dict(categoryorder="array", categoryarray=iw_labels),
    )
    fig_inc.update_yaxes(title_text="Incentive (M Rials)", gridcolor="#e8e8e8",
                         secondary_y=False)
    fig_inc.update_yaxes(title_text="Satisfaction / %", showgrid=False,
                         secondary_y=True)
    figs.append(fig_inc)

    # --- Incentive Funnel (area fill to show conversion gap) ---
    fig_funnel = go.Figure()
    if "snapp_gotmsg_pct" in iw.columns:
        fig_funnel.add_trace(go.Scatter(
            x=iw["week_label"], y=iw["snapp_gotmsg_pct"], name="Got Message %",
            line=dict(color="#2196F3", width=2.5), mode="lines+markers",
            marker=dict(size=6), fill="tozeroy",
            fillcolor="rgba(33,150,243,0.15)",
            hovertemplate="Got Message: %{y:.1f}%<extra></extra>",
        ))
    if "snapp_participation_pct" in iw.columns:
        fig_funnel.add_trace(go.Scatter(
            x=iw["week_label"], y=iw["snapp_participation_pct"],
            name="Participated %",
            line=dict(color=SNAPP_COLOR, width=2.5), mode="lines+markers",
            marker=dict(size=6), fill="tozeroy",
            fillcolor="rgba(0,200,83,0.15)",
            hovertemplate="Participated: %{y:.1f}%<extra></extra>",
        ))
    fig_funnel.update_layout(
        title="Incentive Funnel: Message → Participation (gap = drop-off)",
        height=350, yaxis_title="%", hovermode="x unified",
        plot_bgcolor="white", yaxis=dict(gridcolor="#e8e8e8"),
        xaxis=dict(categoryorder="array", categoryarray=iw_labels),
        legend=dict(orientation="h", y=-0.2),
    )
    figs.append(fig_funnel)

    # --- Incentive Types (lollipop chart) ---
    if len(it) > 0:
        it_sorted = it.sort_values("n", ascending=True)
        fig_types = go.Figure()
        # Stems
        for _, row in it_sorted.iterrows():
            fig_types.add_trace(go.Scatter(
                x=[0, row["n"]], y=[row["reason"]] * 2,
                mode="lines", line=dict(color="#c8e6c9", width=3),
                showlegend=False, hoverinfo="skip",
            ))
        # Dots
        fig_types.add_trace(go.Scatter(
            x=it_sorted["n"], y=it_sorted["reason"],
            mode="markers+text", name="Count",
            marker=dict(color=SNAPP_COLOR, size=14,
                        line=dict(width=1.5, color="white")),
            text=[f"{int(n):,}" for n in it_sorted["n"]],
            textposition="middle right", textfont=dict(size=11),
            hovertemplate="%{y}: %{x:,}<extra></extra>",
        ))
        fig_types.update_layout(
            title="Snapp Incentive Types Received",
            height=max(300, len(it_sorted) * 35 + 100),
            xaxis_title="Count", margin=dict(l=200),
            plot_bgcolor="white", showlegend=False,
            xaxis=dict(gridcolor="#e8e8e8"),
        )
        figs.append(fig_types)

    # --- Unsatisfaction Reasons (butterfly / back-to-back) ---
    if len(ur) > 0:
        snapp_ur = ur[ur["platform"] == "Snapp"].set_index("reason")
        tapsi_ur = ur[ur["platform"] == "Tapsi"].set_index("reason")
        all_reasons = ur["reason"].unique().tolist()
        fig_unsat = go.Figure()
        # Snapp on the right (positive)
        snapp_vals = [int(snapp_ur.loc[r, "n"]) if r in snapp_ur.index else 0
                      for r in all_reasons]
        tapsi_vals = [-int(tapsi_ur.loc[r, "n"]) if r in tapsi_ur.index else 0
                      for r in all_reasons]
        fig_unsat.add_trace(go.Bar(
            y=all_reasons, x=snapp_vals, name="Snapp",
            orientation="h", marker_color=SNAPP_COLOR,
            text=[f"{v:,}" for v in snapp_vals],
            textposition="outside",
            hovertemplate="Snapp — %{y}: %{x:,}<extra></extra>",
        ))
        fig_unsat.add_trace(go.Bar(
            y=all_reasons, x=tapsi_vals, name="Tapsi",
            orientation="h", marker_color=TAPSI_COLOR,
            text=[f"{abs(v):,}" for v in tapsi_vals],
            textposition="outside",
            hovertemplate="Tapsi — %{y}: %{customdata:,}<extra></extra>",
            customdata=[abs(v) for v in tapsi_vals],
        ))
        fig_unsat.update_layout(
            title="Incentive Unsatisfaction Reasons (Snapp → | ← Tapsi)",
            barmode="overlay", height=max(350, len(all_reasons) * 40 + 100),
            xaxis_title="Count", margin=dict(l=180),
            plot_bgcolor="white",
            xaxis=dict(gridcolor="#e8e8e8"),
            legend=dict(orientation="h", y=-0.15),
        )
        figs.append(fig_unsat)

    # --- Incentive by City (dot plot with satisfaction) ---
    fig_ic = go.Figure()
    # Incentive amount as dots
    fig_ic.add_trace(go.Scatter(
        y=ic["city"], x=ic["snapp_incentive_avg"],
        mode="markers", name="Avg Incentive (Rials)",
        marker=dict(color=SNAPP_COLOR, size=13, symbol="circle",
                    line=dict(width=1, color="white")),
        hovertemplate="Incentive: %{x:,.0f} Rials<extra>%{y}</extra>",
    ))
    fig_ic.update_layout(
        title="Average Snapp Incentive by City",
        height=650, xaxis_title="Rials",
        margin=dict(l=120), plot_bgcolor="white",
        xaxis=dict(gridcolor="#e0e0e0"),
        yaxis=dict(dtick=1),
    )
    figs.append(fig_ic)

    return figs


# ============================================================
# PAGE 4: OPERATIONS & DEMOGRAPHICS
# ============================================================
def page4_operations(views):
    figs = []
    nav = views["vw_NavigationUsage"]
    nav_wk = add_persian_week_labels(views["vw_NavigationByWeek"])
    demo = views["vw_Demographics"]
    rs = add_persian_week_labels(views["vw_RideShareByCityWeek"])

    # --- Navigation App Usage (side-by-side donut charts per context) ---
    app_colors = {"Google Map": "#4285F4", "Waze": "#33CCFF", "Neshan": "#FF5722",
                  "Balad": "#9C27B0", "No Navigation App": "#95a5a6"}
    contexts = list(nav["context"].unique())
    if len(contexts) >= 2:
        fig_nav = make_subplots(
            rows=1, cols=len(contexts),
            specs=[[{"type": "pie"}] * len(contexts)],
            subplot_titles=contexts,
        )
        for i, ctx_name in enumerate(contexts):
            sub = nav[nav["context"] == ctx_name].sort_values("n", ascending=False)
            pie_colors = [app_colors.get(a.strip(), "#7f8c8d") for a in sub["nav_app"]]
            fig_nav.add_trace(go.Pie(
                labels=sub["nav_app"], values=sub["n"],
                marker=dict(colors=pie_colors),
                hole=0.45, textinfo="label+percent",
                textposition="outside", textfont=dict(size=11),
                hovertemplate="%{label}: %{value:,} (%{percent})<extra></extra>",
            ), row=1, col=i + 1)
        fig_nav.update_layout(
            title="Navigation App Usage by Context",
            height=400, showlegend=False,
        )
        figs.append(fig_nav)
    else:
        for ctx_name in contexts:
            sub = nav[nav["context"] == ctx_name].sort_values("n", ascending=False)
            pie_colors = [app_colors.get(a.strip(), "#7f8c8d") for a in sub["nav_app"]]
            fig_nav = go.Figure(go.Pie(
                labels=sub["nav_app"], values=sub["n"],
                marker=dict(colors=pie_colors),
                hole=0.45, textinfo="label+percent",
                textposition="outside",
            ))
            fig_nav.update_layout(
                title=f"Navigation App Usage ({ctx_name})", height=400,
            )
            figs.append(fig_nav)

    # --- Navigation by Week (stacked area — market share view) ---
    top_apps = nav_wk.groupby("nav_app")["n"].sum().nlargest(5).index
    nav_top = nav_wk[nav_wk["nav_app"].isin(top_apps)]
    fig_nav_wk = go.Figure()
    for app in top_apps:
        sub = nav_top[nav_top["nav_app"] == app].sort_values("yearweek")
        fig_nav_wk.add_trace(go.Scatter(
            x=sub["week_label"], y=sub["pct"], name=app.strip(),
            mode="lines", stackgroup="one",
            line=dict(width=0.5, color=app_colors.get(app.strip(), "#7f8c8d")),
            fillcolor=app_colors.get(app.strip(), "#7f8c8d"),
            hovertemplate=f"{app.strip()}: %{{y:.1f}}%<extra></extra>",
        ))
    fig_nav_wk.update_layout(
        title="Navigation App Market Share Over Time",
        height=420, yaxis_title="% of Drivers",
        hovermode="x unified", plot_bgcolor="white",
        yaxis=dict(gridcolor="#e8e8e8"),
        xaxis=dict(categoryorder="array",
                   categoryarray=nav_wk.sort_values("yearweek")["week_label"].unique().tolist()),
        legend=dict(orientation="h", y=-0.15),
    )
    figs.append(fig_nav_wk)

    # --- Demographics Distributions ---
    demo_colors = {
        "cooperation_type": "#3498db",
        "age_group": "#2196F3",
        "city": "#3498db",
    }
    for dim in ["cooperation_type", "age_group", "gender", "city"]:
        sub = demo[demo["dimension"] == dim].copy()

        # Gender → donut chart
        if dim == "gender":
            gender_colors = {"Male": "#2196F3", "Female": "#E91E63",
                             "مرد": "#2196F3", "زن": "#E91E63"}
            total = sub["n"].sum()
            sub = sub.sort_values("n", ascending=True)
            bar_colors = [gender_colors.get(c.strip(), "#95a5a6") for c in sub["category"]]
            fig_demo = go.Figure(go.Bar(
                y=sub["category"], x=sub["n"], orientation="h",
                marker_color=bar_colors,
                text=[f"{int(n):,}  ({n/total*100:.1f}%)" for n in sub["n"]],
                textposition="inside", textfont=dict(size=15, color="white"),
                insidetextanchor="middle",
                hovertemplate="%{y}: %{x:,}<extra></extra>",
            ))
            fig_demo.update_layout(
                title="Distribution: Gender",
                height=200, margin=dict(l=100, r=40),
                plot_bgcolor="white", xaxis=dict(gridcolor="#e8e8e8"),
                xaxis_title="Count",
            )
            figs.append(fig_demo)
            continue

        if dim == "city":
            sub = sub[sub["category"].isin(CITY_ORDER)]
            sub["category"] = pd.Categorical(
                sub["category"], categories=list(reversed(CITY_ORDER)), ordered=True)
            sub = sub.sort_values("category")
        else:
            sub = sub.sort_values("n", ascending=True)

        total = sub["n"].sum()
        fig_demo = go.Figure()
        # Lollipop style
        for _, row in sub.iterrows():
            fig_demo.add_trace(go.Scatter(
                x=[0, row["n"]], y=[row["category"]] * 2,
                mode="lines", line=dict(color="#e0e0e0", width=4),
                showlegend=False, hoverinfo="skip",
            ))
        fig_demo.add_trace(go.Scatter(
            y=sub["category"], x=sub["n"],
            mode="markers+text", showlegend=False,
            marker=dict(color=demo_colors.get(dim, "#3498db"), size=12,
                        line=dict(width=1, color="white")),
            text=[f"{int(n):,} ({n/total*100:.1f}%)" for n in sub["n"]],
            textposition="middle right", textfont=dict(size=10),
            hovertemplate="%{y}: %{x:,}<extra></extra>",
        ))
        fig_demo.update_layout(
            title=f"Distribution: {dim.replace('_', ' ').title()}",
            height=max(300, len(sub) * 30 + 100),
            xaxis_title="Count", margin=dict(l=150),
            plot_bgcolor="white", xaxis=dict(gridcolor="#e8e8e8"),
        )
        figs.append(fig_demo)

    # --- Ride Share: Aggregate weekly (stacked bar — cleaner than overlapping area) ---
    label_cats = rs["week_label"].cat.categories.tolist()
    rs_wk = rs.groupby(["yearweek", "week_label"], observed=True).agg(
        snapp_rides=("snapp_rides_total", "sum"),
        tapsi_rides=("joint_tapsi_rides", "sum"),
        total=("total_rides", "sum")
    ).reset_index().sort_values("yearweek")
    rs_wk["week_label"] = pd.Categorical(rs_wk["week_label"], categories=label_cats, ordered=True)
    rs_wk["snapp_pct"] = rs_wk["snapp_rides"] / rs_wk["total"] * 100
    rs_wk["tapsi_pct"] = rs_wk["tapsi_rides"] / rs_wk["total"] * 100

    fig_rs = go.Figure()
    ordered_labels = rs_wk.sort_values("yearweek")["week_label"].tolist()
    fig_rs.add_trace(go.Scatter(
        x=rs_wk["week_label"], y=rs_wk["snapp_pct"], name="Snapp %",
        mode="lines+markers", line=dict(color=SNAPP_COLOR, width=2.5),
        marker=dict(size=5, line=dict(width=1, color="white")),
        hovertemplate="Snapp: %{y:.1f}%<extra></extra>",
    ))
    fig_rs.add_trace(go.Scatter(
        x=rs_wk["week_label"], y=rs_wk["tapsi_pct"], name="Tapsi %",
        mode="lines+markers", line=dict(color=TAPSI_COLOR, width=2.5),
        marker=dict(size=5, line=dict(width=1, color="white")),
        hovertemplate="Tapsi: %{y:.1f}%<extra></extra>",
    ))
    fig_rs.update_layout(
        title="Ride Share: Snapp vs Tapsi Over Time",
        height=380,
        yaxis_title="% of Total Rides",
        hovermode="x unified", plot_bgcolor="white",
        yaxis=dict(gridcolor="#e8e8e8"),
        xaxis=dict(categoryorder="array", categoryarray=ordered_labels),
        legend=dict(orientation="h", y=-0.15),
    )
    figs.append(fig_rs)

    return figs


# ============================================================
# PAGE 5: SURVEY EXPLORER
# ============================================================
def page5_survey(views):
    figs = []
    la = views["vw_LongSurveyAnswers"]
    lr = views["vw_LongRareSurveyAnswers"]

    # Color gradient for percentage bars
    def pct_color(pct, max_pct):
        """Return a color from light blue to deep blue based on % share."""
        ratio = min(pct / max(max_pct, 1), 1.0)
        r = int(220 - 180 * ratio)
        g = int(235 - 155 * ratio)
        b = int(255 - 50 * ratio)
        return f"rgb({r},{g},{b})"

    for label, df in [("Main Survey Questions", la), ("Rare Survey Questions", lr)]:
        questions = df.groupby("question")["n"].sum().nlargest(12).index
        for q in questions:
            sub = df[df["question"] == q].sort_values("pct", ascending=True)
            if len(sub) > 15:
                sub = sub.tail(15)
            max_pct = sub["pct"].max()
            bar_colors = [pct_color(p, max_pct) for p in sub["pct"]]

            fig_q = go.Figure(go.Bar(
                y=sub["answer"], x=sub["pct"], orientation="h",
                marker_color=bar_colors,
                marker_line=dict(color="rgba(30,80,180,0.3)", width=0.5),
                text=[f"{p:.1f}% (n={int(n):,})" for p, n in zip(sub["pct"], sub["n"])],
                textposition="outside", textfont=dict(size=10),
                hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
            ))
            fig_q.update_layout(
                title=f"{q}",
                height=max(280, len(sub) * 30 + 80),
                xaxis_title="% of Responses",
                margin=dict(l=220, r=100),
                plot_bgcolor="white",
                xaxis=dict(gridcolor="#e8e8e8"),
            )
            figs.append(fig_q)

    return figs


# ============================================================
# PAGE 6: RIDE SHARE BY CITY
# ============================================================
def page6_rideshare_city(views):
    figs = []
    rs = add_persian_week_labels(views["vw_RideShareByCityWeek"])
    ordered_labels = rs.sort_values("yearweek")["week_label"].unique().tolist()

    for city in CITY_ORDER:
        city_data = rs[rs["city"] == city].sort_values("yearweek").copy()
        if len(city_data) == 0:
            continue
        city_data["snapp_pct"] = city_data["snapp_rides_total"] / city_data["total_rides"] * 100
        city_data["tapsi_pct"] = city_data["joint_tapsi_rides"] / city_data["total_rides"] * 100

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=city_data["week_label"], y=city_data["snapp_pct"], name="Snapp %",
            mode="lines+markers", line=dict(color=SNAPP_COLOR, width=2),
            marker=dict(size=5, line=dict(width=1, color="white")),
            hovertemplate="Snapp: %{y:.1f}%<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=city_data["week_label"], y=city_data["tapsi_pct"], name="Tapsi %",
            mode="lines+markers", line=dict(color=TAPSI_COLOR, width=2),
            marker=dict(size=5, line=dict(width=1, color="white")),
            hovertemplate="Tapsi: %{y:.1f}%<extra></extra>",
        ))
        n_total = int(city_data["total_rides"].sum())
        fig.update_layout(
            title=f"{city} — Ride Share (n={n_total:,} total rides)",
            height=320,
            yaxis_title="% of Rides",
            hovermode="x unified", plot_bgcolor="white",
            yaxis=dict(gridcolor="#e8e8e8"),
            xaxis=dict(categoryorder="array", categoryarray=ordered_labels),
            legend=dict(orientation="h", y=-0.2),
            margin=dict(t=40, b=60),
        )
        figs.append(fig)

    return figs


# ============================================================
# ASSEMBLE HTML
# ============================================================
def build_html(all_pages):
    """Build a multi-page HTML dashboard with tab navigation."""

    page_names = [
        "Executive Overview",
        "Satisfaction Deep-Dive",
        "Incentive Analysis",
        "Operations & Demographics",
        "Survey Explorer",
        "Ride Share by City",
    ]

    # Convert each figure to HTML div
    page_divs = []
    for i, (name, figs) in enumerate(zip(page_names, all_pages)):
        chart_htmls = []
        for fig in figs:
            fig.update_layout(
                plot_bgcolor="white",
                paper_bgcolor=CARD_BG,
                font=dict(family="Segoe UI, sans-serif"),
            )
            chart_htmls.append(fig.to_html(
                full_html=False, include_plotlyjs=False,
                config={"displayModeBar": True, "responsive": True}
            ))
        display = "block" if i == 0 else "none"
        page_divs.append(
            f'<div id="page{i}" class="page-content" style="display:{display}">'
            + "\n".join(f'<div class="chart-card">{h}</div>' for h in chart_htmls)
            + "</div>"
        )

    # Tab buttons
    tab_buttons = "\n".join(
        f'<button class="tab-btn{"  active" if i == 0 else ""}" '
        f'onclick="showPage({i})">{name}</button>'
        for i, name in enumerate(page_names)
    )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Driver Survey Dashboard</title>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {BG_COLOR};
            color: #2c3e50;
        }}
        .header {{
            background: linear-gradient(135deg, #2c3e50, #34495e);
            color: white;
            padding: 20px 30px;
            text-align: center;
        }}
        .header h1 {{ font-size: 24px; margin-bottom: 5px; }}
        .header p {{ font-size: 13px; opacity: 0.8; }}
        .tab-bar {{
            background: #34495e;
            padding: 0 20px;
            display: flex;
            gap: 2px;
            overflow-x: auto;
        }}
        .tab-btn {{
            background: transparent;
            color: #bdc3c7;
            border: none;
            padding: 12px 24px;
            font-size: 14px;
            cursor: pointer;
            white-space: nowrap;
            border-bottom: 3px solid transparent;
            transition: all 0.2s;
        }}
        .tab-btn:hover {{ color: white; background: rgba(255,255,255,0.05); }}
        .tab-btn.active {{
            color: white;
            border-bottom-color: {SNAPP_COLOR};
            background: rgba(255,255,255,0.1);
        }}
        .page-content {{ padding: 20px; max-width: 1400px; margin: 0 auto; }}
        .chart-card {{
            background: {CARD_BG};
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            margin-bottom: 20px;
            padding: 15px;
            overflow: hidden;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Driver Survey Dashboard</h1>
        <p>Cab_Studies Database | 238,261 responses | Auto-generated from SQL Views</p>
    </div>
    <div class="tab-bar">
        {tab_buttons}
    </div>
    {"".join(page_divs)}
    <script>
        function showPage(idx) {{
            document.querySelectorAll('.page-content').forEach((p, i) => {{
                p.style.display = i === idx ? 'block' : 'none';
            }});
            document.querySelectorAll('.tab-btn').forEach((b, i) => {{
                b.classList.toggle('active', i === idx);
            }});
            // Trigger plotly resize for newly visible charts
            setTimeout(() => window.dispatchEvent(new Event('resize')), 100);
        }}
    </script>
</body>
</html>"""
    return html


# ============================================================
# MAIN
# ============================================================
def main():
    print("=" * 50)
    print("  Building Driver Survey Dashboard")
    print("=" * 50)

    print("\nLoading data from SQL Server...")
    views = load_all_views()

    print("\nBuilding Page 1: Executive Overview...")
    p1 = page1_executive(views)
    print(f"  {len(p1)} charts")

    print("Building Page 2: Satisfaction Deep-Dive...")
    p2 = page2_satisfaction(views)
    print(f"  {len(p2)} charts")

    print("Building Page 3: Incentive Analysis...")
    p3 = page3_incentive(views)
    print(f"  {len(p3)} charts")

    print("Building Page 4: Operations & Demographics...")
    p4 = page4_operations(views)
    print(f"  {len(p4)} charts")

    print("Building Page 5: Survey Explorer...")
    p5 = page5_survey(views)
    print(f"  {len(p5)} charts")

    print("Building Page 6: Ride Share by City...")
    p6 = page6_rideshare_city(views)
    print(f"  {len(p6)} charts")

    total = len(p1)+len(p2)+len(p3)+len(p4)+len(p5)+len(p6)
    print(f"\nAssembling HTML ({total} total charts)...")
    html = build_html([p1, p2, p3, p4, p5, p6])

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\nDashboard saved to: {OUTPUT}")
    print(f"File size: {os.path.getsize(OUTPUT) / 1024 / 1024:.1f} MB")
    print("\nOpening in browser...")
    os.startfile(OUTPUT)


if __name__ == "__main__":
    main()

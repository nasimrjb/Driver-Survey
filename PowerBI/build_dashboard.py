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
    ws = views["vw_WeeklySatisfaction"].sort_values("yearweek")
    nps = views["vw_WeeklyNPS"].sort_values("yearweek")

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
                x=ws["yearweek"], y=ws[col], name=name,
                line=dict(color=color, dash=dash, width=2),
                hovertemplate=f"{name}: %{{y:.2f}}<extra></extra>"
            ), secondary_y=False)

    # Joint driver % on secondary axis
    fig_sat.add_trace(go.Scatter(
        x=ws["yearweek"], y=ws["joint_driver_pct"], name="Joint %",
        line=dict(color="#7f8c8d", dash="dot", width=1.5),
        opacity=0.6
    ), secondary_y=True)

    fig_sat.update_layout(
        title="Weekly Satisfaction Trends (1-5 scale)",
        height=400, hovermode="x unified",
        legend=dict(orientation="h", y=-0.15),
        margin=dict(t=40, b=60)
    )
    fig_sat.update_yaxes(title_text="Satisfaction (1-5)", range=[1, 5], secondary_y=False)
    fig_sat.update_yaxes(title_text="Joint %", secondary_y=True)
    figs.append(fig_sat)

    # --- Weekly NPS (filter to weeks with data, connect gaps) ---
    nps_valid = nps.dropna(subset=["snapp_nps"])
    fig_nps = go.Figure()
    fig_nps.add_trace(go.Scatter(
        x=nps_valid["yearweek"], y=nps_valid["snapp_nps"], name="Snapp NPS",
        line=dict(color=SNAPP_COLOR, width=3), mode="lines+markers",
        marker=dict(size=5),
        fill="tozeroy", fillcolor="rgba(0,200,83,0.1)",
        connectgaps=True
    ))
    nps_tapsi = nps.dropna(subset=["tapsi_nps"])
    if len(nps_tapsi) >= 3:
        fig_nps.add_trace(go.Scatter(
            x=nps_tapsi["yearweek"], y=nps_tapsi["tapsi_nps"], name="Tapsi NPS",
            line=dict(color=TAPSI_COLOR, width=3), mode="lines+markers",
            marker=dict(size=5),
            fill="tozeroy", fillcolor="rgba(255,109,0,0.1)",
            connectgaps=True
        ))
    fig_nps.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
    fig_nps.update_layout(
        title=f"Weekly NPS Score (Snapp: {len(nps_valid)} weeks, Tapsi: {len(nps_tapsi)} weeks — rare question)",
        height=350, hovermode="x unified",
        yaxis_title="NPS",
        legend=dict(orientation="h", y=-0.15)
    )
    figs.append(fig_nps)

    # --- NPS Decomposition (Stacked Area, filter to valid weeks) ---
    nps_decomp = nps.dropna(subset=["snapp_detractor_pct"])
    fig_nps_decomp = go.Figure()
    fig_nps_decomp.add_trace(go.Scatter(
        x=nps_decomp["yearweek"], y=nps_decomp["snapp_detractor_pct"], name="Detractors",
        stackgroup="one", line=dict(color="#e74c3c"), fillcolor="rgba(231,76,60,0.4)"
    ))
    fig_nps_decomp.add_trace(go.Scatter(
        x=nps_decomp["yearweek"], y=nps_decomp["snapp_passive_pct"], name="Passives",
        stackgroup="one", line=dict(color="#f39c12"), fillcolor="rgba(243,156,18,0.4)"
    ))
    fig_nps_decomp.add_trace(go.Scatter(
        x=nps_decomp["yearweek"], y=nps_decomp["snapp_promoter_pct"], name="Promoters",
        stackgroup="one", line=dict(color="#27ae60"), fillcolor="rgba(39,174,96,0.4)"
    ))
    fig_nps_decomp.update_layout(
        title="Snapp NPS Decomposition (Detractors / Passives / Promoters %)",
        height=350, yaxis_title="%", hovermode="x unified",
        legend=dict(orientation="h", y=-0.15)
    )
    figs.append(fig_nps_decomp)

    # --- Response Count ---
    fig_resp = go.Figure(go.Bar(
        x=ws["yearweek"], y=ws["response_count"],
        marker_color="#3498db", opacity=0.7, name="Responses"
    ))
    fig_resp.update_layout(
        title="Weekly Response Count",
        height=250, yaxis_title="Responses",
        margin=dict(t=40, b=30)
    )
    figs.append(fig_resp)

    return figs


# ============================================================
# PAGE 2: SATISFACTION DEEP-DIVE
# ============================================================
def page2_satisfaction(views):
    figs = []
    sc = views["vw_SatisfactionByCity"].sort_values("n", ascending=False).head(15)
    sd = views["vw_SatisfactionByDemographics"]
    hm = views["vw_HoneymoonEffect"]
    scw = views["vw_SatisfactionByCityWeek"]

    # --- Top Cities Satisfaction ---
    fig_city = go.Figure()
    for col, name, color in [
        ("snapp_fare_sat", "Snapp Fare", SNAPP_COLOR),
        ("snapp_income_sat", "Snapp Income", "#2ecc71"),
        ("snapp_req_sat", "Snapp Request", "#1abc9c"),
        ("tapsi_fare_sat", "Tapsi Fare", TAPSI_COLOR),
    ]:
        if col in sc.columns:
            fig_city.add_trace(go.Bar(
                y=sc["city"], x=sc[col], name=name,
                orientation="h", marker_color=color, opacity=0.85
            ))
    fig_city.update_layout(
        title="Satisfaction by City (Top 15)",
        barmode="group", height=500,
        xaxis_title="Satisfaction (1-5)", xaxis_range=[1, 5],
        legend=dict(orientation="h", y=-0.12),
        margin=dict(l=120)
    )
    figs.append(fig_city)

    # --- Satisfaction by Demographics ---
    for dim_name in ["cooperation_type", "driver_type", "age_group", "gender"]:
        sub = sd[sd["dimension"] == dim_name].copy()
        if len(sub) == 0:
            continue
        fig_dim = go.Figure()
        for col, name, color in [
            ("snapp_fare_sat", "Snapp Fare", SNAPP_COLOR),
            ("snapp_income_sat", "Snapp Income", "#2ecc71"),
            ("tapsi_fare_sat", "Tapsi Fare", TAPSI_COLOR),
        ]:
            if col in sub.columns:
                fig_dim.add_trace(go.Bar(
                    x=sub["category"], y=sub[col], name=name,
                    marker_color=color, opacity=0.85,
                    text=[f"n={int(n)}" for n in sub["n"]],
                    textposition="outside"
                ))
        fig_dim.update_layout(
            title=f"Satisfaction by {dim_name.replace('_', ' ').title()}",
            barmode="group", height=350,
            yaxis_title="Satisfaction (1-5)", yaxis_range=[1, 5],
            legend=dict(orientation="h", y=-0.15)
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
    hm_sorted = hm.copy()
    hm_sorted["sort_key"] = hm_sorted["tenure"].apply(
        lambda x: tenure_order.index(x.strip()) if x.strip() in tenure_order else 99
    )
    hm_sorted = hm_sorted.sort_values("sort_key")

    fig_hon = go.Figure()
    for col, name, color in [
        ("snapp_fare_sat", "Fare Sat", SNAPP_COLOR),
        ("snapp_income_sat", "Income Sat", "#2ecc71"),
        ("snapp_overall_sat", "Overall Sat", "#3498db"),
    ]:
        if col in hm_sorted.columns:
            series = hm_sorted[col]
            # Skip traces with >50% null values
            if series.notna().sum() < len(series) * 0.5:
                continue
            fig_hon.add_trace(go.Scatter(
                x=hm_sorted["tenure"], y=hm_sorted[col], name=name,
                mode="lines+markers", line=dict(color=color, width=2.5),
                marker=dict(size=8), connectgaps=True
            ))
    fig_hon.update_layout(
        title="Honeymoon Effect: Satisfaction by Snapp Tenure",
        height=400, yaxis_title="Satisfaction (1-5)", yaxis_range=[1, 5],
        xaxis_title="Tenure", legend=dict(orientation="h", y=-0.2),
        xaxis_tickangle=-30
    )
    figs.append(fig_hon)

    # --- City x Week Heatmap (top 12 cities) ---
    top_cities = scw.groupby("city")["n"].sum().nlargest(12).index.tolist()
    hm_data = scw[scw["city"].isin(top_cities)].pivot_table(
        index="city", columns="yearweek", values="snapp_fare_sat"
    )
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
        height=450, margin=dict(l=120),
        xaxis_title="Year-Week"
    )
    figs.append(fig_heatmap)

    return figs


# ============================================================
# PAGE 3: INCENTIVE ANALYSIS
# ============================================================
def page3_incentive(views):
    figs = []
    iw = views["vw_IncentiveByWeek"].sort_values("yearweek")
    it = views["vw_WideIncentiveTypes"]
    ur = views["vw_WideUnsatisfactionReasons"]
    ic = views["vw_IncentiveByCity"].sort_values("n", ascending=False).head(15)

    # --- Incentive Amount + Satisfaction over time ---
    fig_inc = make_subplots(specs=[[{"secondary_y": True}]])
    fig_inc.add_trace(go.Bar(
        x=iw["yearweek"], y=iw["snapp_incentive_avg_mrial"],
        name="Snapp Incentive (M Rials)",
        marker_color=SNAPP_COLOR, opacity=0.5
    ), secondary_y=False)
    fig_inc.add_trace(go.Scatter(
        x=iw["yearweek"], y=iw["snapp_inc_sat_avg"],
        name="Incentive Satisfaction",
        line=dict(color="#e74c3c", width=2.5), mode="lines+markers"
    ), secondary_y=True)
    if "snapp_commfree_pct" in iw.columns:
        fig_inc.add_trace(go.Scatter(
            x=iw["yearweek"], y=iw["snapp_commfree_pct"],
            name="Comm-Free %",
            line=dict(color="#9b59b6", dash="dash", width=2)
        ), secondary_y=True)
    fig_inc.update_layout(
        title="Incentive ROI: Spend vs Satisfaction",
        height=400, hovermode="x unified",
        legend=dict(orientation="h", y=-0.15)
    )
    fig_inc.update_yaxes(title_text="Incentive (M Rials)", secondary_y=False)
    fig_inc.update_yaxes(title_text="Satisfaction / %", secondary_y=True)
    figs.append(fig_inc)

    # --- Incentive Funnel ---
    fig_funnel = make_subplots(specs=[[{"secondary_y": False}]])
    for col, name, color in [
        ("snapp_gotmsg_pct", "Got Message %", "#3498db"),
        ("snapp_participation_pct", "Participated %", "#2ecc71"),
    ]:
        if col in iw.columns:
            fig_funnel.add_trace(go.Scatter(
                x=iw["yearweek"], y=iw[col], name=name,
                line=dict(color=color, width=2.5), mode="lines+markers"
            ))
    fig_funnel.update_layout(
        title="Incentive Funnel: Message -> Participation",
        height=300, yaxis_title="%", hovermode="x unified",
        legend=dict(orientation="h", y=-0.2)
    )
    figs.append(fig_funnel)

    # --- Incentive Types ---
    if len(it) > 0:
        it_sorted = it.sort_values("n", ascending=True)
        fig_types = go.Figure(go.Bar(
            y=it_sorted["reason"], x=it_sorted["n"],
            orientation="h", marker_color=SNAPP_COLOR,
            text=[f"{int(n):,}" for n in it_sorted["n"]],
            textposition="outside"
        ))
        fig_types.update_layout(
            title="Snapp Incentive Types Received",
            height=300, xaxis_title="Count",
            margin=dict(l=200)
        )
        figs.append(fig_types)

    # --- Unsatisfaction Reasons ---
    if len(ur) > 0:
        fig_unsat = go.Figure()
        for platform, color in [("Snapp", SNAPP_COLOR), ("Tapsi", TAPSI_COLOR)]:
            sub = ur[ur["platform"] == platform].sort_values("n", ascending=True)
            fig_unsat.add_trace(go.Bar(
                y=sub["reason"], x=sub["n"], name=platform,
                orientation="h", marker_color=color, opacity=0.85
            ))
        fig_unsat.update_layout(
            title="Incentive Unsatisfaction Reasons",
            barmode="group", height=350,
            xaxis_title="Count", margin=dict(l=150),
            legend=dict(orientation="h", y=-0.15)
        )
        figs.append(fig_unsat)

    # --- Incentive by City ---
    fig_ic = go.Figure()
    fig_ic.add_trace(go.Bar(
        y=ic["city"], x=ic["snapp_incentive_avg"],
        name="Avg Incentive (Rials)", orientation="h",
        marker_color=SNAPP_COLOR, opacity=0.7
    ))
    fig_ic.update_layout(
        title="Average Snapp Incentive by City (Top 15)",
        height=450, xaxis_title="Rials",
        margin=dict(l=120)
    )
    figs.append(fig_ic)

    return figs


# ============================================================
# PAGE 4: OPERATIONS & DEMOGRAPHICS
# ============================================================
def page4_operations(views):
    figs = []
    nav = views["vw_NavigationUsage"]
    nav_wk = views["vw_NavigationByWeek"]
    demo = views["vw_Demographics"]
    rs = views["vw_RideShareByCityWeek"]

    # --- Navigation App Usage ---
    for ctx_name in nav["context"].unique():
        sub = nav[nav["context"] == ctx_name].sort_values("n", ascending=True)
        colors = {"Google Map": "#4285F4", "Waze": "#33CCFF", "Neshan": "#FF5722",
                  "Balad": "#9C27B0", "No Navigation App": "#95a5a6"}
        bar_colors = [colors.get(a.strip(), "#7f8c8d") for a in sub["nav_app"]]
        fig_nav = go.Figure(go.Bar(
            y=sub["nav_app"], x=sub["n"], orientation="h",
            marker_color=bar_colors,
            text=[f"{int(n):,}" for n in sub["n"]],
            textposition="outside"
        ))
        fig_nav.update_layout(
            title=f"Navigation App Usage ({ctx_name})",
            height=300, xaxis_title="Count",
            margin=dict(l=150)
        )
        figs.append(fig_nav)

    # --- Navigation by Week ---
    top_apps = nav_wk.groupby("nav_app")["n"].sum().nlargest(5).index
    nav_top = nav_wk[nav_wk["nav_app"].isin(top_apps)]
    fig_nav_wk = go.Figure()
    app_colors = {"Google Map": "#4285F4", "Waze": "#33CCFF", "Neshan": "#FF5722",
                  "Balad": "#9C27B0", "No Navigation App": "#95a5a6"}
    for app in top_apps:
        sub = nav_top[nav_top["nav_app"] == app].sort_values("yearweek")
        fig_nav_wk.add_trace(go.Scatter(
            x=sub["yearweek"], y=sub["pct"], name=app.strip(),
            mode="lines", line=dict(
                color=app_colors.get(app.strip(), "#7f8c8d"), width=2.5
            )
        ))
    fig_nav_wk.update_layout(
        title="Navigation App Adoption Over Time",
        height=400, yaxis_title="% of Drivers",
        hovermode="x unified",
        legend=dict(orientation="h", y=-0.15)
    )
    figs.append(fig_nav_wk)

    # --- Demographics Distributions ---
    for dim in ["cooperation_type", "age_group", "gender", "city"]:
        sub = demo[demo["dimension"] == dim].sort_values("n", ascending=True)
        if dim == "city":
            sub = sub.tail(15)  # top 15 cities
        fig_demo = go.Figure(go.Bar(
            y=sub["category"], x=sub["n"], orientation="h",
            marker_color="#3498db", opacity=0.8,
            text=[f"{int(n):,}" for n in sub["n"]],
            textposition="outside"
        ))
        fig_demo.update_layout(
            title=f"Distribution: {dim.replace('_', ' ').title()}",
            height=max(250, len(sub) * 25 + 100),
            xaxis_title="Count", margin=dict(l=150)
        )
        figs.append(fig_demo)

    # --- Ride Share: Aggregate weekly ---
    rs_wk = rs.groupby("yearweek").agg(
        snapp_rides=("snapp_rides_total", "sum"),
        tapsi_rides=("joint_tapsi_rides", "sum"),
        total=("total_rides", "sum")
    ).reset_index().sort_values("yearweek")
    rs_wk["snapp_pct"] = rs_wk["snapp_rides"] / rs_wk["total"] * 100
    rs_wk["tapsi_pct"] = rs_wk["tapsi_rides"] / rs_wk["total"] * 100

    fig_rs = go.Figure()
    fig_rs.add_trace(go.Scatter(
        x=rs_wk["yearweek"], y=rs_wk["snapp_pct"], name="Snapp %",
        fill="tozeroy", line=dict(color=SNAPP_COLOR)
    ))
    fig_rs.add_trace(go.Scatter(
        x=rs_wk["yearweek"], y=rs_wk["tapsi_pct"], name="Tapsi %",
        fill="tozeroy", line=dict(color=TAPSI_COLOR)
    ))
    fig_rs.update_layout(
        title="Ride Share: Snapp vs Tapsi Over Time",
        height=350, yaxis_title="% of Total Rides",
        hovermode="x unified",
        legend=dict(orientation="h", y=-0.15)
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

    for label, df in [("Main Survey Questions", la), ("Rare Survey Questions", lr)]:
        questions = df.groupby("question")["n"].sum().nlargest(12).index
        for q in questions:
            sub = df[df["question"] == q].sort_values("n", ascending=True)
            if len(sub) > 15:
                sub = sub.tail(15)
            fig_q = go.Figure(go.Bar(
                y=sub["answer"], x=sub["pct"], orientation="h",
                marker_color="#3498db", opacity=0.85,
                text=[f"{p:.1f}% (n={int(n)})" for p, n in zip(sub["pct"], sub["n"])],
                textposition="outside"
            ))
            fig_q.update_layout(
                title=f"{q}",
                height=max(280, len(sub) * 28 + 80),
                xaxis_title="% of Responses",
                margin=dict(l=200, r=80)
            )
            figs.append(fig_q)

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
        "Survey Explorer"
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

    print(f"\nAssembling HTML ({len(p1)+len(p2)+len(p3)+len(p4)+len(p5)} total charts)...")
    html = build_html([p1, p2, p3, p4, p5])

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\nDashboard saved to: {OUTPUT}")
    print(f"File size: {os.path.getsize(OUTPUT) / 1024 / 1024:.1f} MB")
    print("\nOpening in browser...")
    os.startfile(OUTPUT)


if __name__ == "__main__":
    main()

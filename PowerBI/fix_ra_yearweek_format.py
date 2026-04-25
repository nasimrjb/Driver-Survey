path = r"D:\Work\Driver Survey\PowerBI\create_views.sql"

with open(path, "r", encoding="utf-8") as f:
    sql = f.read()

FMT    = "CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,\n        yearweek AS yearweek_sort,"
FMT_SM = "CAST(sm.yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(sm.yearweek%100 AS VARCHAR), 2) AS yearweek,\n    sm.yearweek AS yearweek_sort,"

changes = []

# RA-1 SatReview — agg CTE
old = "    SELECT\n        yearweek, weeknumber, city, cooperation_type,\n        COUNT(*) AS n,"
new = "    SELECT\n        " + FMT + "\n        weeknumber, city, cooperation_type,\n        COUNT(*) AS n,"
assert old in sql, "SatReview agg CTE not found"
sql = sql.replace(old, new, 1)
changes.append("SatReview agg CTE")

# RA-1 SatReview — all_drv CTE
old = "        yearweek, weeknumber, city, 'All Drivers' AS cooperation_type,"
new = "        " + FMT + "\n        weeknumber, city, 'All Drivers' AS cooperation_type,"
assert old in sql, "SatReview all_drv CTE not found"
sql = sql.replace(old, new, 1)
changes.append("SatReview all_drv CTE")

# RA-2 CitiesOverview — outer SELECT
old = "SELECT\n    yearweek, weeknumber, city,\n    COUNT(*) AS E_n,"
new = "SELECT\n    " + FMT + "\n    weeknumber, city,\n    COUNT(*) AS E_n,"
assert old in sql, "CitiesOverview outer SELECT not found"
sql = sql.replace(old, new, 1)
changes.append("CitiesOverview outer SELECT")

# RA-3 RideShare — outer SELECT
old = "SELECT\n    yearweek, weeknumber, city,\n    COUNT(*) AS total_Res,"
new = "SELECT\n    " + FMT + "\n    weeknumber, city,\n    COUNT(*) AS total_Res,"
assert old in sql, "RideShare outer SELECT not found"
sql = sql.replace(old, new, 1)
changes.append("RideShare outer SELECT")

# RA-4 PersonaPartTime — outer SELECT
old = "SELECT\n    yearweek, weeknumber, city,\n    COUNT(*) AS total_Res,\n    SUM(CASE WHEN is_joint=1"
new = "SELECT\n    " + FMT + "\n    weeknumber, city,\n    COUNT(*) AS total_Res,\n    SUM(CASE WHEN is_joint=1"
assert old in sql, "PersonaPartTime outer SELECT not found"
sql = sql.replace(old, new, 1)
changes.append("PersonaPartTime outer SELECT")

# RA-5 IncentiveAmounts — Snapp branch (also fix PARTITION BY)
old = (
    "SELECT\n"
    "    yearweek, weeknumber, city, 'Snapp' AS platform,\n"
    "    snapp_incentive_rial_details AS incentive_range,\n"
    "    COUNT(*) AS n_range,\n"
    "    SUM(COUNT(*)) OVER (PARTITION BY weeknumber, city) AS n_total,\n"
    "    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY weeknumber, city),0) AS pct\n"
    "FROM [Cab].[DriverSurvey_ShortMain]\n"
    "WHERE city IS NOT NULL AND snapp_incentive_rial_details IS NOT NULL\n"
    "GROUP BY yearweek, weeknumber, city, snapp_incentive_rial_details"
)
new = (
    "SELECT\n"
    "    " + FMT + "\n"
    "    weeknumber, city, 'Snapp' AS platform,\n"
    "    snapp_incentive_rial_details AS incentive_range,\n"
    "    COUNT(*) AS n_range,\n"
    "    SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total,\n"
    "    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY yearweek, city),0) AS pct\n"
    "FROM [Cab].[DriverSurvey_ShortMain]\n"
    "WHERE city IS NOT NULL AND snapp_incentive_rial_details IS NOT NULL\n"
    "GROUP BY yearweek, weeknumber, city, snapp_incentive_rial_details"
)
assert old in sql, "IncentiveAmounts Snapp branch not found"
sql = sql.replace(old, new, 1)
changes.append("IncentiveAmounts Snapp branch")

# RA-5 IncentiveAmounts — Tapsi branch
old = (
    "SELECT\n"
    "    yearweek, weeknumber, city, 'Tapsi' AS platform,\n"
    "    tapsi_incentive_rial_details AS incentive_range,\n"
    "    COUNT(*) AS n_range,\n"
    "    SUM(COUNT(*)) OVER (PARTITION BY weeknumber, city) AS n_total,\n"
    "    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY weeknumber, city),0) AS pct\n"
    "FROM [Cab].[DriverSurvey_ShortMain]\n"
    "WHERE city IS NOT NULL AND tapsi_incentive_rial_details IS NOT NULL"
)
new = (
    "SELECT\n"
    "    " + FMT + "\n"
    "    weeknumber, city, 'Tapsi' AS platform,\n"
    "    tapsi_incentive_rial_details AS incentive_range,\n"
    "    COUNT(*) AS n_range,\n"
    "    SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total,\n"
    "    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY yearweek, city),0) AS pct\n"
    "FROM [Cab].[DriverSurvey_ShortMain]\n"
    "WHERE city IS NOT NULL AND tapsi_incentive_rial_details IS NOT NULL"
)
assert old in sql, "IncentiveAmounts Tapsi branch not found"
sql = sql.replace(old, new, 1)
changes.append("IncentiveAmounts Tapsi branch")

# RA-6 IncentiveDuration — Snapp branch
old = (
    "SELECT yearweek, weeknumber, city, 'Snapp' AS platform,\n"
    "    snapp_incentive_active_duration AS duration_bucket,\n"
    "    COUNT(*) AS n_range,\n"
    "    SUM(COUNT(*)) OVER (PARTITION BY weeknumber, city) AS n_total,\n"
    "    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY weeknumber, city),0) AS pct\n"
    "FROM [Cab].[DriverSurvey_ShortMain]\n"
    "WHERE city IS NOT NULL AND snapp_incentive_active_duration IS NOT NULL\n"
    "GROUP BY yearweek, weeknumber, city, snapp_incentive_active_duration"
)
new = (
    "SELECT " + FMT + "\n"
    "    weeknumber, city, 'Snapp' AS platform,\n"
    "    snapp_incentive_active_duration AS duration_bucket,\n"
    "    COUNT(*) AS n_range,\n"
    "    SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total,\n"
    "    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY yearweek, city),0) AS pct\n"
    "FROM [Cab].[DriverSurvey_ShortMain]\n"
    "WHERE city IS NOT NULL AND snapp_incentive_active_duration IS NOT NULL\n"
    "GROUP BY yearweek, weeknumber, city, snapp_incentive_active_duration"
)
assert old in sql, "IncentiveDuration Snapp not found"
sql = sql.replace(old, new, 1)
changes.append("IncentiveDuration Snapp branch")

# RA-6 IncentiveDuration — Tapsi branch
old = (
    "SELECT yearweek, weeknumber, city, 'Tapsi' AS platform,\n"
    "    tapsi_incentive_active_duration AS duration_bucket,\n"
    "    COUNT(*) AS n_range,\n"
    "    SUM(COUNT(*)) OVER (PARTITION BY weeknumber, city) AS n_total,\n"
    "    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY weeknumber, city),0) AS pct\n"
    "FROM [Cab].[DriverSurvey_ShortMain]\n"
    "WHERE city IS NOT NULL AND tapsi_incentive_active_duration IS NOT NULL\n"
    "GROUP BY yearweek, weeknumber, city, tapsi_incentive_active_duration"
)
new = (
    "SELECT " + FMT + "\n"
    "    weeknumber, city, 'Tapsi' AS platform,\n"
    "    tapsi_incentive_active_duration AS duration_bucket,\n"
    "    COUNT(*) AS n_range,\n"
    "    SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total,\n"
    "    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY yearweek, city),0) AS pct\n"
    "FROM [Cab].[DriverSurvey_ShortMain]\n"
    "WHERE city IS NOT NULL AND tapsi_incentive_active_duration IS NOT NULL\n"
    "GROUP BY yearweek, weeknumber, city, tapsi_incentive_active_duration"
)
assert old in sql, "IncentiveDuration Tapsi not found"
sql = sql.replace(old, new, 1)
changes.append("IncentiveDuration Tapsi branch")

# RA-7 Persona — 6 CTEs
for dim in ['Activity Type', 'Age Group', 'Education', 'Marital Status', 'Gender', 'Cooperation Type']:
    old = "    SELECT yearweek, weeknumber, city, '" + dim + "' AS dimension,"
    new = "    SELECT " + FMT + "\n        weeknumber, city, '" + dim + "' AS dimension,"
    assert old in sql, f"Persona CTE '{dim}' not found"
    sql = sql.replace(old, new, 1)
    changes.append(f"Persona CTE '{dim}'")

# RA-7 Persona — also fix PARTITION BY weeknumber,city in the n_total window
sql = sql.replace(
    "COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY weeknumber, city) AS n_total\n    FROM [Cab].[DriverSurvey_ShortMain]",
    "COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total\n    FROM [Cab].[DriverSurvey_ShortMain]"
)
changes.append("Persona PARTITION BY fix (all 6 CTEs)")

# RA-8 CommFree — Snapp branch
old = "SELECT yearweek, weeknumber, city, 'Snapp' AS platform,\n    COUNT(*) AS n,\n    SUM(CASE WHEN snapp_gotmessage"
new = "SELECT " + FMT + "\n    weeknumber, city, 'Snapp' AS platform,\n    COUNT(*) AS n,\n    SUM(CASE WHEN snapp_gotmessage"
assert old in sql, "CommFree Snapp not found"
sql = sql.replace(old, new, 1)
changes.append("CommFree Snapp branch")

# RA-8 CommFree — Tapsi branch
old = "SELECT yearweek, weeknumber, city, 'Tapsi' AS platform,\n    SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END) AS n,"
new = "SELECT " + FMT + "\n    weeknumber, city, 'Tapsi' AS platform,\n    SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END) AS n,"
assert old in sql, "CommFree Tapsi not found"
sql = sql.replace(old, new, 1)
changes.append("CommFree Tapsi branch")

# RA-9 CSRare
old = "    sm.yearweek, sm.weeknumber, sm.city,\n    COUNT(*) AS n,\n    AVG(TRY_CAST(sr.snapp_CS_satisfaction_overall"
new = "    " + FMT_SM + "\n    sm.weeknumber, sm.city,\n    COUNT(*) AS n,\n    AVG(TRY_CAST(sr.snapp_CS_satisfaction_overall"
assert old in sql, "CSRare not found"
sql = sql.replace(old, new, 1)
changes.append("CSRare SELECT")

# RA-10 NavReco
old = "    sm.yearweek, sm.weeknumber, sm.city,\n    COUNT(*) AS n,\n    AVG(TRY_CAST(sr.snapp_recommend"
new = "    " + FMT_SM + "\n    sm.weeknumber, sm.city,\n    COUNT(*) AS n,\n    AVG(TRY_CAST(sr.snapp_recommend"
assert old in sql, "NavReco not found"
sql = sql.replace(old, new, 1)
changes.append("NavReco SELECT")

with open(path, "w", encoding="utf-8") as f:
    f.write(sql)

print(f"Applied {len(changes)} changes:")
for c in changes:
    print(f"  OK {c}")

"""Add _sort columns to views with categorical string columns used in Power BI matrix visuals."""

path = r"D:\Work\Driver Survey\PowerBI\create_views.sql"

with open(path, "r", encoding="utf-8") as f:
    sql = f.read()

changes = []

# ── RANGE SORT CASE (shared between Snapp & Tapsi in IncentiveAmounts) ─────────
RANGE_SORT = """\
    CASE incentive_range
        WHEN '<20k'       THEN 1
        WHEN '20_40k'     THEN 2
        WHEN '40_60k'     THEN 3
        WHEN '<50k'       THEN 4
        WHEN '50_100k'    THEN 5
        WHEN '50_200k'    THEN 6
        WHEN '50_250k'    THEN 7
        WHEN '60_80k'     THEN 8
        WHEN '80_100k'    THEN 9
        WHEN '< 100k'     THEN 10
        WHEN '100_150k'   THEN 11
        WHEN '100_200k'   THEN 12
        WHEN '100_250k'   THEN 13
        WHEN '150_200k'   THEN 14
        WHEN '200_300k'   THEN 15
        WHEN '200_400k'   THEN 16
        WHEN '250_500k'   THEN 17
        WHEN '300_500k'   THEN 18
        WHEN '400_600k'   THEN 19
        WHEN '500_750k'   THEN 20
        WHEN '>500k'      THEN 21
        WHEN '600_800k'   THEN 22
        WHEN '750k_1m'    THEN 23
        WHEN '800k_1m'    THEN 24
        WHEN '1m_1.25m'   THEN 25
        WHEN '>1m'        THEN 26
        WHEN '1.25m_1.5m' THEN 27
        WHEN '>1.5m'      THEN 28
        ELSE 99
    END AS incentive_range_sort,"""

DUR_SORT = """\
    CASE duration_bucket
        WHEN 'Few Hours' THEN 1
        WHEN '1 Day'     THEN 2
        WHEN '1_6 Days'  THEN 3
        WHEN '7 Days'    THEN 4
        WHEN '>7 Days'   THEN 5
        ELSE 99
    END AS duration_bucket_sort,"""

YW_FMT = "CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,"

# ─────────────────────────────────────────────────────────────────────────────
# 1. vw_RA_IncentiveAmounts — restructure into CTE + outer SELECT
# ─────────────────────────────────────────────────────────────────────────────
old_ia = """\
CREATE OR ALTER VIEW [Cab].[vw_RA_IncentiveAmounts] AS
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
    weeknumber, city, 'Snapp' AS platform,
    snapp_incentive_rial_details AS incentive_range,
    COUNT(*) AS n_range,
    SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total,
    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY yearweek, city),0) AS pct
FROM [Cab].[DriverSurvey_ShortMain]
WHERE city IS NOT NULL AND snapp_incentive_rial_details IS NOT NULL
GROUP BY yearweek, weeknumber, city, snapp_incentive_rial_details
UNION ALL
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
    weeknumber, city, 'Tapsi' AS platform,
    tapsi_incentive_rial_details AS incentive_range,
    COUNT(*) AS n_range,
    SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total,
    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY yearweek, city),0) AS pct
FROM [Cab].[DriverSurvey_ShortMain]
WHERE city IS NOT NULL AND tapsi_incentive_rial_details IS NOT NULL
  AND TRY_CAST(active_joint AS INT) = 1
GROUP BY yearweek, weeknumber, city, tapsi_incentive_rial_details;"""

new_ia = """\
CREATE OR ALTER VIEW [Cab].[vw_RA_IncentiveAmounts] AS
WITH raw AS (
    SELECT yearweek, weeknumber, city, 'Snapp' AS platform,
        snapp_incentive_rial_details AS incentive_range,
        COUNT(*) AS n_range,
        SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND snapp_incentive_rial_details IS NOT NULL
    GROUP BY yearweek, weeknumber, city, snapp_incentive_rial_details
    UNION ALL
    SELECT yearweek, weeknumber, city, 'Tapsi' AS platform,
        tapsi_incentive_rial_details AS incentive_range,
        COUNT(*) AS n_range,
        SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND tapsi_incentive_rial_details IS NOT NULL
      AND TRY_CAST(active_joint AS INT) = 1
    GROUP BY yearweek, weeknumber, city, tapsi_incentive_rial_details
)
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    weeknumber, city, platform, incentive_range,
""" + RANGE_SORT + """
    n_range, n_total,
    100.0 * n_range / NULLIF(n_total, 0) AS pct
FROM raw;"""

assert old_ia in sql, "IncentiveAmounts block not found"
sql = sql.replace(old_ia, new_ia, 1)
changes.append("vw_RA_IncentiveAmounts: restructured to CTE + incentive_range_sort")

# ─────────────────────────────────────────────────────────────────────────────
# 2. vw_RA_IncentiveDuration — restructure into CTE + outer SELECT
# ─────────────────────────────────────────────────────────────────────────────
old_id = """\
CREATE OR ALTER VIEW [Cab].[vw_RA_IncentiveDuration] AS
SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
    weeknumber, city, 'Snapp' AS platform,
    snapp_incentive_active_duration AS duration_bucket,
    COUNT(*) AS n_range,
    SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total,
    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY yearweek, city),0) AS pct
FROM [Cab].[DriverSurvey_ShortMain]
WHERE city IS NOT NULL AND snapp_incentive_active_duration IS NOT NULL
GROUP BY yearweek, weeknumber, city, snapp_incentive_active_duration
UNION ALL
SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
    weeknumber, city, 'Tapsi' AS platform,
    tapsi_incentive_active_duration AS duration_bucket,
    COUNT(*) AS n_range,
    SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total,
    100.0 * COUNT(*) / NULLIF(SUM(COUNT(*)) OVER (PARTITION BY yearweek, city),0) AS pct
FROM [Cab].[DriverSurvey_ShortMain]
WHERE city IS NOT NULL AND tapsi_incentive_active_duration IS NOT NULL
GROUP BY yearweek, weeknumber, city, tapsi_incentive_active_duration;"""

new_id = """\
CREATE OR ALTER VIEW [Cab].[vw_RA_IncentiveDuration] AS
WITH raw AS (
    SELECT yearweek, weeknumber, city, 'Snapp' AS platform,
        snapp_incentive_active_duration AS duration_bucket,
        COUNT(*) AS n_range,
        SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND snapp_incentive_active_duration IS NOT NULL
    GROUP BY yearweek, weeknumber, city, snapp_incentive_active_duration
    UNION ALL
    SELECT yearweek, weeknumber, city, 'Tapsi' AS platform,
        tapsi_incentive_active_duration AS duration_bucket,
        COUNT(*) AS n_range,
        SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND tapsi_incentive_active_duration IS NOT NULL
    GROUP BY yearweek, weeknumber, city, tapsi_incentive_active_duration
)
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    weeknumber, city, platform, duration_bucket,
""" + DUR_SORT + """
    n_range, n_total,
    100.0 * n_range / NULLIF(n_total, 0) AS pct
FROM raw;"""

assert old_id in sql, "IncentiveDuration block not found"
sql = sql.replace(old_id, new_id, 1)
changes.append("vw_RA_IncentiveDuration: restructured to CTE + duration_bucket_sort")

# ─────────────────────────────────────────────────────────────────────────────
# 3. vw_RA_Persona — add category_sort to each of the 6 CTEs
# ─────────────────────────────────────────────────────────────────────────────

# 3a. activity CTE
old = """\
        CAST(active_time AS NVARCHAR(100)) AS category,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND active_time IS NOT NULL
    GROUP BY yearweek, weeknumber, city, active_time),"""
new = """\
        CAST(active_time AS NVARCHAR(100)) AS category,
        CASE active_time
            WHEN 'few hours/month' THEN 1
            WHEN '<20hour/mo'      THEN 2
            WHEN '5_20hour/week'   THEN 3
            WHEN '20_40h/week'     THEN 4
            WHEN '>40h/week'       THEN 5
            WHEN '8_12hour/day'    THEN 6
            WHEN '>12h/day'        THEN 7
            ELSE 99
        END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND active_time IS NOT NULL
    GROUP BY yearweek, weeknumber, city, active_time),"""
assert old in sql, "Persona activity CTE not found"
sql = sql.replace(old, new, 1)
changes.append("Persona activity CTE: category_sort (Activity Type)")

# 3b. age_grp CTE
old = """\
        CAST(age_group AS NVARCHAR(100)) AS category,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND age_group IS NOT NULL
    GROUP BY yearweek, weeknumber, city, age_group),"""
new = """\
        CAST(age_group AS NVARCHAR(100)) AS category,
        CASE age_group
            WHEN '18_to_35'     THEN 1
            WHEN 'more_than_35' THEN 2
            ELSE 99
        END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND age_group IS NOT NULL
    GROUP BY yearweek, weeknumber, city, age_group),"""
assert old in sql, "Persona age_grp CTE not found"
sql = sql.replace(old, new, 1)
changes.append("Persona age_grp CTE: category_sort (Age Group)")

# 3c. edu CTE
old = """\
        CAST(edu AS NVARCHAR(100)) AS category,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND edu IS NOT NULL
    GROUP BY yearweek, weeknumber, city, edu),"""
new = """\
        CAST(edu AS NVARCHAR(100)) AS category,
        CASE TRY_CAST(edu AS INT) WHEN 0 THEN 1 WHEN 1 THEN 2 ELSE 99 END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND edu IS NOT NULL
    GROUP BY yearweek, weeknumber, city, edu),"""
assert old in sql, "Persona edu CTE not found"
sql = sql.replace(old, new, 1)
changes.append("Persona edu CTE: category_sort (Education)")

# 3d. marr CTE
old = """\
        CAST(marr_stat AS NVARCHAR(100)) AS category,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND marr_stat IS NOT NULL
    GROUP BY yearweek, weeknumber, city, marr_stat),"""
new = """\
        CAST(marr_stat AS NVARCHAR(100)) AS category,
        CASE TRY_CAST(marr_stat AS INT) WHEN 0 THEN 1 WHEN 1 THEN 2 ELSE 99 END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND marr_stat IS NOT NULL
    GROUP BY yearweek, weeknumber, city, marr_stat),"""
assert old in sql, "Persona marr CTE not found"
sql = sql.replace(old, new, 1)
changes.append("Persona marr CTE: category_sort (Marital Status)")

# 3e. gen CTE
old = """\
        CAST(gender AS NVARCHAR(100)) AS category,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND gender IS NOT NULL
    GROUP BY yearweek, weeknumber, city, gender),"""
new = """\
        CAST(gender AS NVARCHAR(100)) AS category,
        CASE gender WHEN 'Female' THEN 1 WHEN 'Male' THEN 2 ELSE 99 END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND gender IS NOT NULL
    GROUP BY yearweek, weeknumber, city, gender),"""
assert old in sql, "Persona gen CTE not found"
sql = sql.replace(old, new, 1)
changes.append("Persona gen CTE: category_sort (Gender)")

# 3f. coop CTE
old = """\
        CAST(cooperation_type AS NVARCHAR(100)) AS category,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND cooperation_type IS NOT NULL
    GROUP BY yearweek, weeknumber, city, cooperation_type)"""
new = """\
        CAST(cooperation_type AS NVARCHAR(100)) AS category,
        CASE cooperation_type WHEN 'Full-Time' THEN 1 WHEN 'Part-Time' THEN 2 ELSE 99 END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND cooperation_type IS NOT NULL
    GROUP BY yearweek, weeknumber, city, cooperation_type)"""
assert old in sql, "Persona coop CTE not found"
sql = sql.replace(old, new, 1)
changes.append("Persona coop CTE: category_sort (Cooperation Type)")

# ─────────────────────────────────────────────────────────────────────────────
with open(path, "w", encoding="utf-8") as f:
    f.write(sql)

print(f"Applied {len(changes)} changes:")
for c in changes:
    print(f"  OK  {c}")

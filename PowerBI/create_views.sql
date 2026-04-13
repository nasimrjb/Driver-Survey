-- ============================================================
-- Driver Survey SQL Views for Power BI
-- Database: Cab_Studies | Schema: Cab
-- Generated from survey_analysis_v7, survey_routine_analysis,
-- trend_insights, and survey_metrics_summary
-- ============================================================

USE [Cab_Studies];
GO

-- ============================================================
-- HELPER: Computed columns as a base view
-- Adds driver_type, numeric casts, yearweek
-- ============================================================
IF OBJECT_ID(N'Cab.vw_ShortBase', N'V') IS NOT NULL DROP VIEW Cab.vw_ShortBase;
GO
CREATE VIEW Cab.vw_ShortBase AS
SELECT
    -- === FROM ShortMain (m) ===
    CAST(m.recordID AS INT) AS recordID,
    TRY_CAST(m.[datetime] AS DATETIME) AS response_datetime,
    m.[_source_file],
    TRY_CAST(m.weeknumber AS INT) AS weeknumber,
    -- yearweek: e.g. 2501 for week 1 of 2025 (ISO 8601 year + week)
    -- Uses the stored weeknumber (now ISO week) and the ISO year which
    -- handles year-boundary correctly (e.g. Dec 31 may be ISO week 1 of next year).
    CASE WHEN TRY_CAST(m.[datetime] AS DATETIME) IS NOT NULL
              AND TRY_CAST(m.weeknumber AS INT) IS NOT NULL
         THEN CASE
              -- If ISO week >= 52 but month is January, the ISO year is previous year
              WHEN TRY_CAST(m.weeknumber AS INT) >= 52
                   AND MONTH(TRY_CAST(m.[datetime] AS DATETIME)) = 1
              THEN ((YEAR(TRY_CAST(m.[datetime] AS DATETIME)) - 1) % 100) * 100
                   + TRY_CAST(m.weeknumber AS INT)
              -- If ISO week = 1 but month is December, the ISO year is next year
              WHEN TRY_CAST(m.weeknumber AS INT) = 1
                   AND MONTH(TRY_CAST(m.[datetime] AS DATETIME)) = 12
              THEN ((YEAR(TRY_CAST(m.[datetime] AS DATETIME)) + 1) % 100) * 100
                   + TRY_CAST(m.weeknumber AS INT)
              ELSE (YEAR(TRY_CAST(m.[datetime] AS DATETIME)) % 100) * 100
                   + TRY_CAST(m.weeknumber AS INT)
              END
         ELSE NULL END AS yearweek,
    -- driver_type
    CASE WHEN TRY_CAST(m.tapsi_ride AS FLOAT) = 0 OR m.tapsi_ride IS NULL OR m.tapsi_ride = ''
         THEN 'Snapp Exclusive' ELSE 'Joint' END AS driver_type,
    -- demographics
    m.city,
    m.gender,
    m.age,
    m.age_group,
    m.education,
    m.marital_status,
    m.original_job,
    m.active_time,
    m.cooperation_type,
    -- platform tenure
    m.snapp_age,
    m.tapsi_age,
    -- joint/signup flags
    TRY_CAST(m.joint_by_signup AS INT) AS joint_by_signup,
    TRY_CAST(m.active_joint AS INT) AS active_joint,
    -- rides
    TRY_CAST(m.snapp_ride AS FLOAT) AS snapp_ride,
    TRY_CAST(m.tapsi_ride AS FLOAT) AS tapsi_ride,
    -- satisfaction scores (1-5) from ShortMain
    TRY_CAST(m.snapp_fare_satisfaction AS INT) AS snapp_fare_satisfaction,
    TRY_CAST(m.snapp_req_count_satisfaction AS INT) AS snapp_req_count_satisfaction,
    TRY_CAST(m.snapp_income_satisfaction AS INT) AS snapp_income_satisfaction,
    TRY_CAST(m.tapsi_fare_satisfaction AS INT) AS tapsi_fare_satisfaction,
    TRY_CAST(m.tapsi_req_count_satisfaction AS INT) AS tapsi_req_count_satisfaction,
    TRY_CAST(m.tapsi_income_satisfaction AS INT) AS tapsi_income_satisfaction,
    -- overall satisfaction: snapp from ShortMain, tapsi from ShortRare
    TRY_CAST(m.snapp_overall_satisfaction_snapp AS INT) AS snapp_overall_satisfaction,
    TRY_CAST(r.tapsi_overall_satisfaction_tapsi AS INT) AS tapsi_overall_satisfaction,
    -- NPS / recommendation (0-10) from ShortRare
    TRY_CAST(r.snapp_recommend AS INT) AS snapp_recommend,
    TRY_CAST(r.tapsidriver_tapsi_recommend AS INT) AS tapsi_recommend,
    TRY_CAST(r.snapp_refer_others AS INT) AS snapp_refer_others,
    TRY_CAST(r.tapsi_refer_others AS INT) AS tapsi_refer_others,
    -- incentive from ShortMain
    TRY_CAST(m.snapp_incentive AS FLOAT) AS snapp_incentive,
    TRY_CAST(m.tapsi_incentive AS FLOAT) AS tapsi_incentive,
    TRY_CAST(m.snapp_overall_incentive_satisfaction AS INT) AS snapp_incentive_satisfaction,
    TRY_CAST(m.tapsi_overall_incentive_satisfaction AS INT) AS tapsi_incentive_satisfaction,
    m.snapp_gotmessage_text_incentive,
    m.tapsi_gotmessage_text_incentive,
    m.snapp_incentive_participation,
    m.tapsi_incentive_participation,
    m.snapp_incentive_rial_details,
    m.tapsi_incentive_rial_details,
    m.snapp_incentive_active_duration,
    m.tapsi_incentive_active_duration,
    m.snapp_joining_bonus,
    m.tapsi_joining_bonus,
    TRY_CAST(m.snapp_incentive_category AS FLOAT) AS snapp_incentive_category,
    TRY_CAST(m.tapsi_incentive_category AS FLOAT) AS tapsi_incentive_category,
    -- commission-free
    TRY_CAST(m.snapp_commfree AS FLOAT) AS snapp_commfree,
    TRY_CAST(m.tapsi_commfree AS FLOAT) AS tapsi_commfree,
    TRY_CAST(m.snapp_commfree_disc_ride AS FLOAT) AS snapp_commfree_disc_ride,
    TRY_CAST(m.tapsi_commfree_disc_ride AS FLOAT) AS tapsi_commfree_disc_ride,
    -- LOC
    TRY_CAST(m.snapp_LOC AS FLOAT) AS snapp_LOC,
    TRY_CAST(m.tapsi_LOC AS FLOAT) AS tapsi_LOC,
    -- wheel
    TRY_CAST(m.wheel AS FLOAT) AS wheel,
    -- navigation
    m.snapp_last_trip_navigation,
    m.tapsi_navigation_type,
    -- === FROM ShortRare (r) ===
    -- CS satisfaction
    TRY_CAST(r.snapp_CS_satisfaction_overall AS INT) AS snapp_CS_satisfaction_overall,
    TRY_CAST(r.snapp_CS_satisfaction_waittime AS INT) AS snapp_CS_satisfaction_waittime,
    TRY_CAST(r.snapp_CS_satisfaction_solution AS INT) AS snapp_CS_satisfaction_solution,
    TRY_CAST(r.snapp_CS_satisfaction_behaviour AS INT) AS snapp_CS_satisfaction_behaviour,
    TRY_CAST(r.snapp_CS_satisfaction_relevance AS INT) AS snapp_CS_satisfaction_relevance,
    TRY_CAST(r.tapsi_CS_satisfaction_overall AS INT) AS tapsi_CS_satisfaction_overall,
    TRY_CAST(r.tapsi_CS_satisfaction_waittime AS INT) AS tapsi_CS_satisfaction_waittime,
    TRY_CAST(r.tapsi_CS_satisfaction_solution AS INT) AS tapsi_CS_satisfaction_solution,
    TRY_CAST(r.tapsi_CS_satisfaction_behaviour AS INT) AS tapsi_CS_satisfaction_behaviour,
    TRY_CAST(r.tapsi_CS_satisfaction_relevance AS INT) AS tapsi_CS_satisfaction_relevance,
    -- Speed satisfaction
    TRY_CAST(r.snapp_speed_satisfaction AS INT) AS snapp_speed_satisfaction,
    TRY_CAST(r.tapsi_speed_satisfaction AS INT) AS tapsi_speed_satisfaction,
    -- Navigation recommendation scores (0-10)
    TRY_CAST(r.recommendation_googlemap AS INT) AS recommendation_googlemap,
    TRY_CAST(r.recommendation_waze AS INT) AS recommendation_waze,
    TRY_CAST(r.recommendation_neshan AS INT) AS recommendation_neshan,
    TRY_CAST(r.recommendation_balad AS INT) AS recommendation_balad,
    TRY_CAST(r.snapp_navigation_app_satisfaction AS INT) AS snapp_nav_app_satisfaction,
    TRY_CAST(r.tapsi_in_app_navigation_satisfaction AS INT) AS tapsi_nav_app_satisfaction,
    -- GPS
    r.gps_problem,
    r.snapp_gps_stage,
    r.tapsi_gps_stage,
    r.fixlocation_familiar,
    r.fixlocation_use,
    TRY_CAST(r.fixlocation_satisfaction AS INT) AS fixlocation_satisfaction,
    -- SnappCarFix / TapsiGarage
    r.snappcarfix_familiar,
    r.snappcarfix_use_ever,
    r.snappcarfix_use_lastmo,
    TRY_CAST(r.snappcarfix_satisfaction_overall AS INT) AS snappcarfix_sat_overall,
    TRY_CAST(r.snappcarfix_recommend AS INT) AS snappcarfix_recommend,
    r.tapsigarage_familiar,
    r.tapsigarage_use_ever,
    r.tapsigarage_use_lastmo,
    TRY_CAST(r.tapsigarage_satisfaction_overall AS INT) AS tapsigarage_sat_overall,
    TRY_CAST(r.tapsigarage_recommend AS INT) AS tapsigarage_recommend,
    -- edu/marr numeric
    TRY_CAST(m.edu AS FLOAT) AS edu_numeric,
    TRY_CAST(m.marr_stat AS FLOAT) AS marr_stat_numeric
FROM [Cab].[DriverSurvey_ShortMain] m
LEFT JOIN [Cab].[DriverSurvey_ShortRare] r ON CAST(m.recordID AS INT) = CAST(r.recordID AS INT)
;
GO


-- ============================================================
-- 1. WEEKLY SATISFACTION TRENDS
-- ============================================================
IF OBJECT_ID(N'Cab.vw_WeeklySatisfaction', N'V') IS NOT NULL DROP VIEW Cab.vw_WeeklySatisfaction;
GO
CREATE VIEW Cab.vw_WeeklySatisfaction AS
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    COUNT(*) AS response_count,
    -- Snapp satisfaction
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)) AS snapp_fare_sat_avg,
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)) AS snapp_req_sat_avg,
    AVG(CAST(snapp_income_satisfaction AS FLOAT)) AS snapp_income_sat_avg,
    AVG(CAST(snapp_overall_satisfaction AS FLOAT)) AS snapp_overall_sat_avg,
    -- Tapsi satisfaction
    AVG(CAST(tapsi_fare_satisfaction AS FLOAT)) AS tapsi_fare_sat_avg,
    AVG(CAST(tapsi_req_count_satisfaction AS FLOAT)) AS tapsi_req_sat_avg,
    AVG(CAST(tapsi_income_satisfaction AS FLOAT)) AS tapsi_income_sat_avg,
    AVG(CAST(tapsi_overall_satisfaction AS FLOAT)) AS tapsi_overall_sat_avg,
    -- Satisfaction gap (Snapp - Tapsi)
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)) - AVG(CAST(tapsi_fare_satisfaction AS FLOAT)) AS fare_sat_gap,
    AVG(CAST(snapp_income_satisfaction AS FLOAT)) - AVG(CAST(tapsi_income_satisfaction AS FLOAT)) AS income_sat_gap,
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)) - AVG(CAST(tapsi_req_count_satisfaction AS FLOAT)) AS req_sat_gap,
    -- Rides
    AVG(snapp_ride) AS snapp_ride_avg,
    AVG(tapsi_ride) AS tapsi_ride_avg,
    -- Joint driver rate
    AVG(CAST(active_joint AS FLOAT)) * 100 AS joint_driver_pct,
    -- Incentive
    AVG(snapp_incentive) AS snapp_incentive_avg,
    AVG(tapsi_incentive) AS tapsi_incentive_avg,
    -- Cooperation type
    SUM(CASE WHEN cooperation_type = 'Full-Time' THEN 1 ELSE 0 END) * 100.0 / COUNT(*) AS fulltime_pct
FROM Cab.vw_ShortBase
WHERE yearweek IS NOT NULL
GROUP BY yearweek
HAVING COUNT(*) >= 100
;
GO


-- ============================================================
-- 2. WEEKLY NPS SCORES
-- ============================================================
IF OBJECT_ID(N'Cab.vw_WeeklyNPS', N'V') IS NOT NULL DROP VIEW Cab.vw_WeeklyNPS;
GO
CREATE VIEW Cab.vw_WeeklyNPS AS
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    -- Snapp NPS
    COUNT(snapp_recommend) AS snapp_nps_n,
    SUM(CASE WHEN snapp_recommend >= 9 THEN 1 ELSE 0 END) * 100.0
        / NULLIF(COUNT(snapp_recommend), 0) AS snapp_promoter_pct,
    SUM(CASE WHEN snapp_recommend <= 6 THEN 1 ELSE 0 END) * 100.0
        / NULLIF(COUNT(snapp_recommend), 0) AS snapp_detractor_pct,
    SUM(CASE WHEN snapp_recommend BETWEEN 7 AND 8 THEN 1 ELSE 0 END) * 100.0
        / NULLIF(COUNT(snapp_recommend), 0) AS snapp_passive_pct,
    (SUM(CASE WHEN snapp_recommend >= 9 THEN 1 ELSE 0 END)
     - SUM(CASE WHEN snapp_recommend <= 6 THEN 1 ELSE 0 END))
     * 100.0 / NULLIF(COUNT(snapp_recommend), 0) AS snapp_nps,
    -- Tapsi NPS
    COUNT(tapsi_recommend) AS tapsi_nps_n,
    SUM(CASE WHEN tapsi_recommend >= 9 THEN 1 ELSE 0 END) * 100.0
        / NULLIF(COUNT(tapsi_recommend), 0) AS tapsi_promoter_pct,
    SUM(CASE WHEN tapsi_recommend <= 6 THEN 1 ELSE 0 END) * 100.0
        / NULLIF(COUNT(tapsi_recommend), 0) AS tapsi_detractor_pct,
    SUM(CASE WHEN tapsi_recommend BETWEEN 7 AND 8 THEN 1 ELSE 0 END) * 100.0
        / NULLIF(COUNT(tapsi_recommend), 0) AS tapsi_passive_pct,
    (SUM(CASE WHEN tapsi_recommend >= 9 THEN 1 ELSE 0 END)
     - SUM(CASE WHEN tapsi_recommend <= 6 THEN 1 ELSE 0 END))
     * 100.0 / NULLIF(COUNT(tapsi_recommend), 0) AS tapsi_nps
FROM Cab.vw_ShortBase
WHERE yearweek IS NOT NULL
GROUP BY yearweek
HAVING COUNT(*) >= 100
;
GO


-- ============================================================
-- 3. SATISFACTION BY CITY
-- ============================================================
IF OBJECT_ID(N'Cab.vw_SatisfactionByCity', N'V') IS NOT NULL DROP VIEW Cab.vw_SatisfactionByCity;
GO
CREATE VIEW Cab.vw_SatisfactionByCity AS
SELECT
    city,
    COUNT(*) AS n,
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)) AS snapp_fare_sat,
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)) AS snapp_req_sat,
    AVG(CAST(snapp_income_satisfaction AS FLOAT)) AS snapp_income_sat,
    AVG(CAST(snapp_overall_satisfaction AS FLOAT)) AS snapp_overall_sat,
    AVG(CAST(tapsi_fare_satisfaction AS FLOAT)) AS tapsi_fare_sat,
    AVG(CAST(tapsi_req_count_satisfaction AS FLOAT)) AS tapsi_req_sat,
    AVG(CAST(tapsi_income_satisfaction AS FLOAT)) AS tapsi_income_sat,
    AVG(CAST(tapsi_overall_satisfaction AS FLOAT)) AS tapsi_overall_sat,
    -- NPS by city
    (SUM(CASE WHEN snapp_recommend >= 9 THEN 1 ELSE 0 END)
     - SUM(CASE WHEN snapp_recommend <= 6 THEN 1 ELSE 0 END))
     * 100.0 / NULLIF(COUNT(snapp_recommend), 0) AS snapp_nps,
    (SUM(CASE WHEN tapsi_recommend >= 9 THEN 1 ELSE 0 END)
     - SUM(CASE WHEN tapsi_recommend <= 6 THEN 1 ELSE 0 END))
     * 100.0 / NULLIF(COUNT(tapsi_recommend), 0) AS tapsi_nps,
    -- Joint rate
    AVG(CAST(active_joint AS FLOAT)) * 100 AS joint_pct,
    -- Avg rides
    AVG(snapp_ride) AS snapp_ride_avg,
    AVG(tapsi_ride) AS tapsi_ride_avg,
    -- Avg LOC
    AVG(snapp_LOC) AS snapp_LOC_avg,
    AVG(tapsi_LOC) AS tapsi_LOC_avg,
    -- Incentive
    AVG(snapp_incentive) AS snapp_incentive_avg,
    -- Got message rate
    SUM(CASE WHEN snapp_gotmessage_text_incentive = 'Yes' THEN 1 ELSE 0 END) * 100.0
        / COUNT(*) AS snapp_gotmsg_pct,
    -- Full-time %
    SUM(CASE WHEN cooperation_type = 'Full-Time' THEN 1 ELSE 0 END) * 100.0
        / COUNT(*) AS fulltime_pct
FROM Cab.vw_ShortBase
WHERE city IS NOT NULL AND city != ''
GROUP BY city
HAVING COUNT(*) >= 20
;
GO


-- ============================================================
-- 4. SATISFACTION BY CITY AND WEEK (for heatmap)
-- ============================================================
IF OBJECT_ID(N'Cab.vw_SatisfactionByCityWeek', N'V') IS NOT NULL DROP VIEW Cab.vw_SatisfactionByCityWeek;
GO
CREATE VIEW Cab.vw_SatisfactionByCityWeek AS
SELECT
    city,
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    COUNT(*) AS n,
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)) AS snapp_fare_sat,
    AVG(CAST(snapp_income_satisfaction AS FLOAT)) AS snapp_income_sat,
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)) AS snapp_req_sat,
    AVG(CAST(snapp_overall_satisfaction AS FLOAT)) AS snapp_overall_sat,
    AVG(CAST(tapsi_fare_satisfaction AS FLOAT)) AS tapsi_fare_sat,
    AVG(CAST(tapsi_overall_satisfaction AS FLOAT)) AS tapsi_overall_sat,
    AVG(snapp_ride) AS snapp_ride_avg,
    AVG(tapsi_ride) AS tapsi_ride_avg,
    AVG(CAST(active_joint AS FLOAT)) * 100 AS joint_pct
FROM Cab.vw_ShortBase
WHERE city IS NOT NULL AND city != '' AND yearweek IS NOT NULL
GROUP BY city, yearweek
HAVING COUNT(*) >= 10
;
GO


-- ============================================================
-- 5. SATISFACTION BY DEMOGRAPHICS
-- ============================================================
IF OBJECT_ID(N'Cab.vw_SatisfactionByDemographics', N'V') IS NOT NULL DROP VIEW Cab.vw_SatisfactionByDemographics;
GO
CREATE VIEW Cab.vw_SatisfactionByDemographics AS
-- By age group
SELECT
    'age_group' AS dimension,
    age_group AS category,
    COUNT(*) AS n,
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)) AS snapp_fare_sat,
    AVG(CAST(snapp_income_satisfaction AS FLOAT)) AS snapp_income_sat,
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)) AS snapp_req_sat,
    AVG(CAST(snapp_overall_satisfaction AS FLOAT)) AS snapp_overall_sat,
    AVG(CAST(tapsi_fare_satisfaction AS FLOAT)) AS tapsi_fare_sat,
    AVG(CAST(tapsi_overall_satisfaction AS FLOAT)) AS tapsi_overall_sat,
    AVG(snapp_ride) AS snapp_ride_avg,
    AVG(tapsi_ride) AS tapsi_ride_avg
FROM Cab.vw_ShortBase
WHERE age_group IS NOT NULL AND age_group != ''
GROUP BY age_group
HAVING COUNT(*) >= 10

UNION ALL

-- By cooperation type
SELECT
    'cooperation_type', cooperation_type, COUNT(*),
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)),
    AVG(CAST(snapp_income_satisfaction AS FLOAT)),
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)),
    AVG(CAST(snapp_overall_satisfaction AS FLOAT)),
    AVG(CAST(tapsi_fare_satisfaction AS FLOAT)),
    AVG(CAST(tapsi_overall_satisfaction AS FLOAT)),
    AVG(snapp_ride), AVG(tapsi_ride)
FROM Cab.vw_ShortBase
WHERE cooperation_type IS NOT NULL AND cooperation_type != ''
GROUP BY cooperation_type
HAVING COUNT(*) >= 10

UNION ALL

-- By driver type (Joint vs Exclusive)
SELECT
    'driver_type', driver_type, COUNT(*),
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)),
    AVG(CAST(snapp_income_satisfaction AS FLOAT)),
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)),
    AVG(CAST(snapp_overall_satisfaction AS FLOAT)),
    AVG(CAST(tapsi_fare_satisfaction AS FLOAT)),
    AVG(CAST(tapsi_overall_satisfaction AS FLOAT)),
    AVG(snapp_ride), AVG(tapsi_ride)
FROM Cab.vw_ShortBase
GROUP BY driver_type
HAVING COUNT(*) >= 10

UNION ALL

-- By gender
SELECT
    'gender', gender, COUNT(*),
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)),
    AVG(CAST(snapp_income_satisfaction AS FLOAT)),
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)),
    AVG(CAST(snapp_overall_satisfaction AS FLOAT)),
    AVG(CAST(tapsi_fare_satisfaction AS FLOAT)),
    AVG(CAST(tapsi_overall_satisfaction AS FLOAT)),
    AVG(snapp_ride), AVG(tapsi_ride)
FROM Cab.vw_ShortBase
WHERE gender IS NOT NULL AND gender != ''
GROUP BY gender
HAVING COUNT(*) >= 10
;
GO


-- ============================================================
-- 6. HONEYMOON EFFECT (Satisfaction by Snapp Tenure)
-- ============================================================
IF OBJECT_ID(N'Cab.vw_HoneymoonEffect', N'V') IS NOT NULL DROP VIEW Cab.vw_HoneymoonEffect;
GO
CREATE VIEW Cab.vw_HoneymoonEffect AS
SELECT
    snapp_age AS tenure,
    COUNT(*) AS n,
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)) AS snapp_fare_sat,
    AVG(CAST(snapp_income_satisfaction AS FLOAT)) AS snapp_income_sat,
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)) AS snapp_req_sat,
    AVG(CAST(snapp_overall_satisfaction AS FLOAT)) AS snapp_overall_sat,
    AVG(CAST(snapp_recommend AS FLOAT)) AS snapp_recommend_avg,
    AVG(snapp_ride) AS snapp_ride_avg,
    AVG(snapp_LOC) AS snapp_LOC_avg
FROM Cab.vw_ShortBase
WHERE snapp_age IS NOT NULL AND snapp_age != ''
GROUP BY snapp_age
HAVING COUNT(*) >= 50
;
GO


-- ============================================================
-- 7. DEMOGRAPHICS OVERVIEW
-- ============================================================
IF OBJECT_ID(N'Cab.vw_Demographics', N'V') IS NOT NULL DROP VIEW Cab.vw_Demographics;
GO
CREATE VIEW Cab.vw_Demographics AS
-- Age
SELECT 'age' AS dimension, age AS category, COUNT(*) AS n
FROM Cab.vw_ShortBase WHERE age IS NOT NULL AND age != '' GROUP BY age
UNION ALL
-- Age group
SELECT 'age_group', age_group, COUNT(*)
FROM Cab.vw_ShortBase WHERE age_group IS NOT NULL AND age_group != '' GROUP BY age_group
UNION ALL
-- Gender
SELECT 'gender', gender, COUNT(*)
FROM Cab.vw_ShortBase WHERE gender IS NOT NULL AND gender != '' GROUP BY gender
UNION ALL
-- Cooperation type
SELECT 'cooperation_type', cooperation_type, COUNT(*)
FROM Cab.vw_ShortBase WHERE cooperation_type IS NOT NULL AND cooperation_type != '' GROUP BY cooperation_type
UNION ALL
-- Education
SELECT 'education', education, COUNT(*)
FROM Cab.vw_ShortBase WHERE education IS NOT NULL AND education != '' GROUP BY education
UNION ALL
-- Marital status
SELECT 'marital_status', marital_status, COUNT(*)
FROM Cab.vw_ShortBase WHERE marital_status IS NOT NULL AND marital_status != '' GROUP BY marital_status
UNION ALL
-- Top occupations
SELECT 'original_job', original_job, COUNT(*)
FROM Cab.vw_ShortBase WHERE original_job IS NOT NULL AND original_job != '' GROUP BY original_job
UNION ALL
-- Active time
SELECT 'active_time', active_time, COUNT(*)
FROM Cab.vw_ShortBase WHERE active_time IS NOT NULL AND active_time != '' GROUP BY active_time
UNION ALL
-- City
SELECT 'city', city, COUNT(*)
FROM Cab.vw_ShortBase WHERE city IS NOT NULL AND city != '' GROUP BY city
;
GO


-- ============================================================
-- 8. INCENTIVE ANALYSIS
-- ============================================================
IF OBJECT_ID(N'Cab.vw_IncentiveByWeek', N'V') IS NOT NULL DROP VIEW Cab.vw_IncentiveByWeek;
GO
CREATE VIEW Cab.vw_IncentiveByWeek AS
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    COUNT(*) AS n,
    -- Incentive amounts
    AVG(snapp_incentive) AS snapp_incentive_avg,
    AVG(snapp_incentive) / 1000000.0 AS snapp_incentive_avg_mrial,
    AVG(tapsi_incentive) AS tapsi_incentive_avg,
    -- Incentive satisfaction
    AVG(CAST(snapp_incentive_satisfaction AS FLOAT)) AS snapp_inc_sat_avg,
    AVG(CAST(tapsi_incentive_satisfaction AS FLOAT)) AS tapsi_inc_sat_avg,
    -- Got message rate
    SUM(CASE WHEN snapp_gotmessage_text_incentive = 'Yes' THEN 1 ELSE 0 END) * 100.0
        / COUNT(*) AS snapp_gotmsg_pct,
    SUM(CASE WHEN tapsi_gotmessage_text_incentive = 'Yes' THEN 1 ELSE 0 END) * 100.0
        / NULLIF(SUM(CAST(active_joint AS INT)), 0) AS tapsi_gotmsg_pct,
    -- Participation rate (among those who got message)
    SUM(CASE WHEN snapp_incentive_participation = 'Yes' THEN 1 ELSE 0 END) * 100.0
        / NULLIF(SUM(CASE WHEN snapp_gotmessage_text_incentive = 'Yes' THEN 1 ELSE 0 END), 0)
        AS snapp_participation_pct,
    -- Commission-free share
    SUM(snapp_commfree) AS snapp_commfree_total,
    SUM(snapp_ride) AS snapp_ride_total,
    CASE WHEN SUM(snapp_ride) > 0
         THEN SUM(snapp_commfree) * 100.0 / SUM(snapp_ride)
         ELSE NULL END AS snapp_commfree_pct,
    -- Wheel
    AVG(wheel) AS wheel_avg
FROM Cab.vw_ShortBase
WHERE yearweek IS NOT NULL
GROUP BY yearweek
HAVING COUNT(*) >= 100
;
GO


-- ============================================================
-- 9. INCENTIVE BY CITY (for routine analysis)
-- ============================================================
IF OBJECT_ID(N'Cab.vw_IncentiveByCity', N'V') IS NOT NULL DROP VIEW Cab.vw_IncentiveByCity;
GO
CREATE VIEW Cab.vw_IncentiveByCity AS
SELECT
    city,
    COUNT(*) AS n,
    SUM(CAST(active_joint AS INT)) AS n_joint,
    -- Snapp incentive
    AVG(snapp_incentive) AS snapp_incentive_avg,
    AVG(CAST(snapp_incentive_satisfaction AS FLOAT)) AS snapp_inc_sat_avg,
    SUM(CASE WHEN snapp_gotmessage_text_incentive = 'Yes' THEN 1 ELSE 0 END) * 100.0
        / COUNT(*) AS snapp_gotmsg_pct,
    SUM(CASE WHEN snapp_incentive_participation = 'Yes' THEN 1 ELSE 0 END) * 100.0
        / NULLIF(SUM(CASE WHEN snapp_gotmessage_text_incentive = 'Yes' THEN 1 ELSE 0 END), 0)
        AS snapp_participation_pct,
    -- Tapsi incentive (joint drivers only)
    AVG(CASE WHEN active_joint = 1 THEN tapsi_incentive END) AS tapsi_incentive_avg,
    AVG(CASE WHEN active_joint = 1 THEN CAST(tapsi_incentive_satisfaction AS FLOAT) END) AS tapsi_inc_sat_avg,
    -- Commission-free
    CASE WHEN SUM(snapp_ride) > 0
         THEN SUM(snapp_commfree) * 100.0 / SUM(snapp_ride) ELSE NULL END AS snapp_commfree_pct,
    -- Satisfaction
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)) AS snapp_fare_sat,
    AVG(CAST(snapp_income_satisfaction AS FLOAT)) AS snapp_income_sat,
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)) AS snapp_req_sat
FROM Cab.vw_ShortBase
WHERE city IS NOT NULL AND city != ''
GROUP BY city
HAVING COUNT(*) >= 5
;
GO


-- ============================================================
-- 10. RIDE SHARE ANALYSIS BY CITY AND WEEK
-- ============================================================
IF OBJECT_ID(N'Cab.vw_RideShareByCityWeek', N'V') IS NOT NULL DROP VIEW Cab.vw_RideShareByCityWeek;
GO
CREATE VIEW Cab.vw_RideShareByCityWeek AS
SELECT
    city,
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    COUNT(*) AS total_respondents,
    SUM(CAST(active_joint AS INT)) AS joint_respondents,
    SUM(CASE WHEN active_joint = 0 THEN 1 ELSE 0 END) AS exclusive_respondents,
    -- Total rides
    SUM(snapp_ride) + SUM(ISNULL(tapsi_ride, 0)) AS total_rides,
    SUM(snapp_ride) AS snapp_rides_total,
    SUM(CASE WHEN active_joint = 0 THEN snapp_ride ELSE 0 END) AS exclusive_snapp_rides,
    SUM(CASE WHEN active_joint = 1 THEN snapp_ride ELSE 0 END) AS joint_snapp_rides,
    SUM(CASE WHEN active_joint = 1 THEN ISNULL(tapsi_ride, 0) ELSE 0 END) AS joint_tapsi_rides,
    -- Ride share percentages
    CASE WHEN (SUM(snapp_ride) + SUM(ISNULL(tapsi_ride, 0))) > 0
         THEN SUM(snapp_ride) * 100.0 / (SUM(snapp_ride) + SUM(ISNULL(tapsi_ride, 0)))
         ELSE NULL END AS snapp_ride_share_pct,
    CASE WHEN (SUM(snapp_ride) + SUM(ISNULL(tapsi_ride, 0))) > 0
         THEN SUM(CASE WHEN active_joint = 1 THEN ISNULL(tapsi_ride, 0) ELSE 0 END)
              * 100.0 / (SUM(snapp_ride) + SUM(ISNULL(tapsi_ride, 0)))
         ELSE NULL END AS tapsi_ride_share_pct
FROM Cab.vw_ShortBase
WHERE city IS NOT NULL AND city != '' AND yearweek IS NOT NULL
GROUP BY city, yearweek
;
GO


-- ============================================================
-- 11. INCENTIVE AMOUNT DISTRIBUTION BY CITY
-- ============================================================
IF OBJECT_ID(N'Cab.vw_IncentiveAmountByCity', N'V') IS NOT NULL DROP VIEW Cab.vw_IncentiveAmountByCity;
GO
CREATE VIEW Cab.vw_IncentiveAmountByCity AS
SELECT
    s.city,
    s.snapp_incentive_rial_details AS amount_bracket,
    'Snapp' AS platform,
    COUNT(*) AS n
FROM Cab.vw_ShortBase s
WHERE s.snapp_incentive_rial_details IS NOT NULL AND s.snapp_incentive_rial_details != ''
      AND s.city IS NOT NULL AND s.city != ''
GROUP BY s.city, s.snapp_incentive_rial_details

UNION ALL

SELECT
    s.city,
    s.tapsi_incentive_rial_details,
    'Tapsi',
    COUNT(*)
FROM Cab.vw_ShortBase s
WHERE s.tapsi_incentive_rial_details IS NOT NULL AND s.tapsi_incentive_rial_details != ''
      AND s.city IS NOT NULL AND s.city != ''
      AND s.active_joint = 1
GROUP BY s.city, s.tapsi_incentive_rial_details
;
GO


-- ============================================================
-- 12. NAVIGATION APP USAGE
-- ============================================================
IF OBJECT_ID(N'Cab.vw_NavigationUsage', N'V') IS NOT NULL DROP VIEW Cab.vw_NavigationUsage;
GO
CREATE VIEW Cab.vw_NavigationUsage AS
SELECT
    'snapp_last_trip' AS context,
    snapp_last_trip_navigation AS nav_app,
    COUNT(*) AS n
FROM Cab.vw_ShortBase
WHERE snapp_last_trip_navigation IS NOT NULL AND snapp_last_trip_navigation != ''
GROUP BY snapp_last_trip_navigation

UNION ALL

SELECT
    'tapsi_default',
    tapsi_navigation_type,
    COUNT(*)
FROM Cab.vw_ShortBase
WHERE tapsi_navigation_type IS NOT NULL AND tapsi_navigation_type != ''
GROUP BY tapsi_navigation_type
;
GO


-- ============================================================
-- 13. NAVIGATION APP USAGE BY WEEK
-- ============================================================
IF OBJECT_ID(N'Cab.vw_NavigationByWeek', N'V') IS NOT NULL DROP VIEW Cab.vw_NavigationByWeek;
GO
CREATE VIEW Cab.vw_NavigationByWeek AS
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    snapp_last_trip_navigation AS nav_app,
    COUNT(*) AS n,
    COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (PARTITION BY yearweek) AS pct
FROM Cab.vw_ShortBase
WHERE snapp_last_trip_navigation IS NOT NULL AND snapp_last_trip_navigation != ''
      AND yearweek IS NOT NULL
GROUP BY yearweek, snapp_last_trip_navigation
;
GO


-- ============================================================
-- 14. WIDE SURVEY: INCENTIVE TYPE BINARY COUNTS
-- ============================================================
IF OBJECT_ID(N'Cab.vw_WideIncentiveTypes', N'V') IS NOT NULL DROP VIEW Cab.vw_WideIncentiveTypes;
GO
CREATE VIEW Cab.vw_WideIncentiveTypes AS
SELECT
    'Snapp' AS platform,
    reason,
    SUM(val) AS n,
    SUM(val) * 100.0 / COUNT(*) AS pct
FROM (
    SELECT recordID,
        'Commission Free on some trips' AS reason,
        TRY_CAST([Snapp Incentive Type__Commission Free on some trips] AS INT) AS val
    FROM [Cab].[DriverSurvey_WideMain]
    WHERE [Snapp Incentive Type__Commission Free on some trips] IS NOT NULL
       AND [Snapp Incentive Type__Commission Free on some trips] != ''
    UNION ALL
    SELECT recordID, 'Pay After Ride',
        TRY_CAST([Snapp Incentive Type__Pay After Ride] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    WHERE [Snapp Incentive Type__Pay After Ride] IS NOT NULL
       AND [Snapp Incentive Type__Pay After Ride] != ''
    UNION ALL
    SELECT recordID, 'Ride-Based Commission-free',
        TRY_CAST([Snapp Incentive Type__Ride-Based Commission-free] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    WHERE [Snapp Incentive Type__Ride-Based Commission-free] IS NOT NULL
       AND [Snapp Incentive Type__Ride-Based Commission-free] != ''
    UNION ALL
    SELECT recordID, 'Earning-based Commission-free',
        TRY_CAST([Snapp Incentive Type__Earning-based Commission-free] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    WHERE [Snapp Incentive Type__Earning-based Commission-free] IS NOT NULL
       AND [Snapp Incentive Type__Earning-based Commission-free] != ''
    UNION ALL
    SELECT recordID, 'Income Guarantee',
        TRY_CAST([Snapp Incentive Type__Income Guarantee] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    WHERE [Snapp Incentive Type__Income Guarantee] IS NOT NULL
       AND [Snapp Incentive Type__Income Guarantee] != ''
    UNION ALL
    SELECT recordID, 'Pay After Income',
        TRY_CAST([Snapp Incentive Type__Pay After Income] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    WHERE [Snapp Incentive Type__Pay After Income] IS NOT NULL
       AND [Snapp Incentive Type__Pay After Income] != ''
) t
GROUP BY reason
;
GO


-- ============================================================
-- 15. WIDE SURVEY: UNSATISFACTION REASONS
-- ============================================================
IF OBJECT_ID(N'Cab.vw_WideUnsatisfactionReasons', N'V') IS NOT NULL DROP VIEW Cab.vw_WideUnsatisfactionReasons;
GO
CREATE VIEW Cab.vw_WideUnsatisfactionReasons AS
SELECT platform, reason, SUM(val) AS n
FROM (
    SELECT 'Snapp' AS platform, 'Not Available' AS reason,
        TRY_CAST([Snapp Last Incentive Unsatisfaction__Not Available] AS INT) AS val
    FROM [Cab].[DriverSurvey_WideMain]
    UNION ALL
    SELECT 'Snapp', 'Improper Amount',
        TRY_CAST([Snapp Last Incentive Unsatisfaction__Improper Amount] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    UNION ALL
    SELECT 'Snapp', 'No Time todo',
        TRY_CAST([Snapp Last Incentive Unsatisfaction__No Time todo] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    UNION ALL
    SELECT 'Snapp', 'Difficult',
        TRY_CAST([Snapp Last Incentive Unsatisfaction__difficult] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    UNION ALL
    SELECT 'Snapp', 'Non Payment',
        TRY_CAST([Snapp Last Incentive Unsatisfaction__Non Payment] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    UNION ALL
    SELECT 'Tapsi', 'Not Available',
        TRY_CAST([Tapsi Incentive Unsatisfaction__Not Available] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    UNION ALL
    SELECT 'Tapsi', 'Improper Amount',
        TRY_CAST([Tapsi Incentive Unsatisfaction__Improper Amount] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    UNION ALL
    SELECT 'Tapsi', 'No Time todo',
        TRY_CAST([Tapsi Incentive Unsatisfaction__No Time todo] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    UNION ALL
    SELECT 'Tapsi', 'Difficult',
        TRY_CAST([Tapsi Incentive Unsatisfaction__difficult] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
    UNION ALL
    SELECT 'Tapsi', 'Non Payment',
        TRY_CAST([Tapsi Incentive Unsatisfaction__Non Payment] AS INT)
    FROM [Cab].[DriverSurvey_WideMain]
) t
GROUP BY platform, reason
;
GO


-- ============================================================
-- 16. LONG SURVEY: QUESTION-ANSWER DISTRIBUTIONS
-- ============================================================
IF OBJECT_ID(N'Cab.vw_LongSurveyAnswers', N'V') IS NOT NULL DROP VIEW Cab.vw_LongSurveyAnswers;
GO
CREATE VIEW Cab.vw_LongSurveyAnswers AS
SELECT
    question,
    answer,
    COUNT(*) AS n,
    COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (PARTITION BY question) AS pct
FROM [Cab].[DriverSurvey_LongMain]
WHERE answer IS NOT NULL AND answer != ''
GROUP BY question, answer
;
GO


-- ============================================================
-- 17. LONG SURVEY: RARE QUESTION-ANSWER DISTRIBUTIONS
-- ============================================================
IF OBJECT_ID(N'Cab.vw_LongRareSurveyAnswers', N'V') IS NOT NULL DROP VIEW Cab.vw_LongRareSurveyAnswers;
GO
CREATE VIEW Cab.vw_LongRareSurveyAnswers AS
SELECT
    question,
    answer,
    COUNT(*) AS n,
    COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (PARTITION BY question) AS pct
FROM [Cab].[DriverSurvey_LongRare]
WHERE answer IS NOT NULL AND answer != ''
GROUP BY question, answer
;
GO


-- ============================================================
-- 18. LONG SURVEY BY CITY (for refusal reasons, CS categories)
-- ============================================================
IF OBJECT_ID(N'Cab.vw_LongSurveyByCity', N'V') IS NOT NULL DROP VIEW Cab.vw_LongSurveyByCity;
GO
CREATE VIEW Cab.vw_LongSurveyByCity AS
SELECT
    lr.question,
    lr.answer,
    sm.city,
    COUNT(*) AS n
FROM [Cab].[DriverSurvey_LongRare] lr
INNER JOIN (
    SELECT DISTINCT recordID, city
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND city != ''
) sm ON lr.recordID = sm.recordID
WHERE lr.answer IS NOT NULL AND lr.answer != ''
GROUP BY lr.question, lr.answer, sm.city
;
GO


-- ============================================================
-- 19. KPI SUMMARY (single-row overview)
-- ============================================================
IF OBJECT_ID(N'Cab.vw_KPISummary', N'V') IS NOT NULL DROP VIEW Cab.vw_KPISummary;
GO
CREATE VIEW Cab.vw_KPISummary AS
SELECT
    COUNT(*) AS total_responses,
    COUNT(DISTINCT yearweek) AS survey_weeks,
    COUNT(DISTINCT city) AS cities,
    -- Joint driver %
    AVG(CAST(active_joint AS FLOAT)) * 100 AS joint_driver_pct,
    -- Full-time %
    SUM(CASE WHEN cooperation_type = 'Full-Time' THEN 1 ELSE 0 END) * 100.0 / COUNT(*) AS fulltime_pct,
    -- Male %
    SUM(CASE WHEN gender = 'Male' THEN 1 ELSE 0 END) * 100.0
        / NULLIF(SUM(CASE WHEN gender IS NOT NULL AND gender != '' THEN 1 ELSE 0 END), 0) AS male_pct,
    -- Under 35 %
    SUM(CASE WHEN age_group = '18_to_35' THEN 1 ELSE 0 END) * 100.0
        / NULLIF(SUM(CASE WHEN age_group IS NOT NULL AND age_group != '' THEN 1 ELSE 0 END), 0) AS under35_pct,
    -- Satisfaction
    AVG(CAST(snapp_fare_satisfaction AS FLOAT)) AS snapp_fare_sat,
    AVG(CAST(snapp_income_satisfaction AS FLOAT)) AS snapp_income_sat,
    AVG(CAST(snapp_req_count_satisfaction AS FLOAT)) AS snapp_req_sat,
    AVG(CAST(tapsi_fare_satisfaction AS FLOAT)) AS tapsi_fare_sat,
    AVG(CAST(tapsi_income_satisfaction AS FLOAT)) AS tapsi_income_sat,
    AVG(CAST(tapsi_req_count_satisfaction AS FLOAT)) AS tapsi_req_sat,
    -- CS satisfaction
    AVG(CAST(snapp_CS_satisfaction_overall AS FLOAT)) AS snapp_cs_sat,
    AVG(CAST(tapsi_CS_satisfaction_overall AS FLOAT)) AS tapsi_cs_sat,
    -- NPS
    (SUM(CASE WHEN snapp_recommend >= 9 THEN 1 ELSE 0 END)
     - SUM(CASE WHEN snapp_recommend <= 6 THEN 1 ELSE 0 END))
     * 100.0 / NULLIF(COUNT(snapp_recommend), 0) AS snapp_nps,
    (SUM(CASE WHEN tapsi_recommend >= 9 THEN 1 ELSE 0 END)
     - SUM(CASE WHEN tapsi_recommend <= 6 THEN 1 ELSE 0 END))
     * 100.0 / NULLIF(COUNT(tapsi_recommend), 0) AS tapsi_nps,
    -- Avg incentive (M Rials)
    AVG(snapp_incentive) / 1000000.0 AS snapp_incentive_avg_mrial,
    AVG(tapsi_incentive) / 1000000.0 AS tapsi_incentive_avg_mrial,
    -- Avg rides
    AVG(snapp_ride) AS snapp_ride_avg,
    AVG(tapsi_ride) AS tapsi_ride_avg
FROM Cab.vw_ShortBase
;
GO


-- ============================================================
-- 20. PERSONA BY CITY (cooperation, age, gender distributions)
-- ============================================================
IF OBJECT_ID(N'Cab.vw_PersonaByCity', N'V') IS NOT NULL DROP VIEW Cab.vw_PersonaByCity;
GO
CREATE VIEW Cab.vw_PersonaByCity AS
SELECT
    city,
    dimension,
    category,
    COUNT(*) AS n,
    COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (PARTITION BY city, dimension) AS pct
FROM (
    SELECT city, 'cooperation_type' AS dimension, cooperation_type AS category
    FROM Cab.vw_ShortBase WHERE cooperation_type IS NOT NULL AND cooperation_type != ''
    UNION ALL
    SELECT city, 'age_group', age_group
    FROM Cab.vw_ShortBase WHERE age_group IS NOT NULL AND age_group != ''
    UNION ALL
    SELECT city, 'gender', gender
    FROM Cab.vw_ShortBase WHERE gender IS NOT NULL AND gender != ''
    UNION ALL
    SELECT city, 'active_time', active_time
    FROM Cab.vw_ShortBase WHERE active_time IS NOT NULL AND active_time != ''
) t
WHERE city IS NOT NULL AND city != ''
GROUP BY city, dimension, category
;
GO


PRINT 'All 20 views created successfully!';
GO

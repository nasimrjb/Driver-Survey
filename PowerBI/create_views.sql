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
    CAST(m.yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(m.yearweek%100 AS VARCHAR), 2) AS yearweek,
    m.yearweek AS yearweek_sort,
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
    yearweek,
    yearweek_sort,
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
GROUP BY yearweek, yearweek_sort
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
    yearweek,
    yearweek_sort,
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
GROUP BY yearweek, yearweek_sort
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
    yearweek,
    yearweek_sort,
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
GROUP BY city, yearweek, yearweek_sort
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
-- Driver type (Joint vs Exclusive)
SELECT 'driver_type', driver_type, COUNT(*)
FROM Cab.vw_ShortBase WHERE driver_type IS NOT NULL AND driver_type != '' GROUP BY driver_type
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
    yearweek,
    yearweek_sort,
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
    -- Participation rate (% of all drivers)
    SUM(CASE WHEN snapp_incentive_participation = 'Yes' THEN 1 ELSE 0 END) * 100.0
        / COUNT(*) AS snapp_participation_pct,
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
GROUP BY yearweek, yearweek_sort
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
        / COUNT(*) AS snapp_participation_pct,
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
    yearweek,
    yearweek_sort,
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
GROUP BY city, yearweek, yearweek_sort
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
    yearweek,
    yearweek_sort,
    snapp_last_trip_navigation AS nav_app,
    COUNT(*) AS n,
    COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (PARTITION BY yearweek) AS pct
FROM Cab.vw_ShortBase
WHERE snapp_last_trip_navigation IS NOT NULL AND snapp_last_trip_navigation != ''
      AND yearweek IS NOT NULL
GROUP BY yearweek, yearweek_sort, snapp_last_trip_navigation
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
WITH raw AS (
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
),
totals AS (
    SELECT platform, SUM(n) AS platform_total
    FROM raw
    GROUP BY platform
)
SELECT r.platform, r.reason, r.n,
       r.n * 100.0 / NULLIF(t.platform_total, 0) AS pct
FROM raw r
JOIN totals t ON r.platform = t.platform
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


-- ============================================================
-- ROUTINE ANALYSIS VIEWS (replicate survey_routine_analysis.py)
-- All views: [Cab] schema, prefix vw_RA_
-- Use these in Power BI matrix reports (city x metric matrices)
-- ============================================================

-- RA-1. SATISFACTION & PARTICIPATION REVIEW
-- Replaces: #3_Sat_* sheets (All / Part-Time / Full-Time variants)
-- N-cutoff:  n (Snapp columns),  n_joint (Jnt_ columns)
-- Fix: TRY_CAST(active_joint AS INT) in base CTE avoids nvarchar→int implicit conversion
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_SatReview] AS
-- BUG FIX: Added UNION ALL to include an 'All Drivers' aggregate row per
-- (weeknumber, city).  The Python script computes "All Drivers" over ALL
-- respondents regardless of cooperation_type, so a simple AVERAGE of the
-- type-specific rows would give an unweighted (wrong) result in DAX.
-- Filter in Power BI: cooperation_type = 'All Drivers' for the All page,
--                     cooperation_type = 'Part-Time'   for the Part-Time page,
--                     cooperation_type = 'Full-Time'   for the Full-Time page.
WITH base AS (
    SELECT
        yearweek, weeknumber, city, cooperation_type,
        TRY_CAST(active_joint AS INT) AS is_joint,
        snapp_gotmessage_text_incentive,
        tapsi_gotmessage_text_incentive,
        snapp_incentive_participation,
        tapsi_incentive_participation,
        TRY_CAST(snapp_overall_incentive_satisfaction AS FLOAT) AS snapp_inc_sat,
        TRY_CAST(tapsi_overall_incentive_satisfaction AS FLOAT) AS tapsi_inc_sat,
        TRY_CAST(snapp_fare_satisfaction           AS FLOAT) AS snapp_fare_sat,
        TRY_CAST(tapsi_fare_satisfaction           AS FLOAT) AS tapsi_fare_sat,
        TRY_CAST(snapp_req_count_satisfaction      AS FLOAT) AS snapp_req_sat,
        TRY_CAST(tapsi_req_count_satisfaction      AS FLOAT) AS tapsi_req_sat,
        TRY_CAST(snapp_income_satisfaction         AS FLOAT) AS snapp_income_sat,
        TRY_CAST(tapsi_income_satisfaction         AS FLOAT) AS tapsi_income_sat
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL
),
agg AS (
    SELECT
        CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
        weeknumber, city, cooperation_type,
        COUNT(*) AS n,
        SUM(CASE WHEN is_joint = 1 THEN 1 ELSE 0 END) AS n_joint,

        -- % Incentive Participation
        100.0 * AVG(CASE WHEN snapp_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            AS Part_pct_Snapp,
        100.0 * SUM(CASE WHEN is_joint=1 AND snapp_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0)
            AS Part_pct_Jnt_Snapp,
        100.0 * SUM(CASE WHEN is_joint=1 AND tapsi_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0)
            AS Part_pct_Jnt_Tapsi,

        -- % Participation Among Who Got Message
        100.0 * SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND snapp_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END),0)
            AS Part_GotMsg_pct_Snapp,
        100.0 * SUM(CASE WHEN is_joint=1 AND snapp_gotmessage_text_incentive='Yes' AND snapp_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN is_joint=1 AND snapp_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END),0)
            AS Part_GotMsg_pct_Jnt_Snapp,
        100.0 * SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tapsi_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END),0)
            AS Part_GotMsg_pct_Jnt_Tapsi,

        -- Avg Incentive Satisfaction (1-5)
        AVG(snapp_inc_sat)                                            AS Incentive_Sat_Snapp,
        AVG(CASE WHEN is_joint=1 THEN snapp_inc_sat ELSE NULL END)    AS Incentive_Sat_Jnt_Snapp,
        AVG(CASE WHEN is_joint=1 THEN tapsi_inc_sat ELSE NULL END)    AS Incentive_Sat_Jnt_Tapsi,

        -- Avg Fare Satisfaction (1-5)
        AVG(snapp_fare_sat)                                           AS Fare_Sat_Snapp,
        AVG(CASE WHEN is_joint=1 THEN snapp_fare_sat ELSE NULL END)   AS Fare_Sat_Jnt_Snapp,
        AVG(CASE WHEN is_joint=1 THEN tapsi_fare_sat ELSE NULL END)   AS Fare_Sat_Jnt_Tapsi,

        -- Avg Request Count Satisfaction (1-5)
        AVG(snapp_req_sat)                                            AS Request_Sat_Snapp,
        AVG(CASE WHEN is_joint=1 THEN snapp_req_sat ELSE NULL END)    AS Request_Sat_Jnt_Snapp,
        AVG(CASE WHEN is_joint=1 THEN tapsi_req_sat ELSE NULL END)    AS Request_Sat_Jnt_Tapsi,

        -- Avg Income Satisfaction (1-5)
        AVG(snapp_income_sat)                                         AS Income_Sat_Snapp,
        AVG(CASE WHEN is_joint=1 THEN snapp_income_sat ELSE NULL END) AS Income_Sat_Jnt_Snapp,
        AVG(CASE WHEN is_joint=1 THEN tapsi_income_sat ELSE NULL END) AS Income_Sat_Jnt_Tapsi
    FROM base
    GROUP BY yearweek, weeknumber, city, cooperation_type
),
all_drv AS (
    -- "All Drivers" synthetic row: same formulas, no cooperation_type filter.
    -- Matches Python's lambda d: d  (no segment filter on the week slice).
    SELECT
        CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
        weeknumber, city, 'All Drivers' AS cooperation_type,
        COUNT(*) AS n,
        SUM(CASE WHEN is_joint = 1 THEN 1 ELSE 0 END) AS n_joint,

        100.0 * AVG(CASE WHEN snapp_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            AS Part_pct_Snapp,
        100.0 * SUM(CASE WHEN is_joint=1 AND snapp_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0)
            AS Part_pct_Jnt_Snapp,
        100.0 * SUM(CASE WHEN is_joint=1 AND tapsi_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0)
            AS Part_pct_Jnt_Tapsi,

        100.0 * SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND snapp_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END),0)
            AS Part_GotMsg_pct_Snapp,
        100.0 * SUM(CASE WHEN is_joint=1 AND snapp_gotmessage_text_incentive='Yes' AND snapp_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN is_joint=1 AND snapp_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END),0)
            AS Part_GotMsg_pct_Jnt_Snapp,
        100.0 * SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tapsi_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
            / NULLIF(SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END),0)
            AS Part_GotMsg_pct_Jnt_Tapsi,

        AVG(snapp_inc_sat)                                            AS Incentive_Sat_Snapp,
        AVG(CASE WHEN is_joint=1 THEN snapp_inc_sat ELSE NULL END)    AS Incentive_Sat_Jnt_Snapp,
        AVG(CASE WHEN is_joint=1 THEN tapsi_inc_sat ELSE NULL END)    AS Incentive_Sat_Jnt_Tapsi,

        AVG(snapp_fare_sat)                                           AS Fare_Sat_Snapp,
        AVG(CASE WHEN is_joint=1 THEN snapp_fare_sat ELSE NULL END)   AS Fare_Sat_Jnt_Snapp,
        AVG(CASE WHEN is_joint=1 THEN tapsi_fare_sat ELSE NULL END)   AS Fare_Sat_Jnt_Tapsi,

        AVG(snapp_req_sat)                                            AS Request_Sat_Snapp,
        AVG(CASE WHEN is_joint=1 THEN snapp_req_sat ELSE NULL END)    AS Request_Sat_Jnt_Snapp,
        AVG(CASE WHEN is_joint=1 THEN tapsi_req_sat ELSE NULL END)    AS Request_Sat_Jnt_Tapsi,

        AVG(snapp_income_sat)                                         AS Income_Sat_Snapp,
        AVG(CASE WHEN is_joint=1 THEN snapp_income_sat ELSE NULL END) AS Income_Sat_Jnt_Snapp,
        AVG(CASE WHEN is_joint=1 THEN tapsi_income_sat ELSE NULL END) AS Income_Sat_Jnt_Tapsi
    FROM base
    GROUP BY yearweek, weeknumber, city
)
SELECT * FROM agg
UNION ALL
SELECT * FROM all_drv;
GO


-- RA-2. CITIES OVERVIEW
-- Replaces: #12_Cities_Overview
-- N-cutoff:  E_n, F_n, G_n (three independent groups)
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_CitiesOverview] AS
WITH src AS (
    SELECT yearweek, weeknumber, city,
        TRY_CAST(active_joint AS INT)    AS is_joint,
        TRY_CAST(snapp_LOC   AS FLOAT)   AS snapp_loc_f,
        TRY_CAST(tapsi_LOC   AS FLOAT)   AS tapsi_loc_f,
        snapp_gotmessage_text_incentive,
        tapsi_gotmessage_text_incentive
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL
)
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
    weeknumber, city,
    COUNT(*) AS E_n,
    SUM(CASE WHEN is_joint = 1 THEN 1 ELSE 0 END) AS F_n,
    SUM(CASE WHEN tapsi_loc_f > 0   THEN 1 ELSE 0 END) AS G_n,

    -- E-group: all drivers
    100.0 * AVG(CASE WHEN is_joint = 1    THEN 1.0 ELSE 0.0 END) AS pct_Joint,
    100.0 * AVG(CASE WHEN tapsi_loc_f > 0 THEN 1.0 ELSE 0.0 END) AS pct_Dual_SU,
    AVG(snapp_loc_f) AS AvgLOC_All_Snapp,
    100.0 * AVG(CASE WHEN snapp_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END)
        AS GotMsg_All_Snapp,

    -- F-group: joint drivers
    AVG(CASE WHEN is_joint=1 THEN snapp_loc_f ELSE NULL END) AS AvgLOC_Joint_Snapp,
    100.0 * SUM(CASE WHEN is_joint=1 AND snapp_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0) AS GotMsg_Joint_Snapp,
    100.0 * SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0) AS GotMsg_Joint_Cmpt,

    -- G-group: competitor-signup drivers
    AVG(CASE WHEN is_joint=1        THEN tapsi_loc_f ELSE NULL END) AS AvgLOC_Joint_Cmpt,
    AVG(CASE WHEN tapsi_loc_f > 0   THEN tapsi_loc_f ELSE NULL END) AS AvgLOC_Joint_Cmpt_SU,
    100.0 * SUM(CASE WHEN tapsi_loc_f > 0 AND tapsi_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN tapsi_loc_f > 0 THEN 1.0 ELSE 0.0 END),0)
        AS GotMsg_Joint_Cmpt_SU
FROM src
GROUP BY yearweek, weeknumber, city;
GO


-- RA-3. RIDE SHARE
-- Replaces: #13_RideShare
-- N-cutoff:  total_Res
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_RideShare] AS
WITH src AS (
    SELECT yearweek, weeknumber, city,
        TRY_CAST(active_joint AS INT)  AS is_joint,
        TRY_CAST(snapp_ride   AS FLOAT) AS snapp_f,
        TRY_CAST(tapsi_ride   AS FLOAT) AS tapsi_f
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL
)
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
    weeknumber, city,
    COUNT(*) AS total_Res,
    SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END) AS Joint_Res,
    SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END) AS Ex_drivers,
    ISNULL(SUM(snapp_f),0) + ISNULL(SUM(tapsi_f),0)             AS Total_Ride,
    ISNULL(SUM(snapp_f),0)                                       AS Total_Ride_Snapp,
    ISNULL(SUM(CASE WHEN is_joint=0 THEN snapp_f ELSE 0 END),0) AS Ex_Ride_Snapp,
    ISNULL(SUM(CASE WHEN is_joint=1 THEN snapp_f ELSE 0 END),0) AS Jnt_Snapp_Ride,
    ISNULL(SUM(CASE WHEN is_joint=1 THEN tapsi_f ELSE 0 END),0) AS Jnt_Tapsi_Ride,
    100.0 * ISNULL(SUM(snapp_f),0)
        / NULLIF(ISNULL(SUM(snapp_f),0)+ISNULL(SUM(tapsi_f),0),0) AS All_Snapp_pct,
    100.0 * ISNULL(SUM(CASE WHEN is_joint=0 THEN snapp_f ELSE 0 END),0)
        / NULLIF(ISNULL(SUM(snapp_f),0)+ISNULL(SUM(tapsi_f),0),0) AS Ex_Drivers_Snapp_pct,
    100.0 * ISNULL(SUM(CASE WHEN is_joint=1 THEN snapp_f ELSE 0 END),0)
        / NULLIF(ISNULL(SUM(snapp_f),0)+ISNULL(SUM(tapsi_f),0),0) AS Jnt_at_Snapp_pct,
    100.0 * ISNULL(SUM(CASE WHEN is_joint=1 THEN tapsi_f ELSE 0 END),0)
        / NULLIF(ISNULL(SUM(snapp_f),0)+ISNULL(SUM(tapsi_f),0),0) AS Jnt_at_Tapsi_pct
FROM src
GROUP BY yearweek, weeknumber, city;
GO


-- RA-4. PERSONA PART-TIME
-- Replaces: #15_Persona_PartTime
-- N-cutoff:  total_Res (also Joint_Res for PT_pct_Joint, Ex_drivers for PT_pct_Exclusive)
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_PersonaPartTime] AS
WITH src AS (
    SELECT yearweek, weeknumber, city, cooperation_type,
        TRY_CAST(active_joint AS INT)  AS is_joint,
        TRY_CAST(snapp_ride   AS FLOAT) AS snapp_f,
        TRY_CAST(tapsi_ride   AS FLOAT) AS tapsi_f
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL
)
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
    weeknumber, city,
    COUNT(*) AS total_Res,
    SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END) AS Joint_Res,
    SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END) AS Ex_drivers,
    100.0 * SUM(CASE WHEN is_joint=1 AND cooperation_type='Part-Time' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0) AS PT_pct_Joint,
    100.0 * SUM(CASE WHEN is_joint=0 AND cooperation_type='Part-Time' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=0 THEN 1.0 ELSE 0.0 END),0) AS PT_pct_Exclusive,
    ISNULL(SUM(CASE WHEN is_joint=1 THEN snapp_f ELSE 0 END),0)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0) AS RidePerBoarded_Snapp,
    ISNULL(SUM(CASE WHEN is_joint=1 THEN tapsi_f ELSE 0 END),0)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0) AS RidePerBoarded_Tapsi,
    ISNULL(SUM(snapp_f),0) / NULLIF(COUNT(*),0) AS AvgAllRides
FROM src
GROUP BY yearweek, weeknumber, city;
GO


-- RA-5. INCENTIVE AMOUNTS (long format)
-- Replaces: #1_Snapp_Incentive_Amt and #2_Tapsi_Incentive_Amt
-- N-cutoff:  n_total  |  Filter platform = 'Tapsi' for sheet #2
-- Matrix: city = rows, incentive_range = columns, pct = values
GO
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
    END AS incentive_range_sort,
    n_range, n_total,
    100.0 * n_range / NULLIF(n_total, 0) AS pct
FROM raw;
GO


-- RA-6. INCENTIVE DURATION (long format)
-- Replaces: #4_Incentive_Duration
-- N-cutoff:  n_total
-- Matrix: city = rows, duration_bucket = columns, pct = values
GO
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
    CASE duration_bucket
        WHEN 'Few Hours' THEN 1
        WHEN '1 Day'     THEN 2
        WHEN '1_6 Days'  THEN 3
        WHEN '7 Days'    THEN 4
        WHEN '>7 Days'   THEN 5
        ELSE 99
    END AS duration_bucket_sort,
    n_range, n_total,
    100.0 * n_range / NULLIF(n_total, 0) AS pct
FROM raw;
GO


-- RA-7. PERSONA (long format — all demographic dimensions)
-- Replaces: all #15_Persona sub-sheets
-- N-cutoff:  n_total
-- Slicer on 'dimension'; Matrix: city = rows, category = columns, pct = values
-- Fix: CAST(edu/marr_stat AS NVARCHAR) prevents UNION ALL type-precedence error when
--      those columns are stored as numeric in ShortMain
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_Persona] AS
WITH activity AS (
    SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
        weeknumber, city, 'Activity Type' AS dimension,
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
    GROUP BY yearweek, weeknumber, city, active_time),
age_grp AS (
    SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
        weeknumber, city, 'Age Group' AS dimension,
        CAST(age_group AS NVARCHAR(100)) AS category,
        CASE age_group
            WHEN '18_to_35'     THEN 1
            WHEN 'more_than_35' THEN 2
            ELSE 99
        END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND age_group IS NOT NULL
    GROUP BY yearweek, weeknumber, city, age_group),
edu AS (
    SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
        weeknumber, city, 'Education' AS dimension,
        CAST(edu AS NVARCHAR(100)) AS category,
        CASE TRY_CAST(edu AS INT) WHEN 0 THEN 1 WHEN 1 THEN 2 ELSE 99 END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND edu IS NOT NULL
    GROUP BY yearweek, weeknumber, city, edu),
marr AS (
    SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
        weeknumber, city, 'Marital Status' AS dimension,
        CAST(marr_stat AS NVARCHAR(100)) AS category,
        CASE TRY_CAST(marr_stat AS INT) WHEN 0 THEN 1 WHEN 1 THEN 2 ELSE 99 END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND marr_stat IS NOT NULL
    GROUP BY yearweek, weeknumber, city, marr_stat),
gen AS (
    SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
        weeknumber, city, 'Gender' AS dimension,
        CAST(gender AS NVARCHAR(100)) AS category,
        CASE gender WHEN 'Female' THEN 1 WHEN 'Male' THEN 2 ELSE 99 END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND gender IS NOT NULL
    GROUP BY yearweek, weeknumber, city, gender),
coop AS (
    SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
        weeknumber, city, 'Cooperation Type' AS dimension,
        CAST(cooperation_type AS NVARCHAR(100)) AS category,
        CASE cooperation_type WHEN 'Full-Time' THEN 1 WHEN 'Part-Time' THEN 2 ELSE 99 END AS category_sort,
        COUNT(*) AS n, SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL AND cooperation_type IS NOT NULL
    GROUP BY yearweek, weeknumber, city, cooperation_type)
SELECT *, 100.0 * n / NULLIF(n_total,0) AS pct FROM activity
UNION ALL SELECT *, 100.0 * n / NULLIF(n_total,0) AS pct FROM age_grp
UNION ALL SELECT *, 100.0 * n / NULLIF(n_total,0) AS pct FROM edu
UNION ALL SELECT *, 100.0 * n / NULLIF(n_total,0) AS pct FROM marr
UNION ALL SELECT *, 100.0 * n / NULLIF(n_total,0) AS pct FROM gen
UNION ALL SELECT *, 100.0 * n / NULLIF(n_total,0) AS pct FROM coop;
GO


-- RA-8. COMMISSION-FREE INCENTIVE (Updated - Page #18)
-- Replaces: #18_CommFree_Snapp and #18_CommFree_Tapsi
-- Type-breakdown columns (multi-select) come from WideMain; all others from ShortMain.
-- N-cutoff:  n
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_CommFree] AS
WITH src AS (
    SELECT m.yearweek, m.weeknumber, m.city,
        TRY_CAST(m.active_joint  AS INT)    AS is_joint,
        TRY_CAST(m.snapp_commfree AS FLOAT) AS snapp_cf,
        TRY_CAST(m.tapsi_commfree AS FLOAT) AS tapsi_cf,
        TRY_CAST(m.snapp_ride AS FLOAT)     AS snapp_ride,
        TRY_CAST(m.tapsi_ride AS FLOAT)     AS tapsi_ride,
        m.snapp_gotmessage_text_incentive,
        m.tapsi_gotmessage_text_incentive,
        CAST(m.snapp_incentive_category AS NVARCHAR(100)) AS snapp_inc_cat,
        CAST(m.tapsi_incentive_category AS NVARCHAR(100)) AS tapsi_inc_cat,
        m.snapp_incentive_participation,
        m.tapsi_incentive_participation,
        -- incentive type binary flags (multi-select) live in WideMain
        TRY_CAST(w.[Snapp Incentive Type__Pay After Ride]              AS INT) AS sn_pay_ride,
        TRY_CAST(w.[Snapp Incentive Type__Earning-based Commission-free] AS INT) AS sn_earn_cf,
        TRY_CAST(w.[Snapp Incentive Type__Ride-Based Commission-free]  AS INT) AS sn_ride_cf,
        TRY_CAST(w.[Snapp Incentive Type__Income Guarantee]            AS INT) AS sn_inc_guar,
        TRY_CAST(w.[Snapp Incentive Type__Pay After Income]            AS INT) AS sn_pay_inc,
        TRY_CAST(w.[Snapp Incentive Type__Commission Free on some trips] AS INT) AS sn_cf_some,
        TRY_CAST(w.[Tapsi Incentive Type__Pay After Ride]              AS INT) AS tp_pay_ride,
        TRY_CAST(w.[Tapsi Incentive Type__Earning-based Commission-free] AS INT) AS tp_earn_cf,
        TRY_CAST(w.[Tapsi Incentive Type__Ride-Based Commission-free]  AS INT) AS tp_ride_cf,
        TRY_CAST(w.[Tapsi Incentive Type__Income Guarantee]            AS INT) AS tp_inc_guar,
        TRY_CAST(w.[Tapsi Incentive Type__Pay After Income]            AS INT) AS tp_pay_inc,
        TRY_CAST(w.[Tapsi Incentive Type__Commission Free on some trips] AS INT) AS tp_cf_some
    FROM [Cab].[DriverSurvey_ShortMain] m
    LEFT JOIN [Cab].[DriverSurvey_WideMain] w
        ON CAST(m.recordID AS INT) = CAST(w.recordID AS INT)
    WHERE m.city IS NOT NULL
)
SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
    weeknumber, city, 'Snapp' AS platform,
    COUNT(*) AS n,
    -- message receipt by category
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' THEN 1 ELSE 0 END) AS Who_Got_Message,
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND snapp_inc_cat='Money' THEN 1 ELSE 0 END) AS GotMsg_Money,
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND snapp_inc_cat='Free-Commission' THEN 1 ELSE 0 END) AS GotMsg_FreeComm,
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND snapp_inc_cat='Money & Free-commission' THEN 1 ELSE 0 END) AS GotMsg_Money_FreeComm,
    -- message receipt by incentive type (binary flags)
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND sn_pay_ride=1 THEN 1 ELSE 0 END) AS GotMsg_PayRide,
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND sn_earn_cf=1 THEN 1 ELSE 0 END) AS GotMsg_EarnCF,
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND sn_ride_cf=1 THEN 1 ELSE 0 END) AS GotMsg_RideCF,
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND sn_inc_guar=1 THEN 1 ELSE 0 END) AS GotMsg_IncGuar,
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND sn_pay_inc=1 THEN 1 ELSE 0 END) AS GotMsg_PayInc,
    SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' AND sn_cf_some=1 THEN 1 ELSE 0 END) AS GotMsg_CFSome,
    -- participation & CF ride metrics
    SUM(CASE WHEN snapp_cf > 0 THEN 1 ELSE 0 END) AS Free_Comm_Drivers,
    SUM(CASE WHEN snapp_incentive_participation='Yes' THEN 1 ELSE 0 END) AS Participated,
    100.0 * SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END) / NULLIF(COUNT(*),0) AS pct_Got_Message,
    100.0 * SUM(CASE WHEN snapp_cf > 0 THEN 1.0 ELSE 0.0 END) / NULLIF(COUNT(*),0) AS pct_Free_Comm_Ride,
    100.0 * SUM(CASE WHEN snapp_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN snapp_gotmessage_text_incentive='Yes' THEN 1 ELSE 0 END),0) AS pct_Participated,
    AVG(CASE WHEN snapp_cf > 0 THEN snapp_cf END) AS Avg_CF_Rides,
    AVG(snapp_ride) AS Avg_Total_Rides,
    AVG(CASE WHEN snapp_ride > 0 THEN 100.0 * snapp_cf / snapp_ride END) AS Avg_pct_CF_RideShare
FROM src
GROUP BY yearweek, weeknumber, city
UNION ALL
SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort,
    weeknumber, city, 'Tapsi' AS platform,
    SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END) AS n,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' THEN 1 ELSE 0 END) AS Who_Got_Message,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tapsi_inc_cat='Money' THEN 1 ELSE 0 END) AS GotMsg_Money,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tapsi_inc_cat='Free-Commission' THEN 1 ELSE 0 END) AS GotMsg_FreeComm,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tapsi_inc_cat='Money & Free-commission' THEN 1 ELSE 0 END) AS GotMsg_Money_FreeComm,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tp_pay_ride=1 THEN 1 ELSE 0 END) AS GotMsg_PayRide,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tp_earn_cf=1 THEN 1 ELSE 0 END) AS GotMsg_EarnCF,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tp_ride_cf=1 THEN 1 ELSE 0 END) AS GotMsg_RideCF,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tp_inc_guar=1 THEN 1 ELSE 0 END) AS GotMsg_IncGuar,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tp_pay_inc=1 THEN 1 ELSE 0 END) AS GotMsg_PayInc,
    SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' AND tp_cf_some=1 THEN 1 ELSE 0 END) AS GotMsg_CFSome,
    SUM(CASE WHEN is_joint=1 AND tapsi_cf > 0 THEN 1 ELSE 0 END) AS Free_Comm_Drivers,
    SUM(CASE WHEN is_joint=1 AND tapsi_incentive_participation='Yes' THEN 1 ELSE 0 END) AS Participated,
    100.0 * SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0) AS pct_Got_Message,
    100.0 * SUM(CASE WHEN is_joint=1 AND tapsi_cf > 0 THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1.0 ELSE 0.0 END),0) AS pct_Free_Comm_Ride,
    100.0 * SUM(CASE WHEN is_joint=1 AND tapsi_incentive_participation='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' THEN 1 ELSE 0 END),0) AS pct_Participated,
    AVG(CASE WHEN is_joint=1 AND tapsi_cf > 0 THEN tapsi_cf END) AS Avg_CF_Rides,
    AVG(CASE WHEN is_joint=1 THEN tapsi_ride END) AS Avg_Total_Rides,
    AVG(CASE WHEN is_joint=1 AND tapsi_ride > 0 THEN 100.0 * tapsi_cf / tapsi_ride END) AS Avg_pct_CF_RideShare
FROM src
GROUP BY yearweek, weeknumber, city;
GO


-- RA-9. CUSTOMER SUPPORT SATISFACTION (from ShortRare)
-- Replaces: #CS_Sat_Snapp and #CS_Sat_Tapsi sheets
-- N-cutoff:  n
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_CSRare] AS
SELECT
    CAST(sm.yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(sm.yearweek%100 AS VARCHAR), 2) AS yearweek,
    sm.yearweek AS yearweek_sort,
    sm.weeknumber, sm.city,
    COUNT(*) AS n,
    AVG(TRY_CAST(sr.snapp_CS_satisfaction_overall   AS FLOAT)) AS Snapp_CS_Overall,
    AVG(TRY_CAST(sr.snapp_CS_satisfaction_waittime  AS FLOAT)) AS Snapp_CS_WaitTime,
    AVG(TRY_CAST(sr.snapp_CS_satisfaction_solution  AS FLOAT)) AS Snapp_CS_Solution,
    AVG(TRY_CAST(sr.snapp_CS_satisfaction_behaviour AS FLOAT)) AS Snapp_CS_Behaviour,
    AVG(TRY_CAST(sr.snapp_CS_satisfaction_relevance AS FLOAT)) AS Snapp_CS_Relevance,
    100.0 * AVG(CASE WHEN sr.snapp_CS_solved='Yes' THEN 1.0 ELSE 0.0 END) AS Snapp_CS_Solved_pct,
    AVG(TRY_CAST(sr.tapsi_CS_satisfaction_overall   AS FLOAT)) AS Tapsi_CS_Overall,
    AVG(TRY_CAST(sr.tapsi_CS_satisfaction_waittime  AS FLOAT)) AS Tapsi_CS_WaitTime,
    AVG(TRY_CAST(sr.tapsi_CS_satisfaction_solution  AS FLOAT)) AS Tapsi_CS_Solution,
    AVG(TRY_CAST(sr.tapsi_CS_satisfaction_behaviour AS FLOAT)) AS Tapsi_CS_Behaviour,
    AVG(TRY_CAST(sr.tapsi_CS_satisfaction_relevance AS FLOAT)) AS Tapsi_CS_Relevance,
    100.0 * AVG(CASE WHEN sr.tapsi_CS_solved='Yes' THEN 1.0 ELSE 0.0 END) AS Tapsi_CS_Solved_pct
FROM [Cab].[DriverSurvey_ShortRare]  sr
JOIN [Cab].[DriverSurvey_ShortMain]  sm ON sm.recordID = sr.recordID
WHERE sm.city IS NOT NULL
GROUP BY sm.yearweek, sm.weeknumber, sm.city;
GO


-- RA-10. NAVIGATION & NPS RECOMMENDATION SCORES (from ShortRare)
-- Replaces: #NavReco_Scores and #Reco_NPS sheets
-- N-cutoff:  n
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_NavReco] AS
SELECT
    CAST(sm.yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(sm.yearweek%100 AS VARCHAR), 2) AS yearweek,
    sm.yearweek AS yearweek_sort,
    sm.weeknumber, sm.city,
    COUNT(*) AS n,
    AVG(TRY_CAST(sr.snapp_recommend               AS FLOAT)) AS Snapp_NPS,
    AVG(TRY_CAST(sr.snappdriver_tapsi_recommend    AS FLOAT)) AS Tapsi_NPS_SnapDriver,
    AVG(TRY_CAST(sr.tapsidriver_tapsi_recommend    AS FLOAT)) AS Tapsi_NPS_TapsiDriver,
    AVG(TRY_CAST(sr.recommendation_googlemap       AS FLOAT)) AS Reco_GoogleMap,
    AVG(TRY_CAST(sr.recommendation_waze            AS FLOAT)) AS Reco_Waze,
    AVG(TRY_CAST(sr.recommendation_neshan          AS FLOAT)) AS Reco_Neshan,
    AVG(TRY_CAST(sr.recommendation_balad           AS FLOAT)) AS Reco_Balad,
    AVG(TRY_CAST(sr.snapp_navigation_app_satisfaction    AS FLOAT)) AS Snapp_Nav_Sat,
    AVG(TRY_CAST(sr.tapsi_in_app_navigation_satisfaction AS FLOAT)) AS Tapsi_Nav_Sat
FROM [Cab].[DriverSurvey_ShortRare]  sr
JOIN [Cab].[DriverSurvey_ShortMain]  sm ON sm.recordID = sr.recordID
WHERE sm.city IS NOT NULL
GROUP BY sm.yearweek, sm.weeknumber, sm.city;
GO


-- ============================================================
-- ROUTINE ANALYSIS VIEWS - BATCH 2 (Pages #5/#6, #8, #9, #14, #16, #17, #19)
-- ============================================================


-- RA-11. INCENTIVE TYPE DISTRIBUTION (Pages #5 & #6)
-- Wide format: pcts by type for Excl-Snapp / Jnt-Snapp / Jnt-Tapsi + avg CF rides
-- Type binary flags (multi-select) joined from WideMain.
-- N-cutoff: n_excl (for Snapp Excl metrics), n_joint (for Joint metrics)
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_IncentiveTypes] AS
WITH src AS (
    SELECT m.yearweek, m.weeknumber, m.city,
        TRY_CAST(m.active_joint AS INT) AS is_joint,
        m.snapp_gotmessage_text_incentive,
        m.tapsi_gotmessage_text_incentive,
        TRY_CAST(m.snapp_commfree AS FLOAT) AS snapp_cf,
        TRY_CAST(m.tapsi_commfree AS FLOAT) AS tapsi_cf,
        TRY_CAST(w.[Snapp Incentive Type__Pay After Ride]               AS INT) AS sn_pay_ride,
        TRY_CAST(w.[Snapp Incentive Type__Ride-Based Commission-free]   AS INT) AS sn_ride_cf,
        TRY_CAST(w.[Snapp Incentive Type__Earning-based Commission-free] AS INT) AS sn_earn_cf,
        TRY_CAST(w.[Snapp Incentive Type__Income Guarantee]             AS INT) AS sn_inc_guar,
        TRY_CAST(w.[Snapp Incentive Type__Pay After Income]             AS INT) AS sn_pay_inc,
        TRY_CAST(w.[Snapp Incentive Type__Commission Free on some trips] AS INT) AS sn_cf_some,
        TRY_CAST(w.[Tapsi Incentive Type__Pay After Ride]               AS INT) AS tp_pay_ride,
        TRY_CAST(w.[Tapsi Incentive Type__Ride-Based Commission-free]   AS INT) AS tp_ride_cf,
        TRY_CAST(w.[Tapsi Incentive Type__Earning-based Commission-free] AS INT) AS tp_earn_cf,
        TRY_CAST(w.[Tapsi Incentive Type__Income Guarantee]             AS INT) AS tp_inc_guar,
        TRY_CAST(w.[Tapsi Incentive Type__Pay After Income]             AS INT) AS tp_pay_inc,
        TRY_CAST(w.[Tapsi Incentive Type__Commission Free on some trips] AS INT) AS tp_cf_some
    FROM [Cab].[DriverSurvey_ShortMain] m
    LEFT JOIN [Cab].[DriverSurvey_WideMain] w
        ON CAST(m.recordID AS INT) = CAST(w.recordID AS INT)
    WHERE m.city IS NOT NULL
)
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    weeknumber, city,
    COUNT(*) AS n,
    SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END) AS n_joint,
    SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END) AS n_excl,
    -- Got-message rates by segment
    100.0*SUM(CASE WHEN is_joint=0 AND snapp_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END),0) AS pct_GotMsg_Excl_Snapp,
    100.0*SUM(CASE WHEN is_joint=1 AND snapp_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_GotMsg_Jnt_Snapp,
    100.0*SUM(CASE WHEN is_joint=1 AND tapsi_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_GotMsg_Jnt_Tapsi,
    100.0*SUM(CASE WHEN is_joint=1 AND snapp_gotmessage_text_incentive='Yes' AND tapsi_gotmessage_text_incentive='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_GotMsg_Both,
    100.0*SUM(CASE WHEN is_joint=1
                    AND (snapp_gotmessage_text_incentive='Yes' OR tapsi_gotmessage_text_incentive='Yes')
                    AND NOT (snapp_gotmessage_text_incentive='Yes' AND tapsi_gotmessage_text_incentive='Yes')
                   THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_GotMsg_Diff,
    -- Snapp Exclusive type breakdown (base = n_excl)
    100.0*SUM(CASE WHEN is_joint=0 AND sn_pay_ride=1 THEN 1.0 ELSE 0.0 END)  / NULLIF(SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END),0) AS pct_PayRide_Excl,
    100.0*SUM(CASE WHEN is_joint=0 AND sn_earn_cf=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END),0) AS pct_EarnCF_Excl,
    100.0*SUM(CASE WHEN is_joint=0 AND sn_ride_cf=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END),0) AS pct_RideCF_Excl,
    100.0*SUM(CASE WHEN is_joint=0 AND sn_inc_guar=1 THEN 1.0 ELSE 0.0 END)  / NULLIF(SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END),0) AS pct_IncGuar_Excl,
    100.0*SUM(CASE WHEN is_joint=0 AND sn_pay_inc=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END),0) AS pct_PayInc_Excl,
    100.0*SUM(CASE WHEN is_joint=0 AND sn_cf_some=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=0 THEN 1 ELSE 0 END),0) AS pct_CFSome_Excl,
    -- Snapp Joint type breakdown (base = n_joint)
    100.0*SUM(CASE WHEN is_joint=1 AND sn_pay_ride=1 THEN 1.0 ELSE 0.0 END)  / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_PayRide_JntSn,
    100.0*SUM(CASE WHEN is_joint=1 AND sn_earn_cf=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_EarnCF_JntSn,
    100.0*SUM(CASE WHEN is_joint=1 AND sn_ride_cf=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_RideCF_JntSn,
    100.0*SUM(CASE WHEN is_joint=1 AND sn_inc_guar=1 THEN 1.0 ELSE 0.0 END)  / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_IncGuar_JntSn,
    100.0*SUM(CASE WHEN is_joint=1 AND sn_pay_inc=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_PayInc_JntSn,
    100.0*SUM(CASE WHEN is_joint=1 AND sn_cf_some=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_CFSome_JntSn,
    -- Tapsi Joint type breakdown (base = n_joint)
    100.0*SUM(CASE WHEN is_joint=1 AND tp_pay_ride=1 THEN 1.0 ELSE 0.0 END)  / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_PayRide_JntTp,
    100.0*SUM(CASE WHEN is_joint=1 AND tp_earn_cf=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_EarnCF_JntTp,
    100.0*SUM(CASE WHEN is_joint=1 AND tp_ride_cf=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_RideCF_JntTp,
    100.0*SUM(CASE WHEN is_joint=1 AND tp_inc_guar=1 THEN 1.0 ELSE 0.0 END)  / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_IncGuar_JntTp,
    100.0*SUM(CASE WHEN is_joint=1 AND tp_pay_inc=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_PayInc_JntTp,
    100.0*SUM(CASE WHEN is_joint=1 AND tp_cf_some=1 THEN 1.0 ELSE 0.0 END)   / NULLIF(SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END),0) AS pct_CFSome_JntTp,
    -- Average commission-free rides (among drivers who did CF rides)
    AVG(CASE WHEN snapp_cf > 0 THEN snapp_cf END) AS Avg_CF_Rides_Snapp,
    AVG(CASE WHEN is_joint=1 AND tapsi_cf > 0 THEN tapsi_cf END) AS Avg_CF_Rides_Tapsi
FROM src
GROUP BY yearweek, weeknumber, city;
GO


-- RA-12. INCENTIVE DISSATISFACTION BY CITY (Page #8)
-- n_sn_low_sat = drivers who cited any Snapp dissatisfaction reason (multi-select from WideMain).
-- n_tp_low_sat = joint drivers who cited any Tapsi dissatisfaction reason.
-- N-cutoff: n_sn_low_sat (Snapp), n_tp_low_sat (Tapsi)
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_IncentiveUnsatCity] AS
WITH src AS (
    SELECT m.yearweek, m.weeknumber, m.city,
        TRY_CAST(m.active_joint AS INT) AS is_joint,
        -- Snapp unsatisfaction reasons (multi-select) from WideMain
        TRY_CAST(w.[Snapp Last Incentive Unsatisfaction__Not Available]  AS INT) AS sn_no_time,
        TRY_CAST(w.[Snapp Last Incentive Unsatisfaction__Improper Amount] AS INT) AS sn_imp_amt,
        TRY_CAST(w.[Snapp Last Incentive Unsatisfaction__No Time todo]   AS INT) AS sn_low_time,
        TRY_CAST(w.[Snapp Last Incentive Unsatisfaction__difficult]      AS INT) AS sn_hard,
        TRY_CAST(w.[Snapp Last Incentive Unsatisfaction__Non Payment]    AS INT) AS sn_nonpay,
        -- Tapsi unsatisfaction reasons from WideMain
        TRY_CAST(w.[Tapsi Incentive Unsatisfaction__Not Available]       AS INT) AS tp_no_time,
        TRY_CAST(w.[Tapsi Incentive Unsatisfaction__Improper Amount]     AS INT) AS tp_imp_amt,
        TRY_CAST(w.[Tapsi Incentive Unsatisfaction__No Time todo]        AS INT) AS tp_low_time,
        TRY_CAST(w.[Tapsi Incentive Unsatisfaction__difficult]           AS INT) AS tp_hard,
        TRY_CAST(w.[Tapsi Incentive Unsatisfaction__Non Payment]         AS INT) AS tp_nonpay
    FROM [Cab].[DriverSurvey_ShortMain] m
    LEFT JOIN [Cab].[DriverSurvey_WideMain] w
        ON CAST(m.recordID AS INT) = CAST(w.recordID AS INT)
    WHERE m.city IS NOT NULL
),
agg AS (
    SELECT yearweek, weeknumber, city, is_joint,
        sn_no_time, sn_imp_amt, sn_low_time, sn_hard, sn_nonpay,
        tp_no_time, tp_imp_amt, tp_low_time, tp_hard, tp_nonpay,
        -- low-sat flag: cited any reason
        CASE WHEN (sn_no_time=1 OR sn_imp_amt=1 OR sn_low_time=1 OR sn_hard=1 OR sn_nonpay=1)
             THEN 1 ELSE 0 END AS sn_low_sat,
        CASE WHEN (tp_no_time=1 OR tp_imp_amt=1 OR tp_low_time=1 OR tp_hard=1 OR tp_nonpay=1)
             THEN 1 ELSE 0 END AS tp_low_sat
    FROM src
)
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    weeknumber, city,
    COUNT(*) AS n_all,
    SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END) AS n_joint,
    SUM(sn_low_sat) AS n_sn_low_sat,
    SUM(CASE WHEN is_joint=1 THEN tp_low_sat ELSE 0 END) AS n_tp_low_sat,
    -- All Snapp dissatisfaction reasons (base = n_sn_low_sat)
    100.0*SUM(CASE WHEN sn_low_sat=1 AND sn_no_time=1  THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(sn_low_sat),0) AS pct_Sn_NoTime,
    100.0*SUM(CASE WHEN sn_low_sat=1 AND sn_imp_amt=1  THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(sn_low_sat),0) AS pct_Sn_ImpAmt,
    100.0*SUM(CASE WHEN sn_low_sat=1 AND sn_low_time=1 THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(sn_low_sat),0) AS pct_Sn_LowTime,
    100.0*SUM(CASE WHEN sn_low_sat=1 AND sn_hard=1     THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(sn_low_sat),0) AS pct_Sn_HardToDo,
    100.0*SUM(CASE WHEN sn_low_sat=1 AND sn_nonpay=1   THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(sn_low_sat),0) AS pct_Sn_NonPay,
    -- Joint Tapsi dissatisfaction reasons (base = n_tp_low_sat)
    100.0*SUM(CASE WHEN is_joint=1 AND tp_low_sat=1 AND tp_no_time=1  THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(CASE WHEN is_joint=1 THEN tp_low_sat ELSE 0 END),0) AS pct_Tp_NoTime,
    100.0*SUM(CASE WHEN is_joint=1 AND tp_low_sat=1 AND tp_imp_amt=1  THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(CASE WHEN is_joint=1 THEN tp_low_sat ELSE 0 END),0) AS pct_Tp_ImpAmt,
    100.0*SUM(CASE WHEN is_joint=1 AND tp_low_sat=1 AND tp_low_time=1 THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(CASE WHEN is_joint=1 THEN tp_low_sat ELSE 0 END),0) AS pct_Tp_LowTime,
    100.0*SUM(CASE WHEN is_joint=1 AND tp_low_sat=1 AND tp_hard=1     THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(CASE WHEN is_joint=1 THEN tp_low_sat ELSE 0 END),0) AS pct_Tp_HardToDo,
    100.0*SUM(CASE WHEN is_joint=1 AND tp_low_sat=1 AND tp_nonpay=1   THEN 1.0 ELSE 0.0 END) / NULLIF(SUM(CASE WHEN is_joint=1 THEN tp_low_sat ELSE 0 END),0) AS pct_Tp_NonPay
FROM agg
GROUP BY yearweek, weeknumber, city;
GO


-- RA-13. INCENTIVE DISSATISFACTION - NATIONAL LEVEL (Page #9)
-- Long format: segment = All Snapp / Joint Snapp / Joint Tapsi; no city breakdown
-- N-cutoff: n_low_sat per segment
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_IncentiveUnsatNational] AS
WITH src AS (
    SELECT yearweek, weeknumber,
        TRY_CAST(active_joint AS INT) AS is_joint,
        TRY_CAST(snapp_incentive_satisfaction AS FLOAT) AS sn_sat,
        TRY_CAST(tapsi_incentive_satisfaction AS FLOAT) AS tp_sat,
        TRY_CAST(snapp_lincentive_unsatisfaction_not_available  AS INT) AS sn_no_time,
        TRY_CAST(snapp_lincentive_unsatisfaction_improper_amount AS INT) AS sn_imp_amt,
        TRY_CAST(snapp_lincentive_unsatisfaction_no_time        AS INT) AS sn_low_time,
        TRY_CAST(snapp_lincentive_unsatisfaction_difficult      AS INT) AS sn_hard,
        TRY_CAST(snapp_lincentive_unsatisfaction_non_payment    AS INT) AS sn_nonpay,
        TRY_CAST(snapp_lincentive_unsatisfaction_other          AS INT) AS sn_other,
        TRY_CAST(tapsi_lincentive_unsatisfaction_not_available  AS INT) AS tp_no_time,
        TRY_CAST(tapsi_lincentive_unsatisfaction_improper_amount AS INT) AS tp_imp_amt,
        TRY_CAST(tapsi_lincentive_unsatisfaction_no_time        AS INT) AS tp_low_time,
        TRY_CAST(tapsi_lincentive_unsatisfaction_difficult      AS INT) AS tp_hard,
        TRY_CAST(tapsi_lincentive_unsatisfaction_non_payment    AS INT) AS tp_nonpay,
        TRY_CAST(tapsi_lincentive_unsatisfaction_other          AS INT) AS tp_other
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL
),
all_snapp AS (
    SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort, weeknumber,
        'All Snapp' AS segment, 1 AS segment_sort,
        COUNT(*) AS n,
        SUM(CASE WHEN sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END) AS n_low_sat,
        100.0*SUM(CASE WHEN sn_sat<=3 AND sn_no_time=1  THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_NoTime,
        100.0*SUM(CASE WHEN sn_sat<=3 AND sn_imp_amt=1  THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_ImpAmt,
        100.0*SUM(CASE WHEN sn_sat<=3 AND sn_low_time=1 THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_LowTime,
        100.0*SUM(CASE WHEN sn_sat<=3 AND sn_hard=1     THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_HardToDo,
        100.0*SUM(CASE WHEN sn_sat<=3 AND sn_nonpay=1   THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_NonPay,
        100.0*SUM(CASE WHEN sn_sat<=3 AND sn_other=1    THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_Other
    FROM src GROUP BY yearweek, weeknumber
),
jnt_snapp AS (
    SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort, weeknumber,
        'Joint Snapp' AS segment, 2 AS segment_sort,
        SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END) AS n,
        SUM(CASE WHEN is_joint=1 AND sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END) AS n_low_sat,
        100.0*SUM(CASE WHEN is_joint=1 AND sn_sat<=3 AND sn_no_time=1  THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_NoTime,
        100.0*SUM(CASE WHEN is_joint=1 AND sn_sat<=3 AND sn_imp_amt=1  THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_ImpAmt,
        100.0*SUM(CASE WHEN is_joint=1 AND sn_sat<=3 AND sn_low_time=1 THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_LowTime,
        100.0*SUM(CASE WHEN is_joint=1 AND sn_sat<=3 AND sn_hard=1     THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_HardToDo,
        100.0*SUM(CASE WHEN is_joint=1 AND sn_sat<=3 AND sn_nonpay=1   THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_NonPay,
        100.0*SUM(CASE WHEN is_joint=1 AND sn_sat<=3 AND sn_other=1    THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND sn_sat IS NOT NULL AND sn_sat<=3 THEN 1 ELSE 0 END),0) AS pct_Other
    FROM src GROUP BY yearweek, weeknumber
),
jnt_tapsi AS (
    SELECT CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort, weeknumber,
        'Joint Tapsi' AS segment, 3 AS segment_sort,
        SUM(CASE WHEN is_joint=1 THEN 1 ELSE 0 END) AS n,
        SUM(CASE WHEN is_joint=1 AND tp_sat IS NOT NULL AND tp_sat<=3 THEN 1 ELSE 0 END) AS n_low_sat,
        100.0*SUM(CASE WHEN is_joint=1 AND tp_sat<=3 AND tp_no_time=1  THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tp_sat IS NOT NULL AND tp_sat<=3 THEN 1 ELSE 0 END),0) AS pct_NoTime,
        100.0*SUM(CASE WHEN is_joint=1 AND tp_sat<=3 AND tp_imp_amt=1  THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tp_sat IS NOT NULL AND tp_sat<=3 THEN 1 ELSE 0 END),0) AS pct_ImpAmt,
        100.0*SUM(CASE WHEN is_joint=1 AND tp_sat<=3 AND tp_low_time=1 THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tp_sat IS NOT NULL AND tp_sat<=3 THEN 1 ELSE 0 END),0) AS pct_LowTime,
        100.0*SUM(CASE WHEN is_joint=1 AND tp_sat<=3 AND tp_hard=1     THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tp_sat IS NOT NULL AND tp_sat<=3 THEN 1 ELSE 0 END),0) AS pct_HardToDo,
        100.0*SUM(CASE WHEN is_joint=1 AND tp_sat<=3 AND tp_nonpay=1   THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tp_sat IS NOT NULL AND tp_sat<=3 THEN 1 ELSE 0 END),0) AS pct_NonPay,
        100.0*SUM(CASE WHEN is_joint=1 AND tp_sat<=3 AND tp_other=1    THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tp_sat IS NOT NULL AND tp_sat<=3 THEN 1 ELSE 0 END),0) AS pct_Other
    FROM src GROUP BY yearweek, weeknumber
)
SELECT * FROM all_snapp
UNION ALL SELECT * FROM jnt_snapp
UNION ALL SELECT * FROM jnt_tapsi;
GO


-- RA-14. NAVIGATION APP USAGE BY CITY (Page #14)
-- Long format: platform = Snapp / Tapsi; pct by app; Tapsi base = joint drivers only
-- N-cutoff: n
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_Navigation] AS
WITH src AS (
    SELECT yearweek, weeknumber, city,
        TRY_CAST(active_joint AS INT) AS is_joint,
        snapp_last_trip_navigation,
        tapsi_navigation_type
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL
),
snapp_nav AS (
    SELECT
        CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort, weeknumber, city, 'Snapp' AS platform,
        SUM(CASE WHEN snapp_last_trip_navigation IS NOT NULL THEN 1 ELSE 0 END) AS n,
        100.0*SUM(CASE WHEN snapp_last_trip_navigation='Neshan'            THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN snapp_last_trip_navigation IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_Neshan,
        100.0*SUM(CASE WHEN snapp_last_trip_navigation='Balad'             THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN snapp_last_trip_navigation IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_Balad,
        100.0*SUM(CASE WHEN snapp_last_trip_navigation='No Navigation App' THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN snapp_last_trip_navigation IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_None,
        100.0*SUM(CASE WHEN snapp_last_trip_navigation='Google Map'        THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN snapp_last_trip_navigation IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_GoogleMap,
        100.0*SUM(CASE WHEN snapp_last_trip_navigation='Waze'              THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN snapp_last_trip_navigation IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_Waze,
        CAST(NULL AS FLOAT) AS pct_InAppNav,
        100.0*SUM(CASE WHEN snapp_last_trip_navigation IS NOT NULL
                        AND snapp_last_trip_navigation NOT IN ('Neshan','Balad','No Navigation App','Google Map','Waze')
                       THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN snapp_last_trip_navigation IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_Other
    FROM src
    GROUP BY yearweek, weeknumber, city
),
tapsi_nav AS (
    SELECT
        CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
        yearweek AS yearweek_sort, weeknumber, city, 'Tapsi' AS platform,
        SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type IS NOT NULL THEN 1 ELSE 0 END) AS n,
        100.0*SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type='Neshan'            THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_Neshan,
        100.0*SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type='Balad'             THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_Balad,
        100.0*SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type='No Navigation App' THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_None,
        CAST(NULL AS FLOAT) AS pct_GoogleMap,
        CAST(NULL AS FLOAT) AS pct_Waze,
        100.0*SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type='In-App Navigation' THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_InAppNav,
        100.0*SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type IS NOT NULL
                        AND tapsi_navigation_type NOT IN ('Neshan','Balad','No Navigation App','In-App Navigation')
                       THEN 1.0 ELSE 0.0 END)/NULLIF(SUM(CASE WHEN is_joint=1 AND tapsi_navigation_type IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_Other
    FROM src
    GROUP BY yearweek, weeknumber, city
)
SELECT * FROM snapp_nav
UNION ALL SELECT * FROM tapsi_nav;
GO


-- RA-15. REFERRAL / JOINING BONUS BY CITY (Page #16)
-- Snapp joining bonus: all drivers; Tapsi joining bonus: joint drivers only
-- N-cutoff: n_Snapp, n_Tapsi
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_Referral] AS
WITH src AS (
    SELECT yearweek, weeknumber, city,
        TRY_CAST(active_joint AS INT) AS is_joint,
        snapp_joining_bonus,
        tapsi_joining_bonus
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL
)
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    weeknumber, city,
    SUM(CASE WHEN snapp_joining_bonus IS NOT NULL THEN 1 ELSE 0 END) AS n_Snapp,
    SUM(CASE WHEN snapp_joining_bonus='Yes' THEN 1 ELSE 0 END) AS joining_Snapp,
    100.0*SUM(CASE WHEN snapp_joining_bonus='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN snapp_joining_bonus IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_Joining_Snapp,
    SUM(CASE WHEN is_joint=1 AND tapsi_joining_bonus IS NOT NULL THEN 1 ELSE 0 END) AS n_Tapsi,
    SUM(CASE WHEN is_joint=1 AND tapsi_joining_bonus='Yes' THEN 1 ELSE 0 END) AS joining_Tapsi,
    100.0*SUM(CASE WHEN is_joint=1 AND tapsi_joining_bonus='Yes' THEN 1.0 ELSE 0.0 END)
        / NULLIF(SUM(CASE WHEN is_joint=1 AND tapsi_joining_bonus IS NOT NULL THEN 1 ELSE 0 END),0) AS pct_Joining_Tapsi
FROM src
GROUP BY yearweek, weeknumber, city;
GO


-- RA-16. TAPSI INACTIVITY BEFORE INCENTIVE BY CITY (Page #17)
-- Joint drivers only; long format with inactivity bucket + sort key
-- N-cutoff: n (per bucket)
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_TapsiInactivity] AS
WITH src AS (
    SELECT yearweek, weeknumber, city,
        TRY_CAST(active_joint AS INT) AS is_joint,
        tapsi_inactive_b4_incentive
    FROM [Cab].[DriverSurvey_ShortMain]
    WHERE city IS NOT NULL
)
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    weeknumber, city,
    tapsi_inactive_b4_incentive AS inactivity_bucket,
    CASE tapsi_inactive_b4_incentive
        WHEN 'Same Day'           THEN 1
        WHEN '1_3 Day Before'     THEN 2
        WHEN '3_7 Days Before'    THEN 3
        WHEN '8_14 Days Before'   THEN 4
        WHEN '15_30 Days_Before'  THEN 5
        WHEN '1_2 Month Before'   THEN 6
        WHEN '2_3 Month Before'   THEN 7
        WHEN '3_6Month Before'    THEN 8
        WHEN '>6 Month Before'    THEN 9
        ELSE 99
    END AS bucket_sort,
    COUNT(*) AS n,
    SUM(COUNT(*)) OVER (PARTITION BY yearweek, city) AS n_total
FROM src
WHERE is_joint = 1
  AND tapsi_inactive_b4_incentive IS NOT NULL
GROUP BY yearweek, weeknumber, city, tapsi_inactive_b4_incentive;
GO


-- RA-17. LUCKY WHEEL USAGE (Page #19)
-- wheel column stores Rial amount won; 0 = did not use
-- N-cutoff: n
GO
CREATE OR ALTER VIEW [Cab].[vw_RA_LuckyWheel] AS
SELECT
    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,
    yearweek AS yearweek_sort,
    weeknumber, city,
    COUNT(*) AS n,
    SUM(CASE WHEN TRY_CAST(wheel AS FLOAT) > 0 THEN 1 ELSE 0 END) AS n_users,
    100.0*SUM(CASE WHEN TRY_CAST(wheel AS FLOAT) > 0 THEN 1.0 ELSE 0.0 END) / NULLIF(COUNT(*),0) AS pct_usage,
    AVG(CASE WHEN TRY_CAST(wheel AS FLOAT) > 0 THEN TRY_CAST(wheel AS FLOAT) END) AS avg_wheel_amount
FROM [Cab].[DriverSurvey_ShortMain]
WHERE city IS NOT NULL
GROUP BY yearweek, weeknumber, city;
GO


PRINT 'All 37 views created successfully (20 dashboard views + 17 routine analysis views)!';
GO

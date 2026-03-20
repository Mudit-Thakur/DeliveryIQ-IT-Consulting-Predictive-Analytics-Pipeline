-- ============================================================
-- FILE: SQL/04_analytics_ctes.sql
-- PURPOSE: High-impact business analytics using CTEs.
--          Each query answers a specific business question.
-- ============================================================

USE ITConsultingDB;
GO

-- ══════════════════════════════════════════════════════════════
-- QUERY 1: DELAY + FORECAST ALERTS
-- Business Question: Which projects are delayed and which
-- are predicted to breach the client's SLA?
-- ══════════════════════════════════════════════════════════════
WITH DelayAnalysis AS (
    SELECT
        P.ProjectID,
        C.ClientName,
        P.Sector,
        P.ProjectType,
        -- Actual delay vs planned date
        DATEDIFF(DAY, P.PlannedEndDate, P.ActualEndDate) AS DelayDays,
        -- Predictive forecast delay
        P.ForecastDelay,
        -- Alert: 1 = forecast exceeds client SLA
        CASE
            WHEN P.ForecastDelay > C.SLA_Days THEN 1
            ELSE 0
        END AS ForecastAlert,
        C.SLA_Days
    FROM Projects P
    INNER JOIN Clients C ON P.ClientID = C.ClientID
    WHERE P.PlannedEndDate IS NOT NULL
      AND P.ActualEndDate  IS NOT NULL
),
HighRiskPerProject AS (
    SELECT
        ProjectID,
        COUNT(*) AS HighRiskCount
    FROM Risks
    WHERE RiskImpact = 'High'
    GROUP BY ProjectID
)
SELECT
    D.ProjectID,
    D.ClientName,
    D.Sector,
    D.ProjectType,
    D.DelayDays,
    D.ForecastDelay,
    D.ForecastAlert,
    D.SLA_Days,
    ISNULL(H.HighRiskCount, 0) AS HighRiskCount
FROM DelayAnalysis D
LEFT JOIN HighRiskPerProject H ON D.ProjectID = H.ProjectID
ORDER BY D.ForecastAlert DESC, D.DelayDays DESC;


-- ══════════════════════════════════════════════════════════════
-- QUERY 2: SECTOR ANALYSIS
-- Business Question: Which sectors have the highest average
-- delay and the most overdue projects?
-- ══════════════════════════════════════════════════════════════
WITH SectorStats AS (
    SELECT
        Sector,
        ProjectID,
        DATEDIFF(DAY, PlannedEndDate, ActualEndDate) AS DelayDays,
        -- OverdueFlag = 1 if delayed more than 10 days
        CASE
            WHEN DATEDIFF(DAY, PlannedEndDate, ActualEndDate) > 10
            THEN 1 ELSE 0
        END AS OverdueFlag
    FROM Projects
    WHERE PlannedEndDate IS NOT NULL
      AND ActualEndDate  IS NOT NULL
)
SELECT
    Sector,
    COUNT(ProjectID)                          AS TotalProjects,
    AVG(DelayDays)                            AS AvgDelayDays,
    SUM(OverdueFlag)                          AS OverdueProjects,
    -- Percentage of projects overdue
    ROUND(
        SUM(OverdueFlag) * 100.0 / COUNT(ProjectID), 1
    )                                         AS PercentOverdue
FROM SectorStats
GROUP BY Sector
ORDER BY AvgDelayDays DESC;


-- ══════════════════════════════════════════════════════════════
-- QUERY 3: CLIENT ANALYSIS
-- Business Question: Which clients have the highest delays
-- and budget overruns? (Only clients with 2+ projects)
-- ══════════════════════════════════════════════════════════════
WITH ClientStats AS (
    SELECT
        P.ProjectID,
        C.ClientName,
        C.Industry,
        C.Region,
        DATEDIFF(DAY, P.PlannedEndDate, P.ActualEndDate) AS DelayDays,
        -- Budget overrun as % of budget
        CASE
            WHEN P.Budget > 0
            THEN ROUND(P.SpentHours * 100.0 / P.Budget, 2)
            ELSE 0
        END AS BudgetOverrunPercent
    FROM Projects P
    INNER JOIN Clients C ON P.ClientID = C.ClientID
    WHERE P.PlannedEndDate IS NOT NULL
      AND P.ActualEndDate  IS NOT NULL
)
SELECT
    ClientName,
    Industry,
    Region,
    COUNT(ProjectID)             AS NumProjects,
    AVG(DelayDays)               AS AvgDelayDays,
    ROUND(AVG(BudgetOverrunPercent), 1) AS AvgBudgetOverrunPct
FROM ClientStats
GROUP BY ClientName, Industry, Region
HAVING COUNT(ProjectID) > 1       -- only clients with multiple projects
ORDER BY AvgDelayDays DESC;


-- ══════════════════════════════════════════════════════════════
-- QUERY 4: PROJECT TYPE vs SECTOR ANALYSIS
-- Business Question: Which project type + sector combinations
-- have the worst delivery performance?
-- ══════════════════════════════════════════════════════════════
WITH ProjectTypeSector AS (
    SELECT
        ProjectType,
        Sector,
        ProjectID,
        DATEDIFF(DAY, PlannedEndDate, ActualEndDate) AS DelayDays,
        CASE
            WHEN DATEDIFF(DAY, PlannedEndDate, ActualEndDate) > 10
            THEN 1 ELSE 0
        END AS OverdueFlag
    FROM Projects
    WHERE PlannedEndDate IS NOT NULL
      AND ActualEndDate  IS NOT NULL
)
SELECT
    ProjectType,
    Sector,
    COUNT(ProjectID)   AS NumProjects,
    AVG(DelayDays)     AS AvgDelayDays,
    SUM(OverdueFlag)   AS OverdueProjects
FROM ProjectTypeSector
GROUP BY ProjectType, Sector
ORDER BY AvgDelayDays DESC, OverdueProjects DESC;


-- ══════════════════════════════════════════════════════════════
-- QUERY 5: EMPLOYEE UTILIZATION
-- Business Question: Which employees are overloaded (>120%)
-- or under-utilized (<80%)?
-- ══════════════════════════════════════════════════════════════
WITH EmployeeStats AS (
    SELECT
        E.EmployeeID,
        E.Name,
        E.Role,
        E.Skill,
        E.Location,
        T.ProjectID,
        SUM(T.SpentHours)    AS TotalSpentPerProject,
        SUM(T.AssignedHours) AS TotalAssignedPerProject
    FROM Teams T
    INNER JOIN Employees E ON T.EmployeeID = E.EmployeeID
    GROUP BY E.EmployeeID, E.Name, E.Role, E.Skill, E.Location, T.ProjectID
)
SELECT
    EmployeeID,
    Name,
    Role,
    Skill,
    Location,
    COUNT(ProjectID)                                           AS NumProjects,
    SUM(TotalSpentPerProject)                                  AS TotalSpentHours,
    SUM(TotalAssignedPerProject)                               AS TotalAssignedHours,
    -- Utilization ratio: >1 = overloaded, <1 = underutilized
    ROUND(
        CAST(SUM(TotalSpentPerProject) AS FLOAT)
        / NULLIF(SUM(TotalAssignedPerProject), 0),
        2
    )                                                          AS UtilizationRatio,
    -- Label for dashboard filtering
    CASE
        WHEN CAST(SUM(TotalSpentPerProject) AS FLOAT)
             / NULLIF(SUM(TotalAssignedPerProject), 0) > 1.2
        THEN 'Overloaded'
        WHEN CAST(SUM(TotalSpentPerProject) AS FLOAT)
             / NULLIF(SUM(TotalAssignedPerProject), 0) < 0.8
        THEN 'Under-Utilized'
        ELSE 'Optimal'
    END                                                        AS UtilizationStatus
FROM EmployeeStats
GROUP BY EmployeeID, Name, Role, Skill, Location
ORDER BY UtilizationRatio DESC;


-- ══════════════════════════════════════════════════════════════
-- QUERY 6: RISK SUMMARY
-- Business Question: Which projects have the most high-impact
-- risks and need immediate executive attention?
-- ══════════════════════════════════════════════════════════════
WITH RiskSummary AS (
    SELECT
        P.ProjectID,
        C.ClientName,
        P.Sector,
        P.ProjectType,
        COUNT(R.RiskType)                                  AS TotalRisks,
        SUM(CASE WHEN R.RiskImpact = 'High'   THEN 1 ELSE 0 END) AS HighRisks,
        SUM(CASE WHEN R.RiskImpact = 'Medium' THEN 1 ELSE 0 END) AS MediumRisks,
        SUM(CASE WHEN R.RiskStatus = 'Open'   THEN 1 ELSE 0 END) AS OpenRisks
    FROM Projects P
    INNER JOIN Clients C ON P.ClientID = C.ClientID
    LEFT JOIN  Risks R   ON P.ProjectID = R.ProjectID
    GROUP BY P.ProjectID, C.ClientName, P.Sector, P.ProjectType
)
SELECT *
FROM RiskSummary
WHERE HighRisks > 0
ORDER BY HighRisks DESC, TotalRisks DESC;


-- ══════════════════════════════════════════════════════════════
-- QUERY 7: MONTHLY DELAY TREND (ACTUAL vs FORECAST)
-- Business Question: Is project delivery performance improving
-- or worsening over time?
-- ══════════════════════════════════════════════════════════════
WITH MonthlyTrend AS (
    SELECT
        Sector,
        ProjectID,
        DATEPART(YEAR,  PlannedEndDate) AS PlanYear,
        DATEPART(MONTH, PlannedEndDate) AS PlanMonth,
        DATEDIFF(DAY, PlannedEndDate, ActualEndDate) AS DelayDays,
        ForecastDelay
    FROM Projects
    WHERE PlannedEndDate IS NOT NULL
      AND ActualEndDate  IS NOT NULL
)
SELECT
    Sector,
    PlanYear,
    PlanMonth,
    -- Friendly month-year label for Power BI
    CONCAT(
        RIGHT('0' + CAST(PlanMonth AS VARCHAR), 2),
        '-', PlanYear
    )                          AS MonthYear,
    COUNT(ProjectID)           AS NumProjects,
    AVG(DelayDays)             AS AvgActualDelay,
    AVG(ForecastDelay)         AS AvgForecastDelay
FROM MonthlyTrend
GROUP BY Sector, PlanYear, PlanMonth
ORDER BY PlanYear, PlanMonth, AvgActualDelay DESC;


-- ══════════════════════════════════════════════════════════════
-- QUERY 8: TOP 10 FORECAST ALERT PROJECTS
-- Business Question: Which projects must be escalated NOW
-- because forecast delay exceeds their SLA?
-- ══════════════════════════════════════════════════════════════
WITH AlertProjects AS (
    SELECT
        P.ProjectID,
        C.ClientName,
        P.Sector,
        P.ProjectType,
        P.ForecastDelay,
        C.SLA_Days,
        -- Days over SLA
        P.ForecastDelay - C.SLA_Days AS DaysOverSLA
    FROM Projects P
    INNER JOIN Clients C ON P.ClientID = C.ClientID
    WHERE P.ForecastDelay > C.SLA_Days
)
SELECT TOP 10
    ProjectID,
    ClientName,
    Sector,
    ProjectType,
    ForecastDelay,
    SLA_Days,
    DaysOverSLA
FROM AlertProjects
ORDER BY DaysOverSLA DESC;


-- ══════════════════════════════════════════════════════════════
-- QUERY 9: EMPLOYEE OVERLOAD ALERTS
-- Business Question: Which employees are at risk of burnout
-- or causing project delays due to overallocation?
-- ══════════════════════════════════════════════════════════════
WITH EmployeeLoad AS (
    SELECT
        T.EmployeeID,
        E.Name,
        E.Role,
        E.Location,
        SUM(T.SpentHours)    AS TotalSpent,
        SUM(T.AssignedHours) AS TotalAssigned,
        ROUND(
            CAST(SUM(T.SpentHours) AS FLOAT)
            / NULLIF(SUM(T.AssignedHours), 0),
            2
        ) AS UtilizationRatio
    FROM Teams T
    INNER JOIN Employees E ON T.EmployeeID = E.EmployeeID
    GROUP BY T.EmployeeID, E.Name, E.Role, E.Location
)
SELECT
    EmployeeID,
    Name,
    Role,
    Location,
    TotalSpent,
    TotalAssigned,
    UtilizationRatio
FROM EmployeeLoad
WHERE UtilizationRatio > 1   -- 100% threshold = overloaded
ORDER BY UtilizationRatio DESC;

PRINT '✅ All 9 analytics queries are ready for Power BI.';
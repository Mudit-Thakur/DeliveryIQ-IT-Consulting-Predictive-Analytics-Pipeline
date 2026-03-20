-- Active: 1772513754324@@127.0.0.1@1433
-- ============================================================
-- FILE: SQL/05_create_views.sql
-- PURPOSE: Create SQL views from the CTE queries above.
--          Power BI connects to these views directly.
-- ============================================================

USE ITConsultingDB;
GO

-- ── View 1: Project Delay + Forecast Alerts ──────────────────
CREATE OR ALTER VIEW vw_ProjectForecast AS
WITH DelayAnalysis AS (
    SELECT
        P.ProjectID,
        C.ClientName,
        P.Sector,
        P.ProjectType,
        DATEDIFF(DAY, P.PlannedEndDate, P.ActualEndDate) AS DelayDays,
        P.ForecastDelay,
        CASE
            WHEN P.ForecastDelay > C.SLA_Days THEN 1 ELSE 0
        END AS ForecastAlert,
        C.SLA_Days
    FROM Projects P
    INNER JOIN Clients C ON P.ClientID = C.ClientID
    WHERE P.PlannedEndDate IS NOT NULL
      AND P.ActualEndDate  IS NOT NULL
),
HighRisk AS (
    SELECT ProjectID, COUNT(*) AS HighRiskCount
    FROM Risks
    WHERE RiskImpact = 'High'
    GROUP BY ProjectID
)
SELECT
    D.*,
    ISNULL(H.HighRiskCount, 0) AS HighRiskCount
FROM DelayAnalysis D
LEFT JOIN HighRisk H ON D.ProjectID = H.ProjectID;
GO

-- ── View 2: Sector Summary ───────────────────────────────────
CREATE OR ALTER VIEW vw_SectorSummary AS
SELECT
    Sector,
    COUNT(ProjectID)    AS TotalProjects,
    AVG(DelayDays)      AS AvgDelayDays,
    SUM(CASE WHEN DelayDays > 10 THEN 1 ELSE 0 END) AS OverdueProjects,
    ROUND(
        SUM(CASE WHEN DelayDays > 10 THEN 1 ELSE 0 END)
        * 100.0 / COUNT(ProjectID), 1
    ) AS PercentOverdue
FROM Projects
WHERE PlannedEndDate IS NOT NULL
  AND ActualEndDate  IS NOT NULL
GROUP BY Sector;
GO

-- ── View 3: Employee Utilization ─────────────────────────────
CREATE OR ALTER VIEW vw_EmployeeUtilization AS
SELECT
    T.EmployeeID,
    E.Name,
    E.Role,
    E.Skill,
    E.Location,
    COUNT(T.ProjectID)   AS NumProjects,
    SUM(T.SpentHours)    AS TotalSpent,
    SUM(T.AssignedHours) AS TotalAssigned,
    ROUND(
        CAST(SUM(T.SpentHours) AS FLOAT)
        / NULLIF(SUM(T.AssignedHours), 0),
        2
    ) AS UtilizationRatio
FROM Teams T
INNER JOIN Employees E ON T.EmployeeID = E.EmployeeID
GROUP BY T.EmployeeID, E.Name, E.Role, E.Skill, E.Location;
GO

-- ── View 4: Monthly Trend ────────────────────────────────────
CREATE OR ALTER VIEW vw_MonthlyTrend AS
SELECT
    Sector,
    DATEPART(YEAR,  PlannedEndDate) AS PlanYear,
    DATEPART(MONTH, PlannedEndDate) AS PlanMonth,
    CONCAT(
        RIGHT('0' + CAST(DATEPART(MONTH, PlannedEndDate) AS VARCHAR), 2),
        '-', DATEPART(YEAR, PlannedEndDate)
    ) AS MonthYear,
    COUNT(ProjectID)   AS NumProjects,
    AVG(DelayDays)     AS AvgActualDelay,
    AVG(ForecastDelay) AS AvgForecastDelay
FROM Projects
WHERE PlannedEndDate IS NOT NULL
  AND ActualEndDate  IS NOT NULL
GROUP BY
    Sector,
    DATEPART(YEAR,  PlannedEndDate),
    DATEPART(MONTH, PlannedEndDate);
GO

PRINT '✅ All 4 views created. Connect Power BI to ITConsultingDB.';
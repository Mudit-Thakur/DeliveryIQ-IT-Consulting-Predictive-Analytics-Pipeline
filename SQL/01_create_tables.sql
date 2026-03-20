-- ============================================================
-- FILE: SQL/01_create_tables.sql
-- PURPOSE: Create the ITConsultingDB database and all 5 tables
-- RUN IN: SSMS connected to your MSSQL Server instance
-- ============================================================

-- Step 1: Create the database
CREATE DATABASE ITConsultingDB;
GO

-- Step 2: Switch to the new database
USE ITConsultingDB;
GO

-- ──────────────────────────────────────────────────────────────
-- TABLE 1: Projects
-- Core project details including dates, budget, hours, and
-- our predictive ForecastDelay column.
-- ──────────────────────────────────────────────────────────────
CREATE TABLE Projects (
    ProjectID       INT             NOT NULL,
    ClientID        INT             NOT NULL,
    Sector          VARCHAR(50)     NOT NULL,
    ProjectType     VARCHAR(100)    NOT NULL,
    StartDate       DATE            NULL,
    PlannedEndDate  DATE            NULL,
    ActualEndDate   DATE            NULL,
    DelayDays       INT             NULL,
    Budget          FLOAT           NOT NULL,
    SpentHours      INT             NOT NULL,
    ForecastDelay   INT             NULL
);

-- ──────────────────────────────────────────────────────────────
-- TABLE 2: Clients
-- Client company profile, region, contract value, and SLA.
-- ──────────────────────────────────────────────────────────────
CREATE TABLE Clients (
    ClientID        INT             NOT NULL,
    ClientName      VARCHAR(200)    NOT NULL,
    Industry        VARCHAR(50)     NOT NULL,
    Region          VARCHAR(50)     NOT NULL,
    ContractValue   FLOAT           NOT NULL,
    SLA_Days        INT             NOT NULL
);

-- ──────────────────────────────────────────────────────────────
-- TABLE 3: Employees
-- Employee skill profile, role, experience, and location.
-- ──────────────────────────────────────────────────────────────
CREATE TABLE Employees (
    EmployeeID      INT             NOT NULL,
    Name            VARCHAR(200)    NOT NULL,
    Role            VARCHAR(100)    NOT NULL,
    ExperienceYears INT             NOT NULL,
    Skill           VARCHAR(50)     NOT NULL,
    Location        VARCHAR(100)    NOT NULL
);

-- ──────────────────────────────────────────────────────────────
-- TABLE 4: Teams
-- Maps employees to projects with their assigned vs spent hours.
-- ──────────────────────────────────────────────────────────────
CREATE TABLE Teams (
    ProjectID       INT             NOT NULL,
    EmployeeID      INT             NOT NULL,
    Role            VARCHAR(100)    NOT NULL,
    AssignedHours   INT             NOT NULL,
    SpentHours      INT             NOT NULL
);

-- ──────────────────────────────────────────────────────────────
-- TABLE 5: Risks
-- Risk log per project with type, impact level, and status.
-- ──────────────────────────────────────────────────────────────
CREATE TABLE Risks (
    ProjectID       INT             NOT NULL,
    RiskType        VARCHAR(100)    NOT NULL,
    RiskImpact      VARCHAR(50)     NOT NULL,
    RiskStatus      VARCHAR(50)     NOT NULL
);

PRINT '✅ All 5 tables created in ITConsultingDB.';
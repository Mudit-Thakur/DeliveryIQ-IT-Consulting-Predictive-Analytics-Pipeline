-- ============================================================
-- FILE: SQL/03_add_keys.sql
-- PURPOSE: Add Primary Keys and Foreign Keys to enforce
--          data integrity and enable Power BI relationships.
-- ============================================================

USE ITConsultingDB;
GO

-- ──────────────────────────────────────────────────────────────
-- PRIMARY KEYS
-- A primary key uniquely identifies each row in a table.
-- ──────────────────────────────────────────────────────────────

ALTER TABLE Projects
    ADD CONSTRAINT PK_Projects PRIMARY KEY (ProjectID);

ALTER TABLE Clients
    ADD CONSTRAINT PK_Clients PRIMARY KEY (ClientID);

ALTER TABLE Employees
    ADD CONSTRAINT PK_Employees PRIMARY KEY (EmployeeID);

-- Teams has a composite primary key: ProjectID + EmployeeID
-- An employee can only appear once per project
ALTER TABLE Teams
    ADD CONSTRAINT PK_Teams PRIMARY KEY (ProjectID, EmployeeID);

PRINT 'Primary keys added.';

-- ──────────────────────────────────────────────────────────────
-- FOREIGN KEYS
-- A foreign key links one table to another.
-- This enforces referential integrity (no orphan records).
-- ──────────────────────────────────────────────────────────────

-- Projects.ClientID → Clients.ClientID
-- A project must belong to a valid client
ALTER TABLE Projects
    ADD CONSTRAINT FK_Project_Client
    FOREIGN KEY (ClientID) REFERENCES Clients(ClientID);

-- Teams.ProjectID → Projects.ProjectID
-- A team assignment must belong to a valid project
ALTER TABLE Teams
    ADD CONSTRAINT FK_Team_Project
    FOREIGN KEY (ProjectID) REFERENCES Projects(ProjectID);

-- Teams.EmployeeID → Employees.EmployeeID
-- A team member must be a valid employee
ALTER TABLE Teams
    ADD CONSTRAINT FK_Team_Employee
    FOREIGN KEY (EmployeeID) REFERENCES Employees(EmployeeID);

-- Risks.ProjectID → Projects.ProjectID
-- A risk must belong to a valid project
ALTER TABLE Risks
    ADD CONSTRAINT FK_Risk_Project
    FOREIGN KEY (ProjectID) REFERENCES Projects(ProjectID);

PRINT '✅ All Primary Keys and Foreign Keys created.';
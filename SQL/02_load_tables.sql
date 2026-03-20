-- ============================================================
-- FILE: SQL/02_load_tables.sql
-- NOTE: Data loading is now handled by Python/load_to_sql.py
--       BULK INSERT has been removed — it was too fragile
--       for CSVs with commas in text fields and Windows
--       line endings (\r\n).
--
-- RUN THIS FILE only to verify row counts after running
-- the Python loader script.
-- ============================================================

USE ITConsultingDB;
GO

-- ──────────────────────────────────────────────────────────────
-- VERIFY: Run this after python Python/load_to_sql.py
-- Expected:
--   Projects  ~100 rows
--   Clients    20  rows
--   Employees  50  rows
--   Teams     ~600 rows
--   Risks     ~150 rows
-- ──────────────────────────────────────────────────────────────
SELECT 'Projects'  AS TableName, COUNT(*) AS TotalRows FROM Projects
UNION ALL
SELECT 'Clients',                COUNT(*)              FROM Clients
UNION ALL
SELECT 'Employees',              COUNT(*)              FROM Employees
UNION ALL
SELECT 'Teams',                  COUNT(*)              FROM Teams
UNION ALL
SELECT 'Risks',                  COUNT(*)              FROM Risks;
GO
# ============================================================
# FILE: Python/clean_explore.py
# PURPOSE: Clean messy CSVs, explore data, and print
#          high-level business insights.
# RUN THIS: After generate_data.py
# ============================================================

import pandas as pd

# ── Load all CSVs ─────────────────────────────────────────────
print("Loading CSVs...")

projects = pd.read_csv(
    "CSV/projects.csv",
    parse_dates=["StartDate", "PlannedEndDate", "ActualEndDate"]
)
clients   = pd.read_csv("CSV/clients.csv")
employees = pd.read_csv("CSV/employees.csv")
teams     = pd.read_csv("CSV/teams.csv")
risks     = pd.read_csv("CSV/risks.csv")

print(f"  Projects:  {projects.shape}")
print(f"  Clients:   {clients.shape}")
print(f"  Employees: {employees.shape}")
print(f"  Teams:     {teams.shape}")
print(f"  Risks:     {risks.shape}")

# ──────────────────────────────────────────────────────────────
# STEP 1: REMOVE DUPLICATES
# ──────────────────────────────────────────────────────────────
print("\n--- STEP 1: Remove Duplicates ---")

before = len(projects)
projects   = projects.drop_duplicates(subset=["ProjectID"])
teams      = teams.drop_duplicates(subset=["ProjectID", "EmployeeID"])
print(f"  Projects: removed {before - len(projects)} duplicates")

# ──────────────────────────────────────────────────────────────
# STEP 2: HANDLE MISSING VALUES
# ──────────────────────────────────────────────────────────────
print("\n--- STEP 2: Handle Missing Values ---")

print("  Missing values before cleaning:")
print(projects[["StartDate","PlannedEndDate","ActualEndDate"]].isnull().sum())

# Fill missing StartDate with the minimum start date
projects["StartDate"] = projects["StartDate"].fillna(
    projects["StartDate"].min()
)
# Fill missing PlannedEndDate with StartDate + 90 days
projects["PlannedEndDate"] = projects["PlannedEndDate"].fillna(
    projects["StartDate"] + pd.Timedelta(days=90)
)
# Fill missing ActualEndDate with PlannedEndDate (assume on time)
projects["ActualEndDate"] = projects["ActualEndDate"].fillna(
    projects["PlannedEndDate"]
)

print("  Missing values after cleaning:")
print(projects[["StartDate","PlannedEndDate","ActualEndDate"]].isnull().sum())

# ──────────────────────────────────────────────────────────────
# STEP 3: RECALCULATE DELAY (in case it was affected by missing data)
# ──────────────────────────────────────────────────────────────
print("\n--- STEP 3: Recalculate Delay Days ---")

projects["DelayDays"] = (
    projects["ActualEndDate"] - projects["PlannedEndDate"]
).dt.days

print("  DelayDays recalculated.")

# ──────────────────────────────────────────────────────────────
# STEP 4: BUSINESS INSIGHTS (Print in console)
# ──────────────────────────────────────────────────────────────
print("\n====================================")
print(" BUSINESS INSIGHTS (Pre-SQL)")
print("====================================")

# Insight 1: Average delay by sector
print("\n[Insight 1] Average Delay by Sector:")
print(
    projects.groupby("Sector")["DelayDays"]
    .mean()
    .sort_values(ascending=False)
    .round(1)
    .to_string()
)

# Insight 2: % of overdue projects
total = len(projects)
overdue = (projects["DelayDays"] > 10).sum()
pct_overdue = round(overdue / total * 100, 1)
print(f"\n[Insight 2] % Overdue Projects (>10 days): {pct_overdue}%")

# Insight 3: Top 5 most delayed projects
print("\n[Insight 3] Top 5 Most Delayed Projects:")
print(
    projects[["ProjectID","Sector","ProjectType","DelayDays"]]
    .sort_values("DelayDays", ascending=False)
    .head(5)
    .to_string(index=False)
)

# Insight 4: Employee utilization
teams["Utilization"] = teams["SpentHours"] / teams["AssignedHours"]
overloaded = teams[teams["Utilization"] > 1.2]
print(f"\n[Insight 4] Overloaded Employees (>120% utilization): {overloaded['EmployeeID'].nunique()}")

# Insight 5: Risk counts
print("\n[Insight 5] Risk Distribution by Impact:")
print(risks["RiskImpact"].value_counts().to_string())

# ──────────────────────────────────────────────────────────────
# STEP 5: OVERWRITE CLEANED CSVs
# Cleaned data goes back into the CSV folder for SQL loading
# ──────────────────────────────────────────────────────────────
projects.to_csv("CSV/projects.csv",   index=False)
teams.to_csv("CSV/teams.csv",         index=False)

print("\n✅ Cleaned CSVs saved. Next step → SQL/01_create_tables.sql")
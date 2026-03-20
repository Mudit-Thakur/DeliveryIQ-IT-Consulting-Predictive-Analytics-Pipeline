# -*- coding: utf-8 -*-
# ============================================================
# FILE: Python/generate_data.py
# PURPOSE: Generate messy, realistic IT consulting datasets.
#
# KEY CHANGE vs previous version:
#   - Seed is now time-based so every run produces different data.
#   - Dataset size (projects / clients / employees) also randomizes.
#   - Saves CSV/run_metadata.json so the email script can read
#     exact KPIs and put them in the email body automatically.
#
# REPRODUCING A SPECIFIC RUN:
#   Every run prints its seed value.
#   To replay: set USE_FIXED_SEED = True and paste the seed into
#   FIXED_SEED below, then rerun.
# ============================================================

import sys
import io
import os
import time

# ── Force UTF-8 on Windows terminal ──────────────────────────
sys.stdout = io.TextIOWrapper(
    sys.stdout.buffer,
    encoding="utf-8",
    errors="replace",
    line_buffering=True,
)

import pandas as pd
import numpy as np
import random
import json
import datetime
from faker import Faker

# ══════════════════════════════════════════════════════════════
# SEED CONFIGURATION
#
# USE_FIXED_SEED = False  -->  new random data every single run
# USE_FIXED_SEED = True   -->  same data every run (for debugging)
# ══════════════════════════════════════════════════════════════

USE_FIXED_SEED = False
FIXED_SEED     = 42

if USE_FIXED_SEED:
    SEED = FIXED_SEED
    print("Seed mode  : FIXED ({})".format(SEED))
else:
    SEED = int(time.time())
    print("Seed mode  : DYNAMIC")
    print("Seed value : {}  <-- paste into FIXED_SEED to replay this run".format(SEED))

random.seed(SEED)
np.random.seed(SEED % (2**32))
fake = Faker()
Faker.seed(SEED)

# ──────────────────────────────────────────────────────────────
# DYNAMIC SCALE: dataset size changes every run
# ──────────────────────────────────────────────────────────────

NUM_PROJECTS  = random.randint(80,  150)
NUM_CLIENTS   = random.randint(15,  30)
NUM_EMPLOYEES = random.randint(40,  70)

print("Run config : {} projects / {} clients / {} employees".format(
    NUM_PROJECTS, NUM_CLIENTS, NUM_EMPLOYEES
))

# ── Output folder ─────────────────────────────────────────────
output_folder = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..", "CSV"
)
os.makedirs(output_folder, exist_ok=True)
print("CSV folder : ready")
print("")

# ──────────────────────────────────────────────────────────────
# TABLE 1: PROJECTS
# ──────────────────────────────────────────────────────────────

sectors = [
    "Auto", "Energy", "Transportation",
    "Logistics", "Finance", "BFSI", "IT"
]
project_types = [
    "Custom Development", "ERP Implementation",
    "Maintenance", "Migration", "Cloud Adoption"
]

projects = []

for pid in range(1, NUM_PROJECTS + 1):
    start_date  = fake.date_between(start_date="-2y", end_date="today")
    planned_end = start_date + pd.Timedelta(days=random.randint(30, 180))
    actual_end  = planned_end + pd.Timedelta(days=random.randint(-10, 60))
    delay_days  = (actual_end - planned_end).days
    budget      = round(random.uniform(100_000, 2_000_000), 2)
    spent_hours = random.randint(500, 10_000)

    # ForecastDelay simulates a predictive model output
    # In a real pipeline this comes from a trained regression model
    forecast_delay = delay_days + random.randint(-5, 10)

    projects.append({
        "ProjectID":      pid,
        "ClientID":       random.randint(1, NUM_CLIENTS),
        "Sector":         random.choice(sectors),
        "ProjectType":    random.choice(project_types),
        "StartDate":      start_date,
        "PlannedEndDate": planned_end,
        "ActualEndDate":  actual_end,
        "DelayDays":      delay_days,
        "Budget":         budget,
        "SpentHours":     spent_hours,
        "ForecastDelay":  forecast_delay,
    })

projects_df = pd.DataFrame(projects)

# Introduce messiness: ~5% missing dates
for col in ["StartDate", "PlannedEndDate", "ActualEndDate"]:
    n_missing = max(1, int(len(projects_df) * 0.05))
    idx = projects_df.sample(n=n_missing).index
    projects_df.loc[idx, col] = pd.NaT

# Random number of duplicate rows (3-7)
n_dupes = random.randint(3, 7)
projects_df = pd.concat(
    [projects_df, projects_df.sample(n=n_dupes)],
    ignore_index=True,
)
print("  Projects  : {} rows ({} base + {} duplicates)".format(
    len(projects_df), NUM_PROJECTS, n_dupes
))

# ──────────────────────────────────────────────────────────────
# TABLE 2: CLIENTS
# ──────────────────────────────────────────────────────────────

regions = ["North", "South", "East", "West"]

clients = []
for cid in range(1, NUM_CLIENTS + 1):
    clients.append({
        "ClientID":      cid,
        "ClientName":    fake.company(),
        "Industry":      random.choice(sectors),
        "Region":        random.choice(regions),
        "ContractValue": round(random.uniform(100_000, 5_000_000), 2),
        "SLA_Days":      random.choice([30, 60, 90, 120]),
    })

clients_df = pd.DataFrame(clients)
print("  Clients   : {} rows".format(len(clients_df)))

# ──────────────────────────────────────────────────────────────
# TABLE 3: EMPLOYEES
# ──────────────────────────────────────────────────────────────

roles     = ["Developer", "Tester", "Project Manager", "Business Analyst", "DevOps"]
skills    = ["Python", "SQL", "PowerBI", "Tableau", "Excel", "Java"]
locations = ["Bangalore", "Pune", "Hyderabad", "Chennai"]

employees = []
for eid in range(1, NUM_EMPLOYEES + 1):
    employees.append({
        "EmployeeID":      eid,
        "Name":            fake.name(),
        "Role":            random.choice(roles),
        "ExperienceYears": random.randint(1, 15),
        "Skill":           random.choice(skills),
        "Location":        random.choice(locations),
    })

employees_df = pd.DataFrame(employees)
print("  Employees : {} rows".format(len(employees_df)))

# ──────────────────────────────────────────────────────────────
# TABLE 4: TEAMS
# ──────────────────────────────────────────────────────────────

teams = []

for pid in range(1, NUM_PROJECTS + 1):
    team_size  = random.randint(3, min(10, NUM_EMPLOYEES))
    member_ids = random.sample(list(employees_df["EmployeeID"]), team_size)

    for mid in member_ids:
        assigned = random.randint(100, 500)
        spent    = max(0, assigned + random.randint(-50, 100))

        teams.append({
            "ProjectID":     pid,
            "EmployeeID":    mid,
            "Role":          employees_df.loc[
                                 employees_df["EmployeeID"] == mid,
                                 "Role"
                             ].values[0],
            "AssignedHours": assigned,
            "SpentHours":    spent,
        })

teams_df = pd.DataFrame(teams)

n_team_dupes = random.randint(5, 15)
teams_df = pd.concat(
    [teams_df, teams_df.sample(n=n_team_dupes)],
    ignore_index=True,
)
print("  Teams     : {} rows ({} base + {} duplicates)".format(
    len(teams_df), len(teams_df) - n_team_dupes, n_team_dupes
))

# ──────────────────────────────────────────────────────────────
# TABLE 5: RISKS
# ──────────────────────────────────────────────────────────────

risk_types    = ["Schedule Risk", "Budget Risk", "Technical Risk", "Client Risk"]
impact_levels = ["Low", "Medium", "High"]
statuses      = ["Open", "Closed", "Mitigated"]

risks = []
for pid in range(1, NUM_PROJECTS + 1):
    for _ in range(random.choice([0, 1, 2, 3])):
        risks.append({
            "ProjectID":  pid,
            "RiskType":   random.choice(risk_types),
            "RiskImpact": random.choice(impact_levels),
            "RiskStatus": random.choice(statuses),
        })

risks_df = pd.DataFrame(risks)
print("  Risks     : {} rows".format(len(risks_df)))

# ──────────────────────────────────────────────────────────────
# SAVE ALL CSVs
# lineterminator="\n" prevents Windows CRLF issues in MSSQL
# ──────────────────────────────────────────────────────────────

save_args = dict(index=False, encoding="utf-8", lineterminator="\n")

projects_df.to_csv(  os.path.join(output_folder, "projects.csv"),  **save_args)
clients_df.to_csv(   os.path.join(output_folder, "clients.csv"),   **save_args)
employees_df.to_csv( os.path.join(output_folder, "employees.csv"), **save_args)
teams_df.to_csv(     os.path.join(output_folder, "teams.csv"),     **save_args)
risks_df.to_csv(     os.path.join(output_folder, "risks.csv"),     **save_args)

# ──────────────────────────────────────────────────────────────
# SAVE RUN METADATA (JSON)
#
# The pipeline reads this file to build the email body
# automatically. Every value here ends up in the email.
# ──────────────────────────────────────────────────────────────

valid_projects = projects_df.dropna(subset=["DelayDays"])
overdue_count  = int((valid_projects["DelayDays"] > 10).sum())

metadata = {
    "run_timestamp":    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "seed":             SEED,
    "num_projects":     NUM_PROJECTS,
    "num_clients":      NUM_CLIENTS,
    "num_employees":    NUM_EMPLOYEES,
    "total_risks":      len(risks_df),
    "avg_delay_days":   round(float(valid_projects["DelayDays"].mean()), 1),
    "pct_overdue":      round(overdue_count / NUM_PROJECTS * 100, 1),
    "overdue_count":    overdue_count,
    "high_risk_count":  int((risks_df["RiskImpact"] == "High").sum()),
    "open_risk_count":  int((risks_df["RiskStatus"] == "Open").sum()),
    "forecast_alerts":  int(
        (projects_df["ForecastDelay"] > 30).sum()
    ),
    "top_delayed_sector": str(
        valid_projects.groupby("Sector")["DelayDays"]
        .mean()
        .idxmax()
    ),
    "top_delayed_type": str(
        valid_projects.groupby("ProjectType")["DelayDays"]
        .mean()
        .idxmax()
    ),
}

meta_path = os.path.join(output_folder, "run_metadata.json")
with open(meta_path, "w", encoding="utf-8") as f:
    json.dump(metadata, f, indent=2)

print("")
print("Metadata saved --> CSV/run_metadata.json")
print("All 5 CSVs     --> CSV/ folder")
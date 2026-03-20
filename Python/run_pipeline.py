# -*- coding: utf-8 -*-
# ============================================================
# FILE: Python/run_pipeline.py
# PURPOSE: Run the full pipeline in one command.
#
#   Step 1 -> Generate dynamic datasets (different every run)
#   Step 2 -> Push DataFrames into MSSQL via SQLAlchemy
#   Step 3 -> Add keys + create analytics views
#   Step 4 -> Generate executive slide deck
#   Step 5 -> Send email with HTML report + PPTX attachment
#
# WHAT CHANGED:
#   - Data is now completely different every run (dynamic seed)
#   - Dataset size (projects/clients/employees) also randomizes
#   - Email step added at the end with HTML body from live KPIs
#   - All emoji replaced with [OK]/[WARN]/[FAIL] (Windows-safe)
#   - UTF-8 forced at top (prevents cp1252 crash on Windows)
#
# CONFIGURE BEFORE RUNNING:
#   1. SQL_SERVER  -> your MSSQL instance name (line ~60)
#   2. SEND_EMAIL  -> set True when ready to email (line ~70)
#   3. Open send_email.py and fill in your email credentials
# ============================================================

import sys
import io

# ── Force UTF-8 on Windows terminal ──────────────────────────
# Must be the FIRST executable line -- before any other import.
# Prevents UnicodeEncodeError on cp1252 Windows terminals.
sys.stdout = io.TextIOWrapper(
    sys.stdout.buffer,
    encoding="utf-8",
    errors="replace",
    line_buffering=True,
)

import subprocess
import os
import json
import pandas as pd

# ══════════════════════════════════════════════════════════════
# CONFIGURATION  <-- only edit this section
# ══════════════════════════════════════════════════════════════

# Your SQL Server instance name.
# Open SSMS. The name in the "Server name" login box is this value.
#
# Common values:
#   (localdb)\MSSQLLocalDB    <- LocalDB (Visual Studio default)
#   (localdb)\ProjectsV13     <- another LocalDB variant
#   localhost\SQLEXPRESS       <- SQL Server Express
#   localhost                  <- SQL Server Developer/Standard
#
# To find your LocalDB name, run in Command Prompt:
#   sqllocaldb info
SQL_SERVER   = r"(localdb)\MSSQLLocalDB"   # <-- update this
SQL_DATABASE = "your_database_name"

# Set to True once you have filled in email credentials
# in Python/send_email.py
SEND_EMAIL = False    # <-- change to True when ready

# ══════════════════════════════════════════════════════════════
# LOGGING  (plain ASCII -- no emoji)
# ══════════════════════════════════════════════════════════════

DIVIDER = "=" * 55

def banner(msg):
    print("\n" + DIVIDER)
    print("  " + msg)
    print(DIVIDER)

def step(n, total, msg):
    print("\n[{}/{}] {}".format(n, total, msg))

def ok(msg):
    print("  [OK]   " + msg)

def warn(msg):
    print("  [WARN] " + msg)

def fail(msg):
    print("  [FAIL] " + msg)

def info(msg):
    print("  " + msg)

# ══════════════════════════════════════════════════════════════
# HELPER: resolve path to a sibling script
# Works whether you run from Python/ or from the project root.
# ══════════════════════════════════════════════════════════════

def find_script(filename):
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(here, filename),
        os.path.join(here, "Python", filename),
    ]
    return next((p for p in candidates if os.path.exists(p)), None)

# ══════════════════════════════════════════════════════════════
# HELPER: run a child Python script as subprocess
# Passes PYTHONIOENCODING=utf-8 so child also uses UTF-8.
# Returns True on success, False on failure.
# ══════════════════════════════════════════════════════════════

def run_script(script_path, label):
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    result = subprocess.run(
        [sys.executable, script_path],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=os.path.dirname(script_path),
        env=env,
    )

    for line in result.stdout.strip().splitlines():
        info(line)

    if result.returncode != 0:
        fail("{} failed:".format(label))
        for line in result.stderr.strip().splitlines():
            info("  " + line)
        return False

    return True

# ══════════════════════════════════════════════════════════════
# HELPER: SQLAlchemy engine with multi-driver fallback
# ══════════════════════════════════════════════════════════════

def get_engine(server, database):
    from sqlalchemy import create_engine, text
    import urllib.parse

    drivers = [
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "ODBC Driver 11 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]

    last_error = None
    for driver in drivers:
        conn_str = (
            "DRIVER={{{d}}};"
            "SERVER={s};"
            "DATABASE={db};"
            "Trusted_Connection=yes;"
        ).format(d=driver, s=server, db=database)

        url = "mssql+pyodbc:///?odbc_connect=" + urllib.parse.quote_plus(conn_str)

        try:
            engine = create_engine(url, fast_executemany=True)
            with engine.connect() as conn:
                conn.execute(text("SELECT 1"))
            ok("ODBC driver: {}".format(driver))
            return engine
        except Exception as e:
            last_error = e
            continue

    raise RuntimeError("No ODBC driver worked. Last error: {}".format(last_error))

# ══════════════════════════════════════════════════════════════
# HELPER: DDL via raw pyodbc (CREATE/ALTER -- not data)
# ══════════════════════════════════════════════════════════════

def run_ddl(cursor, sql, label=""):
    try:
        cursor.execute(sql)
        if label:
            ok(label)
    except Exception as e:
        msg = str(e)
        ignorable = [
            "already an object named", "already exists",
            "There is already", "Violation of PRIMARY KEY",
        ]
        if any(x in msg for x in ignorable):
            if label:
                warn("{} -- already exists, skipped.".format(label))
        else:
            if label:
                warn("{}: {}".format(label, msg[:120]))

# ══════════════════════════════════════════════════════════════
# STEP 1: GENERATE DYNAMIC DATA
# ══════════════════════════════════════════════════════════════

def step1_generate():
    step(1, 5, "Generating dynamic datasets...")
    info("Every run produces different data (time-based seed).")

    script = find_script("generate_data.py")
    if not script:
        fail("Cannot find generate_data.py")
        sys.exit(1)

    success = run_script(script, "generate_data.py")
    if not success:
        sys.exit(1)

    ok("Dynamic datasets generated.")

# ══════════════════════════════════════════════════════════════
# STEP 2: LOAD DATAFRAMES INTO MSSQL
#
# Why DataFrame.to_sql() instead of BULK INSERT?
# -----------------------------------------------
# BULK INSERT reads files from SQL Server's file system, not
# Python's. Via pyodbc it always fails because:
#   - The path is resolved on SQL Server's machine
#   - GO is SSMS-only syntax, not valid SQL
#   - cursor.execute() rejects multi-statement batches
#
# Solution: pandas already has the data in memory.
# to_sql() hands it directly to SQLAlchemy as INSERT statements.
# No file paths. No GO. No batch problems. Re-runnable safely.
# ══════════════════════════════════════════════════════════════

def step2_load_mssql():
    step(2, 5, "Loading data into MSSQL ({})...".format(SQL_DATABASE))

    here    = os.path.dirname(os.path.abspath(__file__))
    csv_dir = os.path.join(here, "..", "CSV")
    if not os.path.isdir(csv_dir):
        csv_dir = os.path.join(here, "CSV")

    tables = [
        ("projects.csv",  "Projects"),
        ("clients.csv",   "Clients"),
        ("employees.csv", "Employees"),
        ("teams.csv",     "Teams"),
        ("risks.csv",     "Risks"),
    ]

    try:
        engine = get_engine(SQL_SERVER, SQL_DATABASE)
    except RuntimeError as e:
        fail("Cannot connect to SQL Server.")
        info("")
        info("Diagnose your instance:")
        info("  Open Command Prompt and run:  sqllocaldb info")
        info("  Copy the name and set SQL_SERVER at the top of this file.")
        info("")
        info("Current SQL_SERVER = '{}'".format(SQL_SERVER))
        info("")
        info("Common values:")
        info("  (localdb)\\MSSQLLocalDB   <- LocalDB default")
        info("  localhost\\SQLEXPRESS      <- SQL Express")
        info("  localhost                  <- Developer/Standard")
        info("")
        info("FALLBACK: import CSVs manually via SSMS:")
        info("  Right-click database -> Tasks -> Import Flat File")
        return False

    for csv_file, table_name in tables:
        csv_path = os.path.join(csv_dir, csv_file)
        if not os.path.exists(csv_path):
            warn("{} not found -- skipping {}.".format(csv_file, table_name))
            continue

        try:
            df = pd.read_csv(csv_path)
            for col in df.columns:
                if "Date" in col:
                    df[col] = pd.to_datetime(df[col], errors="coerce")

            # if_exists="replace" -> drop + recreate each time
            # Safe to rerun without manual truncation
            df.to_sql(
                name=table_name, con=engine,
                if_exists="replace", index=False,
                chunksize=500, schema="dbo",
            )
            ok("{}: {:,} rows.".format(table_name, len(df)))

        except Exception as e:
            warn("Failed to load {}: {}".format(table_name, str(e)[:120]))

    engine.dispose()
    return True

# ══════════════════════════════════════════════════════════════
# STEP 3: ADD KEYS + CREATE ANALYTICS VIEWS
# ══════════════════════════════════════════════════════════════

def step3_keys_and_views():
    step(3, 5, "Adding keys and creating analytics views...")

    import pyodbc

    drivers = [
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "ODBC Driver 11 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]

    conn = None
    for drv in drivers:
        try:
            conn = pyodbc.connect(
                "DRIVER={{{}}};"
                "SERVER={};"
                "DATABASE={};"
                "Trusted_Connection=yes;".format(drv, SQL_SERVER, SQL_DATABASE),
                autocommit=True,
            )
            break
        except Exception:
            continue

    if conn is None:
        warn("DDL connection failed. Run SQL/03_add_keys.sql manually in SSMS.")
        return

    cursor = conn.cursor()

    # Primary Keys
    for label, sql in [
        ("PK_Projects",  """IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
                             WHERE CONSTRAINT_NAME='PK_Projects')
                            ALTER TABLE Projects ADD CONSTRAINT PK_Projects PRIMARY KEY (ProjectID)"""),
        ("PK_Clients",   """IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
                             WHERE CONSTRAINT_NAME='PK_Clients')
                            ALTER TABLE Clients ADD CONSTRAINT PK_Clients PRIMARY KEY (ClientID)"""),
        ("PK_Employees", """IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
                             WHERE CONSTRAINT_NAME='PK_Employees')
                            ALTER TABLE Employees ADD CONSTRAINT PK_Employees PRIMARY KEY (EmployeeID)"""),
    ]:
        run_ddl(cursor, sql, label)

    # Foreign Keys
    for label, sql in [
        ("FK_Project_Client", """IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
                                  WHERE CONSTRAINT_NAME='FK_Project_Client')
                                 ALTER TABLE Projects ADD CONSTRAINT FK_Project_Client
                                 FOREIGN KEY (ClientID) REFERENCES Clients(ClientID)"""),
        ("FK_Team_Project",   """IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
                                  WHERE CONSTRAINT_NAME='FK_Team_Project')
                                 ALTER TABLE Teams ADD CONSTRAINT FK_Team_Project
                                 FOREIGN KEY (ProjectID) REFERENCES Projects(ProjectID)"""),
        ("FK_Risk_Project",   """IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
                                  WHERE CONSTRAINT_NAME='FK_Risk_Project')
                                 ALTER TABLE Risks ADD CONSTRAINT FK_Risk_Project
                                 FOREIGN KEY (ProjectID) REFERENCES Projects(ProjectID)"""),
    ]:
        run_ddl(cursor, sql, label)

    # Analytics views
    views = [
        ("vw_ProjectForecast", """
            CREATE OR ALTER VIEW vw_ProjectForecast AS
            WITH D AS (
                SELECT P.ProjectID, C.ClientName, P.Sector, P.ProjectType,
                       DATEDIFF(DAY,P.PlannedEndDate,P.ActualEndDate) AS DelayDays,
                       P.ForecastDelay,
                       CASE WHEN P.ForecastDelay > C.SLA_Days THEN 1 ELSE 0 END AS ForecastAlert,
                       C.SLA_Days
                FROM Projects P INNER JOIN Clients C ON P.ClientID=C.ClientID
                WHERE P.PlannedEndDate IS NOT NULL AND P.ActualEndDate IS NOT NULL
            ),
            H AS (
                SELECT ProjectID, COUNT(*) AS HighRiskCount
                FROM Risks WHERE RiskImpact='High' GROUP BY ProjectID
            )
            SELECT D.*, ISNULL(H.HighRiskCount,0) AS HighRiskCount
            FROM D LEFT JOIN H ON D.ProjectID=H.ProjectID"""),

        ("vw_SectorSummary", """
            CREATE OR ALTER VIEW vw_SectorSummary AS
            SELECT Sector,
                   COUNT(ProjectID) AS TotalProjects,
                   AVG(DelayDays)   AS AvgDelayDays,
                   SUM(CASE WHEN DelayDays>10 THEN 1 ELSE 0 END) AS OverdueProjects,
                   ROUND(SUM(CASE WHEN DelayDays>10 THEN 1 ELSE 0 END)*100.0/COUNT(ProjectID),1) AS PercentOverdue
            FROM Projects
            WHERE PlannedEndDate IS NOT NULL AND ActualEndDate IS NOT NULL
            GROUP BY Sector"""),

        ("vw_EmployeeUtilization", """
            CREATE OR ALTER VIEW vw_EmployeeUtilization AS
            SELECT T.EmployeeID, E.Name, E.Role, E.Skill, E.Location,
                   COUNT(T.ProjectID) AS NumProjects,
                   SUM(T.SpentHours)  AS TotalSpent,
                   SUM(T.AssignedHours) AS TotalAssigned,
                   ROUND(CAST(SUM(T.SpentHours) AS FLOAT)/NULLIF(SUM(T.AssignedHours),0),2) AS UtilizationRatio
            FROM Teams T INNER JOIN Employees E ON T.EmployeeID=E.EmployeeID
            GROUP BY T.EmployeeID, E.Name, E.Role, E.Skill, E.Location"""),

        ("vw_MonthlyTrend", """
            CREATE OR ALTER VIEW vw_MonthlyTrend AS
            SELECT Sector,
                   DATEPART(YEAR,PlannedEndDate)  AS PlanYear,
                   DATEPART(MONTH,PlannedEndDate) AS PlanMonth,
                   CONCAT(RIGHT('0'+CAST(DATEPART(MONTH,PlannedEndDate) AS VARCHAR),2),
                          '-',DATEPART(YEAR,PlannedEndDate)) AS MonthYear,
                   COUNT(ProjectID)   AS NumProjects,
                   AVG(DelayDays)     AS AvgActualDelay,
                   AVG(ForecastDelay) AS AvgForecastDelay
            FROM Projects
            WHERE PlannedEndDate IS NOT NULL AND ActualEndDate IS NOT NULL
            GROUP BY Sector, DATEPART(YEAR,PlannedEndDate), DATEPART(MONTH,PlannedEndDate)"""),
    ]

    for label, sql in views:
        run_ddl(cursor, sql.strip(), label)

    cursor.close()
    conn.close()
    ok("Keys and views ready.")

# ══════════════════════════════════════════════════════════════
# STEP 4: GENERATE SLIDE DECK
# ══════════════════════════════════════════════════════════════

def step4_slides():
    step(4, 5, "Generating executive slide deck...")

    script = find_script("generate_pptx.py")
    if not script:
        warn("generate_pptx.py not found. Skipping slide deck.")
        return False

    success = run_script(script, "generate_pptx.py")
    if success:
        ok("Executive_Summary.pptx saved.")
    return success

# ══════════════════════════════════════════════════════════════
# STEP 5: SEND EMAIL
# ══════════════════════════════════════════════════════════════

def step5_email():
    step(5, 5, "Sending email report...")

    if not SEND_EMAIL:
        info("SEND_EMAIL = False  -- skipping.")
        info("To enable: open run_pipeline.py and set SEND_EMAIL = True")
        info("Then fill in your credentials in Python/send_email.py")
        return

    script = find_script("send_email.py")
    if not script:
        warn("send_email.py not found. Skipping email.")
        return

    success = run_script(script, "send_email.py")
    if success:
        ok("Email dispatched.")

# ══════════════════════════════════════════════════════════════
# LOAD RUN METADATA FOR SUMMARY DISPLAY
# ══════════════════════════════════════════════════════════════

def load_meta_for_summary():
    here     = os.path.dirname(os.path.abspath(__file__))
    meta_path = os.path.join(here, "..", "CSV", "run_metadata.json")
    if not os.path.isdir(os.path.join(here, "..", "CSV")):
        meta_path = os.path.join(here, "CSV", "run_metadata.json")

    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    banner("IT CONSULTING ANALYTICS PIPELINE")
    info("SQL Server  : " + SQL_SERVER)
    info("Database    : " + SQL_DATABASE)
    info("Email step  : {}".format("ENABLED" if SEND_EMAIL else "disabled (SEND_EMAIL=False)"))

    step1_generate()

    loaded = step2_load_mssql()

    if loaded:
        step3_keys_and_views()

    step4_slides()

    step5_email()

    # ── Final summary ─────────────────────────────────────────
    meta = load_meta_for_summary()

    banner("PIPELINE COMPLETE")

    if meta:
        info("Run timestamp   : " + str(meta.get("run_timestamp", "N/A")))
        info("Seed            : " + str(meta.get("seed",           "N/A")))
        info("Projects        : " + str(meta.get("num_projects",   "N/A")))
        info("Clients         : " + str(meta.get("num_clients",    "N/A")))
        info("Employees       : " + str(meta.get("num_employees",  "N/A")))
        info("Avg delay       : " + str(meta.get("avg_delay_days", "N/A")) + " days")
        info("% Overdue       : " + str(meta.get("pct_overdue",    "N/A")) + "%")
        info("Forecast alerts : " + str(meta.get("forecast_alerts","N/A")))
        info("High risks      : " + str(meta.get("high_risk_count","N/A")))
    print()

    if loaded:
        info("MSSQL           : " + SQL_DATABASE + " on " + SQL_SERVER)
    else:
        info("MSSQL           : NOT loaded (see warnings above)")

    info("Slides          : Reports/Executive_Summary.pptx")
    print()
    info("NEXT STEPS:")
    info("  1. Open Power BI Desktop")
    info("  2. Home -> Get Data -> SQL Server")
    info("  3. Server   : " + SQL_SERVER)
    info("  4. Database : " + SQL_DATABASE)
    info("  5. Import mode -> load tables + 4 views")
    print()
    if not SEND_EMAIL:
        info("To enable email:")
        info("  1. Open Python/send_email.py")
        info("  2. Set EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECIPIENTS")
        info("  3. Set SEND_EMAIL = True in this file")
    print()
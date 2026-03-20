# ============================================================
# FILE: Python/load_to_sql.py
# FIX 1: numpy.int64 converted to Python int before insert
# FIX 2: CSV path is now relative to THIS script file, so
#         it works whether you run it from Python/ or from
#         the project root or anywhere else.
# ============================================================

import pandas as pd
import pyodbc
import os

# ──────────────────────────────────────────────────────────────
# PATH SETUP
# __file__  = full path to this script, e.g.:
#   C:\...\ITConsultingPipeline\Python\load_to_sql.py
# script_dir = the folder this script lives in (Python/)
# project_dir = one level up from Python/ = ITConsultingPipeline/
# csv_dir    = project_dir/CSV/
#
# This means no matter WHERE you run the script from
# (Python/ folder, project root, anywhere), it always
# finds the CSV folder correctly.
# ──────────────────────────────────────────────────────────────
script_dir  = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(script_dir)
csv_dir     = os.path.join(project_dir, "CSV")

# ──────────────────────────────────────────────────────────────
# CONNECTION STRING
# r"..." is a raw string — backslash is treated literally,
# not as an escape character like \n or \t.
# ──────────────────────────────────────────────────────────────
SERVER   = r"MUDITSPC\SQLEXPRESS"
DATABASE = "ITConsultingDB"

CONN_STR = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    f"Trusted_Connection=yes;"
)

# ──────────────────────────────────────────────────────────────
# HELPER: run_query
# ──────────────────────────────────────────────────────────────
def run_query(conn, sql):
    cursor = conn.cursor()
    cursor.execute(sql)
    conn.commit()
    cursor.close()

# ──────────────────────────────────────────────────────────────
# HELPER: load_table
# ──────────────────────────────────────────────────────────────
def load_table(conn, csv_path, table_name, columns, dtypes=None):
    print(f"\n  Loading {table_name}...")
    print(f"  Reading: {csv_path}")

    # Pandas reads CSV cleanly — handles quoted fields, line endings
    df = pd.read_csv(csv_path, dtype=dtypes)

    # Keep only the columns we need
    df = df[columns]

    # NaN → None so SQL gets NULL not the string "nan"
    df = df.where(pd.notnull(df), None)

    # Truncate table before loading
    run_query(conn, f"TRUNCATE TABLE {table_name};")

    # Build parameterised INSERT
    placeholders = ", ".join(["?"] * len(columns))
    col_list     = ", ".join(columns)
    insert_sql   = (
        f"INSERT INTO {table_name} ({col_list}) "
        f"VALUES ({placeholders})"
    )

    # Convert numpy types → native Python types
    # pyodbc does not understand numpy.int64 or numpy.float64.
    # .item() converts them to plain int/float.
    # hasattr check ensures strings and None pass through untouched.
    rows = [
        tuple(
            x.item() if hasattr(x, "item") else x
            for x in row
        )
        for row in df.itertuples(index=False, name=None)
    ]

    # Batch insert
    cursor = conn.cursor()
    cursor.fast_executemany = True
    cursor.executemany(insert_sql, rows)
    conn.commit()
    cursor.close()

    print(f"  ✅ {table_name}: {len(rows)} rows inserted.")

# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
def main():
    print("=" * 50)
    print("  IT Consulting — Python SQL Loader")
    print("=" * 50)
    print(f"\nCSV folder : {csv_dir}")
    print(f"Server     : {SERVER}")
    print(f"Database   : {DATABASE}")

    # Verify CSV folder exists before even trying to connect
    if not os.path.exists(csv_dir):
        print(f"\n  ❌ CSV folder not found: {csv_dir}")
        print("  Run Python/generate_data.py first to create the CSVs.")
        return

    # List CSV files found
    csv_files = [f for f in os.listdir(csv_dir) if f.endswith(".csv")]
    print(f"\nCSV files found: {csv_files}")
    if not csv_files:
        print("  ❌ No CSV files found. Run generate_data.py first.")
        return

    # Connect
    print(f"\n Connecting to SQL Server...")
    try:
        conn = pyodbc.connect(CONN_STR)
        print("  ✅ Connected.")
    except Exception as e:
        print(f"\n  ❌ Connection failed: {e}")
        print("\n  TROUBLESHOOTING:")
        print(r"  Try SERVER = r'(localdb)\MSSQLLocalDB'")
        print(r"  Or  SERVER = r'localhost\SQLEXPRESS'")
        print(r"  Or  SERVER = 'localhost'")
        print("  Check SSMS Connect dialog for the exact server name.")
        return

    # Truncate in child-first order
    print("\nClearing existing data...")
    for tbl in ["Risks", "Teams", "Projects", "Clients", "Employees"]:
        run_query(conn, f"TRUNCATE TABLE {tbl};")
        print(f"  Cleared: {tbl}")

    # ── Load all tables ────────────────────────────────────────
    load_table(
        conn,
        csv_path   = os.path.join(csv_dir, "projects.csv"),
        table_name = "Projects",
        columns    = [
            "ProjectID", "ClientID", "Sector", "ProjectType",
            "StartDate", "PlannedEndDate", "ActualEndDate",
            "DelayDays", "Budget", "SpentHours", "ForecastDelay"
        ],
        dtypes = {
            "ProjectID":     "Int64",
            "ClientID":      "Int64",
            "DelayDays":     "Int64",
            "SpentHours":    "Int64",
            "ForecastDelay": "Int64",
            "Budget":        "float64"
        }
    )

    load_table(
        conn,
        csv_path   = os.path.join(csv_dir, "clients.csv"),
        table_name = "Clients",
        columns    = [
            "ClientID", "ClientName", "Industry",
            "Region", "ContractValue", "SLA_Days"
        ],
        dtypes = {
            "ClientID":      "Int64",
            "SLA_Days":      "Int64",
            "ContractValue": "float64"
        }
    )

    load_table(
        conn,
        csv_path   = os.path.join(csv_dir, "employees.csv"),
        table_name = "Employees",
        columns    = [
            "EmployeeID", "Name", "Role",
            "ExperienceYears", "Skill", "Location"
        ],
        dtypes = {
            "EmployeeID":      "Int64",
            "ExperienceYears": "Int64"
        }
    )

    load_table(
        conn,
        csv_path   = os.path.join(csv_dir, "teams.csv"),
        table_name = "Teams",
        columns    = [
            "ProjectID", "EmployeeID", "Role",
            "AssignedHours", "SpentHours"
        ],
        dtypes = {
            "ProjectID":     "Int64",
            "EmployeeID":    "Int64",
            "AssignedHours": "Int64",
            "SpentHours":    "Int64"
        }
    )

    load_table(
        conn,
        csv_path   = os.path.join(csv_dir, "risks.csv"),
        table_name = "Risks",
        columns    = [
            "ProjectID", "RiskType", "RiskImpact", "RiskStatus"
        ],
        dtypes = {
            "ProjectID": "Int64"
        }
    )

    # ── Verify row counts ──────────────────────────────────────
    print("\n" + "=" * 50)
    print("  VERIFICATION — Row Counts")
    print("=" * 50)
    cursor = conn.cursor()
    for tbl in ["Projects", "Clients", "Employees", "Teams", "Risks"]:
        cursor.execute(f"SELECT COUNT(*) FROM {tbl}")
        count  = cursor.fetchone()[0]
        status = "✅" if count > 0 else "❌ EMPTY"
        print(f"  {tbl:<12}: {count:>5} rows  {status}")
    cursor.close()
    conn.close()

    print("\n✅ Done. Next step: run SQL/03_add_keys.sql in SSMS.")

if __name__ == "__main__":
    main()
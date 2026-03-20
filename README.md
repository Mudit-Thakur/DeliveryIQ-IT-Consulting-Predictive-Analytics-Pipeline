# 🏢 DeliveryIQ — IT Consulting Predictive Analytics Pipeline

> **End-to-end automated analytics pipeline** that simulates real-world IT consulting project delivery intelligence — from raw messy data to executive dashboards and automated email reports.

---

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-2.0+-150458?style=for-the-badge&logo=pandas&logoColor=white)
![SQL Server](https://img.shields.io/badge/SQL%20Server-MSSQL-CC2927?style=for-the-badge&logo=microsoftsqlserver&logoColor=white)
![Power BI](https://img.shields.io/badge/Power%20BI-Dashboard-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![NumPy](https://img.shields.io/badge/NumPy-1.24+-013243?style=for-the-badge&logo=numpy&logoColor=white)
![Faker](https://img.shields.io/badge/Faker-Data%20Gen-FF6B6B?style=for-the-badge&logoColor=white)
![SMTP](https://img.shields.io/badge/SMTP-Email%20Automation-0078D4?style=for-the-badge&logo=gmail&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

---

## 📌 What Is This?

**DeliveryIQ** is a fully automated, end-to-end predictive analytics pipeline built to simulate how senior data analysts solve real business problems at IT consulting firms like TCS, Infosys, Accenture, Cognizant, Wipro, and Tech Mahindra.

Every run produces a completely different dataset. Every number on every dashboard, every slide, and every email is pulled from live data — nothing is hardcoded.

### The Business Problem It Solves

> *"IT consulting projects consistently miss deadlines, causing client dissatisfaction and cost overruns. Which projects are at risk? Which sectors are failing? Which employees are overloaded? What should leadership do about it — right now?"*

---

## 🚀 Live Run Snapshot

From a real pipeline execution on `2026-03-20`:

| KPI | Value |
|-----|-------|
| Total Projects Analysed | 106 |
| Average Delay | 30.3 days |
| Projects Overdue | 88 (83.0%) |
| Forecast SLA Alerts | 69 |
| High-Impact Risks | 46 |
| Open Risks | 50 |
| Most Delayed Sector | BFSI (39.6 day avg) |
| Highest Delay Type | Custom Development (32.6d actual / 36.9d forecast) |

---

## 🗂️ Project Architecture

```
DeliveryIQ/
│
├── Python/
│   ├── generate_data.py       ← Dynamic dataset generator (new data every run)
│   ├── clean_explore.py       ← EDA, deduplication, null handling
│   ├── generate_pptx.py       ← Auto-builds 7-slide executive deck from live data
│   ├── send_email.py          ← HTML email + PPTX attachment via SMTP
│   └── run_pipeline.py        ← Master orchestrator (runs all steps in order)
│
├── CSV/                       ← Auto-generated each run (5 tables + metadata)
│   ├── projects.csv
│   ├── clients.csv
│   ├── employees.csv
│   ├── teams.csv
│   ├── risks.csv
│   └── run_metadata.json      ← Live KPIs consumed by PPTX and email
│
├── SQL/
│   ├── 01_create_tables.sql   ← Schema creation
│   ├── 02_load_tables.sql     ← Data loading
│   ├── 03_add_keys.sql        ← PK / FK relationships
│   ├── 04_analytics_ctes.sql  ← 9 high-impact CTE queries
│   └── 05_create_views.sql    ← 4 analytics views for Power BI
│
├── PowerBI/
│   └── dashboard.pbix         ← 8-page consultant-style dashboard
│
├── Reports/
│   └── Executive_Summary.pptx ← Auto-generated each run
│
└── README.md
```

---

## ⚙️ Full Pipeline Flow

```
generate_data.py  (dynamic seed → fresh data every run)
        ↓
CSV/ folder  (5 messy tables + run_metadata.json)
        ↓
MSSQL Server  (relational schema → CTEs → 4 analytics views)
        ↓
Power BI  (8-page consultant dashboard → DAX measures → heatmaps)
        ↓
generate_pptx.py  (7-slide deck → every number from live metadata)
        ↓
send_email.py  (HTML email → color-coded KPIs → PPTX attached)
```

One command runs everything:

```bash
python Python/run_pipeline.py
```

---

## 📊 Dashboard Pages (Power BI — 8 Pages)

| Page | Purpose | Key Visuals |
|------|---------|-------------|
| 1 — Executive Summary | Leadership snapshot | KPI cards, Forecast Alert table |
| 2 — Sector Analysis | Which sectors are failing | Bar charts, heatmap matrix |
| 3 — Client Analysis | Delays and budget overruns by client | Clustered bar, conditional table |
| 4 — Project Type Analysis | Actual vs forecast by type | Stacked bar, clustered bar, table |
| 5 — Employee Utilization | Overload detection | Heatmap matrix, distribution chart |
| 6 — Risk Analysis | Risk exposure and project register | Donut, stacked column, risk table |
| 7 — Delay Trend | Month-over-month actual vs forecast | Line charts, rolling 3M average |
| 8 — Recommendations | Auto-generated executive actions | Tagged recommendation cards |

---

## 🧮 SQL Analytics (9 High-Impact CTE Queries)

Each query answers a specific business question:

```
Query 1  → Which projects are delayed and forecast to breach SLA?
Query 2  → Which sectors have the highest average delay?
Query 3  → Which clients have the most delays and budget overruns?
Query 4  → Which project type + sector combinations perform worst?
Query 5  → Which employees are overloaded or under-utilized?
Query 6  → Which projects have the most high-impact risks?
Query 7  → How does actual delay trend month-over-month?
Query 8  → Which projects need immediate escalation?
Query 9  → Which employees are at overload risk (>120% utilization)?
```

---

## 📐 Data Model (Star Schema)

```
                DimSector (1)
                     |
        ┌────────────┼──────────────────┐
        ↓            ↓                  ↓
  vw_SectorSummary  vw_MonthlyTrend   Projects (1)
                                        |
                          ┌─────────────┼──────────────┐
                          ↓             ↓              ↓
                        Teams         Risks          Clients
                          ↓
                       Employees (1)
                          ↓
                 vw_EmployeeUtilization

  vw_ProjectForecast  ← filtered via slicer (no direct relationship)
```

---

## 📁 Dataset Details

All data is **100% synthetically generated** using Python Faker. No real client, employee, or project data is used anywhere.

| Table | Rows (per run) | Description |
|-------|---------------|-------------|
| Projects | 80–150 | Core project details with ForecastDelay column |
| Clients | 15–30 | Client companies, industries, SLA commitments |
| Employees | 40–70 | Roles, skills, locations |
| Teams | 500–900 | Project-employee allocations with utilization data |
| Risks | 100–250 | Risk type, impact level, status per project |

Sectors covered: `Auto` · `Energy` · `Transportation` · `Logistics` · `Finance` · `BFSI` · `IT`

---

## 📧 Automated Email Report

After every pipeline run, an HTML email is sent automatically containing:

- Color-coded KPI cards (Red / Orange / Green based on live thresholds)
- Detailed findings table with sector and project type insights
- 4 auto-generated recommendations based on actual run numbers
- Executive_Summary.pptx attached

---

## 📑 Auto-Generated Slide Deck

The `generate_pptx.py` script builds a 7-slide executive deck where **every single number is pulled from `run_metadata.json`** — the live KPI file written by `generate_data.py` on each run.

| Slide | Content |
|-------|---------|
| 1 | Title — run timestamp, seed, project count |
| 2 | Executive Summary — KPI cards + top 5 delayed projects |
| 3 | Sector Performance — horizontal bar chart by avg delay |
| 4 | Project Type Analysis — actual vs forecast dual bars |
| 5 | Risk Analysis — impact and status distribution |
| 6 | Key Recommendations — 5 auto-generated actions |
| 7 | Footer — run metadata for reproducibility |

Sample output from `2026-03-20` run:

```
Avg Delay:         30.3 days
% Overdue:         83.0%
Forecast Alerts:   69 projects
High-Risk Items:   46
Worst Sector:      BFSI (39.6 day avg)
Worst Type:        Custom Development (32.6d actual / 36.9d forecast)
```

---

## 🔄 Dynamic Data — Different Every Run

Every pipeline execution produces a completely different dataset:

```python
# Time-based seed — changes every second
SEED = int(time.time())

# Dataset size also randomizes
NUM_PROJECTS  = random.randint(80,  150)
NUM_CLIENTS   = random.randint(15,  30)
NUM_EMPLOYEES = random.randint(40,  70)
```

To reproduce a specific run, copy the seed from the terminal output and set:

```python
USE_FIXED_SEED = True
FIXED_SEED     = 1774005593   # paste your seed here
```

---

## 🛠️ Tech Stack

| Layer | Tool | Purpose |
|-------|------|---------|
| Data Generation | Python + Faker + NumPy | Synthetic messy datasets |
| Data Manipulation | Pandas | Cleaning, EDA, CSV export |
| Database | Microsoft SQL Server | Relational schema, CTE analytics |
| Analytics | SQL CTEs + Views | 9 high-impact business queries |
| Visualisation | Power BI Desktop | 8-page consultant dashboard |
| Calculations | DAX | KPI measures, heatmaps, color rules |
| Reporting | Python-pptx | Auto-generated executive slides |
| Email | smtplib + MIME | HTML report with attachment |
| Orchestration | Python subprocess | End-to-end pipeline automation |

---

## ⚡ Quick Start

### 1. Clone the repo

```bash
git clone https://github.com/yourusername/DeliveryIQ.git
cd DeliveryIQ
```

### 2. Install dependencies

```bash
pip install pandas numpy faker python-pptx sqlalchemy pyodbc
```

### 3. Set up SQL Server

Run these files in SSMS in order:

```
SQL/01_create_tables.sql
SQL/02_load_tables.sql
SQL/03_add_keys.sql
SQL/04_analytics_ctes.sql
SQL/05_create_views.sql
```

### 4. Configure your SQL Server instance

Open `Python/run_pipeline.py` and update line 60:

```python
SQL_SERVER = r"(localdb)\MSSQLLocalDB"   # ← your instance name
```

To find your instance name, run in Command Prompt:

```bash
sqllocaldb info
```

### 5. Run the pipeline

```bash
python Python/run_pipeline.py
```

### 6. Connect Power BI

```
Power BI Desktop
  → Home → Get Data → SQL Server
  → Server:    your SQL_SERVER value
  → Database:  ITConsultingDB
  → Mode:      Import
  → Load all 5 tables + 4 views
```

### 7. Enable email (optional)

Open `Python/send_email.py` and fill in:

```python
EMAIL_SENDER      = "your@gmail.com"
EMAIL_PASSWORD    = "your-16-char-app-password"
EMAIL_RECIPIENTS  = ["recipient@example.com"]
```

Then in `run_pipeline.py` set:

```python
SEND_EMAIL = True
```

---

## 🔐 Security Notes

```
Never commit these values to GitHub:

  send_email.py   →  EMAIL_SENDER, EMAIL_PASSWORD
  run_pipeline.py →  SQL_SERVER (minor but clean anyway)
  CSV/            →  excluded via .gitignore (generated data)
  Reports/        →  excluded via .gitignore (generated slides)
```

All data in this project is **synthetically generated**. No real credentials, no real company data, no real employees.

---

## 🎯 Business Impact Demonstrated

| Insight | Business Value |
|---------|---------------|
| ForecastDelay alerts | Proactive escalation before SLA breach |
| Sector heatmaps | Target resource reallocation where it matters |
| Employee utilization | Prevent burnout and project delays |
| Risk exposure score | Weighted priority for PMO attention |
| Rolling 3M trend | Distinguish noise from real performance changes |

---

## 📈 Interview Talking Points

> *"I built a predictive analytics pipeline that generates realistic messy consulting data in Python, loads it into MSSQL with a star schema, runs nine CTE-based analytics queries, visualizes everything in an 8-page consultant-style Power BI dashboard, and automatically emails an executive slide deck after every run — with every number dynamically pulled from live data."*

---

## 🗺️ Roadmap / Future Improvements

- [ ] Replace `ForecastDelay` simulation with a trained `scikit-learn` regression model
- [ ] Add Azure Data Factory for cloud-based pipeline orchestration
- [ ] Implement cost overrun prediction as a separate ML model
- [ ] Add Power BI Service scheduled refresh
- [ ] Expand to include a second domain (Finance or Healthcare)

---

## 👤 Author

Built by **Mudit** as a portfolio project demonstrating end-to-end data analytics engineering skills.

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://linkedin.com/in/yourprofile)
[![GitHub](https://img.shields.io/badge/GitHub-Follow-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/yourusername)

---

## 📄 License

This project is licensed under the MIT License. See `LICENSE` for details.

---

*Generated pipeline output samples are from live runs and change with every execution.*
*All data is 100% synthetic — no real client, employee, or project information is used.*

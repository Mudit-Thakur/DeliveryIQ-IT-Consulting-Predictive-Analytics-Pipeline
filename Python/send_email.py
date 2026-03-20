
# -*- coding: utf-8 -*-
# ============================================================
# FILE: Python/send_email.py
# PURPOSE: Send the executive slide deck by email after every
#          pipeline run, with a professional HTML email body
#          built automatically from run_metadata.json.
#
# SUPPORTS:
#   - Gmail (App Password)
#   - Outlook / Office 365
#   - Any SMTP server
#
# SETUP (read this once):
#   Gmail:
#     1. Turn on 2-Step Verification on your Google account.
#     2. Go to: myaccount.google.com -> Security -> App Passwords
#     3. Create a new App Password (select "Mail" + "Windows Computer")
#     4. Paste the 16-character password into EMAIL_PASSWORD below.
#     5. Set SMTP_SERVER = "smtp.gmail.com", SMTP_PORT = 587
#
#   Outlook / Office 365:
#     1. Set SMTP_SERVER = "smtp.office365.com", SMTP_PORT = 587
#     2. Use your normal Outlook password (or app password if MFA is on).
#
#   Other providers: check their SMTP settings page.
#
# SECURITY NOTE:
#   Never commit this file to GitHub with real credentials.
#   Use environment variables in production (shown at the bottom).
# ============================================================

import sys
import io
import os

# ── Force UTF-8 on Windows ───────────────────────────────────
sys.stdout = io.TextIOWrapper(
    sys.stdout.buffer,
    encoding="utf-8",
    errors="replace",
    line_buffering=True,
)

import smtplib
import json
import datetime
from email.mime.multipart  import MIMEMultipart
from email.mime.text       import MIMEText
from email.mime.base       import MIMEBase
from email                 import encoders

# ══════════════════════════════════════════════════════════════
# EMAIL CONFIGURATION  <-- edit these values
# ══════════════════════════════════════════════════════════════

# Sender credentials
EMAIL_SENDER   = "your_email@gmail.com"        # your Gmail / Outlook address
EMAIL_PASSWORD = "your_app_password_here"      # 16-char App Password (Gmail)
                                               # or normal password (Outlook)

# Recipients  (add as many as you want)
EMAIL_RECIPIENTS = [
    "recipient1@example.com",
    # "recipient2@example.com",   # uncomment to add more
]

# SMTP settings
SMTP_SERVER = "smtp.gmail.com"   # Gmail
# SMTP_SERVER = "smtp.office365.com"   # Outlook / Office 365
SMTP_PORT   = 587                # TLS port (works for both Gmail and Outlook)

# ── File paths ────────────────────────────────────────────────
# These paths are relative to where run_pipeline.py lives.
# Adjust if your folder structure is different.
_here        = os.path.dirname(os.path.abspath(__file__))
_project_root = os.path.join(_here, "..")

METADATA_PATH = os.path.join(_project_root, "CSV",     "run_metadata.json")
PPTX_PATH     = os.path.join(_project_root, "Reports", "Executive_Summary.pptx")

# ══════════════════════════════════════════════════════════════
# HELPER: load run metadata
# ══════════════════════════════════════════════════════════════

def load_metadata():
    """
    Read the JSON metadata file that generate_data.py creates.
    If the file does not exist, return safe default values so
    the email still sends even without real numbers.
    """
    defaults = {
        "run_timestamp":    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "num_projects":     "N/A",
        "num_clients":      "N/A",
        "num_employees":    "N/A",
        "avg_delay_days":   "N/A",
        "pct_overdue":      "N/A",
        "overdue_count":    "N/A",
        "high_risk_count":  "N/A",
        "open_risk_count":  "N/A",
        "forecast_alerts":  "N/A",
        "top_delayed_sector": "N/A",
        "top_delayed_type": "N/A",
        "seed":             "N/A",
    }

    if not os.path.exists(METADATA_PATH):
        print("  [WARN] run_metadata.json not found. Using default values.")
        return defaults

    try:
        with open(METADATA_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        # Fill in any missing keys with defaults
        return {**defaults, **data}
    except Exception as e:
        print("  [WARN] Could not read metadata: {}".format(e))
        return defaults

# ══════════════════════════════════════════════════════════════
# HELPER: build professional HTML email body
# ══════════════════════════════════════════════════════════════

def build_html_body(meta):
    """
    Build a clean, professional HTML email body using the KPIs
    from run_metadata.json.

    Layout:
      - Header bar with title and run timestamp
      - KPI summary table (4 key numbers)
      - Detailed findings section
      - Recommendations section
      - Footer with seed info for reproducibility
    """

    # Color thresholds for KPI cells
    # We make the avg delay cell red if > 20, orange if > 10, green otherwise
    delay = meta.get("avg_delay_days", 0)
    if isinstance(delay, (int, float)):
        if delay > 20:
            delay_color = "#D9534F"
        elif delay > 10:
            delay_color = "#E87722"
        else:
            delay_color = "#5CB85C"
    else:
        delay_color = "#888888"

    pct_overdue = meta.get("pct_overdue", 0)
    if isinstance(pct_overdue, (int, float)):
        overdue_color = "#D9534F" if pct_overdue > 30 else (
                        "#E87722" if pct_overdue > 15 else "#5CB85C")
    else:
        overdue_color = "#888888"

    html = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>IT Consulting Analytics Report</title>
</head>
<body style="margin:0; padding:0; background-color:#f4f4f4;
             font-family: 'Segoe UI', Arial, sans-serif; color:#333333;">

  <!-- Outer wrapper -->
  <table width="100%" cellpadding="0" cellspacing="0"
         style="background-color:#f4f4f4; padding: 30px 0;">
    <tr>
      <td align="center">

        <!-- Email card -->
        <table width="680" cellpadding="0" cellspacing="0"
               style="background:#ffffff; border-radius:8px;
                      box-shadow: 0 2px 8px rgba(0,0,0,0.10);">

          <!-- ── HEADER ── -->
          <tr>
            <td style="background-color:#0072C6; border-radius:8px 8px 0 0;
                       padding: 28px 36px;">
              <h1 style="margin:0; color:#ffffff; font-size:22px;
                         font-weight:700; letter-spacing:0.5px;">
                IT Consulting Predictive Analytics
              </h1>
              <p style="margin:6px 0 0; color:#cce4f7; font-size:13px;">
                Automated Pipeline Report &nbsp;|&nbsp;
                Run: {run_timestamp}
              </p>
            </td>
          </tr>

          <!-- ── INTRO ── -->
          <tr>
            <td style="padding: 28px 36px 0;">
              <p style="margin:0; font-size:14px; line-height:1.7; color:#555;">
                The IT Consulting Analytics pipeline has completed successfully.
                This report summarises project delivery performance, resource
                utilisation, and risk exposure across
                <strong>{num_projects} projects</strong>,
                <strong>{num_clients} clients</strong>, and
                <strong>{num_employees} employees</strong>.
                The full executive slide deck is attached as a PowerPoint file.
              </p>
            </td>
          </tr>

          <!-- ── KPI CARDS ── -->
          <tr>
            <td style="padding: 24px 36px 0;">
              <h2 style="margin:0 0 14px; font-size:15px; color:#0072C6;
                         text-transform:uppercase; letter-spacing:1px;
                         border-bottom: 2px solid #0072C6; padding-bottom:6px;">
                Key Performance Indicators
              </h2>
              <table width="100%" cellpadding="0" cellspacing="10">
                <tr>
                  <!-- KPI 1: Avg Delay -->
                  <td width="25%" style="background:{delay_color};
                              border-radius:6px; padding:16px 12px;
                              text-align:center;">
                    <div style="color:#fff; font-size:28px; font-weight:700;
                                line-height:1;">{avg_delay_days}</div>
                    <div style="color:#ffe; font-size:11px; margin-top:5px;">
                      Avg Delay (Days)
                    </div>
                  </td>
                  <!-- KPI 2: % Overdue -->
                  <td width="25%" style="background:{overdue_color};
                              border-radius:6px; padding:16px 12px;
                              text-align:center;">
                    <div style="color:#fff; font-size:28px; font-weight:700;
                                line-height:1;">{pct_overdue}%</div>
                    <div style="color:#ffe; font-size:11px; margin-top:5px;">
                      Projects Overdue
                    </div>
                  </td>
                  <!-- KPI 3: Forecast Alerts -->
                  <td width="25%" style="background:#E87722;
                              border-radius:6px; padding:16px 12px;
                              text-align:center;">
                    <div style="color:#fff; font-size:28px; font-weight:700;
                                line-height:1;">{forecast_alerts}</div>
                    <div style="color:#ffe; font-size:11px; margin-top:5px;">
                      Forecast Alerts
                    </div>
                  </td>
                  <!-- KPI 4: High Risk -->
                  <td width="25%" style="background:#D9534F;
                              border-radius:6px; padding:16px 12px;
                              text-align:center;">
                    <div style="color:#fff; font-size:28px; font-weight:700;
                                line-height:1;">{high_risk_count}</div>
                    <div style="color:#ffe; font-size:11px; margin-top:5px;">
                      High-Risk Items
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ── DETAILED FINDINGS ── -->
          <tr>
            <td style="padding: 24px 36px 0;">
              <h2 style="margin:0 0 14px; font-size:15px; color:#0072C6;
                         text-transform:uppercase; letter-spacing:1px;
                         border-bottom: 2px solid #0072C6; padding-bottom:6px;">
                Detailed Findings
              </h2>
              <table width="100%" cellpadding="8" cellspacing="0"
                     style="border-collapse:collapse; font-size:13px;">

                <tr style="background:#f0f4f8;">
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             color:#555; width:55%;">
                    Total Projects Analysed
                  </td>
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             font-weight:700; color:#333;">
                    {num_projects}
                  </td>
                </tr>

                <tr>
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             color:#555;">
                    Projects Overdue (&gt;10 days late)
                  </td>
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             font-weight:700; color:#D9534F;">
                    {overdue_count}
                  </td>
                </tr>

                <tr style="background:#f0f4f8;">
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             color:#555;">
                    Open Risks Requiring Attention
                  </td>
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             font-weight:700; color:#E87722;">
                    {open_risk_count}
                  </td>
                </tr>

                <tr>
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             color:#555;">
                    Sector with Highest Average Delay
                  </td>
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             font-weight:700; color:#333;">
                    {top_delayed_sector}
                  </td>
                </tr>

                <tr style="background:#f0f4f8;">
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             color:#555;">
                    Project Type with Highest Average Delay
                  </td>
                  <td style="padding:10px 14px; border-bottom:1px solid #e0e0e0;
                             font-weight:700; color:#333;">
                    {top_delayed_type}
                  </td>
                </tr>

                <tr>
                  <td style="padding:10px 14px; color:#555;">
                    Total Employees Monitored
                  </td>
                  <td style="padding:10px 14px; font-weight:700; color:#333;">
                    {num_employees}
                  </td>
                </tr>

              </table>
            </td>
          </tr>

          <!-- ── RECOMMENDATIONS ── -->
          <tr>
            <td style="padding: 24px 36px 0;">
              <h2 style="margin:0 0 14px; font-size:15px; color:#0072C6;
                         text-transform:uppercase; letter-spacing:1px;
                         border-bottom: 2px solid #0072C6; padding-bottom:6px;">
                Automated Recommendations
              </h2>

              <!-- Recommendation 1 -->
              <table width="100%" cellpadding="0" cellspacing="0"
                     style="margin-bottom:10px;">
                <tr>
                  <td width="6" style="background:#D9534F;
                               border-radius:3px 0 0 3px;">&nbsp;</td>
                  <td style="background:#fff5f5; padding:12px 16px;
                             border:1px solid #f5c6cb; border-left:none;
                             border-radius:0 3px 3px 0; font-size:13px;
                             color:#333; line-height:1.6;">
                    <strong>CRITICAL:</strong> {forecast_alerts} projects
                    are forecast to breach their SLA. Escalate to PMO
                    immediately and assign dedicated project managers.
                  </td>
                </tr>
              </table>

              <!-- Recommendation 2 -->
              <table width="100%" cellpadding="0" cellspacing="0"
                     style="margin-bottom:10px;">
                <tr>
                  <td width="6" style="background:#E87722;
                               border-radius:3px 0 0 3px;">&nbsp;</td>
                  <td style="background:#fff8f0; padding:12px 16px;
                             border:1px solid #ffd59e; border-left:none;
                             border-radius:0 3px 3px 0; font-size:13px;
                             color:#333; line-height:1.6;">
                    <strong>HIGH PRIORITY:</strong> {open_risk_count} risks
                    are currently open. The {top_delayed_sector} sector
                    requires immediate resource review and milestone
                    check-ins.
                  </td>
                </tr>
              </table>

              <!-- Recommendation 3 -->
              <table width="100%" cellpadding="0" cellspacing="0"
                     style="margin-bottom:10px;">
                <tr>
                  <td width="6" style="background:#0072C6;
                               border-radius:3px 0 0 3px;">&nbsp;</td>
                  <td style="background:#f0f7ff; padding:12px 16px;
                             border:1px solid #b3d4f0; border-left:none;
                             border-radius:0 3px 3px 0; font-size:13px;
                             color:#333; line-height:1.6;">
                    <strong>ACTION:</strong> {top_delayed_type} projects
                    show the highest average delays. Standardise delivery
                    milestones and conduct a root-cause review for this
                    project type.
                  </td>
                </tr>
              </table>

              <!-- Recommendation 4 -->
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="6" style="background:#5CB85C;
                               border-radius:3px 0 0 3px;">&nbsp;</td>
                  <td style="background:#f5fff5; padding:12px 16px;
                             border:1px solid #b2dfb2; border-left:none;
                             border-radius:0 3px 3px 0; font-size:13px;
                             color:#333; line-height:1.6;">
                    <strong>MONITORING:</strong> Review the attached Power BI
                    dashboard weekly. Schedule a stakeholder review for all
                    projects in the red zone.
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ── ATTACHMENT NOTE ── -->
          <tr>
            <td style="padding: 24px 36px 0;">
              <table width="100%" cellpadding="12" cellspacing="0"
                     style="background:#f8f9fa; border:1px solid #dee2e6;
                            border-radius:6px; font-size:13px; color:#555;">
                <tr>
                  <td>
                    <strong style="color:#333;">Attached:</strong>
                    Executive_Summary.pptx &nbsp;|&nbsp;
                    Auto-generated executive slide deck with KPI cards,
                    sector analysis, risk summary, and recommendations.
                    Open in Microsoft PowerPoint or Google Slides.
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ── FOOTER ── -->
          <tr>
            <td style="padding: 24px 36px 28px;">
              <hr style="border:none; border-top:1px solid #e0e0e0; margin:0 0 16px;">
              <p style="margin:0; font-size:11px; color:#aaa; line-height:1.6;">
                This report was generated automatically by the IT Consulting
                Analytics Pipeline.<br>
                Run timestamp: {run_timestamp} &nbsp;|&nbsp;
                Seed: {seed} &nbsp;|&nbsp;
                To replay this exact dataset set
                <code style="background:#f4f4f4; padding:1px 4px;
                             border-radius:3px;">USE_FIXED_SEED = True</code>
                and
                <code style="background:#f4f4f4; padding:1px 4px;
                             border-radius:3px;">FIXED_SEED = {seed}</code>
                in generate_data.py.
              </p>
            </td>
          </tr>

        </table>
        <!-- end card -->

      </td>
    </tr>
  </table>

</body>
</html>
""".format(
        run_timestamp      = meta["run_timestamp"],
        num_projects       = meta["num_projects"],
        num_clients        = meta["num_clients"],
        num_employees      = meta["num_employees"],
        avg_delay_days     = meta["avg_delay_days"],
        pct_overdue        = meta["pct_overdue"],
        overdue_count      = meta["overdue_count"],
        high_risk_count    = meta["high_risk_count"],
        open_risk_count    = meta["open_risk_count"],
        forecast_alerts    = meta["forecast_alerts"],
        top_delayed_sector = meta["top_delayed_sector"],
        top_delayed_type   = meta["top_delayed_type"],
        seed               = meta["seed"],
        delay_color        = delay_color,
        overdue_color      = overdue_color,
    )

    return html

# ══════════════════════════════════════════════════════════════
# MAIN: build and send the email
# ══════════════════════════════════════════════════════════════

def send_report_email():
    """
    Build the email, attach the PPTX, and send via SMTP TLS.
    Returns True on success, False on failure.
    """
    print("Building email...")

    meta = load_metadata()

    # ── Build the message object ──────────────────────────────
    # MIMEMultipart("alternative") = email with both plain text
    # and HTML versions (email clients show the best one they support)
    msg = MIMEMultipart("mixed")

    msg["Subject"] = (
        "IT Consulting Analytics Report | "
        "{} Projects | {:.0f}% Overdue | {} Alerts | {}".format(
            meta["num_projects"],
            meta["pct_overdue"] if isinstance(meta["pct_overdue"], (int, float)) else 0,
            meta["forecast_alerts"],
            meta["run_timestamp"],
        )
    )
    msg["From"]    = EMAIL_SENDER
    msg["To"]      = ", ".join(EMAIL_RECIPIENTS)

    # ── Plain text fallback (shown by clients that block HTML) ─
    plain_text = (
        "IT Consulting Analytics Pipeline - Automated Report\n"
        "====================================================\n\n"
        "Run: {run_timestamp}\n"
        "Projects:        {num_projects}\n"
        "Clients:         {num_clients}\n"
        "Employees:       {num_employees}\n\n"
        "Avg Delay:       {avg_delay_days} days\n"
        "Overdue:         {pct_overdue}% ({overdue_count} projects)\n"
        "Forecast Alerts: {forecast_alerts}\n"
        "High Risk Items: {high_risk_count}\n"
        "Open Risks:      {open_risk_count}\n\n"
        "Worst Sector:    {top_delayed_sector}\n"
        "Worst Proj Type: {top_delayed_type}\n\n"
        "The executive slide deck is attached.\n"
        "Seed: {seed}\n"
    ).format(**meta)

    # Attach both plain and HTML versions
    alt_part = MIMEMultipart("alternative")
    alt_part.attach(MIMEText(plain_text, "plain", "utf-8"))
    alt_part.attach(MIMEText(build_html_body(meta), "html", "utf-8"))
    msg.attach(alt_part)

    # ── Attach the PPTX file ──────────────────────────────────
    if os.path.exists(PPTX_PATH):
        try:
            with open(PPTX_PATH, "rb") as f:
                pptx_data = f.read()

            attachment = MIMEBase(
                "application",
                "vnd.openxmlformats-officedocument"
                ".presentationml.presentation"
            )
            attachment.set_payload(pptx_data)
            encoders.encode_base64(attachment)
            attachment.add_header(
                "Content-Disposition",
                'attachment; filename="Executive_Summary.pptx"',
            )
            msg.attach(attachment)
            print("  [OK]   PPTX attached ({:.1f} KB)".format(
                os.path.getsize(PPTX_PATH) / 1024
            ))
        except Exception as e:
            print("  [WARN] Could not attach PPTX: {}".format(e))
    else:
        print("  [WARN] PPTX not found at: {}".format(PPTX_PATH))
        print("         Email will send without attachment.")

    # ── Send via SMTP TLS ─────────────────────────────────────
    print("Connecting to {} port {}...".format(SMTP_SERVER, SMTP_PORT))

    try:
        # smtplib.SMTP → plain connection
        # .starttls()  → upgrades to encrypted TLS
        # .login()     → authenticates
        # .sendmail()  → sends
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            server.ehlo()           # introduce ourselves to server
            server.starttls()       # encrypt the connection
            server.ehlo()           # re-introduce after TLS
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(
                EMAIL_SENDER,
                EMAIL_RECIPIENTS,
                msg.as_bytes(),     # as_bytes handles UTF-8 subjects correctly
            )

        print("  [OK]   Email sent to: {}".format(", ".join(EMAIL_RECIPIENTS)))
        return True

    except smtplib.SMTPAuthenticationError:
        print("  [FAIL] Authentication failed.")
        print("         Gmail: use an App Password, not your normal password.")
        print("         Go to: myaccount.google.com -> Security -> App Passwords")
        return False

    except smtplib.SMTPException as e:
        print("  [FAIL] SMTP error: {}".format(e))
        return False

    except Exception as e:
        print("  [FAIL] Unexpected error: {}".format(e))
        return False


# ── Allow running this file directly for testing ─────────────
if __name__ == "__main__":
    print("=" * 55)
    print("  SEND REPORT EMAIL (standalone test)")
    print("=" * 55)
    print("")

    # Quick config check before attempting to send
    if "your_email" in EMAIL_SENDER or "your_app_password" in EMAIL_PASSWORD:
        print("  [FAIL] You have not configured your email credentials.")
        print("")
        print("  Open send_email.py and update:")
        print("    EMAIL_SENDER   = 'your_actual@gmail.com'")
        print("    EMAIL_PASSWORD = 'your 16-char app password'")
        print("    EMAIL_RECIPIENTS = ['recipient@example.com']")
        print("")
        print("  Gmail App Password setup:")
        print("    1. myaccount.google.com -> Security")
        print("    2. Enable 2-Step Verification")
        print("    3. App Passwords -> create one for Mail")
        print("    4. Paste the 16-char password here")
        sys.exit(1)

    success = send_report_email()
    if success:
        print("")
        print("Done. Check your inbox.")
    else:
        print("")
        print("Email failed. See errors above.")
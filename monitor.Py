"""
================================================================
Dravio — NEET UG Counselling Notification Monitor
================================================================
Version     : 2.0.0
Description : Monitors 5 official NEET UG counselling sites for
              new notifications. Sends email alerts + logs to
              Google Sheets automatically.
Sites       : Gujarat (ACPUGMEC), Maharashtra (CET Cell),
              MP (DME), Rajasthan (rajugneet2026), MCC (AIQ)
================================================================
"""

import os
import sys
import json
import time
import hashlib
import smtplib
import logging
import traceback
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import requests
import schedule
import gspread
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------------
# Load .env file
# ----------------------------------------------------------------
load_dotenv()

# ----------------------------------------------------------------
# Logging Setup
# ----------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%d-%m-%Y %I:%M:%S %p",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("dravio_monitor.log", encoding="utf-8")
    ]
)
log = logging.getLogger("DravioMonitor")

# ----------------------------------------------------------------
# Configuration from .env
# ----------------------------------------------------------------
GMAIL_SENDER            = os.getenv("GMAIL_SENDER", "")
GMAIL_APP_PASSWORD      = os.getenv("GMAIL_APP_PASSWORD", "")
ALERT_RECIPIENTS        = [e.strip() for e in os.getenv("ALERT_RECIPIENTS", "").split(",") if e.strip()]
GOOGLE_SHEET_ID         = os.getenv("GOOGLE_SHEET_ID", "")
CREDENTIALS_FILE        = os.getenv("CREDENTIALS_FILE", "credentials.json")
WORKSHEET_NAME          = os.getenv("WORKSHEET_NAME", "Notification Log")
CHECK_INTERVAL_MINUTES  = int(os.getenv("CHECK_INTERVAL_MINUTES", "60"))
MAX_RETRIES             = int(os.getenv("MAX_RETRIES", "3"))
RESTART_WAIT_MINUTES    = int(os.getenv("RESTART_WAIT_MINUTES", "5"))

STATE_FILE              = "last_state.json"

# ----------------------------------------------------------------
# Sites Configuration
# Each site has:
#   - url        : The exact URL to monitor
#   - name       : Display name
#   - state      : State/body label
#   - filter_tag : CSS tag/class to extract (None = full page text)
#   - filter_kw  : Keywords to filter relevant content (None = all)
# ----------------------------------------------------------------
SITES = [
    {
        "name"       : "Gujarat (ACPUGMEC)",
        "state"      : "Gujarat",
        "url"        : "https://medadmgujarat.org/ug/home.aspx",
        # Gujarat has date-wise notifications in the main content area
        "filter_tag" : None,
        "filter_kw"  : None,
    },
    {
        "name"       : "Maharashtra (CET Cell)",
        "state"      : "Maharashtra",
        "url"        : "https://cetcell.mahacet.org/",
        # Maharashtra: We only care about Medical Education tab content
        # The notifications table has Course Name + Subject + Published Date
        # We filter rows containing medical keywords
        "filter_tag" : None,
        "filter_kw"  : ["medical", "health science", "mbbs", "neet", "nursing",
                        "bds", "mbbs", "ayush", "dmer", "health"],
    },
    {
        "name"       : "Madhya Pradesh (DME)",
        "state"      : "Madhya Pradesh",
        "url"        : "https://dme.mponline.gov.in/Portal/Services/DMEMP/DMEUG/Profile/Instructions.aspx",
        # MP: "Latest Instructions" tab — currently empty but new instructions will appear here
        "filter_tag" : None,
        "filter_kw"  : None,
    },
    {
        "name"       : "Rajasthan (rajugneet2026)",
        "state"      : "Rajasthan",
        "url"        : "https://www.rajugneet2026.com/",
        # Rajasthan: Recent Posts section has new notifications
        "filter_tag" : None,
        "filter_kw"  : None,
    },
    {
        "name"       : "MCC (AIQ / Central Counselling)",
        "state"      : "All India",
        "url"        : "https://mcc.nic.in/current-events-ug/",
        # MCC: Current Events UG table — Title + Year + View/Download
        "filter_tag" : None,
        "filter_kw"  : None,
    },
]

# ----------------------------------------------------------------
# Google Sheets Setup
# ----------------------------------------------------------------
SHEET_HEADERS = [
    "Sr No",
    "Date",
    "Time (IST)",
    "State / Body",
    "Site Name",
    "URL",
    "Change Detected",
    "Status",
    "Notified Via",
    "Remarks"
]

def get_worksheet():
    """Connect to Google Sheets and return the target worksheet."""
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds     = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        client    = gspread.authorize(creds)
        sheet     = client.open_by_key(GOOGLE_SHEET_ID)

        try:
            ws = sheet.worksheet(WORKSHEET_NAME)
        except gspread.WorksheetNotFound:
            ws = sheet.add_worksheet(title=WORKSHEET_NAME, rows=2000, cols=12)
            ws.append_row(SHEET_HEADERS)
            # Format header: dark blue bg, white bold text
            ws.format("A1:J1", {
                "backgroundColor" : {"red": 0.1, "green": 0.27, "blue": 0.43},
                "textFormat"      : {
                    "bold"            : True,
                    "foregroundColor" : {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "fontSize"        : 11
                },
                "horizontalAlignment": "CENTER"
            })
            log.info("✅ Created new Google Sheet worksheet with headers.")
        return ws

    except Exception as e:
        log.error(f"❌ Google Sheets connection failed: {e}")
        return None


def log_to_sheet(ws, site, change_summary, status="New Notification"):
    """Append a new row to Google Sheet for a detected change."""
    if not ws:
        log.warning("⚠️  Skipping sheet log — worksheet unavailable.")
        return

    try:
        existing   = ws.get_all_values()
        sr_no      = len(existing)  # 1 = header, so first data row = 1

        now        = datetime.now()
        row        = [
            sr_no,
            now.strftime("%d-%m-%Y"),
            now.strftime("%I:%M %p"),
            site["state"],
            site["name"],
            site["url"],
            change_summary[:500],
            status,
            "Email",
            "Auto-detected by Dravio Monitor"
        ]
        ws.append_row(row)

        # Highlight new row in light yellow
        row_num = len(existing) + 1
        ws.format(f"A{row_num}:J{row_num}", {
            "backgroundColor": {"red": 1.0, "green": 0.98, "blue": 0.80}
        })
        log.info(f"✅ Logged to Google Sheet: Row {row_num}")

    except Exception as e:
        log.error(f"❌ Failed to log to Google Sheet: {e}")


# ----------------------------------------------------------------
# Email Alert
# ----------------------------------------------------------------
def send_email(subject, html_body, recipients=None):
    """Send an HTML email alert."""
    if recipients is None:
        recipients = ALERT_RECIPIENTS
    try:
        msg            = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = f"Dravio Monitor <{GMAIL_SENDER}>"
        msg["To"]      = ", ".join(recipients)
        msg.attach(MIMEText(html_body, "html"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_SENDER, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_SENDER, recipients, msg.as_string())

        log.info(f"✅ Email sent to: {', '.join(recipients)}")

    except Exception as e:
        log.error(f"❌ Email sending failed: {e}")


def build_notification_email(site, change_summary):
    """Build formatted HTML email for a new counselling notification."""
    now     = datetime.now().strftime("%d %B %Y, %I:%M %p")
    subject = f"🔔 Dravio Alert: New Notification — {site['name']}"

    html = f"""
    <html>
    <body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f4f4;">

      <!-- Header -->
      <div style="background:#1a3c6e;padding:18px 24px;">
        <h1 style="color:#fff;margin:0;font-size:20px;">🔔 Dravio — Counselling Notification Alert</h1>
        <p style="color:#a8c4e0;margin:4px 0 0;font-size:12px;">
          NEET UG MBBS Counselling Monitor · Automated Alert
        </p>
      </div>

      <!-- Body -->
      <div style="max-width:620px;margin:20px auto;background:#fff;
                  border-radius:8px;overflow:hidden;
                  box-shadow:0 2px 8px rgba(0,0,0,0.1);">

        <!-- Alert Banner -->
        <div style="background:#fff3cd;border-left:5px solid #ffc107;padding:14px 20px;">
          <p style="margin:0;font-size:14px;font-weight:bold;color:#856404;">
            ⚠️ A new notification has been detected on an official counselling site!
          </p>
        </div>

        <!-- Details Table -->
        <div style="padding:24px;">
          <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         width:32%;font-weight:bold;color:#495057;">🏛️ State / Body</td>
              <td style="padding:10px;border:1px solid #dee2e6;
                         font-weight:bold;font-size:15px;color:#1a3c6e;">
                {site['state']}
              </td>
            </tr>
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">📋 Authority</td>
              <td style="padding:10px;border:1px solid #dee2e6;color:#212529;">
                {site['name']}
              </td>
            </tr>
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">🔗 Official URL</td>
              <td style="padding:10px;border:1px solid #dee2e6;">
                <a href="{site['url']}" style="color:#1a3c6e;">{site['url']}</a>
              </td>
            </tr>
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">📝 Change Detected</td>
              <td style="padding:10px;border:1px solid #dee2e6;color:#212529;">
                {change_summary[:400]}
              </td>
            </tr>
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">🕐 Detected At</td>
              <td style="padding:10px;border:1px solid #dee2e6;color:#212529;">
                {now} (IST)
              </td>
            </tr>
          </table>

          <!-- Action Button -->
          <div style="text-align:center;margin:24px 0 10px;">
            <a href="{site['url']}"
               style="background:#1a3c6e;color:#fff;padding:12px 32px;
                      border-radius:5px;text-decoration:none;
                      font-size:14px;font-weight:bold;display:inline-block;">
              🌐 View Official Site
            </a>
          </div>

          <p style="color:#888;font-size:11px;text-align:center;margin-top:12px;">
            Review the notification and update Dravio platform if required.
          </p>
        </div>

        <!-- Footer -->
        <div style="background:#f8f9fa;padding:14px 20px;
                    border-top:1px solid #dee2e6;text-align:center;">
          <p style="margin:0;color:#6c757d;font-size:11px;">
            This is an automated alert from <strong>Dravio Monitor</strong><br>
            Monitoring NEET UG MBBS Counselling Sites across India<br>
            Gujarat · Maharashtra · Madhya Pradesh · Rajasthan · MCC (AIQ)
          </p>
        </div>
      </div>

    </body>
    </html>
    """
    return subject, html


def build_failure_email(error_details):
    """Build HTML email for monitor failure/crash alert."""
    now     = datetime.now().strftime("%d %B %Y, %I:%M %p")
    subject = "🚨 URGENT: Dravio Monitor Has Stopped Working"

    html = f"""
    <html>
    <body style="font-family:Arial,sans-serif;background:#f4f4f4;margin:0;padding:0;">

      <div style="background:#c0392b;padding:18px 24px;">
        <h1 style="color:#fff;margin:0;font-size:20px;">🚨 Dravio Monitor — Critical Failure Alert</h1>
      </div>

      <div style="max-width:620px;margin:20px auto;background:#fff;
                  border-radius:8px;overflow:hidden;
                  box-shadow:0 2px 8px rgba(0,0,0,0.1);">

        <div style="background:#fdecea;border-left:5px solid #c0392b;padding:14px 20px;">
          <p style="margin:0;font-size:14px;font-weight:bold;color:#c0392b;">
            ⚠️ The Dravio monitoring script has encountered a critical error and stopped running!
          </p>
        </div>

        <div style="padding:24px;">
          <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         width:32%;font-weight:bold;color:#495057;">🕐 Failed At</td>
              <td style="padding:10px;border:1px solid #dee2e6;">{now} (IST)</td>
            </tr>
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">❌ Error Details</td>
              <td style="padding:10px;border:1px solid #dee2e6;
                         font-family:monospace;font-size:12px;color:#c0392b;">
                {error_details[:600]}
              </td>
            </tr>
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">⏰ Auto Restart</td>
              <td style="padding:10px;border:1px solid #dee2e6;">
                Script will attempt to restart in {RESTART_WAIT_MINUTES} minutes
              </td>
            </tr>
          </table>

          <div style="background:#fdecea;border-radius:5px;padding:12px 16px;margin-top:16px;">
            <p style="margin:0;font-size:13px;color:#c0392b;">
              <strong>Action Required:</strong> If the monitor does not recover automatically,
              please manually restart it on your server to resume counselling site monitoring.
            </p>
          </div>
        </div>

        <div style="background:#f8f9fa;padding:12px 20px;
                    border-top:1px solid #dee2e6;text-align:center;">
          <p style="margin:0;color:#888;font-size:11px;">
            Automated failure notification from <strong>Dravio Monitor</strong>
          </p>
        </div>
      </div>

    </body>
    </html>
    """
    return subject, html


# ----------------------------------------------------------------
# Page Fetching & Parsing
# ----------------------------------------------------------------
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


def fetch_page_text(site):
    """
    Fetch a site's page and return relevant extracted text.
    Applies site-specific filtering where configured.
    SSL verification disabled for Indian govt sites with expired certs.
    """
    url       = site["url"]
    filter_kw = site.get("filter_kw")

    # Many Indian govt sites (medadmgujarat, dme.mponline etc.) have
    # expired or self-signed SSL certificates — disable verification safely
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    resp = requests.get(url, headers=HEADERS, timeout=30, verify=False)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    # Remove noise tags
    for tag in soup(["script", "style", "nav", "footer", "header",
                     "meta", "link", "noscript"]):
        tag.decompose()

    raw_text = soup.get_text(separator="\n", strip=True)

    # Apply keyword filter for Maharashtra (Medical Education tab only)
    if filter_kw:
        lines         = raw_text.splitlines()
        filtered      = [
            line for line in lines
            if any(kw.lower() in line.lower() for kw in filter_kw)
        ]
        return "\n".join(filtered) if filtered else raw_text

    return raw_text


def get_hash(text):
    """Return MD5 hash of text content."""
    return hashlib.md5(text.encode("utf-8")).hexdigest()


def get_change_summary(old_text, new_text):
    """Generate a plain-English change summary."""
    old_lines = set(old_text.splitlines())
    new_lines = set(new_text.splitlines())
    added     = [
        l.strip() for l in (new_lines - old_lines)
        if len(l.strip()) > 15
    ]
    if added:
        preview = " | ".join(added[:4])
        return f"New content added: {preview[:400]}"
    return "Page content has changed — a new notification may have been posted."


# ----------------------------------------------------------------
# State Management
# ----------------------------------------------------------------
def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


# ----------------------------------------------------------------
# Main Check Function
# ----------------------------------------------------------------
def check_all_sites():
    """Check all 5 counselling sites for new notifications."""
    log.info("=" * 60)
    log.info(f"🔍 Starting check — {datetime.now().strftime('%d-%m-%Y %I:%M %p IST')}")
    log.info("=" * 60)

    state      = load_state()
    ws         = get_worksheet()
    any_change = False

    for site in SITES:
        name = site["name"]
        url  = site["url"]
        log.info(f"\n📡 Checking: {name}")
        log.info(f"   URL: {url}")

        retry_count = 0
        success     = False

        while retry_count < MAX_RETRIES and not success:
            try:
                content      = fetch_page_text(site)
                current_hash = get_hash(content)

                if name not in state:
                    # First run — save baseline
                    state[name] = {
                        "hash"         : current_hash,
                        "last_checked" : datetime.now().isoformat(),
                        "snapshot"     : content[:3000]
                    }
                    log.info(f"   ✅ Baseline snapshot saved (first run)")

                else:
                    prev_hash = state[name]["hash"]

                    if current_hash != prev_hash:
                        log.info(f"   🔔 CHANGE DETECTED on {name}!")
                        any_change     = True

                        old_content    = state[name].get("snapshot", "")
                        change_summary = get_change_summary(old_content, content)
                        log.info(f"   📋 Summary: {change_summary[:200]}")

                        # Send email alert
                        subject, html = build_notification_email(site, change_summary)
                        send_email(subject, html)

                        # Log to Google Sheets
                        log_to_sheet(ws, site, change_summary)

                        # Update state
                        state[name] = {
                            "hash"         : current_hash,
                            "last_checked" : datetime.now().isoformat(),
                            "snapshot"     : content[:3000]
                        }
                    else:
                        log.info(f"   ✅ No change detected")
                        state[name]["last_checked"] = datetime.now().isoformat()

                success = True

            except requests.exceptions.RequestException as e:
                retry_count += 1
                log.warning(f"   ⚠️  Fetch failed (attempt {retry_count}/{MAX_RETRIES}): {e}")
                if retry_count < MAX_RETRIES:
                    time.sleep(30)  # Wait 30s before retry

            except Exception as e:
                log.error(f"   ❌ Unexpected error for {name}: {e}")
                log_to_sheet(ws, site, f"Fetch error: {str(e)}", status="Error")
                break

        if not success:
            log.error(f"   ❌ All {MAX_RETRIES} retries failed for {name}")
            log_to_sheet(ws, site,
                         f"Failed after {MAX_RETRIES} retries",
                         status="Fetch Failed")

    save_state(state)

    if not any_change:
        log.info(f"\n✅ All sites checked — No new notifications found.")

    log.info(f"\n⏰ Next check in {CHECK_INTERVAL_MINUTES} minutes...")


# ----------------------------------------------------------------
# Validation Check on Startup
# ----------------------------------------------------------------
def validate_config():
    """Validate all required environment variables before starting."""
    errors = []
    if not GMAIL_SENDER:
        errors.append("GMAIL_SENDER is not set in .env")
    if not GMAIL_APP_PASSWORD or GMAIL_APP_PASSWORD == "xxxx xxxx xxxx xxxx":
        errors.append("GMAIL_APP_PASSWORD is not set in .env")
    if not ALERT_RECIPIENTS:
        errors.append("ALERT_RECIPIENTS is not set in .env")
    if not GOOGLE_SHEET_ID or GOOGLE_SHEET_ID == "your_sheet_id_here":
        errors.append("GOOGLE_SHEET_ID is not set in .env")
    if not os.path.exists(CREDENTIALS_FILE):
        errors.append(f"credentials.json not found at: {CREDENTIALS_FILE}")

    if errors:
        log.error("❌ Configuration errors found:")
        for err in errors:
            log.error(f"   • {err}")
        log.error("Please check your .env file and try again.")
        sys.exit(1)

    log.info("✅ Configuration validated successfully.")


# ----------------------------------------------------------------
# Entry Point with Auto-Restart on Crash
# ----------------------------------------------------------------
if __name__ == "__main__":

    validate_config()

    log.info("=" * 60)
    log.info("   🏥 Dravio Counselling Monitor v2.0 — Starting Up")
    log.info("=" * 60)
    log.info(f"   Monitoring  : {len(SITES)} sites")
    for s in SITES:
        log.info(f"                 • {s['name']}")
    log.info(f"   Interval    : Every {CHECK_INTERVAL_MINUTES} minutes")
    log.info(f"   Alerts to   : {', '.join(ALERT_RECIPIENTS)}")
    log.info(f"   Started At  : {datetime.now().strftime('%d-%m-%Y %I:%M %p IST')}")
    log.info("=" * 60)

    consecutive_failures = 0

    while True:
        try:
            # Run immediately on start
            check_all_sites()
            consecutive_failures = 0

            # Schedule recurring checks
            schedule.every(CHECK_INTERVAL_MINUTES).minutes.do(check_all_sites)

            while True:
                schedule.run_pending()
                time.sleep(30)

        except KeyboardInterrupt:
            log.info("\n⛔ Monitor stopped by user.")
            sys.exit(0)

        except Exception as e:
            consecutive_failures += 1
            error_details = traceback.format_exc()
            log.error(f"💥 Critical crash (failure #{consecutive_failures}):\n{error_details}")

            # Send failure alert email
            try:
                subject, html = build_failure_email(error_details)
                send_email(subject, html)
                log.info("📧 Failure alert email sent.")
            except Exception as mail_err:
                log.error(f"❌ Could not send failure email: {mail_err}")

            # Clear schedule and wait before restarting
            schedule.clear()
            wait_seconds = RESTART_WAIT_MINUTES * 60
            log.info(f"⏳ Waiting {RESTART_WAIT_MINUTES} minutes before restarting...")
            time.sleep(wait_seconds)
            log.info("🔄 Restarting monitor...")
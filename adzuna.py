# ============================================================
#  JOB SCANNER — RSS + ADZUNA (MERGED)  —  June 2026  (v3)
#  Combines:
#    - RSS feeds (politepol/rss.app — LinkedIn, Indeed, Naukri)
#    - Adzuna API (24-hour filter built in)
#  Both write to ONE Google Sheet and send ONE combined
#  Telegram summary at the end of each run.
#
#  CHANGES IN v3:
#    - Storage backend switched fully from Excel to Google Sheets.
#      All openpyxl/Excel code has been removed — there were two
#      half-wired storage paths in the previous version (Excel checks
#      running alongside Google Sheets writes, and a missing
#      `seen_store` variable causing an immediate crash). This version
#      uses Google Sheets only, end to end.
#    - Dedup store restored and now backed by THIS SHEET ONLY: every
#      run reads existing rows from the sheet itself (URLs + normalized
#      title+company signatures) — no separate seen_jobs.json file, no
#      separate Excel file. The sheet IS the single source of truth, so
#      there's nothing to go out of sync.
#    - Clear setup instructions + an explicit upfront check for
#      service_account.json, so missing credentials fail with a plain
#      English message instead of a traceback.
#
#  CARRIED OVER FROM v2:
#    - SENIOR CARVE-OUT: "senior" / "senior analyst" titles pass IF the
#      years mentioned are <= MAX_EXPERIENCE_YEARS. Director / VP /
#      Head of / Principal / Senior Manager / Senior Director remain
#      fully blocked regardless of experience.
#    - Broadened KEYWORDS (Scrum Master, Product Owner, Agile/IT/
#      Associate analyst variants, etc).
#
#  ============================================================
#  SETUP (run once in terminal):
#     pip install feedparser requests gspread oauth2client
#
#  GOOGLE SHEETS SETUP (one-time):
#  1. Go to https://console.cloud.google.com/ → create a project.
#  2. Enable "Google Sheets API" and "Google Drive API" for it.
#  3. Create a Service Account → create a JSON key for it → download it.
#  4. Rename the downloaded file to "service_account.json" and put it
#     in the SAME FOLDER as this script.
#  5. Open the JSON file, find the "client_email" field (looks like
#     xxxx@xxxx.iam.gserviceaccount.com).
#  6. Create a Google Sheet named exactly "Job_Hunt_Tracker", add a
#     tab named exactly "📋 Applications", and Share that sheet with
#     the client_email from step 5 (Editor access).
#  7. Row 1 of the tab should have headers:
#     Title | Company | Location | Source | Date | URL | Status
#
#  THEN:
#  1. Section 1  — paste your RSS feed URLs (politepol/rss.app)
#  2. Section 2  — paste your Adzuna app_id / app_key
#  3. Section 3  — roles, cities, experience cap (shared by both)
#  4. Section 4  — Telegram bot token (optional, shared by both)
#  5. Run:  python job_scanner.py
# ============================================================

import feedparser
import datetime
import requests
import os
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ============================================================
#  ← EDIT THIS  (Section 1) — RSS feed URLs
#  From politepol.com or rss.app. Leave "" to skip a feed.
# ============================================================
RSS_FEEDS = {
    "LinkedIn_BA_Mumbai":      "https://politepaul.com/fd/SU8t29W2Howk.xml",
    "LinkedIn_BA_NaviMumbai":  "https://politepaul.com/fd/BF56TWCeJkDW.xml",
    "LinkedIn_BA_Pune":        "https://politepaul.com/fd/HPMeRONWVt5k.xml",
    "Indeed_BA_Mumbai":        "",
    "Indeed_BA_Pune":          "",
    "Naukri_BA_Mumbai":        "",
    "Naukri_BA_NaviMumbai":    "",
    "Naukri_BA_Pune":          "",
}

# ============================================================
#  ← EDIT THIS  (Section 2) — Adzuna API credentials
#  Sign up free at https://developer.adzuna.com/
#
#  Reads from environment variables first (used by GitHub Actions —
#  see ADZUNA_APP_ID / ADZUNA_APP_KEY secrets), and falls back to the
#  hardcoded values below for local runs. Leave the fallback values
#  as "YOUR_APP_ID" / "YOUR_APP_KEY" to skip Adzuna entirely when no
#  env vars are set.
# ============================================================
ADZUNA_APP_ID  = os.environ.get("ADZUNA_APP_ID", "YOUR_APP_ID")
ADZUNA_APP_KEY = os.environ.get("ADZUNA_APP_KEY", "YOUR_APP_KEY")
ADZUNA_COUNTRY = "in"

ADZUNA_ROLES  = ["business analyst", "data analyst", "product analyst"]
ADZUNA_CITIES = ["mumbai", "navi mumbai", "pune"]
ADZUNA_MAX_DAYS_OLD       = 2     # 2-day lookback so a skipped run doesn't lose postings
ADZUNA_RESULTS_PER_SEARCH = 20

# ============================================================
#  ← EDIT THIS  (Section 3) — Shared filters (used by BOTH sources)
# ============================================================
KEYWORDS = [
    "business analyst",
    "ba ",
    "data analyst",
    "product analyst",
    "systems analyst",
    "agile analyst",
    "it analyst",
    "associate analyst",
    "junior analyst",
    "scrum master",
    "product owner",
    "requirements analyst",
    "process analyst",
]

# Words that block a listing OUTRIGHT, no matter what experience is mentioned.
# These describe a role *level*, not a year count, so a low year requirement
# in the text doesn't change that the role itself is senior leadership.
HARD_BLOCKED = [
    "senior manager",
    "senior director",
    "director",
    "vice president",
    "vp ",
    "head of",
    "principal",
    "lead analyst",
]

# Words that are blocked by DEFAULT, but allowed through if the years
# mentioned in the posting are <= MAX_EXPERIENCE_YEARS. "Senior" titles
# vary a lot in actual seniority across companies, so we let the
# experience-years check (if present) decide instead of the word alone.
SOFT_BLOCKED = [
    "senior analyst",
    "senior",
]

MAX_EXPERIENCE_YEARS = 3     # set to None to disable this check entirely

# ============================================================
#  ← EDIT THIS  (Section 4) — Telegram alerts (optional, shared)
#  Reads from environment variables first (TELEGRAM_TOKEN /
#  TELEGRAM_CHAT_ID secrets in GitHub Actions), falls back to the
#  hardcoded values below for local runs.
# ============================================================
TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_TOKEN", "YOUR_TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "YOUR_TELEGRAM_CHAT_ID")

# ============================================================
#  ← EDIT THIS  (Section 5) — Google Sheets
#  See setup instructions in the header comment above.
# ============================================================
# ============================================================
#  ← EDIT THIS  (Section 5) — Google Sheets
#  See setup instructions in the header comment above.
#  SERVICE_ACCOUNT_FILE can be overridden via the GOOGLE_SERVICE_ACCOUNT_FILE
#  env var (used by the GitHub Actions workflow, which writes the
#  credentials JSON to a temp path at runtime).
# ============================================================
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
SHEET_NAME  = "Job_Hunt_Tracker"          # the Google Sheet's name
TAB_NAME    = "📋 Applications"            # the worksheet/tab name inside it


# ============================================================
#  DO NOT EDIT BELOW THIS LINE
# ============================================================

# ── Experience extraction (takes the UPPER bound) ────────────

RANGE_PATTERN = re.compile(
    r'(\d{1,2})\s*(?:\+|to|-|–)\s*(\d{1,2})?\s*(?:years?|yrs?)\b',
    re.IGNORECASE
)
PLAIN_PATTERN = re.compile(
    r'(\d{1,2})\s*\+?\s*(?:years?|yrs?)\b',
    re.IGNORECASE
)
REVERSE_PATTERN = re.compile(
    r'(?:experience|exp)\s*[:\-]?\s*(\d{1,2})\b',
    re.IGNORECASE
)


def extract_max_years(text: str):
    """Return the HIGHEST number of years mentioned in the text.
    e.g. '5 to 8 years' -> 8, not 5 — since the role needs UP TO 8 years,
    which is what should be checked against the cap, not the lowest figure."""
    text = text or ""
    candidates = []

    for m in RANGE_PATTERN.finditer(text):
        low  = int(m.group(1))
        high = int(m.group(2)) if m.group(2) else low
        candidates.append(high)

    for m in PLAIN_PATTERN.finditer(text):
        candidates.append(int(m.group(1)))

    for m in REVERSE_PATTERN.finditer(text):
        candidates.append(int(m.group(1)))

    return max(candidates) if candidates else None


def title_matches(title: str, snippet: str = "") -> bool:
    """Shared filter used by both RSS and Adzuna results.

    Block logic:
      - HARD_BLOCKED words -> always rejected, regardless of years found.
      - SOFT_BLOCKED words (e.g. "senior") -> rejected UNLESS years are
        found AND those years are <= MAX_EXPERIENCE_YEARS. If "senior"
        appears but no year is mentioned anywhere, we can't verify it's
        a low-experience "senior" role, so it stays rejected (safer
        default than letting unverifiable senior titles through).
    """
    combined = f"{title} {snippet}".lower()

    has_keyword = any(k in combined for k in KEYWORDS)
    if not has_keyword:
        return False

    if any(b in combined for b in HARD_BLOCKED):
        return False

    years_found = extract_max_years(combined) if MAX_EXPERIENCE_YEARS is not None else None

    if any(b in combined for b in SOFT_BLOCKED):
        if years_found is None or years_found > MAX_EXPERIENCE_YEARS:
            return False
        # else: "senior" title but years <= cap -> allowed through

    if MAX_EXPERIENCE_YEARS is not None:
        if years_found is not None and years_found > MAX_EXPERIENCE_YEARS:
            return False

    return True


# ── Dedup (URL + title/company signature, backed by the sheet itself) ─

def normalize_signature(title: str, company: str) -> str:
    """Collapse whitespace/case/punctuation so near-identical reposts
    ('Business  Analyst!' vs 'business analyst') still match."""
    raw = f"{title}|{company}".lower()
    raw = re.sub(r'[^a-z0-9|]+', ' ', raw)
    raw = re.sub(r'\s+', ' ', raw).strip()
    return raw


def load_seen_store(sheet) -> dict:
    """Build the dedup store directly from existing sheet rows.
    Tracks both raw URL (column F) and a normalized title+company
    signature (columns A+B), so a job reposted under a new URL is
    still caught as a duplicate."""
    store = {"urls": set(), "signatures": set()}

    try:
        records = sheet.get_all_values()
    except Exception as e:
        print(f"  ⚠️  Could not read existing rows from the sheet ({e}) — starting with an empty dedup store.")
        return store

    for row in records[1:]:  # skip header row
        title   = row[0] if len(row) > 0 else ""
        company = row[1] if len(row) > 1 else ""
        url     = row[5] if len(row) > 5 else ""

        if url:
            store["urls"].add(url)
        if title and company:
            store["signatures"].add(normalize_signature(title, company))

    return store


def is_duplicate(store: dict, url: str, title: str, company: str) -> bool:
    if url and url in store["urls"]:
        return True
    if normalize_signature(title, company) in store["signatures"]:
        return True
    return False


def mark_seen(store: dict, url: str, title: str, company: str):
    if url:
        store["urls"].add(url)
    store["signatures"].add(normalize_signature(title, company))


# ── Google Sheets helpers ─────────────────────────────────────

def get_sheet():
    """Connects to Google Sheets and returns the worksheet, or None
    with a clear, plain-English explanation if anything's missing."""
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        print(f"  ❌ '{SERVICE_ACCOUNT_FILE}' not found in this folder: {os.getcwd()}")
        print("     You need a Google service account key to use Google Sheets.")
        print("     See the setup instructions in the comment header at the top of this script.\n")
        return None

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
        client = gspread.authorize(creds)
    except Exception as e:
        print(f"  ❌ Could not authorize with Google: {e}")
        print(f"     Check that '{SERVICE_ACCOUNT_FILE}' is a valid service account key.\n")
        return None

    try:
        spreadsheet = client.open(SHEET_NAME)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"  ❌ No Google Sheet named '{SHEET_NAME}' found.")
        print(f"     Create it, and make sure it's shared with your service account's client_email")
        print(f"     (find that email inside '{SERVICE_ACCOUNT_FILE}').\n")
        return None
    except Exception as e:
        print(f"  ❌ Could not open Google Sheet '{SHEET_NAME}': {e}\n")
        return None

    try:
        worksheet = spreadsheet.worksheet(TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  ❌ Sheet '{SHEET_NAME}' was opened, but tab '{TAB_NAME}' is missing.")
        print(f"     Available tabs: {[ws.title for ws in spreadsheet.worksheets()]}\n")
        return None
    except Exception as e:
        print(f"  ❌ Could not open tab '{TAB_NAME}': {e}\n")
        return None

    return worksheet


def add_to_sheet(sheet, title: str, company: str, url: str, source: str, location: str = ""):
    row = [
        title,
        company,
        location,
        source,
        datetime.date.today().strftime("%d-%b-%Y"),
        url,
        "Applied",
    ]
    try:
        sheet.append_row(row, value_input_option="USER_ENTERED")
    except Exception as e:
        print(f"  ❌ Could not write row to Google Sheet: {e}")


def escape_html(text: str) -> str:
    """Escape special characters so Telegram's HTML parser doesn't break
    on job titles/companies containing &, <, > (common in real postings,
    e.g. 'R&D Analyst' or 'Sales <Manager>')."""
    if not text:
        return ""
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;"))


def send_telegram(message: str):
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        return
    try:
        response = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            json={
                "chat_id": TELEGRAM_CHAT_ID,
                "text": message,
                "parse_mode": "HTML",
                "disable_web_page_preview": True
            },
            timeout=20
        )

        print("Telegram status:", response.status_code)
        print(response.text)
    except Exception as e:
        print(f"  Telegram send failed: {e}")


# ── Source 1: RSS feeds ──────────────────────────────────────

def scan_rss_feed(name: str, url: str):
    jobs = []
    if not url.strip():
        return jobs

    try:
        headers = {
            "User-Agent": "Mozilla/5.0"
        }
        response = requests.get(url, headers=headers, timeout=10)

        feed = feedparser.parse(response.content)
        print(f"Feed entries found: {len(feed.entries)}")
    except Exception as e:
        print(f"    ERROR parsing {name}: {e}")
        return jobs

    for entry in feed.entries:
        title   = getattr(entry, "title",   "").strip()
        link    = getattr(entry, "link",    "").strip()
        summary = getattr(entry, "summary", "")
        author  = getattr(entry, "author",  "")
        company = author.strip() if author else name.split("_")[0]

        if not title or not link:
            continue

        jobs.append({
            "title":       title,
            "company":     company,
            "location":    "",
            "url":         link,
            "description": summary,
            "source":      name,
        })
    return jobs


# ── Source 2: Adzuna API ─────────────────────────────────────

def scan_adzuna(role: str, city: str):
    jobs = []
    url = f"https://api.adzuna.com/v1/api/jobs/{ADZUNA_COUNTRY}/search/1"
    params = {
        "app_id":           ADZUNA_APP_ID,
        "app_key":          ADZUNA_APP_KEY,
        "results_per_page": ADZUNA_RESULTS_PER_SEARCH,
        "what":             role,
        "where":            city,
        "max_days_old":     ADZUNA_MAX_DAYS_OLD,
        "sort_by":          "date",
        "content-type":     "application/json",
    }
    try:
        r = requests.get(url, params=params, timeout=15)
    except requests.exceptions.RequestException as e:
        print(f"    Network error: {e}")
        return jobs

    if r.status_code == 401:
        print("    ❌ 401 Unauthorized — check your Adzuna app_id/app_key.")
        return jobs
    if r.status_code != 200:
        print(f"    ❌ HTTP {r.status_code}: {r.text[:120]}")
        return jobs

    data    = r.json()
    results = data.get("results", [])

    for job in results:
        jobs.append({
            "title":       job.get("title", ""),
            "company":     job.get("company", {}).get("display_name", ""),
            "location":    job.get("location", {}).get("display_name", ""),
            "url":         job.get("redirect_url", ""),
            "description": job.get("description", ""),
            "source":      f"Adzuna_{role.replace(' ', '')}_{city.replace(' ', '')}",
        })
    return jobs


# ── MAIN ─────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 62)
    print("  Job Scanner (RSS + Adzuna) —", datetime.date.today().strftime("%d %B %Y"))
    print(f"  Experience cap: ≤ {MAX_EXPERIENCE_YEARS} yrs" if MAX_EXPERIENCE_YEARS else "  No experience cap")
    print("=" * 62)

    sheet = get_sheet()
    sheet_ok = sheet is not None

    if sheet_ok:
        seen_store = load_seen_store(sheet)
    else:
        print("  ⚠️  Continuing without Google Sheets — jobs will be found but NOT logged anywhere,")
        print("     and duplicate detection across runs will not work until this is fixed.\n")
        seen_store = {"urls": set(), "signatures": set()}

    print(f"  Already tracked: {len(seen_store['urls'])} URL(s), {len(seen_store['signatures'])} title+company signature(s)\n")

    new_jobs = []

    # ── Run RSS feeds ────────────────────────────────────────
    active_rss = {k: v for k, v in RSS_FEEDS.items() if v.strip()}
    empty_rss  = [k for k, v in RSS_FEEDS.items() if not v.strip()]

    print(f"  [RSS] Active feeds: {len(active_rss)} / {len(RSS_FEEDS)}")
    if empty_rss:
        print(f"  [RSS] Not set up yet: {', '.join(empty_rss)}")
    print()

    for name, url in active_rss.items():
        print(f"[RSS] Checking {name} ...")

        all_jobs = scan_rss_feed(name, url)

        fetched_count = len(all_jobs)
        filtered_count = 0

        for job in all_jobs:
            if not job["url"]:
                continue
            if is_duplicate(seen_store, job["url"], job["title"], job["company"]):
                continue

            if not title_matches(job["title"], job.get("description", "")):
                continue

            filtered_count += 1

            if sheet_ok:
                add_to_sheet(sheet, job["title"], job["company"], job["url"], job["source"], job["location"])

            mark_seen(seen_store, job["url"], job["title"], job["company"])
            new_jobs.append(job)

        print(f"   Fetched: {fetched_count}")
        print(f"   Matching after filters: {filtered_count}")

    # ── Run Adzuna ───────────────────────────────────────────
    print()
    adzuna_ready = ADZUNA_APP_ID != "YOUR_APP_ID" and ADZUNA_APP_KEY != "YOUR_APP_KEY"

    if not adzuna_ready:
        print("  [Adzuna] Skipped — no app_id/app_key set in Section 2.")
    else:
        print(f"  [Adzuna] Posted within last {ADZUNA_MAX_DAYS_OLD} day(s)")
        for role in ADZUNA_ROLES:
            for city in ADZUNA_CITIES:
                print(f"  [Adzuna] Searching {role} / {city} ...", end="  ")
                results = scan_adzuna(role, city)
                new_count = 0
                for job in results:
                    if not job["url"]:
                        continue
                    if is_duplicate(seen_store, job["url"], job["title"], job["company"]):
                        continue
                    if not title_matches(job["title"], job.get("description", "")):
                        continue
                    if sheet_ok:
                        add_to_sheet(sheet, job["title"], job["company"], job["url"], job["source"], job["location"])
                    mark_seen(seen_store, job["url"], job["title"], job["company"])
                    new_jobs.append(job)
                    new_count += 1
                print(f"{new_count} new matching job(s) (of {len(results)} returned)")

    # ── Combined summary + single Telegram message ──────────
    print("\n" + "=" * 62)
    if new_jobs:
        if sheet_ok:
            print(f"  ✅  {len(new_jobs)} new job(s) logged to Google Sheets (RSS + Adzuna combined):\n")
        else:
            print(f"  ⚠️  {len(new_jobs)} new job(s) found, but NOT saved (Google Sheets issue above):\n")

        for j in new_jobs:
            loc = f" ({j['location']})" if j.get("location") else ""
            print(f"     • {j['title']}  @  {j['company']}{loc}  [{j['source']}]")

        send_telegram(
            f"🎯 <b>{len(new_jobs)} new job(s) found — "
            f"{datetime.date.today().strftime('%d %b')}</b>"
        )

        for j in new_jobs:
            loc = f" ({j['location']})" if j.get("location") else ""

            title_escaped = escape_html(j['title'])
            company_escaped = escape_html(j['company'])
            job_url = escape_html(j['url'])

            msg = (
                f"💼 <b>{title_escaped}</b>\n"
                f"🏢 {company_escaped}{loc}\n"
                f"🔗 {job_url}\n"
                f"📌 {j['source']}"
            )

            send_telegram(msg)
    else:
        print("  ℹ️  No new matching jobs found this run (RSS + Adzuna combined).")
        if not active_rss and not adzuna_ready:
            print("  → Neither RSS feeds nor Adzuna are set up yet. See Sections 1 and 2.")

    print(f"\n  Done. Open your '{SHEET_NAME}' Google Sheet to review.")
    print("=" * 62 + "\n")


if __name__ == "__main__":
    main()
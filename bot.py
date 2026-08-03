"""
Content & Digital Marketing Job Scraper Bot v5.0
========================
منابع رایگان:
  • Remotive.com
  • Jobicy.com
  • Arbeitnow
  • Adzuna (با API key)
  • FindWork.dev
  • Cloudflare Worker

منابع پولی (اختیاری):
  • JSearch via RapidAPI  — پلن رایگان 200 req/ماه

Cover Letter:
  • هر آگهی یک دکمه "ChatGPT Cover Letter" داره
  • کلیک → باز شدن ChatGPT با پرامپت آماده

ذخیره‌سازی اختیاری:
  • Google Sheets (Batch append)

متغیرهای محیطی (GitHub Secrets):
  TELEGRAM_BOT_TOKEN   — اجباری
  TELEGRAM_CHAT_ID     — اجباری
  RAPIDAPI_KEY         — اختیاری
  GSHEET_CREDENTIALS   — اختیاری (JSON)
  GSHEET_ID            — اختیاری
  CF_WORKER_URL        — اختیاری
  ADZUNA_APP_ID        — اختیاری
  ADZUNA_API_KEY       — اختیاری
"""

import html
import json
import logging
import os
import re
import time
import traceback
import urllib.parse
from collections import OrderedDict
from datetime import datetime, timezone
from pathlib import Path

import requests
from dotenv import load_dotenv

try:
    import gspread
    from google.oauth2.service_account import Credentials
    SHEETS_AVAILABLE = True
except ImportError:
    SHEETS_AVAILABLE = False

SCRIPT_DIR = Path(__file__).parent

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger(__name__)

load_dotenv()

TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

if not TELEGRAM_TOKEN:
    raise ValueError("TELEGRAM_BOT_TOKEN not set")
if not TELEGRAM_CHAT_ID:
    raise ValueError("TELEGRAM_CHAT_ID not set")

RAPIDAPI_KEY       = os.environ.get("RAPIDAPI_KEY", "")
GSHEET_CREDENTIALS = os.environ.get("GSHEET_CREDENTIALS", "")
GSHEET_ID          = os.environ.get("GSHEET_ID", "")
GSHEET_SHEET_NAME  = "Jobs"

CF_WORKER_URL    = os.environ.get("CF_WORKER_URL", "")
ADZUNA_APP_ID    = os.environ.get("ADZUNA_APP_ID", "")
ADZUNA_API_KEY   = os.environ.get("ADZUNA_API_KEY", "")

SEEN_JOBS_FILE   = SCRIPT_DIR / "seen_jobs.txt"
MAX_SEEN_JOBS    = 3000
MAX_JOBS_PER_RUN = 20
MIN_FIT_SCORE    = 35
MAX_JOB_AGE_DAYS = 7

JSEARCH_QUERIES = {
    1: [
        "Content Specialist remote worldwide",
        "Digital Marketing Specialist remote worldwide",
        "Social Media Specialist remote worldwide",
    ],
    2: [
        "E-commerce Content Specialist remote worldwide",
    ],
    3: [
        "Content Operations Specialist remote worldwide",
        "Product Content Specialist remote worldwide",
    ],
}

_DEFAULT_SKILLS = [
    "digital marketing", "content strategy", "content management",
    "content operations", "e-commerce content", "ecommerce content",
    "product content", "product listing", "product descriptions",
    "content quality control", "social media management",
    "market research", "competitor analysis", "brand communication",
    "ai-assisted content", "prompt writing", "cms", "basic seo",
    "google workspace", "microsoft office", "photoshop", "illustrator",
]
_user_skills_env = os.environ.get("USER_SKILLS", "")
MY_SKILLS = [s.strip().lower() for s in _user_skills_env.split(",") if s.strip()] if _user_skills_env else _DEFAULT_SKILLS

BLACKLIST_KEYWORDS = [
    "us residents only", "must reside in us", "must be located in us",
    "must be based in the us", "must be based in us",
    "must be authorized to work in the us", "us work authorization required",
    "eu work authorization required", "security clearance required",
    "native english speaker only", "commission only", "unpaid internship",
    "cold calling", "door to door", "full stack", "fullstack",
    "software engineer", "software developer", "data scientist",
    "senior content", "senior marketing", "head of content",
    "head of marketing", "director of content", "director of marketing", "vp of",
    "10+ years", "8+ years", "7+ years",
]

BOOST_KEYWORDS = {
    "e-commerce content": 24, "ecommerce content": 24,
    "content operations": 22, "product content": 22,
    "product listing": 20, "content specialist": 20,
    "digital marketing specialist": 18, "social media specialist": 18,
    "content strategist": 18, "content manager": 15,
    "content editor": 12, "marketing coordinator": 12,
    "content quality": 12, "ai-assisted content": 10,
    "prompt writing": 10, "cms": 8, "mid-level": 10,
    "specialist": 8, "part-time": 8, "contract": 5,
    "remote-first": 8, "async": 5, "flexible": 4,
}

_SKILL_PATTERNS   = {s: re.compile(r"\b" + re.escape(s) + r"\b", re.I) for s in MY_SKILLS}
_BOOST_PATTERNS   = {kw: re.compile(r"\b" + re.escape(kw) + r"\b", re.I) for kw in BOOST_KEYWORDS}
_BLACKLIST_PATTERNS = {kw: re.compile(r"\b" + re.escape(kw.lower()) + r"\b", re.I) for kw in BLACKLIST_KEYWORDS}

# ── Prompt Template ─────────────────────────────────────────────────────────

CL_PROMPT_TEMPLATE = os.environ.get("CL_PROMPT", "")

def load_prompt_template() -> str:
    # اول از متغیر محیطی CL_PROMPT بخون
    if CL_PROMPT_TEMPLATE:
        return CL_PROMPT_TEMPLATE.strip()
    # اگه نیست، از فایل prompt.txt بخون
    try:
        prompt_file = SCRIPT_DIR / "prompt.txt"
        if prompt_file.exists():
            with open(prompt_file, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content:
                    return content
    except Exception as e:
        log.warning(f"Could not load prompt.txt: {e}")
    return "Write a concise, natural cover letter for the '{title}' position at '{company}'. Highlight my experience in digital marketing, content management, e-commerce product information, content quality control, cross-functional coordination, and AI-assisted content. Do not invent experience. Job link: {url}"

# ── Seen Jobs Cache ─────────────────────────────────────────────────────────

def load_seen_jobs() -> OrderedDict:
    seen = OrderedDict()
    if SEEN_JOBS_FILE.exists():
        for line in SEEN_JOBS_FILE.read_text(encoding="utf-8").splitlines():
            if line.strip():
                seen[line.strip()] = True
        log.info(f"Loaded {len(seen)} seen IDs")
    else:
        log.info("No cache — starting fresh")
    return seen

def save_seen_jobs(seen: OrderedDict) -> None:
    ids = list(seen.keys())
    if len(ids) > MAX_SEEN_JOBS:
        ids = ids[-MAX_SEEN_JOBS:]
    SEEN_JOBS_FILE.write_text("\n".join(ids), encoding="utf-8")
    log.info(f"Saved {len(ids)} IDs to cache")

# ── Fit Score ───────────────────────────────────────────────────────────────

def calculate_fit_score(job: dict) -> tuple:
    score = 0
    matched_skills = []
    title    = (job.get("title") or "").lower()
    desc     = (job.get("description") or "").lower()
    combined = f"{title} {desc}"

    for kw, pts in BOOST_KEYWORDS.items():
        if _BOOST_PATTERNS[kw].search(combined):
            score += pts

    for skill in MY_SKILLS:
        if _SKILL_PATTERNS[skill].search(combined):
            matched_skills.append(skill)
            score += 7

    target_title_terms = [
        "content specialist", "content strategist", "content manager",
        "content editor", "content operations", "product content",
        "e-commerce content", "ecommerce content", "digital marketing",
        "social media", "marketing coordinator",
    ]
    if any(term in title for term in target_title_terms):
        score += 12
    if job.get("salary"):
        score += 10
    if job.get("remote"):
        score += 8
    if any(re.search(r"\b" + w + r"\b", title) for w in ["junior", "associate", "entry", "jr"]):
        score += 10

    return min(score, 100), matched_skills[:4]

# ── Free Sources ────────────────────────────────────────────────────────────

def fetch_remotive() -> list:
    endpoints = [
        "https://remotive.com/api/remote-jobs?search=content+specialist&limit=20",
        "https://remotive.com/api/remote-jobs?search=digital+marketing&limit=20",
        "https://remotive.com/api/remote-jobs?search=social+media&limit=15",
    ]
    results = []
    for url in endpoints:
        try:
            resp = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
            resp.raise_for_status()
            for j in resp.json().get("jobs", []):
                results.append({
                    "id":           f"remotive_{j.get('id', '')}",
                    "title":        j.get("title", ""),
                    "company":      j.get("company_name", ""),
                    "description":  j.get("description", ""),
                    "salary":       j.get("salary", ""),
                    "remote":       True,
                    "url":          j.get("url", ""),
                    "source":       "Remotive",
                    "source_emoji": "🌐",
                    "posted_at":    (j.get("publication_date") or "")[:10],
                    "location":     "Remote",
                })
        except Exception as e:
            log.error(f"Remotive error: {e}")
        time.sleep(1)
    log.info(f"Remotive -> {len(results)} jobs")
    return results

def fetch_jobicy() -> list:
    endpoints = [
        "https://jobicy.com/api/v2/remote-jobs?tag=content-marketing&count=20",
        "https://jobicy.com/api/v2/remote-jobs?tag=digital-marketing&count=20",
        "https://jobicy.com/api/v2/remote-jobs?tag=social-media&count=15",
    ]
    results = []
    for url in endpoints:
        try:
            resp = requests.get(url, timeout=15)
            resp.raise_for_status()
            for j in resp.json().get("jobs", []):
                lo = j.get("annualSalaryMin")
                hi = j.get("annualSalaryMax")
                cur = j.get("annualSalaryCurrency", "USD")
                sal = f"{cur} {int(lo):,}-{int(hi):,}/yr" if lo and hi else (f"{cur} {int(lo):,}+/yr" if lo else "")
                results.append({
                    "id":           f"jobicy_{j.get('id', '')}",
                    "title":        j.get("jobTitle", ""),
                    "company":      j.get("companyName", ""),
                    "description":  j.get("jobDescription", ""),
                    "salary":       sal,
                    "remote":       True,
                    "url":          j.get("url", ""),
                    "source":       "Jobicy",
                    "source_emoji": "🟢",
                    "posted_at":    (j.get("pubDate") or "")[:10],
                    "location":     "Remote",
                })
        except Exception as e:
            log.error(f"Jobicy error: {e}")
        time.sleep(1)
    log.info(f"Jobicy -> {len(results)} jobs")
    return results

def fetch_arbeitnow() -> list:
    TARGET_TERMS = ["content specialist", "content manager", "content strategist", "content operations", "digital marketing", "social media", "marketing coordinator", "e-commerce content", "ecommerce content", "product content", "content editor"]
    try:
        resp = requests.get("https://arbeitnow.com/api/job-board-api", timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        results = []
        for j in resp.json().get("data", []):
            if not j.get("remote"):
                continue
            title = (j.get("title") or "").lower()
            desc  = (j.get("description") or "").lower()[:300]
            if not any(t in title or t in desc for t in TARGET_TERMS):
                continue
            results.append({
                "id":           f"arbeitnow_{j.get('slug', '')}",
                "title":        j.get("title", ""),
                "company":      j.get("company_name", ""),
                "description":  j.get("description", ""),
                "salary":       "",
                "remote":       True,
                "url":          j.get("url", ""),
                "source":       "Arbeitnow",
                "source_emoji": "🔷",
                "posted_at":    datetime.now(timezone.utc).strftime("%Y-%m-%d"),
                "location":     "Remote",
            })
        log.info(f"Arbeitnow -> {len(results)} jobs")
        return results
    except Exception as e:
        log.error(f"Arbeitnow error: {e}")
        return []

def fetch_adzuna() -> list:
    if not ADZUNA_APP_ID or not ADZUNA_API_KEY:
        return []
    results = []
    for q in ["content specialist", "digital marketing specialist", "social media specialist"]:
        try:
            resp = requests.get(
                f"https://api.adzuna.com/v1/api/jobs/us/search/1",
                params={"app_id": ADZUNA_APP_ID, "app_key": ADZUNA_API_KEY,
                        "what": q, "what_or": "remote", "max_days_old": 7,
                        "results_per_page": 15, "content-type": "application/json"},
                timeout=15,
            )
            resp.raise_for_status()
            for j in resp.json().get("results", []):
                results.append({
                    "id":           f"adzuna_{j.get('id', '')}",
                    "title":        j.get("title", ""),
                    "company":      (j.get("company") or {}).get("display_name", ""),
                    "description":  j.get("description", ""),
                    "salary":       f"${int(float(j['salary_min'])):,}-${int(float(j.get('salary_max') or j.get('salary_min'))):,}/yr" if j.get("salary_min") else "",
                    "remote":       True,
                    "url":          j.get("redirect_url", ""),
                    "source":       "Adzuna",
                    "source_emoji": "🟡",
                    "posted_at":    (j.get("created") or "")[:10],
                    "location":     j.get("location", {}).get("display_name", "Remote"),
                })
        except Exception as e:
            log.error(f"Adzuna error ({q}): {e}")
        time.sleep(1)
    log.info(f"Adzuna -> {len(results)} jobs")
    return results

def fetch_findwork() -> list:
    TARGET_TERMS = ["content specialist", "content manager", "content strategist", "content operations", "digital marketing", "social media", "marketing coordinator", "product content", "content editor"]
    try:
        resp = requests.get(
            "https://findwork.dev/api/jobs/",
            params={"search": "content marketing", "remote": "true", "order_by": "-date_posted"},
            headers={"User-Agent": "Mozilla/5.0 (compatible; ContentMarketingJobBot/5.0)"},
            timeout=15,
        )
        if resp.status_code == 403:
            log.warning("FindWork.dev: access denied")
            return []
        resp.raise_for_status()
        results = []
        for j in resp.json().get("results", []):
            title = (j.get("role") or "").lower()
            desc  = (j.get("text") or "").lower()[:500]
            if not any(t in title or t in desc for t in TARGET_TERMS):
                continue
            results.append({
                "id":           f"findwork_{j.get('id', '')}",
                "title":        j.get("role", ""),
                "company":      j.get("company_name", ""),
                "description":  j.get("text", ""),
                "salary":       "",
                "remote":       j.get("remote", True),
                "url":          j.get("url", ""),
                "source":       "FindWork",
                "source_emoji": "🟣",
                "posted_at":    (j.get("date_posted") or "")[:10],
                "location":     j.get("location") or "Remote",
            })
        log.info(f"FindWork -> {len(results)} jobs")
        return results
    except Exception as e:
        log.error(f"FindWork error: {e}")
        return []

def fetch_cloudflare_worker() -> list:
    if not CF_WORKER_URL:
        return []
    worker_url = CF_WORKER_URL.rstrip("/")
    if not worker_url.endswith("/jobs"):
        worker_url += "/jobs"
    try:
        resp = requests.get(worker_url, headers={"User-Agent": "ContentMarketingJobBot/5.0"}, timeout=20)
        if resp.status_code in (401, 404):
            log.error(f"CF Worker: {resp.status_code}")
            return []
        resp.raise_for_status()
        data = resp.json()
        if data.get("status") != "ok":
            return []
        jobs = []
        for j in data.get("jobs", []):
            if not j.get("id") or not j.get("title"):
                continue
            jobs.append({
                "id":           str(j.get("id", "")),
                "title":        j.get("title", ""),
                "company":      j.get("company", ""),
                "description":  j.get("description", ""),
                "salary":       j.get("salary", ""),
                "remote":       j.get("remote", True),
                "url":          j.get("url", ""),
                "source":       j.get("source", "CF Worker"),
                "source_emoji": j.get("source_emoji", "☁️"),
                "posted_at":    (j.get("posted_at") or "")[:10],
                "location":     j.get("location", "Remote"),
            })
        log.info(f"CF Worker -> {len(jobs)} jobs")
        return jobs
    except Exception as e:
        log.error(f"CF Worker error: {e}")
        return []

# ── JSearch API (اختیاری) ───────────────────────────────────────────────────

def _should_run_p3() -> bool:
    return datetime.now(timezone.utc).day % 2 == 0

def search_jsearch(query: str) -> list:
    if not RAPIDAPI_KEY:
        return []

    url = "https://jsearch.p.rapidapi.com/search-v2"

    headers = {
        "x-rapidapi-key": RAPIDAPI_KEY,
        "x-rapidapi-host": "jsearch.p.rapidapi.com",
    }

    params = {
        "query": query,
        "date_posted": "week",
        "work_from_home": "true",
    }

    try:
        resp = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=30,
        )

        if resp.status_code == 429:
            log.warning("JSearch rate limit reached")
            return []

        if resp.status_code in (401, 403):
            log.error(f"JSearch authorization error: {resp.status_code}")
            return []

        resp.raise_for_status()

        data = resp.json()

        if data.get("status") != "OK":
            log.warning(f"JSearch unexpected response: {data}")
            return []

        jobs = [_normalize_jsearch(j) for j in data.get("data", [])]
        log.info(f"JSearch -> {len(jobs)} jobs for: {query}")
        return jobs

    except requests.exceptions.Timeout:
        log.warning(f"JSearch timeout for: {query}")
        return []

    except Exception as e:
        log.error(f"JSearch error for '{query}': {e}")
        return []
def _normalize_jsearch(j: dict) -> dict:
    salary = j.get("job_salary_string", "")
    if not salary and j.get("job_min_salary"):
        lo = int(j["job_min_salary"])
        hi = int(j.get("job_max_salary") or lo)
        per = {"year": "/yr", "month": "/mo", "hour": "/hr"}.get((j.get("job_salary_period") or "").lower(), "")
        salary = f"${lo:,}-${hi:,}{per}" if lo != hi else f"${lo:,}+{per}"

    city, country = j.get("job_city") or "", j.get("job_country") or ""
    loc_parts = [p for p in (city, country) if p]
    loc = ", ".join(loc_parts) or "Remote"

    return {
        "id":           j.get("job_id", ""),
        "title":        j.get("job_title", ""),
        "company":      j.get("employer_name", ""),
        "description":  j.get("job_description", ""),
        "salary":       salary,
        "remote":       True,
        "url":          j.get("job_apply_link") or j.get("job_google_link") or "",
        "source":       j.get("job_publisher", "JSearch"),
        "source_emoji": "🔍",
        "posted_at":    (j.get("job_posted_at_datetime_utc") or "")[:10],
        "location":     loc,
    }

# ── Filters ─────────────────────────────────────────────────────────────────

def is_blacklisted(job: dict) -> tuple:
    text = f"{(job.get('title') or '').lower()} {(job.get('description') or '')[:2000].lower()}"
    for kw, pattern in _BLACKLIST_PATTERNS.items():
        if pattern.search(text):
            return True, kw
    return False, ""

def is_too_old(job: dict) -> bool:
    posted = (job.get("posted_at") or "")[:10]
    if not posted:
        return False
    try:
        dt = datetime.strptime(posted, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        return (datetime.now(timezone.utc) - dt).days > MAX_JOB_AGE_DAYS
    except Exception:
        return False

# ── Telegram ────────────────────────────────────────────────────────────────

def _score_bar(score: int) -> str:
    filled = round(score / 10)
    return "█" * filled + "░" * (10 - filled)

def format_job(job: dict, score: int, skills: list) -> str:
    title   = html.escape(job.get("title") or "No Title")
    company = html.escape(job.get("company") or "Unknown")
    salary  = job.get("salary") or ""
    source  = html.escape(job.get("source") or "")
    semoji  = job.get("source_emoji", "🌐")
    posted  = job.get("posted_at") or ""
    loc     = html.escape(job.get("location") or "Remote")

    lines = [
        f"💼 <b>{title}</b>",
        f"🏢 {company}",
        f"📍 {loc}",
    ]
    if salary:
        lines.append(f"💰 <b>{html.escape(str(salary))}</b>")
    lines.append(f"📊 {_score_bar(score)} {score}/100")
    if skills:
        lines.append(f"✅ {', '.join(html.escape(s) for s in skills)}")
    lines.append(f"{semoji} {source}")
    if posted:
        lines.append(f"📅 {posted}")

    return "\n".join(lines)

def build_job_buttons(job: dict) -> dict:
    """ساخت دکمه‌ها: Apply + ChatGPT Cover Letter"""
    url = job.get("url", "")
    if not url:
        return {}

    title   = job.get("title", "")
    company = job.get("company", "")

    template = load_prompt_template()
    prompt   = template.format(title=title, company=company, url=url)
    safe_prompt = urllib.parse.quote(prompt)
    chatgpt_url = f"https://chatgpt.com/?q={safe_prompt}"

    return {"inline_keyboard": [
        [{"text": "📝 Apply Now", "url": url}],
        [{"text": "🤖 ChatGPT Cover Letter", "url": chatgpt_url}]
    ]}

def send_telegram(text: str, reply_markup: dict = None, _retries: int = 3) -> bool:
    api_url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "link_preview_options": {"is_disabled": True},
    }
    if reply_markup:
        payload["reply_markup"] = reply_markup

    for attempt in range(1, _retries + 1):
        try:
            resp = requests.post(api_url, json=payload, timeout=15)
            if resp.ok:
                return True

            if resp.status_code == 429:
                retry_after = resp.json().get("parameters", {}).get("retry_after", 30)
                log.warning(f"Telegram Flood Wait — sleeping {retry_after}s (attempt {attempt}/{_retries})")
                time.sleep(retry_after + 1)
                continue

            log.error(f"Telegram {resp.status_code}: {resp.text[:200]}")
            return False
        except requests.exceptions.Timeout:
            log.warning(f"Telegram timeout (attempt {attempt}/{_retries})")
            if attempt < _retries:
                time.sleep(3)
        except Exception as e:
            log.error(f"Telegram error: {e}")
            return False
    return False

# ── Google Sheets ────────────────────────────────────────────────────────────

def get_sheets_client():
    if not SHEETS_AVAILABLE or not GSHEET_CREDENTIALS or not GSHEET_ID:
        return None
    try:
        creds = Credentials.from_service_account_info(
            json.loads(GSHEET_CREDENTIALS),
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
        )
        log.info("Google Sheets connected")
        return gspread.authorize(creds)
    except Exception as e:
        log.error(f"Sheets auth error: {e}")
        return None

def ensure_sheet_headers(client) -> None:
    if not client:
        return
    try:
        sheet = client.open_by_key(GSHEET_ID).worksheet(GSHEET_SHEET_NAME)
        if not sheet.row_values(1):
            sheet.insert_row(
                ["Job Title", "Company", "Source", "Apply Link", "Posted",
                 "Salary", "Fit Score", "Location", "Saved At (UTC)", "Status", "Cover Letter Link"],
                1,
            )
    except Exception as e:
        log.error(f"Sheet header error: {e}")

def batch_append_to_sheet(client, rows: list) -> None:
    if not client or not rows:
        return
    try:
        sheet = client.open_by_key(GSHEET_ID).worksheet(GSHEET_SHEET_NAME)
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
        log.info(f"Batch appended {len(rows)} rows to Google Sheets")
    except Exception as e:
        log.error(f"Sheet batch append error: {e}")

# ── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    log.info(f"=== Content & Digital Marketing Job Scraper v5.0 started at {now} ===")

    seen_jobs = load_seen_jobs()
    sheets = get_sheets_client()
    ensure_sheet_headers(sheets)

    raw_jobs = []
    source_counts = {}

    # ── منابع رایگان ─────────────────────────────────────────────────────────
    for fn, name in [
        (fetch_remotive, "Remotive"),
        (fetch_jobicy, "Jobicy"),
        (fetch_arbeitnow, "Arbeitnow"),
        (fetch_adzuna, "Adzuna"),
        (fetch_findwork, "FindWork"),
        (fetch_cloudflare_worker, "CF Worker"),
    ]:
        try:
            jobs = fn()
            source_counts[name] = len(jobs)
            raw_jobs.extend(jobs)
        except Exception as e:
            log.error(f"{name} failed: {e}\n{traceback.format_exc()}")
            source_counts[name] = 0

    # ── JSearch (اختیاری) ────────────────────────────────────────────────────
    jsearch_total = 0
    for priority in sorted(JSEARCH_QUERIES.keys()):
        if priority == 3 and not _should_run_p3():
            log.info("Skipping P3 JSearch queries (odd day)")
            continue
        for query in JSEARCH_QUERIES[priority]:
            try:
                jobs = search_jsearch(query)
                jsearch_total += len(jobs)
                raw_jobs.extend(jobs)
            except Exception as e:
                log.error(f"JSearch '{query}': {e}")
            time.sleep(1.5)
    source_counts["JSearch"] = jsearch_total

    # ── فیلتر + امتیازدهی ────────────────────────────────────────────────────
    seen_ids = set()
    title_keys = set()
    stats = {"blacklisted": 0, "seen": 0, "old": 0, "low_score": 0}
    qualified = []

    for job in raw_jobs:
        try:
            jid = job.get("id") or job.get("url") or ""
            title_key = f"{(job.get('title') or '').lower().strip()}|{(job.get('company') or '').lower().strip()}"

            if not jid:
                continue
            if jid in seen_jobs or jid in seen_ids:
                stats["seen"] += 1
                continue
            if title_key in title_keys:
                stats["seen"] += 1
                seen_ids.add(jid)
                seen_jobs[jid] = True
                continue

            seen_ids.add(jid)
            seen_jobs[jid] = True
            title_keys.add(title_key)

            bl, _ = is_blacklisted(job)
            if bl:
                stats["blacklisted"] += 1
                continue

            if is_too_old(job):
                stats["old"] += 1
                continue

            score, skills = calculate_fit_score(job)
            if score < MIN_FIT_SCORE:
                stats["low_score"] += 1
                continue

            qualified.append((job, score, skills))
        except Exception as e:
            log.error(f"Processing error: {e}")

    qualified.sort(key=lambda x: x[1], reverse=True)

    log.info(
        f"Qualified: {len(qualified)} | BL: {stats['blacklisted']} | "
        f"Seen: {stats['seen']} | Old: {stats['old']} | Low: {stats['low_score']}"
    )

    # ── ارسال به تلگرام ──────────────────────────────────────────────────────
    active_sources = {k: v for k, v in source_counts.items() if v > 0}
    sources_line = " | ".join(f"{k}: {v}" for k, v in active_sources.items())

    if not qualified:
        send_telegram(
            f"🔍 <b>Daily Report</b>\n📅 {now}\n\n"
            f"No qualified jobs found.\n\n"
            f"📌 {sources_line or 'No sources'}\n"
            f"⛔ {stats['blacklisted']} filtered | "
            f"📉 {stats['low_score']} low score | "
            f"🔁 {stats['seen']} duplicates | "
            f"🕐 {stats['old']} old"
        )
        save_seen_jobs(seen_jobs)
        return

    send_telegram(
        f"🤖 <b>New Content & Digital Marketing Jobs</b>\n"
        f"📅 {now}\n\n"
        f"✅ <b>{len(qualified)}</b> jobs (sorted by fit)\n"
        f"⛔ {stats['blacklisted']} filtered | "
        f"📉 {stats['low_score']} low | "
        f"🔁 {stats['seen']} dupes\n\n"
        f"📌 {sources_line}\n"
        f"🤖 ChatGPT Cover Letter: ON\n"
        f"➖➖➖➖➖➖➖➖"
    )
    time.sleep(1.5)

    sent = 0
    sheet_rows = []

    for job, score, skills in qualified[:MAX_JOBS_PER_RUN]:
        try:
            buttons = build_job_buttons(job)
            msg = format_job(job, score, skills)

            if send_telegram(msg, reply_markup=buttons if buttons else None):
                sent += 1
                sheet_rows.append([
                    job.get("title", ""), job.get("company", ""),
                    job.get("source", ""), job.get("url", ""),
                    job.get("posted_at", ""), job.get("salary", ""),
                    score, job.get("location", ""),
                    datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"),
                    "New", "ChatGPT URL"
                ])

            time.sleep(1.5)
        except Exception as e:
            log.error(f"Send error: {e}")

    batch_append_to_sheet(sheets, sheet_rows)
    save_seen_jobs(seen_jobs)
    log.info(f"=== Done. Sent {sent}/{len(qualified)} ===")

if __name__ == "__main__":
    main()

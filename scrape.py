#!/usr/bin/env python3
"""
BizDeadlines scraper — regenerates data.js from the ABS xlsx + live conference websites.

Usage:
    python scrape.py                    # full run: rebuild + scrape
    python scrape.py --build-only       # rebuild from xlsx, skip scraping
    python scrape.py --scrape-only      # scrape existing data.js (skip xlsx rebuild)
    python scrape.py --force            # ignore the 30-day cooldown

Run monthly via GitHub Actions (.github/workflows/update-data.yml).
"""

import argparse
import json
import re
import sys
import time
import warnings
from datetime import date, datetime, timedelta
from pathlib import Path

warnings.filterwarnings("ignore")  # suppress LibreSSL noise

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    sys.exit("Missing dependencies. Run: pip install requests beautifulsoup4 lxml openpyxl")

try:
    import openpyxl
    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False

# ── Configuration ────────────────────────────────────────────────────────────

REPO_ROOT   = Path(__file__).parent
XLSX_GLOB   = list(REPO_ROOT.glob("ABS_Journal_List_*.xlsx"))
XLSX_PATH   = XLSX_GLOB[0] if XLSX_GLOB else None
DATA_JS     = REPO_ROOT / "data.js"
TODAY       = date.today()
SCRAPE_GAP  = timedelta(days=30)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# ── Month helpers ────────────────────────────────────────────────────────────

MONTH_NUM = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
    "jan": 1, "feb": 2, "mar": 3, "apr": 4,
    "jun": 6, "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

def first_month(s):
    if not s:
        return None
    if "varies" in s.lower():
        return None
    for name, n in MONTH_NUM.items():
        if name in s.lower():
            return n
    return None

def last_month(s):
    if not s:
        return None
    for part in reversed(re.split(r"[–\-/,]", s.lower())):
        for name, n in MONTH_NUM.items():
            if name in part.strip():
                return n
    return None

def next_deadline_from_typical(typical):
    """Return the next occurrence of the deadline month (never before today)."""
    m = first_month(typical)
    if m is None:
        return None
    d = date(TODAY.year, m, 15)
    if d <= TODAY:
        d = date(TODAY.year + 1, m, 15)
    return d

def conf_dates_from_typical(dl_date, typical_conf):
    if dl_date is None or not typical_conf:
        return None, None
    ms = first_month(typical_conf)
    me = last_month(typical_conf)
    if ms is None:
        return None, None
    y = dl_date.year if ms >= dl_date.month else dl_date.year + 1
    start = date(y, ms, 15)
    end   = date(y, me if me else ms, 18)
    return start, end

# ── Topic assignment ─────────────────────────────────────────────────────────

JOURNAL_TOPIC_RULES = [
    (["the accounting review", "contemporary accounting", "journal of accounting",
      "accounting horizons", "behavioral research in accounting",
      "accounting, organizations", "journal of management accounting",
      "journal of international accounting", "accounting literature",
      "management accounting research", "auditing:",
      "european accounting review", "british accounting review"], ["ACC"]),
    (["review of financial studies", "journal of finance", "financial management",
      "review of finance", "journal of financial economics",
      "journal of financial and quantitative",
      "real estate economics", "journal of real estate",
      "financial analysts", "journal of risk and insurance",
      "european financial management", "european journal of finance",
      "financial review"], ["FIN"]),
    (["journal of marketing research", "journal of marketing",
      "international journal of research in marketing",
      "journal of consumer research", "journal of consumer psychology",
      "journal of the academy of marketing science",
      "journal of advertising", "industrial marketing management",
      "international marketing review", "journal of business-to-business",
      "journal of international marketing", "journal of public policy and marketing",
      "journal of interactive marketing", "marketing letters",
      "marketing science", "journal of service research",
      "annals of tourism", "european journal of marketing"], ["MKT"]),
    (["information systems research", "mis quarterly", "journal of information technology",
      "journal of management information systems", "journal of strategic information systems",
      "information and management", "information systems journal",
      "european journal of information systems",
      "business & information systems engineering"], ["IS"]),
    (["human-computer studies"], ["HCI"]),
    (["ieee transactions on knowledge and data engineering"], ["DS", "TECH"]),
    (["management science", "operations research", "manufacturing and service operations",
      "production and operations management", "journal of operations management",
      "transportation science", "journal of the operational research society",
      "european journal of operational research", "iie transactions",
      "international journal of operations and production management",
      "international journal of production research",
      "decision sciences"], ["OPS"]),
    (["interfaces"], ["OPS", "DS"]),
    (["journal of supply chain management", "journal of purchasing and supply management"], ["SCM"]),
    (["strategic management journal", "global strategy journal",
      "strategic entrepreneurship journal"], ["STR"]),
    (["organization studies", "organization science"], ["OB", "MGT"]),
    (["academy of management", "journal of management",
      "british journal of management", "european management review",
      "management and organization review", "management international review"], ["MGT"]),
    (["human resource management"], ["HR"]),
    (["journal of occupational and organizational psychology",
      "european journal of work and organizational",
      "leadership quarterly"], ["HR", "OB"]),
    (["entrepreneurship theory and practice", "journal of business venturing",
      "strategic entrepreneurship", "small business economics",
      "international small business journal", "family business review",
      "journal of small business management"], ["ENT"]),
    (["ieee transactions on engineering management", "journal of product innovation management",
      "r&d management", "ieee transactions on systems, man",
      "ieee transactions on intelligent transportation"], ["TECH"]),
    (["journal of behavioral decision making"], ["PSY"]),
    (["journal of economic psychology"], ["PSY", "ECON"]),
    (["journal of cultural economics"], ["ECON"]),
    (["journal of international business studies", "journal of world business",
      "asia pacific journal of management", "international business review",
      "journal of international management"], ["IB"]),
    (["communication research"], ["COMM"]),
    (["public administration review", "nonprofit and voluntary sector quarterly"], ["PA"]),
    (["business history", "enterprise & society"], ["HIST", "MGT"]),
    (["group decision and negotiation"], ["OPS", "DS"]),
    (["international journal of project management"], ["MGT", "OPS"]),
    (["journal of quality technology"], ["OPS", "TECH"]),
]

def topics_for(journals, conf_name):
    topics = set()
    for j in journals:
        jl = j["name"].lower()
        for keywords, codes in JOURNAL_TOPIC_RULES:
            if any(k in jl for k in keywords):
                topics.update(codes)
                break
    if not topics:
        cn = conf_name.lower()
        if "accounting" in cn:         topics.add("ACC")
        elif "finance" in cn or "financial" in cn: topics.add("FIN")
        elif "marketing" in cn:        topics.add("MKT")
        elif "information system" in cn: topics.add("IS")
        elif "operations" in cn:       topics.add("OPS")
        elif "management" in cn:       topics.add("MGT")
        else:                          topics.add("MGT")
    return sorted(topics)

# ── Known locations ───────────────────────────────────────────────────────────

KNOWN_LOCATIONS = {
    "AOM Annual Meeting":           "Philadelphia, PA, USA",
    "AAA Annual Meeting":           "Las Vegas, NV, USA",
    "AFA Annual Meeting":           "Washington, DC, USA",
    "ICIS":                         "Lisbon, Portugal",
    "ECIS":                         "Milan, Italy",
    "INFORMS Annual Meeting":       "San Francisco, CA, USA",
    "FMA Annual Meeting":           "Tampa, FL, USA",
    "AIB Annual Meeting":           "Manchester, UK",
    "SMS Annual Conference":        "Berlin, Germany",
    "BHC Annual Meeting":           "London, UK",
    "R&D Management Conference":    "Manchester, UK",
    "DSI Annual Meeting":           "San Francisco, CA, USA",
    "AREUEA Annual Conference":     "San Francisco, CA, USA",
    "INTERACT Conference":          "Tallinn, Estonia",
    "ACEI International Conference":"Rotterdam, Netherlands",
    "EURAM Annual Conference":      "Kristiansand, Norway",
    "ISBE Annual Conference":       "Birmingham, UK",
    "TTRA Annual International Conference": "Tampa, FL, USA",
    "IEEE IEEM":                    "Marina Bay Sands, Singapore",
}

RANK_ORDER = {"A*": 0, "A": 1, "B": 2}

# ── Date extraction from HTML ─────────────────────────────────────────────────

# Compiled date patterns, ordered most→least specific
_DATE_PATTERNS = [
    (re.compile(r"\b(202[5-9])[-./](0[1-9]|1[0-2])[-./](0[1-9]|[12]\d|3[01])\b"), "iso"),
    (re.compile(
        r"\b(\d{1,2})\s+(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)[,\s]+(202[5-9])\b", re.I), "dmy_long"),
    (re.compile(
        r"\b(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{1,2})[,\s]+(202[5-9])\b", re.I), "mdy_long"),
    (re.compile(
        r"\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?[,\s]+(202[5-9])\b",
        re.I), "dmy_short"),
    (re.compile(
        r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+(\d{1,2})[,\s]+(202[5-9])\b",
        re.I), "mdy_short"),
]

_SUBMIT_RE  = re.compile(
    r"submiss|submit|deadline|abstract.*due|paper.*due|due.*paper|"
    r"manuscript|call for paper|cfp|extended", re.I)
_CONF_RE    = re.compile(
    r"conference|annual meeting|congress|symposium|workshop|"
    r"takes place|will be held|venue|location|registration", re.I)


def _parse_date_match(m, fmt):
    try:
        if fmt == "iso":
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if fmt == "dmy_long":
            mo = MONTH_NUM[m.group(2).lower()]
            yr = int(m.group(3)); yr = yr if yr > 100 else 2000 + yr
            return date(yr, mo, int(m.group(1)))
        if fmt == "mdy_long":
            mo = MONTH_NUM[m.group(1).lower()]
            yr = int(m.group(3)); yr = yr if yr > 100 else 2000 + yr
            return date(yr, mo, int(m.group(2)))
        if fmt == "dmy_short":
            mo = MONTH_NUM[m.group(2).lower().rstrip(".")]
            yr = int(m.group(3)); yr = yr if yr > 100 else 2000 + yr
            return date(yr, mo, int(m.group(1)))
        if fmt == "mdy_short":
            mo = MONTH_NUM[m.group(1).lower().rstrip(".")]
            yr = int(m.group(3)); yr = yr if yr > 100 else 2000 + yr
            return date(yr, mo, int(m.group(2)))
    except (ValueError, KeyError):
        return None


def _fetch_text(url, timeout=14):
    """Fetch URL and return plain text (or None on failure)."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout,
                         allow_redirects=True, verify=False)
        if r.status_code != 200:
            return None, r.status_code
        soup = BeautifulSoup(r.text, "lxml")
        for tag in soup(["script", "style", "nav", "footer"]):
            tag.decompose()
        return soup.get_text("\n", strip=True), 200
    except Exception as exc:
        return None, str(exc)[:80]


def _extract_best_dates(text):
    """
    Returns (deadline: date|None, conf_start: date|None, conf_end: date|None).
    All dates are >= TODAY.
    """
    if not text:
        return None, None, None

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    found = {}  # date → {is_sub, raw, ctx}

    for i, line in enumerate(lines):
        ctx = " ".join(lines[max(0, i - 3): i + 4])
        is_sub = bool(_SUBMIT_RE.search(ctx))

        for pattern, fmt in _DATE_PATTERNS:
            for m in pattern.finditer(line):
                d = _parse_date_match(m, fmt)
                if d and d >= date(TODAY.year - 1, 6, 1):
                    if d not in found or (is_sub and not found[d]["is_sub"]):
                        found[d] = {
                            "is_sub": is_sub,
                            "raw": m.group(0),
                            "ctx": line[:120],
                        }

    if not found:
        return None, None, None

    all_dates = sorted(found)

    # Best deadline: earliest upcoming sub-context date
    deadline = None
    for d in all_dates:
        if found[d]["is_sub"] and d >= TODAY:
            deadline = d
            break
    # fallback: most recent past sub date
    if not deadline:
        sub_past = [d for d in all_dates if found[d]["is_sub"] and d < TODAY]
        deadline = sub_past[-1] if sub_past else None

    # Best conf dates: first/last date after deadline (or today)
    # Use the first non-sub future date as anchor when deadline is past,
    # to avoid picking a conference date as the conf_start anchor.
    if deadline and deadline >= TODAY:
        anchor = deadline
    else:
        anchor = TODAY

    # Separate future dates into non-sub (conference) and sub categories
    future_nonsub = [d for d in all_dates if d >= anchor and not found[d]["is_sub"]]
    future_all    = [d for d in all_dates if d >= anchor]

    # Prefer non-submission dates for conf_start (avoids picking deadline == conf_start)
    conf_start = future_nonsub[0] if future_nonsub else (future_all[0] if future_all else None)
    conf_end   = None
    if conf_start:
        nearby = [d for d in future_all
                  if (d - conf_start).days <= 10 and d != conf_start]
        if nearby:
            conf_end = nearby[-1]

    # Sanity: a "deadline" that equals or follows the conference start is actually a conf date
    if deadline and conf_start and deadline >= conf_start:
        deadline = None

    return deadline, conf_start, conf_end


# ── Build conferences from xlsx ───────────────────────────────────────────────

def build_from_xlsx():
    """Parse xlsx → list of conference dicts (no scraping yet)."""
    if not HAVE_OPENPYXL:
        sys.exit("openpyxl required for xlsx rebuild. pip install openpyxl")
    if not XLSX_PATH or not XLSX_PATH.exists():
        sys.exit(f"ABS xlsx not found in {REPO_ROOT}")

    wb = openpyxl.load_workbook(XLSX_PATH)
    ws = wb.active

    raw = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        rating, journal, j_url, conf_name, conf_url, deadline, conf_month, notes = row
        if not conf_name or conf_name == "—" or not conf_url or conf_url == "—":
            continue
        if conf_name not in raw:
            raw[conf_name] = {
                "conf_url": conf_url,
                "deadline_typical": deadline or "",
                "conf_month_typical": conf_month or "",
                "notes": notes or "",
                "ratings": [],
                "journals": [],
            }
        if rating and rating not in raw[conf_name]["ratings"]:
            raw[conf_name]["ratings"].append(rating)
        raw[conf_name]["journals"].append({
            "name": journal or "",
            "url":  j_url  or "",
            "rating": rating or "",
        })

    conferences = []
    for name, c in raw.items():
        dl = next_deadline_from_typical(c["deadline_typical"])
        if dl is None:
            dl = date(TODAY.year + 1, 6, 15)

        cs, ce = conf_dates_from_typical(dl, c["conf_month_typical"])
        if cs is None:
            m = (dl.month + 2) % 12 + 1
            y = dl.year + ((dl.month + 2) // 12)
            cs = date(y, m, 15)
            ce = date(y, m, 18)

        slug = re.sub(r"[^a-z0-9]", "", name.lower())[:18]
        ratings = sorted(c["ratings"], key=lambda x: RANK_ORDER.get(x, 9))

        conferences.append({
            "id":                slug,
            "short":             name,
            "name":              name,
            "url":               c["conf_url"],
            "deadline":          dl.isoformat(),
            "conference_start":  cs.isoformat(),
            "conference_end":    ce.isoformat(),
            "location":          KNOWN_LOCATIONS.get(name, "TBA"),
            "topics":            topics_for(c["journals"], name),
            "ranking":           ratings,
            "note":              c["notes"],
            "deadline_typical":  c["deadline_typical"],
            "conf_month_typical": c["conf_month_typical"],
            "journals":          c["journals"],
            "date_confirmed":    False,
            "conf_confirmed":    False,
        })

    conferences.sort(key=lambda x: x["deadline"])
    return conferences


# ── Scrape and update ─────────────────────────────────────────────────────────

def scrape_and_update(conferences, verbose=True):
    """Attempt to fetch each conference URL and update dates in-place."""
    n = len(conferences)
    updated = 0

    for i, c in enumerate(conferences):
        if verbose:
            print(f"  [{i+1:2d}/{n}] {c['short'][:45]:45s}", end=" ", flush=True)

        text, status = _fetch_text(c["url"])
        if text is None:
            if verbose:
                print(f"SKIP ({status})")
            time.sleep(0.3)
            continue

        dl, cs, ce = _extract_best_dates(text)

        changed = False

        # Reject scraped deadlines that are inside the conference window or after a
        # manually-confirmed conference start (likely a camera-ready/conf date, not submission)
        _cs = date.fromisoformat(c["conference_start"]) if c.get("conference_start") else None
        _ce = date.fromisoformat(c["conference_end"]) if c.get("conference_end") else None
        deadline_in_conf = bool(_cs and _ce and _cs <= (dl or date.min) <= _ce)
        deadline_after_manual_conf = bool(c.get("_manual_conf") and _cs and (dl or date.min) > _cs)

        # Only accept a scraped deadline if it's in the future and passes sanity checks
        if dl and dl >= TODAY and not deadline_in_conf and not deadline_after_manual_conf:
            c["deadline"]        = dl.isoformat()
            c["date_confirmed"]  = True
            changed = True

        # Accept conf dates only if not already manually confirmed, future, and
        # at least 7 days after deadline (avoid deadline == conf_start artifacts)
        if not c.get("_manual_conf") and cs and cs >= TODAY:
            dl_date = date.fromisoformat(c["deadline"])
            # If stored deadline == scraped conf_start, previous scrape confused a conf date
            # for a submission deadline — reset to a typical estimate so the gap check works.
            if dl_date == cs:
                dl_date = next_deadline_from_typical(c.get("deadline_typical", "")) or (cs - timedelta(days=200))
                c["deadline"] = dl_date.isoformat()
                c["date_confirmed"] = False
            gap = (cs - dl_date).days
            # Conference must be 7 days to 18 months after deadline
            if 7 <= gap <= 540:
                c["conference_start"] = cs.isoformat()
                c["conf_confirmed"]   = True
                changed = True
                if ce and ce > cs and (ce - cs).days <= 14:
                    c["conference_end"] = ce.isoformat()
                else:
                    c["conference_end"] = cs.isoformat()

        if verbose:
            if changed:
                marks = ("✓" if c["date_confirmed"] else "") + ("📅" if c["conf_confirmed"] else "")
                print(f"{marks}  dl={c['deadline']}  cs={c['conference_start']}")
            else:
                print("no new dates")

        time.sleep(0.5)

    return updated


# ── Write data.js ─────────────────────────────────────────────────────────────

TOPICS_JS = """// Topic colour palette
const TOPICS = {
  ACC:  { label: "Accounting",          color: "#6366f1" },
  COMM: { label: "Communication",       color: "#8b5cf6" },
  DS:   { label: "Data Science",        color: "#0ea5e9" },
  ECON: { label: "Economics",           color: "#f59e0b" },
  ENT:  { label: "Entrepreneurship",    color: "#10b981" },
  FIN:  { label: "Finance",             color: "#3b82f6" },
  HCI:  { label: "Human-Computer Int.", color: "#06b6d4" },
  HIST: { label: "Business History",    color: "#a78bfa" },
  HR:   { label: "Human Resources",     color: "#f97316" },
  IB:   { label: "Int'l Business",      color: "#14b8a6" },
  IS:   { label: "Information Systems", color: "#2563eb" },
  MGT:  { label: "Management",          color: "#7c3aed" },
  MKT:  { label: "Marketing",           color: "#ec4899" },
  OB:   { label: "Org. Behaviour",      color: "#f43f5e" },
  OPS:  { label: "Operations & OR",     color: "#84cc16" },
  PA:   { label: "Public Admin.",       color: "#6b7280" },
  PSY:  { label: "Psychology",          color: "#a855f7" },
  RE:   { label: "Real Estate",         color: "#78716c" },
  SCM:  { label: "Supply Chain",        color: "#22c55e" },
  SME:  { label: "Small Business",      color: "#fdba74" },
  STR:  { label: "Strategy",            color: "#1d4ed8" },
  TECH: { label: "Technology Mgmt.",    color: "#0284c7" },
};"""


def write_data_js(conferences, scrape_date):
    lines = [
        "// Auto-generated — do not edit by hand.",
        f'// Source: ABS Journal List xlsx + live website scraping.',
        f'// Last scraped: {scrape_date}',
        f'const SCRAPE_DATE = "{scrape_date}";',
        "",
        "const CONFERENCES = [",
    ]
    for c in conferences:
        row = {k: v for k, v in c.items() if not k.startswith("_")}
        lines.append("  " + json.dumps(row, ensure_ascii=False) + ",")
    lines += ["];", "", TOPICS_JS, ""]
    DATA_JS.write_text("\n".join(lines), encoding="utf-8")
    print(f"\n✓ Wrote {len(conferences)} conferences to {DATA_JS}")


# ── Manually verified overrides (applied after xlsx rebuild) ─────────────────
# These were confirmed by visiting each conference website.
# Format: slug → {field: value, ...}
# Re-run `python scrape.py` to discover new ones automatically.
MANUAL_OVERRIDES = {
    # Deadline was Dec 1 2025 (past); conference is Jun 2-5 2026.
    # Next cycle deadline ~Oct-Nov 2026; showing 2026 conference while it's still upcoming.
    "emacannualconferen": {
        "conference_start": "2026-06-02",
        "conference_end":   "2026-06-05",
        "conf_confirmed":   True,
        "deadline":         "2026-11-15",
        "date_confirmed":   False,
    },
}


def apply_overrides(conferences):
    """Apply MANUAL_OVERRIDES to the conference list in-place."""
    idx = {c["id"]: c for c in conferences}
    for slug, fields in MANUAL_OVERRIDES.items():
        if slug in idx:
            idx[slug].update(fields)
            if fields.get("conf_confirmed"):
                idx[slug]["_manual_conf"] = True


# ── Load existing data.js (for --scrape-only) ────────────────────────────────

def load_existing_data_js():
    if not DATA_JS.exists():
        sys.exit(f"{DATA_JS} not found. Run without --scrape-only first.")
    text = DATA_JS.read_text(encoding="utf-8")
    m = re.search(r"const CONFERENCES = (\[.*?\]);", text, re.DOTALL)
    if not m:
        sys.exit("Could not parse CONFERENCES from data.js")
    raw = re.sub(r",\s*([\]\}])", r"\1", m.group(1))
    return json.loads(raw)


def last_scrape_date():
    """Return the SCRAPE_DATE from data.js, or None."""
    if not DATA_JS.exists():
        return None
    m = re.search(r'SCRAPE_DATE\s*=\s*"(\d{4}-\d{2}-\d{2})"', DATA_JS.read_text())
    if m:
        try:
            return date.fromisoformat(m.group(1))
        except ValueError:
            pass
    return None


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--build-only",  action="store_true",
                        help="Rebuild from xlsx, skip web scraping")
    parser.add_argument("--scrape-only", action="store_true",
                        help="Scrape existing data.js, skip xlsx rebuild")
    parser.add_argument("--force",       action="store_true",
                        help="Ignore 30-day cooldown")
    parser.add_argument("--quiet",       action="store_true",
                        help="Less verbose output")
    args = parser.parse_args()

    verbose = not args.quiet

    # ── Cooldown check ────────────────────────────────────────────────────────
    if not args.force and not args.build_only:
        last = last_scrape_date()
        if last and (TODAY - last) < SCRAPE_GAP:
            days_left = (SCRAPE_GAP - (TODAY - last)).days
            print(f"Skipping scrape — last run was {last} ({(TODAY-last).days}d ago).")
            print(f"Run with --force to override, or wait {days_left} more days.")
            if not args.scrape_only:
                # Still rebuild from xlsx so local date estimates stay fresh
                print("Rebuilding from xlsx (no scraping)…")
                conferences = build_from_xlsx()
                write_data_js(conferences, last.isoformat())
            return

    # ── Step 1: build base data ───────────────────────────────────────────────
    if args.scrape_only:
        print("Loading existing data.js…")
        conferences = load_existing_data_js()
    else:
        print(f"Building from xlsx: {XLSX_PATH}")
        conferences = build_from_xlsx()
        apply_overrides(conferences)
        print(f"  → {len(conferences)} conferences ({sum(1 for c in conferences if c['conf_confirmed'])} conf dates pre-confirmed)")

    # ── Step 2: scrape ────────────────────────────────────────────────────────
    if not args.build_only:
        print(f"\nScraping {len(conferences)} conference websites…")
        scrape_and_update(conferences, verbose=verbose)

    # ── Step 3: re-sort and write ─────────────────────────────────────────────
    conferences.sort(key=lambda x: x["deadline"])
    scrape_date = TODAY.isoformat() if not args.build_only else (last_scrape_date() or TODAY).isoformat()
    write_data_js(conferences, scrape_date)

    # ── Summary ───────────────────────────────────────────────────────────────
    dl_conf  = sum(1 for c in conferences if c["date_confirmed"])
    cs_conf  = sum(1 for c in conferences if c["conf_confirmed"])
    print(f"Confirmed deadlines:    {dl_conf}/{len(conferences)}")
    print(f"Confirmed conf dates:   {cs_conf}/{len(conferences)}")
    print(f"Estimated (both):       {len(conferences) - dl_conf - cs_conf + min(dl_conf, cs_conf)}")


if __name__ == "__main__":
    main()

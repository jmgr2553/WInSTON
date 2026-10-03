#!/usr/bin/env python3
"""
hunter.py - finds postdoc and early-career research posts that suit you.

All your settings - the websites searched, the language model, your files and
your PhD timeline - live in config.toml. Copy config.example.toml to start.

STEP 1 (collect): gather ads from several sources and store them in a local
cache file so nothing is downloaded twice. Two kinds of source are supported:
  - SEARCH SITES (e.g. jobs.ac.uk, EURAXESS): the script runs keyword
    searches and reads the result pages.
  - LISTING PAGES (e.g. the Cambridge colleges vacancies page): a single
    page listing current vacancies; the script keeps every link whose text
    looks like a fellowship or postdoc.

STEP 2 (rank): turn your profile, your CV and every ad into "embeddings"
(lists of numbers that encode meaning) and measure how close each ad is to
you. This is fast but rough - a first sieve.

STEP 3 (judge): a language model running on your own computer (Ollama, LM
Studio, llama.cpp...) reads the best ads from step 2 in full. It rates the
field fit and extracts the facts that decide eligibility - the ad's PhD rule,
quoted, with its dates and the start date. The script then compares those
dates with your PhD timeline to decide eligibility. Answers are stored in the
cache, so each ad is only read once.

Everything ends up in a spreadsheet (postdoc_matches.xlsx by default).

Usage:
    python hunter.py                  # all three steps
    python hunter.py --no-llm         # skip step 3
    python hunter.py --rejudge        # step 3 re-reads every ad (e.g. after
                                      # editing your candidate file)
    python hunter.py --model NAME     # use another model for this run
    python hunter.py --config FILE    # use another settings file
    python hunter.py --check-sources  # test each website once, then stop
"""

import argparse
import json
import re
from collections import Counter
import sys
import time
import tomllib
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote, quote_plus, urljoin, urlparse

import pandas as pd
import requests
import trafilatura
from bs4 import BeautifulSoup
from curl_cffi import requests as browser_requests
from dateutil import parser as dateparser

# =============================================================================
# SETTINGS - read from config.toml (see config.example.toml for every option)
# =============================================================================
# The values below are filled in by load_config() when the script starts.
# Nothing here needs editing: change config.toml instead.

DEFAULT_CONFIG = "config.toml"

# Settings that may be left out of config.toml, and their default values.
DEFAULTS = {
    "files": {
        "not_interested": "",
        "cache": "jobs_cache.json",
        "output": "postdoc_matches.xlsx",
    },
    "ranking": {
        "embedding_model": "intfloat/multilingual-e5-base",
        "profile_weight": 0.6,
        "negative_weight": 0.5,
    },
    "llm": {
        "provider": "ollama",
        "model": "qwen3.5:4b",
        "context": 8192,
        "top_n": 40,
        "text_chars": 6000,
        "api_key": "",
    },
    "fit_scale": {
        "own_field": "",
        "close_fields": "",
        "neighbouring_fields": "",
    },
    "network": {
        "seconds_between_downloads": 3,
        "skip_cert_check": [],
    },
}
DEFAULT_BASE_URLS = {"ollama": "http://localhost:11434", "openai": "http://localhost:1234/v1"}


def load_config(path, model=None):
    """Read config.toml and fill in the settings used by the rest of the
    script. 'model' (from --model) overrides the one in the file."""
    global CV_PATH, PROFILE_PATH, CANDIDATE_PATH, NOT_INTERESTED_PATH, CACHE_FILE, OUTPUT_FILE
    global THESIS_SUBMISSION, PHD_AWARD, TITLE_PATTERN, EXCLUDE_TITLE_PATTERN
    global SEARCH_SITES, LISTING_PAGES, LISTING_PATTERNS, ALWAYS_READ_SOURCES, SCHEME_NOTES
    global MODEL_NAME, PROFILE_WEIGHT, NEGATIVE_WEIGHT
    global LLM_PROVIDER, LLM_BASE_URL, LLM_MODEL, LLM_API_KEY, LLM_CONTEXT, LLM_TOP_N, LLM_TEXT_CHARS
    global FIT_SCALE, SECONDS_BETWEEN_DOWNLOADS, SKIP_CERT_CHECK

    path = Path(path)
    if not path.exists():
        sys.exit(f"Settings file {path} not found. To start, copy the example:\n"
                 f"    cp config.example.toml {path}\n"
                 f"then edit it (see README.md).")
    try:
        with path.open("rb") as f:
            cfg = tomllib.load(f)
    except tomllib.TOMLDecodeError as err:
        sys.exit(f"{path} has a mistake in it: {err}")
    for section, values in DEFAULTS.items():
        cfg[section] = {**values, **cfg.get(section, {})}

    def need(section, key):
        try:
            return cfg[section][key]
        except KeyError:
            sys.exit(f"{path}: '{key}' is missing from the [{section}] section.")

    # Relative file names are read from the folder that holds config.toml.
    def in_config_dir(name):
        return path.parent / name if name else None

    files = cfg["files"]
    CV_PATH = in_config_dir(need("files", "cv"))
    PROFILE_PATH = in_config_dir(need("files", "profile"))
    CANDIDATE_PATH = in_config_dir(need("files", "candidate"))
    NOT_INTERESTED_PATH = in_config_dir(files["not_interested"])
    CACHE_FILE = in_config_dir(files["cache"])
    OUTPUT_FILE = in_config_dir(files["output"])

    THESIS_SUBMISSION = need("timeline", "thesis_submission")
    PHD_AWARD = need("timeline", "phd_award")
    for name, value in [("thesis_submission", THESIS_SUBMISSION), ("phd_award", PHD_AWARD)]:
        if not isinstance(value, date):
            sys.exit(f"{path}: [timeline] {name} must be a date without quotes, e.g. 2027-03-31.")

    TITLE_PATTERN = re.compile("|".join(need("titles", "include")), re.IGNORECASE)
    EXCLUDE_TITLE_PATTERN = re.compile("|".join(cfg["titles"].get("exclude", [])) or r"(?!)",
                                       re.IGNORECASE)

    SEARCH_SITES = cfg.get("search_sites", [])
    for site in SEARCH_SITES:
        missing = {"name", "url", "link_pattern"} - set(site)
        if missing:
            sys.exit(f"{path}: a [[search_sites]] entry lacks {', '.join(sorted(missing))}.")
        site.setdefault("first_page", 1)
        site.setdefault("page_step", 1)
        site.setdefault("per_page", 25)
        site.setdefault("max_pages", 3)
        site.setdefault("title_filter", True)
        site.setdefault("queries", [])
    LISTING_PAGES = cfg.get("listing_pages", [])
    for page in LISTING_PAGES:
        if not {"name", "url"} <= set(page):
            sys.exit(f"{path}: every [[listing_pages]] entry needs a name and a url.")
    SEARCH_SITES = [s for s in SEARCH_SITES if s.get("enabled", True)]
    LISTING_PAGES = [p for p in LISTING_PAGES if p.get("enabled", True)]
    if not SEARCH_SITES and not LISTING_PAGES:
        sys.exit(f"{path}: no [[search_sites]] or [[listing_pages]] are switched on.")
    LISTING_PATTERNS = {p["name"]: p["link_pattern"] for p in LISTING_PAGES if p.get("link_pattern")}
    ALWAYS_READ_SOURCES = {p["name"] for p in LISTING_PAGES if p.get("always_read")}
    SCHEME_NOTES = {n["pattern"]: n["note"] for n in cfg.get("scheme_notes", [])}

    ranking = cfg["ranking"]
    MODEL_NAME = ranking["embedding_model"]
    PROFILE_WEIGHT = ranking["profile_weight"]
    NEGATIVE_WEIGHT = ranking["negative_weight"]

    llm = cfg["llm"]
    LLM_PROVIDER = llm["provider"].lower()
    if LLM_PROVIDER not in DEFAULT_BASE_URLS:
        sys.exit(f"{path}: [llm] provider must be \"ollama\" or \"openai\", not \"{llm['provider']}\".")
    LLM_BASE_URL = llm.get("base_url", DEFAULT_BASE_URLS[LLM_PROVIDER]).rstrip("/")
    LLM_MODEL = model or llm["model"]
    LLM_API_KEY = llm["api_key"]
    LLM_CONTEXT = llm["context"]
    LLM_TOP_N = llm["top_n"]
    LLM_TEXT_CHARS = llm["text_chars"]
    FIT_SCALE = cfg["fit_scale"]

    SECONDS_BETWEEN_DOWNLOADS = cfg["network"]["seconds_between_downloads"]
    SKIP_CERT_CHECK = set(cfg["network"]["skip_cert_check"])

# =============================================================================
# Helpers
# =============================================================================


def fetch(url, retries=3):
    """Download a URL the way Chrome would, and return the response.
    If the site answers 429 ("Too Many Requests"), wait and try again."""
    verify = urlparse(url).hostname not in SKIP_CERT_CHECK
    for attempt in range(retries + 1):
        response = browser_requests.get(url, impersonate="chrome", timeout=30, verify=verify)
        if response.status_code == 429 and attempt < retries:
            retry_after = str(response.headers.get("retry-after", ""))
            wait = int(retry_after) if retry_after.isdigit() else 30 * (attempt + 1)
            print(f"      (the site asks us to slow down - waiting {wait} seconds)")
            time.sleep(wait)
            continue
        response.raise_for_status()
        return response


def clean_text(html):
    """Strip the HTML tags and menus from a page and return plain text."""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "nav", "header", "footer", "form"]):
        tag.decompose()
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True))


# A date such as "15th October 2026", "9 Oct 2026", "16 October" (no year),
# or the American order "October 27, 2026".
_MONTH = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?"
# Spanish: "10 de julio de 2026"
SPANISH_MONTHS = {
    "enero": "January", "febrero": "February", "marzo": "March", "abril": "April",
    "mayo": "May", "junio": "June", "julio": "July", "agosto": "August",
    "septiembre": "September", "setiembre": "September", "octubre": "October",
    "noviembre": "November", "diciembre": "December",
}
_MES = "(?:" + "|".join(SPANISH_MONTHS) + ")"
# German: "23. Mai 2026", and the numeric form "15.07.2026"
GERMAN_MONTHS = {
    "januar": "January", "februar": "February", "märz": "March", "april": "April",
    "mai": "May", "juni": "June", "juli": "July", "august": "August",
    "september": "September", "oktober": "October", "november": "November",
    "dezember": "December",
}
_MONAT = "(?:" + "|".join(GERMAN_MONTHS) + ")"
DATE_RE = (r"(?:\d{1,2}(?:st|nd|rd|th)?\s+" + _MONTH + r"(?:,?\s+\d{4})?"
           r"|" + _MONTH + r"\s+\d{1,2}(?!\d)(?:st|nd|rd|th)?(?:,?\s+\d{4})?"
           r"|\d{1,2}\s+de\s+" + _MES + r"(?:\s+de\s+\d{4})?"
           r"|\d{1,2}\.\s*" + _MONAT + r"(?:\s+\d{4})?"
           r"|\d{1,2}\.\d{1,2}\.\d{4})")
LABELLED_DATE = re.compile(
    r"(?:Clos(?:es|ing date)|Application deadline|Deadline(?: for applications)?|"
    r"(?:are |is )?due(?: on| by)?|open until|no later than|"
    r"plazo[^.:]{0,60}?(?:hasta|finaliza|termina)(?: el)?|fecha l[ií]mite[^.:]{0,40}?|"
    r"hasta el|Bewerbungsschluss|Bewerbungsfrist|Einsendeschluss)\s*:?\s*(" + DATE_RE + ")",
    re.IGNORECASE,
)


def parse_date(text):
    """Turn a date string into a date. If it has no year (as on the Oxford
    Gazette), assume the next time that day comes round."""
    today = date.today()
    numeric = re.fullmatch(r"\s*(\d{1,2})\.(\d{1,2})\.(\d{4})\s*", text)
    if numeric:                                       # German "15.07.2026"
        day, month, year = map(int, numeric.groups())
        try:
            return date(year, month, day)
        except ValueError:
            return None
    lowered = text.lower()
    for es, en in SPANISH_MONTHS.items():            # "10 de julio de 2026"
        lowered = lowered.replace(es, en)             # -> "10 de July de 2026"
    for de, en in GERMAN_MONTHS.items():             # "23. Mai 2026"
        lowered = re.sub(rf"\b{de}\b", en, lowered)  # -> "23. May 2026"
    lowered = re.sub(r"(\d)\.\s", r"\1 ", lowered)    # -> "23 May 2026"
    text = re.sub(r"\bde\b", " ", lowered)            # -> "10 July 2026"
    try:
        d = dateparser.parse(text, default=datetime(today.year, 1, 1)).date()
    except (ValueError, OverflowError):
        return None
    if not re.search(r"\d{4}", text) and d < today:
        d = d.replace(year=d.year + 1)
    return d


def find_deadline(text):
    """Look for a labelled deadline ('Closes: ...', 'Application Deadline ...')."""
    m = LABELLED_DATE.search(text)
    return parse_date(m.group(1)) if m else None


def last_date_in(text):
    """The last date mentioned in a short listing entry - on the Gazette and
    the Cambridge colleges page, the closing date comes last."""
    dates = re.findall(DATE_RE, text, re.IGNORECASE)
    return parse_date(dates[-1]) if dates else None


def find_institution(html):
    """Best guess at the employer's name. May come back empty - that's fine."""
    soup = BeautifulSoup(html, "html.parser")
    tag = soup.find(attrs={"class": re.compile(r"employer|institution", re.I)})
    if tag is None:
        tag = soup.find("a", href=re.compile(r"/employer", re.I))
    return tag.get_text(" ", strip=True)[:120] if tag else ""


def links_matching(html, page_url, pattern):
    """{ad address: {"title", "context", "institution"}} for every link whose
    address fits the site's ad pattern. The same ad is often linked twice
    (title, 'more' button), so we keep one entry per address, with the
    longest title. 'context' is the result card around the link (summary,
    employer, deadline), used when the ad page itself isn't downloaded."""
    soup = BeautifulSoup(html, "html.parser")
    links = {}
    for a in soup.find_all("a", href=True):
        m = re.match(pattern, urljoin(page_url, a["href"]))
        if not m:
            continue
        link, title = m.group(0), a.get_text(" ", strip=True)
        if len(title) <= len(links.get(link, {}).get("title", "")):
            continue
        card = a.find_parent("li") or a.find_parent("article")
        context = card.get_text(" ", strip=True) if card else ""
        if len(context) > 4000:     # the "card" was the whole results list
            context, card = "", None
        employer = card.find("a", href=re.compile(r"/organisations/profile/")) if card else None
        links[link] = {
            "title": title,
            "context": context,
            "institution": employer.get_text(" ", strip=True) if employer else "",
        }
    return links


def listing_links(html, page_url, link_pattern=None):
    """(address, title, context) for every link on a listing page whose text
    looks like a fellowship or postdoc. 'context' is the table row or list
    item around the link, which often holds the college and closing date.
    If the page's ads have a recognisable address (link_pattern), only links
    with that address count - this skips menu items such as "Scholarships /
    Fellowships / Grants / Prizes". If such a link's text is too short (e.g.
    "Read more"), the title is rebuilt from the address:
    .../call-for-applications-eume-fellowships-2027-2028 becomes
    "Call For Applications Eume Fellowships 2027 2028"."""
    soup = BeautifulSoup(html, "html.parser")
    items, seen = [], set()
    for a in soup.find_all("a", href=True):
        link = urljoin(page_url, a["href"])
        title = a.get_text(" ", strip=True)
        if not link.startswith("http") or link in seen:
            continue  # skip e-mail links and repeated links to the same ad
        if link_pattern and not re.search(link_pattern, link):
            continue  # not an ad address on this site
        if len(title.split()) < 3:
            if link_pattern:
                title = link.rstrip("/").rsplit("/", 1)[-1].replace("-", " ").title()
            else:
                continue  # short menu items like "Fellows"
        if not TITLE_PATTERN.search(title) or EXCLUDE_TITLE_PATTERN.search(title):
            continue
        seen.add(link)
        box = a.find_parent(["tr", "li"])
        context = box.get_text(" ", strip=True) if box else title
        if len(context) > 600:        # the "box" was a whole list, not one entry
            context = title
        items.append((link, title, context))
    return items


def download_page(url):
    """Download a web page or a PDF. Returns {"html": ...} for pages,
    {"pdf_text": ...} for PDFs (their text, extracted), or None."""
    try:
        response = fetch(url)
    except Exception as err:
        print(f"      (could not download {url[:70]}: {err}; keeping the listing text)")
        return None
    finally:
        time.sleep(SECONDS_BETWEEN_DOWNLOADS)
    kind = response.headers.get("content-type", "").lower()
    if "html" in kind:
        return {"html": response.text}
    if "pdf" in kind or url.lower().endswith(".pdf"):
        try:
            import io
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(response.content))
            text = " ".join(p.extract_text() or "" for p in reader.pages[:15])
            return {"pdf_text": re.sub(r"\s+", " ", text)[:15000]}
        except Exception:
            return None
    return None


PARTICULARS_WORDS = re.compile(
    r"particulars|further (?:details|information)|job description|information for "
    r"applicants|guidance|role profile|ausschreibung|bases|convocatoria", re.IGNORECASE)


def particulars_links(html, page_url):
    """PDF links on an ad page that look like further particulars."""
    soup = BeautifulSoup(html, "html.parser")
    found = []
    for a in soup.find_all("a", href=True):
        url = urljoin(page_url, a["href"])
        if ".pdf" in url.lower() and PARTICULARS_WORDS.search(a.get_text(" ", strip=True) + " " + url):
            if url not in found:
                found.append(url)
    return found


def build_record(link, title, source, memo, today, context="", institution="", download=True):
    """Download the ad (once per run, even if several entries share a page)
    and turn it into a cache entry. If the download fails - or the link is a
    PDF - keep the listing text, so the ad still reaches the spreadsheet."""
    if not download:
        memo.setdefault(link, None)
    if link not in memo:
        memo[link] = download_page(link)
    page = memo[link]            # None, {"html": ...} or {"pdf_text": ...}

    if page and "html" in page:
        html = page["html"]
        full_text = clean_text(html)
        body = trafilatura.extract(html) or full_text   # just the ad, for ranking
        # College fellowships often keep the rules in a "further particulars"
        # PDF linked from the page; read up to two of them.
        for pdf_link in particulars_links(html, link)[:2]:
            if pdf_link not in memo:
                memo[pdf_link] = download_page(pdf_link)
            extra = memo[pdf_link]
            if extra and "pdf_text" in extra:
                body += "\n[From the further particulars]\n" + extra["pdf_text"]
                full_text += " " + extra["pdf_text"]
        if not institution:
            institution = find_institution(html)
    elif page and "pdf_text" in page:
        full_text = body = page["pdf_text"]
    else:
        full_text = body = context or title

    # A date in the listing entry beats one found on the page, because a
    # college page shared by several posts contains several dates.
    deadline = None
    if context:
        deadline = find_deadline(context) or last_date_in(context)
    deadline = deadline or find_deadline(full_text)

    return {
        "title": title,
        "link": link,
        "institution": institution,
        "source": source,
        "deadline": deadline.isoformat() if deadline else None,
        "text": ((context + "\n") if context else "") + body[:20000],
        "first_seen": today,
        # True when only the search-result card was stored (EURAXESS): the
        # full ad is then downloaded in step 3, if the model is to read it.
        "card_only": not download,
    }


def read_cv(path):
    path = Path(path)
    if not path.exists():
        sys.exit(f"CV not found at {path.resolve()} - check [files] cv in config.toml.")
    if path.suffix.lower() == ".pdf":
        from pypdf import PdfReader
        return " ".join(page.extract_text() or "" for page in PdfReader(path).pages)
    return path.read_text(encoding="utf-8")


def read_profile(path):
    path = Path(path)
    if not path.exists():
        sys.exit(f"Profile not found at {path.resolve()} - check [files] profile in config.toml.")
    text = path.read_text(encoding="utf-8").strip()
    if not text:
        sys.exit(f"{path} is empty - write a short description of the posts you want.")
    if len(text.split()) > 300:
        print(f"Note: profile has {len(text.split())} words; the model only reads the first ~350.")
    return text


def strip_boilerplate(texts, min_count=3):
    """Remove sentences that appear in several different ads, such as an
    employer's standard "About our campus..." paragraph or a website's
    "Apply now" text. They say nothing about the post itself but can make
    unrelated ads look alike - e.g. every Hebrew University ad mentions
    "Hebrew", which resembles your profile whatever the ad is about."""
    split = [re.split(r"(?<=[.!?])\s+", t) for t in texts]
    counts = Counter(s.strip() for sentences in split for s in set(sentences)
                     if len(s.split()) >= 6)
    return [" ".join(s for s in sentences if counts.get(s.strip(), 0) < min_count)
            for sentences in split]


def chunk_words(text, size=200):
    """Split long text into pieces of ~200 words. The model only reads about
    350 words at a time, so a whole CV has to be fed to it in pieces."""
    words = text.split()
    return [" ".join(words[i:i + size]) for i in range(0, len(words), size)]


# =============================================================================
# STEP 1 - collect
# =============================================================================


def search_templates(site):
    """(label, address with {page} left open) for each search on a site:
    either keyword searches built from "queries", or searches pasted from
    the browser in "search_urls" (we remove any page number they carry and
    add our own)."""
    templates = []
    for query in site.get("queries", []):
        encoded = quote(query, safe="") if site.get("encode") == "percent" else quote_plus(query)
        templates.append((f"{site['name']}: {query}", site["url"].replace("{query}", encoded)))
    for n, url in enumerate(site.get("search_urls", []), start=1):
        url = re.sub(r"([?&])page=\d+&?", r"\1", url).rstrip("?&")
        url += ("&" if "?" in url else "?") + "page={page}"
        templates.append((f"{site['name']}: search {n}", url))
    return templates


def collect_search_site(site, cache, memo, today):
    new = 0
    first_pages = {}   # first-page results of each search, to spot ignored keywords
    for label, template in search_templates(site):
        print(f"Searching {label}")
        previous_page = set()

        for i in range(site["max_pages"]):
            page_value = site["first_page"] + i * site["page_step"]
            url = template.replace("{page}", str(page_value))
            try:
                html = fetch(url).text
                links = links_matching(html, url, site["link_pattern"])
            except Exception as err:
                print(f"  could not load results page {i + 1}: {err}")
                break
            if i == 0:
                count = re.search(site.get("results_count_re", r"(?!)"), html)
                total = int(count.group(1).replace(",", "")) if count else None
                same_as = next((other for other, seen in first_pages.items()
                                if seen == set(links)), None)
                if (total and total > site.get("max_results", 10**9)) or (links and same_as):
                    reason = (f"{total} results" if total and total > site.get("max_results", 10**9)
                              else f"same results as '{same_as}'")
                    print(f"  the site seems to have ignored the keyword ({reason}) - skipped. "
                          f"Try again later.")
                    break
                first_pages[label] = set(links)
            if not links or set(links) == previous_page:
                break   # no results, or the site ignored the "next page" request
            previous_page = set(links)

            kept = [(l, info) for l, info in links.items()
                    if (not site["title_filter"] or TITLE_PATTERN.search(info["title"]))
                    and not EXCLUDE_TITLE_PATTERN.search(info["title"])]
            first_title = next(iter(links.values()))["title"][:70]
            print(f"  page {i + 1}: {len(links)} results, {len(kept)} kept "
                  f"(first result: {first_title})")

            download = site.get("download_ads", True)
            for link, info in kept:
                if link in cache:
                    continue
                cache[link] = build_record(
                    link, info["title"], label, memo, today,
                    context="" if download else info["context"],
                    institution=info["institution"], download=download)
                new += 1
                print(f"    + {info['title'][:80]}")

            if len(links) < site["per_page"]:
                break   # that was the last page
            time.sleep(site.get("delay", SECONDS_BETWEEN_DOWNLOADS))
    return new


def collect_listing_page(label, url, cache, memo, today, link_pattern=None):
    print(f"Reading {label}")
    try:
        html = fetch(url).text
    except Exception as err:
        print(f"  could not load the page: {err}")
        return 0
    items = listing_links(html, url, link_pattern)
    print(f"  {len(items)} links look like fellowships or postdocs")
    if not items:
        # Helps tell "no vacancies today" apart from "the list is drawn by
        # JavaScript after the page loads, so the script can't see it".
        n_links = len(BeautifulSoup(html, "html.parser").find_all("a", href=True))
        print(f"  (the page has {n_links} links in total)")

    new = 0
    for link, title, context in items:
        # Several posts can share one address (e.g. a college's vacancies
        # page), so the cache key combines the address and the title.
        key = f"{link} | {title}"
        if key in cache:
            continue
        # Gazette entries read "College; Post; salary; deadline".
        institution = title.split(";")[0].strip() if ";" in title else ""
        cache[key] = build_record(link, title, label, memo, today,
                                  context=context, institution=institution)
        new += 1
        print(f"    + {title[:80]}")
    return new


def collect():
    cache = json.loads(CACHE_FILE.read_text(encoding="utf-8")) if CACHE_FILE.exists() else {}
    today = date.today().isoformat()
    memo = {}   # pages downloaded during this run
    new = 0

    for site in SEARCH_SITES:
        new += collect_search_site(site, cache, memo, today)
    for page in LISTING_PAGES:
        new += collect_listing_page(page["name"], page["url"], cache, memo, today,
                                    page.get("link_pattern"))

    CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nStep 1 done: {new} new ads, {len(cache)} in the cache.\n")
    return cache


# =============================================================================
# STEP 2 - rank with embeddings
# =============================================================================


def rank(cache):
    """Return the current ads, each with an embedding score ("fit") and its
    cleaned text ("body"), best first."""
    today = date.today()

    # Drop ads whose deadline has passed (keep those with unknown deadline).
    ads = []
    for key, ad in cache.items():
        deadline = datetime.fromisoformat(ad["deadline"]).date() if ad["deadline"] else None
        if deadline and deadline < today:
            continue
        if EXCLUDE_TITLE_PATTERN.search(ad["title"]):
            continue   # e.g. PhD positions collected before the filter existed
        pattern = LISTING_PATTERNS.get(ad["source"])
        if pattern and not re.search(pattern, ad.get("link", key)):
            continue   # e.g. a menu link collected before the patterns were strict
        ads.append({**ad, "key": key, "link": ad.get("link", key), "deadline_date": deadline})
    if not ads:
        sys.exit("No current ads to rank. Check the sources in config.toml "
                 "(python hunter.py --check-sources).")

    # Read the profile and CV first, so a missing file is reported straight away.
    profile = read_profile(PROFILE_PATH)
    cv_chunks = chunk_words(read_cv(CV_PATH))

    from sentence_transformers import SentenceTransformer
    print(f"Loading embedding model {MODEL_NAME}...")
    model = SentenceTransformer(MODEL_NAME)

    # The e5 models expect "query: " before the thing you are searching FOR
    # and "passage: " before the things you are searching THROUGH.
    queries = ["query: " + profile] + ["query: " + c for c in cv_chunks]
    bodies = strip_boilerplate([ad["text"][:8000] for ad in ads])
    passages = ["passage: " + ad["title"] + ". " + body[:3000]
                for ad, body in zip(ads, bodies)]

    q_vecs = model.encode(queries, normalize_embeddings=True)
    p_vecs = model.encode(passages, normalize_embeddings=True, show_progress_bar=True)

    # Similarity = how closely two vectors point the same way (1 = identical).
    sims = p_vecs @ q_vecs.T                 # one row per ad, one column per query
    profile_sim = sims[:, 0]                 # closeness to the profile
    cv_sim = sims[:, 1:].max(axis=1)         # closeness to the best-matching CV piece
    scores = PROFILE_WEIGHT * profile_sim + (1 - PROFILE_WEIGHT) * cv_sim

    # Subtract closeness to the "not interested" description, so that e.g. a
    # biology postdoc at the Hebrew University - which resembles your profile
    # only because it mentions "Hebrew" - drops down the list.
    if NOT_INTERESTED_PATH and NOT_INTERESTED_PATH.exists():
        negative = NOT_INTERESTED_PATH.read_text(encoding="utf-8").strip()
        if negative:
            n_vec = model.encode(["query: " + negative], normalize_embeddings=True)
            scores = scores - NEGATIVE_WEIGHT * (p_vecs @ n_vec.T)[:, 0]

    for ad, score, body in zip(ads, scores, bodies):
        ad["fit"] = float(score)
        ad["body"] = body
    ads.sort(key=lambda ad: ad["fit"], reverse=True)
    ads = merge_duplicates(ads)
    print(f"\nStep 2 done: {len(ads)} ads ranked.\n")
    return ads


# Words too common in job titles to show that two ads are the same post.
COMMON_WORDS = set("""
research researcher fellow fellows fellowship fellowships junior senior early career
postdoctoral postdoc post doctoral associate assistant stipendiary non year years
college university cambridge oxford london call applications application position
positions arts humanities sciences social studies leverhulme programme program
january february march april june july august september october november december
""".split())


def distinctive_words(text):
    words = re.findall(r"[a-zà-ÿ]{4,}", (text or "").lower())
    return {w for w in words if w not in COMMON_WORDS}


def same_page(a, b):
    """True if two links point to the same page (ignoring http/https, "www.",
    a trailing slash, and tracking tags like ?utm_source=...)."""
    def norm(u):
        u = re.sub(r"^https?://(www\.)?", "", u or "").split("?")[0].split("#")[0]
        return u.rstrip("/").lower()
    return bool(a) and norm(a) == norm(b)


def merge_duplicates(ads):
    """The same post is often advertised in several places (the Magdalene JRF
    on jobs.ac.uk, the Oxford Gazette and the Cambridge colleges page). Two
    ads count as one when they have the same deadline, their titles share a
    distinctive word (e.g. "Magdalene", "Finley"), and their institutions,
    when both are known, don't contradict each other. The best-ranked copy
    is kept; the others' links go into its "Also at" column."""
    kept = []
    for ad in ads:                      # ads are already sorted best first
        twin = None
        for other in kept:
            if not ad["deadline_date"] or ad["deadline_date"] != other["deadline_date"]:
                continue
            if same_page(ad["link"], other["link"]):
                twin = other     # e.g. "Junior Research Fellow" and "Magdalene College... JRF"
                break
            if not distinctive_words(ad["title"]) & distinctive_words(other["title"]):
                continue
            inst_a, inst_b = distinctive_words(ad["institution"]), distinctive_words(other["institution"])
            if inst_a and inst_b and not inst_a & inst_b:
                continue                # e.g. two Leverhulme ECFs at different universities
            twin = other
            break
        if twin:
            twin.setdefault("also_at", []).append(ad["link"])
            # Keep the more informative title ("Magdalene College, Cambridge;
            # 3 year JRF" rather than just "Junior Research Fellow").
            if len(distinctive_words(ad["title"])) > len(distinctive_words(twin["title"])):
                twin["title"] = ad["title"]
                twin["institution"] = twin["institution"] or ad["institution"]
        else:
            kept.append(ad)
    if len(kept) < len(ads):
        print(f"Merged {len(ads) - len(kept)} duplicate ads (same post, several sites).")
    return kept


# =============================================================================
# STEP 3 - a language model reads the best ads
# =============================================================================
# The model EXTRACTS facts (the PhD rule, quoted; its dates; the start date).
# The script then DECIDES eligibility by comparing those dates with your
# timeline (THESIS_SUBMISSION, PHD_AWARD). Small models find rules in a text
# well but compare dates badly, so each side does what it is good at. The
# decision is made when the spreadsheet is written, so changing your
# timeline in the settings updates every verdict without re-running step 3.

LLM_VERSION = 5   # answers stored by an older version of step 3 are re-read

# Step 3 asks the model three short sets of questions per ad:
#   1. ABOUT: what the post is - read from the ad ALONE. The candidate is not
#      shown, because when they were, the model started describing generic
#      college fellowships as if they were about the candidate's own topic.
#   2. RULES: the PhD rule and its dates - also from the ad alone.
#   3. FIT: how well the post suits the candidate - judged from the candidate
#      and the short description produced in step 1, not the whole ad.
# Each answer must fit a template (a "JSON schema") that the model server enforces.

ABOUT_SCHEMA = {
    "type": "object",
    "properties": {
        "summary": {"type": "string"},
        "subjects": {"type": "string"},
        "level": {"type": "string", "enum": [
            "postdoc or early-career fellowship", "PhD or predoctoral",
            "faculty or senior", "non-research post", "unclear"]},
        "start_date": {"type": "string"},
        "deadline": {"type": "string"},
        "other_requirements": {"type": "string"},
    },
    "required": ["summary", "subjects", "level", "start_date", "deadline",
                 "other_requirements"],
}

RULES_SCHEMA = {
    "type": "object",
    "properties": {
        "phd_rule": {"type": "string"},
        "phd_required_at_start": {"type": "boolean"},
        "thesis_submitted_by": {"type": "string"},
        "phd_awarded_by": {"type": "string"},
        "phd_awarded_after": {"type": "string"},
    },
    "required": ["phd_rule", "phd_required_at_start", "thesis_submitted_by",
                 "phd_awarded_by", "phd_awarded_after"],
}

FIT_SCHEMA = {
    "type": "object",
    "properties": {
        "field_fit": {"type": "integer", "minimum": 0, "maximum": 10},
        "other_barrier": {"type": "string"},
    },
    "required": ["field_fit", "other_barrier"],
}

SYSTEM_PROMPT = (
    "You read academic job adverts. Answer only from what you are given; never "
    "invent requirements, subjects or dates. Adverts may be in Spanish, German "
    "or another language; always answer in English."
)

ABOUT_QUESTIONS = """Answer these questions about the advert above:
- summary: one sentence saying what the post is, using only the advert
  (institution, subject, length).
- subjects: the subjects or fields the post is for, in a few words, exactly as
  the advert states them. If it is open to all subjects, write "any subject";
  if to all arts and humanities, "any arts and humanities subject".
- level: the career stage the post is for, judged from the title and the
  description of the duties.
- start_date: when the post starts, as YYYY-MM-DD (1st of the month if only a
  month is given; 1 September of the first year for an academic year such as
  "2027-28"). "" if not stated. Never give the deadline or the end date here.
- deadline: the application deadline as YYYY-MM-DD; "" if not stated.
- other_requirements: requirements other than the PhD (languages, nationality,
  residence, a particular degree or skill), briefly, as the advert states
  them; "" if none."""

RULES_QUESTIONS = """Extract these facts from the advert above:
- phd_rule: copy word for word the sentence that says when applicants must
  have their PhD (submitted, awarded, "in hand", "by the start date",
  "doctorado obtenido entre...", "abgeschlossene Promotion"...). "" if the
  advert has no such sentence.
- phd_required_at_start: true if the PhD must be awarded by the time the post
  starts (including simply "must hold a PhD"); false if the advert accepts
  people still finishing their PhD, or says nothing.
- thesis_submitted_by: a date by which the thesis must be SUBMITTED, as
  YYYY-MM-DD. If the rule says "by the closing date" or "by the application
  deadline", give the application deadline. Otherwise "".
- phd_awarded_by: a date BEFORE which the PhD must be awarded or defended (a
  deadline), as YYYY-MM-DD. If the rule says "by the application deadline" or
  "at the time of application", give the application deadline. For a window
  ("PhD obtained between 1 January 2025 and 31 December 2026") give the END
  of the window. Otherwise "". Never put the start date here.
- phd_awarded_after: an EARLIEST allowed date for the PhD ("not awarded before
  30 September 2025", the START of a window), as YYYY-MM-DD; otherwise ""."""

def fit_questions():
    """The field-fit questions. The scale is described with the fields given
    in [fit_scale] in config.toml; where they are left empty, it falls back on
    general wording (small models judge better with concrete examples)."""
    own = FIT_SCALE["own_field"] or "the candidate's own research field, as described above"
    close = FIT_SCALE["close_fields"] or "fields closely related to the candidate's research"
    near = FIT_SCALE["neighbouring_fields"] or "neighbouring fields where the candidate's work could fit"
    return f"""Answer these questions:
- field_fit (0-10): how well the POST's subjects match the CANDIDATE's
  research. Use this scale:
    10 = the post is specifically for the candidate's own field: {own}.
     8 = the post is specifically for a closely related field: {close}.
     6 = the post is open to any subject, or to any arts and humanities
         subject, or to a neighbouring field: {near}.
     3 = humanities or social sciences, but a different field or period.
     0 = natural sciences, medicine, engineering, business, or non-academic.
  An open "any subject" fellowship is a 6, however well the candidate would
  do in it.
- other_barrier: if the post lists a requirement other than the PhD that the
  candidate clearly does NOT meet (a nationality, residence, language or
  degree the candidate lacks), quote it; otherwise ""."""


# ---- Talking to the model ---------------------------------------------------
# Two kinds of local server are supported ([llm] provider in config.toml):
#   "ollama" - the Ollama app (https://ollama.com).
#   "openai" - any server that copies OpenAI's chat API: LM Studio, llama.cpp's
#              llama-server, vLLM, Jan, LocalAI... The reading window
#              ("context") is then set in the server itself, not here.


def llm_headers():
    return {"Authorization": f"Bearer {LLM_API_KEY}"} if LLM_API_KEY else {}


def llm_ready():
    """Check the model server is running and has the model."""
    if LLM_PROVIDER == "ollama":
        address, how_to_start = f"{LLM_BASE_URL}/api/tags", "Open the Ollama app"
    else:
        address, how_to_start = f"{LLM_BASE_URL}/models", "Start your model server"
    try:
        listing = requests.get(address, headers=llm_headers(), timeout=5).json()
    except (requests.RequestException, ValueError):
        print(f"No model server answers at {LLM_BASE_URL}. {how_to_start} and run "
              f"the script again, or check [llm] base_url in config.toml. Skipping step 3.")
        return False

    if LLM_PROVIDER == "ollama":
        names = [m.get("name", "") for m in listing.get("models", [])]
        if LLM_MODEL not in names and f"{LLM_MODEL}:latest" not in names:
            print(f"Ollama doesn't have the model {LLM_MODEL}. Installed: "
                  f"{', '.join(names) or 'none'}. Change [llm] model in config.toml, "
                  f"or run: ollama pull {LLM_MODEL}. Skipping step 3.")
            return False
    else:
        # Some servers (llama-server) serve a single model under whatever
        # name, so a name that isn't listed is only a warning.
        names = [m.get("id", "") for m in listing.get("data", [])]
        if names and LLM_MODEL not in names:
            print(f"Note: the server lists {', '.join(names)}, not {LLM_MODEL}. "
                  f"Trying anyway; if it fails, change [llm] model in config.toml.")
    return True


def parse_answer(content):
    """The model's reply -> a dict. Reasoning models may "think out loud"
    first, inside <think>...</think>; that part is dropped."""
    content = re.sub(r"<think>.*?</think>", "", content, flags=re.DOTALL)
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", content, re.DOTALL)   # answer wrapped in extra text
        return json.loads(m.group(0)) if m else None


def ask_llm(user_prompt, schema):
    """Send one advert to the model and return its answers as a dict."""
    messages = [{"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt}]
    if LLM_PROVIDER == "ollama":
        return ask_ollama(messages, schema)
    return ask_openai(messages, schema)


def ask_ollama(messages, schema):
    payload = {
        "model": LLM_MODEL,
        "messages": messages,
        "format": schema,
        "stream": False,
        # Qwen models can "think out loud" before answering. That is slower
        # and not needed here, so it's switched off when the model allows it.
        "think": False,
        # temperature 0 = always pick the most likely answer (no randomness).
        "options": {"temperature": 0, "num_ctx": LLM_CONTEXT},
    }
    url = f"{LLM_BASE_URL}/api/chat"
    response = requests.post(url, json=payload, timeout=600)
    if response.status_code == 400 and "think" in response.text.lower():
        payload.pop("think")             # this model doesn't support the switch
        response = requests.post(url, json=payload, timeout=600)
    response.raise_for_status()
    return parse_answer(response.json()["message"]["content"])


def ask_openai(messages, schema):
    payload = {
        "model": LLM_MODEL,
        "messages": messages,
        "temperature": 0,
        # Asks the server to keep the answer to the template, as Ollama does.
        "response_format": {"type": "json_schema",
                            "json_schema": {"name": "answer", "schema": schema}},
        # Switch off "thinking out loud", as for Ollama: servers spell it
        # differently (Ollama and LM Studio: reasoning_effort; llama.cpp:
        # enable_thinking). Thinking makes each answer several times slower
        # and can overflow the reading window.
        "reasoning_effort": "none",
        "chat_template_kwargs": {"enable_thinking": False},
    }
    url = f"{LLM_BASE_URL}/chat/completions"
    while True:
        response = requests.post(url, json=payload, headers=llm_headers(), timeout=600)
        if response.status_code != 400:
            break
        # Drop whatever option this server rejects, and try again.
        rejected = next((k for k in ["reasoning_effort", "chat_template_kwargs", "response_format"]
                         if k in payload and k in response.text), None)
        if not rejected:
            break
        payload.pop(rejected)
        if rejected == "response_format":
            # This server can't enforce a template: describe it in the prompt instead.
            payload["messages"] = messages[:-1] + [{
                "role": "user",
                "content": messages[-1]["content"] + "\n\nReply with a single JSON object "
                           f"that follows this JSON schema, and nothing else:\n{json.dumps(schema)}"}]
    response.raise_for_status()
    return parse_answer(response.json()["choices"][0]["message"]["content"] or "")


def tidy_answer(answer):
    """Make sure every field exists and has a sensible value."""
    answer = dict(answer)
    try:
        fit = max(0, min(10, int(answer.get("field_fit"))))
    except (TypeError, ValueError):
        fit = None
    record = {k: str(answer.get(k, "") or "").strip() for k in
              ["level", "phd_rule", "thesis_submitted_by", "phd_awarded_by",
               "phd_awarded_after", "start_date", "deadline", "other_barrier",
               "summary", "subjects", "other_requirements", "start_note"]}
    record.update({
        "field_fit": fit,
        "phd_required_at_start": bool(answer.get("phd_required_at_start")),
        "model": LLM_MODEL,
        "version": LLM_VERSION,
    })
    return record


def to_date(text):
    """A YYYY-MM-DD string from the model -> a date (None if missing/garbled)."""
    if not text:
        return None
    try:
        return datetime.fromisoformat(text[:10]).date()
    except ValueError:
        return parse_date(text) if re.search(r"\d{4}", text) else None


PHD_WORDS = re.compile(r"ph\.?\s?d|dphil|doctor|doktor|promotion|thesis|dissertation|tesis",
                       re.IGNORECASE)
REFERS_TO_DEADLINE = re.compile(
    r"deadline|closing date|time of (?:the )?application|date of application|when applying|"
    r"bewerbungsschluss|zeitpunkt der bewerbung|plazo|fecha de (?:cierre|solicitud)",
    re.IGNORECASE)
AWARD_WORDS = re.compile(r"award|defen[cd]|complet|conferred|hold|obtain|in hand|"
                         r"abgeschlossen|verteidigt|obtenid|posesi", re.IGNORECASE)


def decide_eligibility(llm, deadline=None):
    """Compare the ad's PhD rule with your timeline. Returns (verdict, note).
    Verdicts: yes / no / check (another barrier) / unclear (no rule found)."""
    def fmt(d):
        return d.strftime("%d %b %Y")

    submitted_by = to_date(llm.get("thesis_submitted_by"))
    awarded_by = to_date(llm.get("phd_awarded_by"))
    awarded_after = to_date(llm.get("phd_awarded_after"))
    start = to_date(llm.get("start_date"))
    quote = llm.get("phd_rule", "")
    if quote and not PHD_WORDS.search(quote):
        quote = ""            # the model quoted something that isn't a PhD rule
    quoted = f' Rule: "{quote[:200]}"' if quote else ""
    if llm.get("start_note") and start:
        quoted += f" ({llm['start_note']})"

    # "PhD completed by the application deadline": the model often leaves the
    # date out because the rule doesn't repeat it, so fill it in here.
    if quote and deadline and not (submitted_by or awarded_by) and REFERS_TO_DEADLINE.search(quote):
        if AWARD_WORDS.search(quote):
            awarded_by = deadline
        else:
            submitted_by = deadline

    if submitted_by and submitted_by < THESIS_SUBMISSION:
        return "no", (f"Thesis must be submitted by {fmt(submitted_by)}; yours: "
                      f"{fmt(THESIS_SUBMISSION)}.{quoted}")
    if awarded_by and awarded_by < PHD_AWARD:
        return "no", (f"PhD must be awarded by {fmt(awarded_by)}; yours: "
                      f"{fmt(PHD_AWARD)}.{quoted}")
    if awarded_after and PHD_AWARD < awarded_after:
        return "no", (f"PhD must be awarded after {fmt(awarded_after)}; yours: "
                      f"{fmt(PHD_AWARD)}.{quoted}")
    if llm.get("phd_required_at_start") and not awarded_by:
        if start and start < PHD_AWARD:
            return "no", (f"PhD needed at start ({fmt(start)}); yours: "
                          f"{fmt(PHD_AWARD)}.{quoted}")
        if not start:
            return "unclear", f"PhD needed at start, but no start date found.{quoted}"

    barrier = llm.get("other_barrier", "")
    if re.search(r"ph\.?d|dphil|doctor|doktor|promotion|thesis|tesis", barrier, re.IGNORECASE):
        barrier = ""   # that's the PhD rule, already handled above
    if barrier:
        return "check", f"Possible other requirement: {barrier[:200]}"

    checked = (submitted_by or awarded_by or awarded_after
               or (llm.get("phd_required_at_start") and start))
    if checked:
        start_note = f" Starts {fmt(start)}." if start else ""
        return "yes", f"Your timeline meets the PhD rule.{start_note}{quoted}"
    if quote and not llm.get("phd_required_at_start"):
        return "yes", f"Open to people still finishing their PhD.{quoted}"
    if quote:
        return "unclear", f"PhD rule without a clear date.{quoted}"
    return "unclear", "No PhD rule found in the ad."


def overall_score(llm, eligible):
    """The script's overall score (0-10): field fit, capped for posts you
    can't take or at the wrong career stage."""
    fit = llm.get("field_fit")
    if fit is None:
        return None
    if eligible == "no" or llm.get("level") in ("PhD or predoctoral", "non-research post"):
        return min(fit, 2)
    if llm.get("level") == "faculty or senior":
        return max(fit - 3, 0)
    return fit


# EURAXESS career-stage labels ("First Stage Researcher (R1)" ...). They are
# removed before the model reads an ad: the model kept treating them as rules.
R_LABELS = re.compile(
    r"(?:Researcher Profile\s*:?)?\s*(?:First Stage|Recognised|Established|Leading)"
    r"\s+Researcher\s*\(R[1-4]\)", re.IGNORECASE)
SITE_DELAY_FOR_FULL_ADS = 10   # seconds; EURAXESS limits how fast it can be read


def judge(ads, cache, rejudge=False):
    """Have the model read the top LLM_TOP_N ads. Answers are saved in the cache after each ad, so stopping halfway
    (Ctrl+C) loses nothing."""
    if not llm_ready():
        return
    if not CANDIDATE_PATH.exists():
        print(f"{CANDIDATE_PATH} not found - it tells the model who you are. Skipping step 3.")
        return
    candidate = CANDIDATE_PATH.read_text(encoding="utf-8").strip()
    profile = read_profile(PROFILE_PATH)
    today = date.today().isoformat()

    def needs_reading(ad):
        stored = cache[ad["key"]].get("llm") or {}
        return (rejudge or stored.get("model") != LLM_MODEL
                or stored.get("version") != LLM_VERSION)

    chosen = ads[:LLM_TOP_N] + [ad for ad in ads[LLM_TOP_N:]
                                if ad["source"] in ALWAYS_READ_SOURCES]
    to_read = [ad for ad in chosen if needs_reading(ad)]
    print(f"Step 3: {LLM_MODEL} reads {len(to_read)} ads "
          f"({len(chosen) - len(to_read)} already read on earlier runs).")
    if to_read:
        print("The model answers three sets of questions per ad: roughly 25-60 seconds per ad.\n")

    for i, ad in enumerate(to_read, start=1):
        record = cache[ad["key"]]
        if record.get("card_only"):
            # Only the EURAXESS result card was stored; fetch the full ad now,
            # since the PhD rule is rarely on the card.
            try:
                html = fetch(ad["link"]).text
                full = trafilatura.extract(html) or clean_text(html)
                record["text"] = (record["text"] + "\n" + full)[:20000]
                record["card_only"] = False
                ad["body"] = record["text"]
            except Exception as err:
                print(f"      (could not download the full ad: {err}; using the card)")
            time.sleep(SITE_DELAY_FOR_FULL_ADS)
        text = R_LABELS.sub("", ad["body"])[:LLM_TEXT_CHARS]
        notes = "".join(f"\nNOTE FROM THE FUNDER'S RULES: {note}\n"
                        for pattern, note in SCHEME_NOTES.items()
                        if re.search(pattern, ad["title"], re.IGNORECASE))
        advert = (
            f"TODAY'S DATE: {today}\n\n"
            f"JOB ADVERT\nTitle: {ad['title']}\n"
            f"Institution: {ad['institution'] or 'not given'}\n"
            f"Text:\n{text}\n{notes}\n"
        )
        started = time.time()
        try:
            about = ask_llm(advert + ABOUT_QUESTIONS, ABOUT_SCHEMA)
            rules = ask_llm(advert + RULES_QUESTIONS, RULES_SCHEMA)
            post = (f"THE POST (as summarised from the advert)\n"
                    f"Title: {ad['title']}\n"
                    f"Institution: {ad['institution'] or 'not given'}\n"
                    f"Summary: {(about or {}).get('summary', '')}\n"
                    f"Subjects: {(about or {}).get('subjects', '')}\n"
                    f"Other requirements: {(about or {}).get('other_requirements', '') or 'none'}\n\n")
            fit = ask_llm(f"CANDIDATE\n{candidate}\n\n"
                             f"WHAT THE CANDIDATE IS LOOKING FOR\n{profile}\n\n"
                             f"{post}{fit_questions()}", FIT_SCHEMA)
        except Exception as err:
            print(f"  [{i}/{len(to_read)}] {ad['title'][:60]} - model error: {err}")
            continue
        if not (about and rules and fit):
            print(f"  [{i}/{len(to_read)}] {ad['title'][:60]} - unreadable answer, skipped")
            continue
        llm = tidy_answer({**about, **rules, **fit})
        if not to_date(llm["start_date"]):
            # No start date found: many fellowships only name an academic
            # year ("2027-28"), which means a start around 1 September.
            year = re.search(r"\b(20\d\d)\s*[-\u2013/]\s*(?:20)?(\d\d)\b",
                             ad["title"] + " " + ad["body"][:3000])
            if year and (int(year.group(2)) - int(year.group(1)[2:])) % 100 == 1:
                llm["start_date"] = f"{year.group(1)}-09-01"
                llm["start_note"] = f"start assumed from '{year.group(0)}'"
        cache[ad["key"]]["llm"] = llm
        CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
        verdict, _ = decide_eligibility(llm, ad["deadline_date"] or to_date(llm.get("deadline")))
        print(f"  [{i}/{len(to_read)}] {ad['title'][:60]} -> field fit {llm['field_fit']}, "
              f"eligible: {verdict} ({time.time() - started:.0f}s)")

    for ad in ads:   # attach stored answers, including ones from earlier runs
        ad["llm"] = cache[ad["key"]].get("llm")
    print("\nStep 3 done.\n")


# =============================================================================
# Output
# =============================================================================


def write_output(ads):
    today = date.today()
    rows = []
    for ad in ads:
        llm = ad.get("llm") or {}
        eligible, note, overall = "", "", None
        d = ad["deadline_date"] or to_date(llm.get("deadline"))  # fall back on the model
        if llm.get("version") == LLM_VERSION:
            eligible, note = decide_eligibility(llm, d)
            overall = overall_score(llm, eligible)
        if d and d < today:
            continue   # the model found a deadline that has already passed
        rows.append({
            "Overall": overall,
            "Field fit (AI)": llm.get("field_fit"),
            "Eligible": eligible,
            "Eligibility note": note,
            "Start (AI)": llm.get("start_date", ""),
            "Level (AI)": llm.get("level", ""),
            "Summary (AI)": llm.get("summary", ""),
            "Subjects (AI)": llm.get("subjects", ""),
            "Fit score": round(ad["fit"], 3),
            "Name": ad["title"],
            "Institution": ad["institution"],
            "Source": ad["source"],
            "Deadline": d.isoformat() if d else "",
            "Days left": (d - today).days if d else "",
            "Link": ad["link"],
            "Also at": " ".join(ad.get("also_at", [])),
            "New this run": ad["first_seen"] == today.isoformat(),
        })

    # Ads the model has read come first, best first; the rest follow in
    # embedding order.
    df = pd.DataFrame(rows)
    df["_read"] = df["Overall"].notna()
    df = df.sort_values(["_read", "Overall", "Fit score"], ascending=False)
    df = df.drop(columns="_read")
    df.to_excel(OUTPUT_FILE, index=False)

    print(f"{len(df)} ads written to {OUTPUT_FILE}\n")
    print("Top 20:")
    for _, r in df.head(20).iterrows():
        overall = "-" if pd.isna(r["Overall"]) else int(r["Overall"])
        eligible = r["Eligible"] or "-"
        print(f"  {overall:>2}  {eligible:<7}  {r['Name'][:65]}  [{r['Deadline'] or 'no deadline'}]")


def check_sources():
    """Load the first results page of each source once and report what the
    script finds there - handy after adding a website to config.toml.
    Nothing is saved."""
    problems = 0
    for site in SEARCH_SITES:
        templates = search_templates(site)
        if not templates:
            print(f"{site['name']}: no queries or search_urls - nothing to search.")
            problems += 1
            continue
        label, template = templates[0]
        url = template.replace("{page}", str(site["first_page"]))
        try:
            links = links_matching(fetch(url).text, url, site["link_pattern"])
        except Exception as err:
            print(f"{label}: could not load {url}\n  {err}")
            problems += 1
            continue
        kept = [i for i in links.values()
                if (not site["title_filter"] or TITLE_PATTERN.search(i["title"]))
                and not EXCLUDE_TITLE_PATTERN.search(i["title"])]
        print(f"{label}: {len(links)} ads on the first page, {len(kept)} with a suitable title")
        for info in (kept or list(links.values()))[:3]:
            print(f"    e.g. {info['title'][:80]}")
        if not links:
            print("  -> no links match link_pattern: check it against an ad's address.")
            problems += 1
        time.sleep(site.get("delay", SECONDS_BETWEEN_DOWNLOADS))
    for page in LISTING_PAGES:
        try:
            html = fetch(page["url"]).text
        except Exception as err:
            print(f"{page['name']}: could not load {page['url']}\n  {err}")
            problems += 1
            continue
        items = listing_links(html, page["url"], page.get("link_pattern"))
        n_links = len(BeautifulSoup(html, "html.parser").find_all("a", href=True))
        print(f"{page['name']}: {len(items)} of {n_links} links look like fellowships or postdocs")
        for _, title, _ in items[:3]:
            print(f"    e.g. {title[:80]}")
        if n_links < 10:
            print("  -> very few links: the page probably draws its list with JavaScript, "
                  "which this script can't run.")
            problems += 1
        time.sleep(SECONDS_BETWEEN_DOWNLOADS)
    print(f"\nChecked {len(SEARCH_SITES) + len(LISTING_PAGES)} sources, "
          f"{problems} with problems. (A page with 0 suitable links may simply have "
          f"no vacancies today.)")


# =============================================================================

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Find postdoc posts that suit you. "
                                 "Settings are read from config.toml.")
    ap.add_argument("--config", default=DEFAULT_CONFIG,
                    help=f"settings file to use (default: {DEFAULT_CONFIG})")
    ap.add_argument("--model", help="language model to use in step 3, instead of "
                    "the one in the settings file")
    ap.add_argument("--no-llm", action="store_true", help="skip step 3")
    ap.add_argument("--rejudge", action="store_true",
                    help="step 3 re-reads every ad, ignoring stored answers")
    ap.add_argument("--check-sources", action="store_true",
                    help="test each website in the settings once, then stop")
    args = ap.parse_args()
    load_config(args.config, model=args.model)

    if args.check_sources:
        check_sources()
        sys.exit()
    cache = collect()
    ads = rank(cache)
    if not args.no_llm:
        judge(ads, cache, rejudge=args.rejudge)
    write_output(ads)
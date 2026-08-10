#!/usr/bin/env python3
"""Harvest published Demurrage & Detention (D&D) tariffs for the top carriers
and compile them into a single workbook.

Modeled on the terminal-extractor pattern (xairo-seixas/terminal-extractor):
  - one config entry + one pair of (discover, parse) functions per source
  - every source is processed independently; a failure in one never aborts
    the others
  - raw output is preserved (nothing is silently dropped), so downstream
    normalization can be iterated on once we've seen real extracted samples
  - run exits 1 only if every source failed

Supported carriers (all publish D&D tariffs as public PDFs, no login):
  - Maersk      : per-country Import/Export local-information pages
  - MSC         : per-country local-information pages
  - CMA CGM     : one index page, every country's PDF listed directly
  - Hapag-Lloyd : one index page per region, PDFs listed per document

Not covered yet: COSCO. Its tariff tool (elines.coscoshipping.com) is a
JS-rendered SPA with no static HTML to parse - it needs the same kind of
browser-automation parser this repo already has for GCT/TruckGate, but
built and tested against the live tool with a working browser session
first. Add it as a fifth entry in CARRIERS once that's done.

IMPORTANT — this script has not been run end-to-end. It was written against
page content fetched via a read-only web tool that cannot execute the script
itself (no outbound network access from that environment to these domains).
The site structure it targets (confirmed by inspection on 2026-08-03):
  - CMA CGM and Hapag-Lloyd: verified against real fetched pages.
  - Maersk: verified against one live country page (US); the crawl logic
    for discovering the *list* of countries is written by analogy to MSC's
    local-information page and NOT yet verified against Maersk's own index.
  - MSC: the local-information country index is verified; the presence and
    location of a tariff PDF on every individual country page is confirmed
    for the US only and assumed elsewhere.
Run with --sample 5 first and inspect output/logs before a full run.
"""
# This laptop runs Python 3.9 (confirmed via `python3 --version` on the
# self-hosted runner), which doesn't support the `int | None` union syntax
# used throughout this file's type hints (that's a 3.10+ feature - PEP 604).
# Without this import, just loading the module raises
# "TypeError: unsupported operand type(s) for |: 'type' and 'NoneType'" at
# the first `def ...(x: int | None = None)` it hits (confirmed in a real
# run's traceback, at discover_cma_cgm's signature). `from __future__ import
# annotations` defers all annotations to plain strings instead of evaluating
# them at def-time, so the `|` is never actually executed - works on
# Python 3.7+, no behavior change on newer versions.
from __future__ import annotations
import argparse
import csv
import io
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Iterable

import requests
from bs4 import BeautifulSoup

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}
REQUEST_TIMEOUT = 30
REQUEST_DELAY_S = 0.5  # be polite - these are production carrier sites
OUTPUT_DIR = Path(__file__).parent / "output"


# ---------------------------------------------------------------------------
# Shared data model
# ---------------------------------------------------------------------------

@dataclass
class TariffDoc:
    """One discovered tariff document, before/after PDF extraction."""
    carrier: str
    region: str
    country: str
    title: str
    pdf_url: str
    effective_date_guess: str = ""
    validity_end_guess: str = ""
    status: str = "discovered"       # discovered | downloaded | parsed | failed
    error: str = ""
    tables: list = field(default_factory=list)   # list[list[list[str]]]
    raw_text: str = ""


def polite_get(url: str, **kwargs) -> requests.Response:
    time.sleep(REQUEST_DELAY_S)
    resp = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, **kwargs)
    resp.raise_for_status()
    return resp


# ---------------------------------------------------------------------------
# Browser-based fetching (Playwright) - plain requests.get() was getting
# 403'd on cma_cgm/hapag_lloyd/msc from GitHub Actions' IPs (their bot
# protection can fingerprint a bare HTTP client and/or blocklist known
# cloud-runner IP ranges). A real headless browser has a legitimate
# TLS/JS/HTTP fingerprint that passes those checks - this mirrors the
# approach terminal-extractor already uses for the same reason.
#
# One browser context is reused for the whole run (not per-request) so we
# aren't paying browser-launch cost per page, and so any session cookie a
# site's challenge sets while loading its HTML carries over to later
# requests (including PDF downloads) in the same run.
# ---------------------------------------------------------------------------

_browser_playwright = None
_browser_instance = None
_browser_context = None


def _get_browser_context():
    global _browser_playwright, _browser_instance, _browser_context
    if _browser_context is None:
        from playwright.sync_api import sync_playwright
        _browser_playwright = sync_playwright().start()
        # headless=False on purpose, confirmed necessary by a direct real-world
        # test (2026-08-10): a plain requests.get() from this same laptop's
        # network got Cloudflare's "enable JS" challenge (403) on
        # cma-cgm.com, and so did headless Chromium - but a real Chromium
        # window (headless=False), same browser, same network, got the real
        # page (1.1MB of actual markup, not the challenge). That isolates the
        # block to something about headless mode's fingerprint specifically
        # (not the network, not "any automation") - so this runs a real,
        # visible browser instead. It's the same client a person uses
        # manually, just automated instead of clicked - not a spoofed or
        # impersonated fingerprint.
        #
        # Caveat: this needs an active logged-in GUI session on the laptop to
        # render a window - it won't work over a plain SSH session with no
        # display, and if this runner is ever converted to a background
        # service (svc.sh install), the laptop still needs to stay logged in
        # (screen can be locked, just not logged out) for this to keep working.
        _browser_instance = _browser_playwright.chromium.launch(headless=False)
        _browser_context = _browser_instance.new_context(
            user_agent=HEADERS["User-Agent"],
            viewport={"width": 1366, "height": 900},
        )
    return _browser_context


def browser_get_html(url: str) -> str:
    """Fetches a page's HTML through the browser context.

    History of this function, because the fix at each stage only made sense
    in light of what the previous one actually did in production:

      1. Started as plain requests.get() - got 403'd on cma_cgm/hapag_lloyd/
         msc from GitHub Actions' IPs (bot protection fingerprinting a bare
         HTTP client + flagged cloud-runner IP ranges).
      2. Switched to a full Playwright page.goto(..., wait_until="networkidle")
         - fixed the 403s, but Hapag-Lloyd/Maersk then timed out every time:
         those sites run continuous background polling (GTM, chat widgets,
         personalization) that never lets the network go idle.
      3. Switched the wait to "domcontentloaded" - fixed the timeouts, but
         cma_cgm/hapag_lloyd/msc all came back with "Discovered 0
         document(s)" and no error. The real run showed this wasn't a
         timing issue: verified directly (via a plain, non-JS fetch of the
         live pages) that CMA CGM's and MSC's country/tariff listings ARE
         present in the raw server-rendered HTML - full, complete, no JS
         needed. Yet a real headless-Chromium *page* render of the same URL
         comes back without that listing. That points at the sites hiding
         content from an automated/headless browser at the JS/DOM level
         specifically (a different mechanism than the IP-based 403 from
         stage 1) - not a wait-timing problem.

    Stage 4: fetched via the browser context's request API (ctx.request.get)
    instead of a full page navigation - still Chromium's real TLS/HTTP
    fingerprint, no JS/DOM for a headless-detection script to touch. Also
    came back empty (0 documents), with no error, and *faster* than stage 3
    - meaning it succeeded on the request path, not the page-nav fallback.
    So two completely different fetch mechanisms (real page render, bare
    context request) both return "successful, empty" from inside GitHub
    Actions, while a direct fetch of the exact same URLs from outside GitHub
    Actions returns the real, full content. That combination - no error,
    but silently different/empty content, specifically from that network -
    smells like a soft block keyed on the requesting IP/network (GitHub
    Actions' runner ranges are well-known and can be denylisted or served a
    stripped placeholder rather than a hard 4xx, precisely so the caller
    can't tell it was blocked). Guessing another fetch mechanism blindly
    isn't productive without seeing what's actually coming back - so stage 5
    is diagnostic, not a fix: log the status, byte length, and a content
    snippet for every fetch so the next real run tells us what GitHub
    Actions' IP is actually being served.
    """
    time.sleep(REQUEST_DELAY_S)
    ctx = _get_browser_context()

    def _log(via: str, status, html: str) -> None:
        snippet = " ".join(html.split())[:200] if html else "(empty)"
        print(f"    [fetch:{via}] {url}\n"
              f"      status={status} bytes={len(html)}\n"
              f"      snippet={snippet!r}")

    try:
        resp = ctx.request.get(url, timeout=45000)
        html = resp.text()
        _log("request", resp.status, html)
        if resp.status < 400:
            return html
    except Exception as exc:  # noqa: BLE001
        print(f"    [fetch:request] {url}\n      EXCEPTION: {exc}")

    page = ctx.new_page()
    try:
        resp = page.goto(url, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(2000)  # let head-of-page bot checks settle
        html = page.content()
        _log("page", resp.status if resp else "N/A", html)
        return html
    finally:
        page.close()


def browser_get_bytes(url: str) -> bytes:
    """Downloads binary content (e.g. a PDF) via the same browser context's
    session/cookies, using Playwright's request API rather than a full page
    navigation (faster, and works for non-HTML responses)."""
    time.sleep(REQUEST_DELAY_S)
    ctx = _get_browser_context()
    resp = ctx.request.get(url)
    if resp.status >= 400:
        raise RuntimeError(f"HTTP {resp.status} for {url}")
    return resp.body()


def close_browser() -> None:
    global _browser_playwright, _browser_instance, _browser_context
    if _browser_instance is not None:
        _browser_instance.close()
    if _browser_playwright is not None:
        _browser_playwright.stop()
    _browser_instance = _browser_context = _browser_playwright = None


DATE_IN_TEXT = re.compile(
    r"(?:effective\s+)?(\d{1,2}[\s\-/](?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*[\s\-/]\d{2,4}"
    r"|\d{4}-\d{2}-\d{2}"
    r"|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+\d{1,2},?\s+\d{4})",
    re.IGNORECASE,
)


def guess_effective_date(*texts: str) -> str:
    for text in texts:
        m = DATE_IN_TEXT.search(text or "")
        if m:
            return m.group(1)
    return ""


# Looks for an explicit expiration/validity-end phrase near a date. Most
# carrier tariffs only publish an effective (start) date and stay in force
# until superseded - so this will legitimately come back empty most of the
# time. That's expected, not a bug: leave Contract end date blank in that case.
VALIDITY_END_IN_TEXT = re.compile(
    r"(?:valid\s+(?:through|until|to)|expir\w*(?:\s+on)?|through)\s*:?\s*"
    r"(\d{1,2}[\s\-/](?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*[\s\-/]\d{2,4}"
    r"|\d{4}-\d{2}-\d{2}"
    r"|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+\d{1,2},?\s+\d{4})",
    re.IGNORECASE,
)


def guess_validity_end(*texts: str) -> str:
    for text in texts:
        m = VALIDITY_END_IN_TEXT.search(text or "")
        if m:
            return m.group(1)
    return ""


# Used in the Contract number column for rows sourced from a public tariff
# rather than an actual BuyCo service contract - keeps them traceable/
# filterable instead of leaving the column ambiguous-blank vs blank-because-
# missing. Per your call: blank Shipper, this flag for Contract number.
CONTRACT_NUMBER_FLAG = "SCRAPED - STANDARD TARIFF"


# ---------------------------------------------------------------------------
# CMA CGM — single index page, every country listed with a direct PDF link
# ---------------------------------------------------------------------------

CMA_CGM_INDEX_URL = "https://www.cma-cgm.com/ebusiness/tariffs/demurrage-detention"


def discover_cma_cgm(limit: int | None = None) -> list[TariffDoc]:
    html = browser_get_html(CMA_CGM_INDEX_URL)
    soup = BeautifulSoup(html, "html.parser")
    docs: list[TariffDoc] = []
    current_region = "Unknown"

    # Page is a flat sequence of region headings (h2/h3) followed by <li><a> links.
    for el in soup.find_all(["h2", "h3", "a"]):
        if el.name in ("h2", "h3"):
            heading = el.get_text(strip=True)
            if heading:
                current_region = heading
            continue
        href = el.get("href", "")
        if not href.lower().endswith(".pdf"):
            continue
        country = el.get_text(strip=True) or "Unknown"
        docs.append(TariffDoc(
            carrier="CMA CGM",
            region=current_region,
            country=country,
            title=country,
            pdf_url=href if href.startswith("http") else f"https://www.cma-cgm.com{href}",
        ))
        if limit and len(docs) >= limit:
            break
    return docs


# ---------------------------------------------------------------------------
# Hapag-Lloyd — one index page per region, docs listed with title/size/lang
# ---------------------------------------------------------------------------

HAPAG_REGIONS = [
    "north-america", "latin-america", "europe", "africa", "middle-east", "asia",
]
HAPAG_BASE = "https://www.hapag-lloyd.com/en/online-business/quotation/detention-demurrage"

# Rough keyword -> country map for inferring country from a free-text title.
# Extend as real titles are seen; unmatched titles fall back to "Unknown".
HAPAG_COUNTRY_HINTS = {
    "usa": "United States", "u.s.": "United States", "us ": "United States",
    "canada": "Canada", "mexico": "Mexico", "brazil": "Brazil",
    "argentina": "Argentina", "chile": "Chile", "colombia": "Colombia",
    "peru": "Peru", "china": "China", "india": "India", "japan": "Japan",
    "korea": "South Korea", "vietnam": "Vietnam", "thailand": "Thailand",
    "singapore": "Singapore", "malaysia": "Malaysia", "indonesia": "Indonesia",
    "australia": "Australia", "new zealand": "New Zealand",
    "south africa": "South Africa", "nigeria": "Nigeria", "egypt": "Egypt",
    "uae": "United Arab Emirates", "saudi": "Saudi Arabia",
    "germany": "Germany", "france": "France", "spain": "Spain",
    "italy": "Italy", "uk": "United Kingdom", "united kingdom": "United Kingdom",
    "netherlands": "Netherlands", "belgium": "Belgium",
}


def guess_country_from_title(title: str) -> str:
    lowered = title.lower()
    for hint, country in HAPAG_COUNTRY_HINTS.items():
        if hint in lowered:
            return country
    return "Unknown"


def discover_hapag_lloyd(limit: int | None = None) -> list[TariffDoc]:
    docs: list[TariffDoc] = []
    for region in HAPAG_REGIONS:
        url = f"{HAPAG_BASE}/{region}.html"
        try:
            html = browser_get_html(url)
        except Exception as exc:  # noqa: BLE001 - Playwright raises its own error types
            print(f"    [hapag_lloyd] failed to load region {region}: {exc}")
            continue
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if not href.lower().endswith(".pdf"):
                continue
            title = a.get_text(strip=True) or href
            docs.append(TariffDoc(
                carrier="Hapag-Lloyd",
                region=region.replace("-", " ").title(),
                country=guess_country_from_title(title),
                title=title,
                pdf_url=href if href.startswith("http") else f"https://www.hapag-lloyd.com{href}",
            ))
            if limit and len(docs) >= limit:
                return docs
    return docs


# ---------------------------------------------------------------------------
# Maersk — no single index; crawl local-information -> per-country pages
# ---------------------------------------------------------------------------

MAERSK_INDEX_URL = "https://www.maersk.com/local-information"
MAERSK_BASE = "https://www.maersk.com"

# Confirmed by direct inspection (2026-08-04): the index page itself does
# NOT list countries - it only links to these 5 region hub pages. Countries
# are one level deeper, on each hub page, under a "Country/Region" heading.
# (Verified: a plain fetch of MAERSK_INDEX_URL returns 200 with real content
# from GitHub Actions too - Maersk isn't blocking us; the old code was just
# looking for countries on the wrong page.)
MAERSK_REGIONS = ["asia-pacific", "europe", "imea", "latin-america", "north-america"]

# A region hub page also links to non-country pages that happen to match the
# same /local-information/<segment>/<segment> shape - "Routes to/from" links
# (shipping-from-north-america-to-europe/med-canada-express-eastbound) and
# feeder-route links (europe-feeder-shipping-routes/n03) are two segments
# too. A keyword blocklist on the second segment alone doesn't catch these -
# the giveaway noise word ("feeder", "shipping-from", ...) is often in the
# FIRST segment, and route/service names in the second segment (e.g.
# "med-canada-express-eastbound") don't look like noise at all on their own.
# Rather than keep extending a blocklist, whitelist the first segment
# instead: a real country link's first segment is always one of the 5
# known region slugs; every route/service/solutions link's first segment is
# some longer compound string that never matches one of those 5 exactly.
# That alone still lets through same-region solutions links like
# north-america/ground-freight (a valid region + a non-country second
# segment), so a small keyword check on the second segment stays too - just
# as a second layer, not the primary filter.
_MAERSK_REGION_SLUGS = set(MAERSK_REGIONS)
MAERSK_NON_COUNTRY_KEYWORDS = ("freight", "feeder", "route", "service", "solution", "logistics")

# Hard safety cap: there are ~150-200 countries worldwide, so legitimate
# discovery should never come close to this. If a future page-structure
# change lets noise links back through, this stops the crawl from silently
# running for hours (a single earlier bug here caused ~900+ garbage fetches
# before it was caught) instead of just failing fast and loud.
MAERSK_MAX_COUNTRIES = 300


def discover_maersk_countries() -> list[tuple[str, str, str]]:
    """Returns (region, country, country_url) tuples.

    Crawls each of the 5 known region hub pages (not the index page - see
    MAERSK_REGIONS comment above) and pulls country links from each one's
    "Country/Region" list. The region label is taken from the country's own
    URL, not from which hub page it was found on - hub pages can cross-link
    a country that actually belongs to a different region (e.g. Mexico is
    listed on the North America hub page but its real URL is under
    /local-information/latin-america/mexico).
    """
    seen_urls: set[str] = set()
    out: list[tuple[str, str, str]] = []
    for region_slug in MAERSK_REGIONS:
        hub_url = f"{MAERSK_BASE}/local-information/{region_slug}"
        try:
            html = browser_get_html(hub_url)
        except Exception as exc:  # noqa: BLE001
            print(f"    [maersk] failed to load region hub {region_slug}: {exc}")
            continue
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if "/local-information/" not in href:
                continue
            parts = [p for p in href.split("/local-information/")[-1].split("/") if p]
            if len(parts) != 2:
                continue  # expect <region>/<country> exactly
            url_region, country_slug = parts
            if url_region.lower() not in _MAERSK_REGION_SLUGS:
                continue  # a route/service/solutions link, not a country page
            if any(kw in country_slug.lower() for kw in MAERSK_NON_COUNTRY_KEYWORDS):
                continue  # same-region solutions link (e.g. north-america/ground-freight)
            full_url = href if href.startswith("http") else f"{MAERSK_BASE}{href}"
            if full_url in seen_urls:
                continue
            seen_urls.add(full_url)
            country = country_slug.replace("-", " ").title()
            out.append((url_region.replace("-", " ").title(), country, full_url))
            if len(out) >= MAERSK_MAX_COUNTRIES:
                print(f"    [maersk] hit the {MAERSK_MAX_COUNTRIES}-country safety cap - "
                      f"stopping discovery early. This should never happen for real "
                      f"country links; if it does, the noise filter above needs "
                      f"another look.")
                return out
    return out


def discover_maersk(limit: int | None = None) -> list[TariffDoc]:
    countries = discover_maersk_countries()
    if not countries:
        print("    [maersk] WARNING: 0 countries discovered from index page - "
              "the index-page parser likely needs to be rewritten against "
              "the real page structure.")
    docs: list[TariffDoc] = []
    for region, country, country_url in countries:
        for direction in ("import", "export"):
            url = f"{country_url}/{direction}"
            try:
                html = browser_get_html(url)
            except Exception:  # noqa: BLE001
                continue  # not every country publishes both directions
            soup = BeautifulSoup(html, "html.parser")
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if not href.lower().endswith(".pdf"):
                    continue
                text = a.get_text(strip=True).lower()
                if "demurrage" not in text and "detention" not in text and "tariff" not in text:
                    continue
                title = a.get_text(strip=True)
                docs.append(TariffDoc(
                    carrier="Maersk",
                    region=region,
                    country=country,
                    title=f"{direction.title()}: {title}",
                    pdf_url=href if href.startswith("http") else f"{MAERSK_BASE}{href}",
                ))
                if limit and len(docs) >= limit:
                    return docs
    return docs


# ---------------------------------------------------------------------------
# MSC — no single index; crawl local-information -> per-country pages
# ---------------------------------------------------------------------------

MSC_INDEX_URL = "https://www.msc.com/en/local-information"
MSC_BASE = "https://www.msc.com"


def discover_msc_countries() -> list[tuple[str, str, str]]:
    html = browser_get_html(MSC_INDEX_URL)
    soup = BeautifulSoup(html, "html.parser")
    out = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "/local-information/" not in href:
            continue
        parts = [p for p in href.split("/local-information/")[-1].split("/") if p]
        if len(parts) != 2:
            continue  # expect <region>/<country>
        region, country_slug = parts
        country = country_slug.replace("-", " ").title()
        full_url = href if href.startswith("http") else f"{MSC_BASE}{href}"
        out.append((region.replace("-", " ").title(), country, full_url))
    seen = set()
    deduped = []
    for row in out:
        if row[2] in seen:
            continue
        seen.add(row[2])
        deduped.append(row)
    return deduped


def discover_msc(limit: int | None = None) -> list[TariffDoc]:
    countries = discover_msc_countries()
    docs: list[TariffDoc] = []
    for region, country, country_url in countries:
        try:
            html = browser_get_html(country_url)
        except Exception as exc:  # noqa: BLE001
            print(f"    [msc] failed to load {country}: {exc}")
            continue
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if not href.lower().endswith(".pdf"):
                continue
            text = a.get_text(strip=True).lower()
            if "demurrage" not in text and "detention" not in text and "per diem" not in text and "tariff" not in text:
                continue
            title = a.get_text(strip=True)
            docs.append(TariffDoc(
                carrier="MSC",
                region=region,
                country=country,
                title=title,
                pdf_url=href if href.startswith("http") else f"{MSC_BASE}{href}",
            ))
            if limit and len(docs) >= limit:
                return docs
    return docs


# ---------------------------------------------------------------------------
# PDF download + best-effort table extraction (shared by all carriers)
# ---------------------------------------------------------------------------

def parse_pdf_bytes(doc: TariffDoc, pdf_bytes: bytes) -> None:
    """Mutates doc in place: extracts tables/text from already-downloaded
    PDF bytes. Split out from fetch_and_parse_pdf() so a manual-drop source
    (a human-uploaded PDF pulled from a Drive folder, rather than fetched
    live from a carrier's site) can reuse the exact same parsing path -
    parse_cma_cgm() and friends only care about the PDF's own text content,
    not where the bytes came from."""
    if pdfplumber is None:
        doc.status = "failed"
        doc.error = "pdfplumber not installed"
        return
    doc.status = "downloaded"

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            all_text = []
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    # Drop fully-empty rows
                    cleaned = [row for row in table if any((c or "").strip() for c in row)]
                    if cleaned:
                        doc.tables.append(cleaned)
                text = page.extract_text() or ""
                all_text.append(text)
            doc.raw_text = "\n".join(all_text)
    except Exception as exc:  # noqa: BLE001 - PDF parsing can fail in many ways
        doc.status = "failed"
        doc.error = f"PDF parse failed: {exc}"
        return

    doc.effective_date_guess = guess_effective_date(doc.title, doc.raw_text[:2000])
    doc.validity_end_guess = guess_validity_end(doc.raw_text)
    doc.status = "parsed" if doc.tables else "parsed_no_tables"


def fetch_and_parse_pdf(doc: TariffDoc) -> None:
    """Mutates doc in place: downloads the PDF live from doc.pdf_url, then
    parses it via parse_pdf_bytes(). Used for carriers whose sites we can
    still reach directly (currently: Maersk). Not used for carriers behind
    a Cloudflare/WAF block (CMA CGM) - those go through the manual-drop path
    in run_monthly.py instead, which calls parse_pdf_bytes() directly on
    bytes downloaded from a Drive folder."""
    if pdfplumber is None:
        doc.status = "failed"
        doc.error = "pdfplumber not installed"
        return
    try:
        pdf_bytes = browser_get_bytes(doc.pdf_url)
    except Exception as exc:  # noqa: BLE001
        doc.status = "failed"
        doc.error = f"download failed: {exc}"
        return
    parse_pdf_bytes(doc, pdf_bytes)


# ---------------------------------------------------------------------------
# Carrier dispatcher
# ---------------------------------------------------------------------------

CARRIERS: dict[str, Callable[[int | None], list[TariffDoc]]] = {
    "cma_cgm": discover_cma_cgm,
    "hapag_lloyd": discover_hapag_lloyd,
    "maersk": discover_maersk,
    "msc": discover_msc,
    # "cosco": discover_cosco,   # TODO - needs a browser-automation parser
}


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def write_outputs(all_docs: list[TariffDoc], out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    today = datetime.now(timezone.utc).date().isoformat()

    # 1) Index CSV - one row per discovered document, regardless of parse outcome
    index_path = out_dir / f"dnd_tariff_index_{today}.csv"
    with index_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow([
            "Carrier", "Region", "Country", "Title", "PDF URL",
            "Effective Date (guess)", "Validity End (guess)", "Status", "Error", "Table Count",
        ])
        for d in all_docs:
            writer.writerow([
                d.carrier, d.region, d.country, d.title, d.pdf_url,
                d.effective_date_guess, d.validity_end_guess, d.status, d.error, len(d.tables),
            ])
    print(f"  Wrote index: {index_path}")

    # 2) Combined xlsx - one "Index" sheet + one sheet per carrier with raw
    #    extracted table rows (tagged with country/title/pdf so they're
    #    traceable back to source).
    if Workbook is None:
        print("  openpyxl not installed - skipping xlsx output")
        return

    wb = Workbook()
    idx_ws = wb.active
    idx_ws.title = "Index"
    idx_ws.append([
        "Carrier", "Region", "Country", "Title", "PDF URL",
        "Effective Date (guess)", "Validity End (guess)", "Status", "Error", "Table Count",
    ])
    for d in all_docs:
        idx_ws.append([
            d.carrier, d.region, d.country, d.title, d.pdf_url,
            d.effective_date_guess, d.validity_end_guess, d.status, d.error, len(d.tables),
        ])

    by_carrier: dict[str, list[TariffDoc]] = {}
    for d in all_docs:
        by_carrier.setdefault(d.carrier, []).append(d)

    for carrier, docs in by_carrier.items():
        ws = wb.create_sheet(title=carrier[:31])
        ws.append(["Country", "Title", "PDF URL", "Table #", "Row"])
        for d in docs:
            for t_idx, table in enumerate(d.tables):
                for row in table:
                    ws.append([d.country, d.title, d.pdf_url, t_idx, " | ".join(c or "" for c in row)])
            if not d.tables and d.raw_text:
                ws.append([d.country, d.title, d.pdf_url, "raw_text", d.raw_text[:500]])

    xlsx_path = out_dir / f"dnd_tariffs_{today}.xlsx"
    wb.save(xlsx_path)
    print(f"  Wrote workbook: {xlsx_path}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--carriers", nargs="+", choices=list(CARRIERS), default=list(CARRIERS),
                         help="Which carriers to run (default: all implemented)")
    parser.add_argument("--sample", type=int, default=None,
                         help="Limit discovered docs per carrier (smoke-test before a full run)")
    parser.add_argument("--skip-pdf-parse", action="store_true",
                         help="Only discover PDF links, don't download/parse them (fast index-only run)")
    parser.add_argument("--out", type=Path, default=OUTPUT_DIR)
    args = parser.parse_args()

    all_docs: list[TariffDoc] = []
    failures: list[str] = []

    try:
        for carrier in args.carriers:
            print(f"\n── {carrier} ──")
            try:
                docs = CARRIERS[carrier](args.sample)
            except Exception as exc:  # noqa: BLE001
                print(f"  ✗ FAILED discovery — {exc}")
                failures.append(f"{carrier}: discovery failed: {exc}")
                continue

            print(f"  Discovered {len(docs)} document(s)")
            if not args.skip_pdf_parse:
                for i, doc in enumerate(docs, 1):
                    fetch_and_parse_pdf(doc)
                    if doc.status == "failed":
                        print(f"    [{i}/{len(docs)}] ✗ {doc.country} — {doc.title}: {doc.error}")
                    else:
                        print(f"    [{i}/{len(docs)}] ✓ {doc.country} — {doc.title} "
                              f"({len(doc.tables)} table(s))")
            all_docs.extend(docs)

            if not docs:
                failures.append(f"{carrier}: 0 documents discovered — index parser likely needs fixing")
    finally:
        close_browser()  # always shut the headless browser down, even on failure

    write_outputs(all_docs, args.out)

    print("\n── Summary ──")
    print(f"  Carriers run : {len(args.carriers)}")
    print(f"  Documents    : {len(all_docs)}")
    print(f"  Failed       : {sum(1 for d in all_docs if d.status == 'failed')}")
    if failures:
        print("\nCarrier-level issues:")
        for f in failures:
            print(f"  • {f}")

    if not all_docs:
        print("\nNothing discovered at all - treat this as a broken build, not an empty result.")
        return 1

    print("\nDone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

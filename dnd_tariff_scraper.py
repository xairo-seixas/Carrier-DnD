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
        _browser_instance = _browser_playwright.chromium.launch(headless=True)
        _browser_context = _browser_instance.new_context(
            user_agent=HEADERS["User-Agent"],
            viewport={"width": 1366, "height": 900},
        )
    return _browser_context


def browser_get_html(url: str) -> str:
    """Loads a page in a real browser and returns its rendered HTML."""
    time.sleep(REQUEST_DELAY_S)
    ctx = _get_browser_context()
    page = ctx.new_page()
    try:
        page.goto(url, wait_until="networkidle", timeout=45000)
        return page.content()
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


def discover_maersk_countries() -> list[tuple[str, str, str]]:
    """Returns (region, country, country_url) tuples.

    UNVERIFIED: written by analogy to MSC's index page. Maersk's local
    information hub may be laid out differently (e.g. JS-driven region
    picker) - if this returns zero countries, the index page needs to be
    inspected directly (view-source) and this function rewritten.
    """
    html = browser_get_html(MAERSK_INDEX_URL)
    soup = BeautifulSoup(html, "html.parser")
    out = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "/local-information/" not in href:
            continue
        parts = [p for p in href.split("/local-information/")[-1].split("/") if p]
        if len(parts) != 2:
            continue  # expect <region>/<country> exactly
        region, country_slug = parts
        country = country_slug.replace("-", " ").title()
        full_url = href if href.startswith("http") else f"{MAERSK_BASE}{href}"
        out.append((region.replace("-", " ").title(), country, full_url))
    # de-dupe
    seen = set()
    deduped = []
    for row in out:
        if row[2] in seen:
            continue
        seen.add(row[2])
        deduped.append(row)
    return deduped


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

def fetch_and_parse_pdf(doc: TariffDoc) -> None:
    """Mutates doc in place: downloads the PDF, extracts tables/text."""
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

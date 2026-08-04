#!/usr/bin/env python3
"""Monthly D&D tariff refresh - the scheduled entry point.

What one run does, in order:
  1. Download the current master workbook from a fixed Google Drive file
     (GDRIVE_FILE_ID) - this is the same workbook BuyCo works out of, not a
     fresh copy of the template each time.
  2. Scrape all implemented carriers (dnd_tariff_scraper.py) and download +
     parse every discovered tariff PDF.
  3. Map each into BuyCo's Sheet1 schema and upsert into the downloaded
     workbook (buyco_mapper.upsert_template):
       - new (carrier, direction, port, penalty type) combo -> append
       - existing combo -> only Contract end date is refreshed, and only if
         a later one was found; every other field is left exactly as-is
       - a combo whose existing row is marked "Manually Reviewed" (col AA)
         -> skipped entirely, no update, no duplicate append
  4. Upload the updated workbook back to the SAME Drive file, then verify
     the round trip by re-downloading and comparing bytes (mirrors
     terminal_slot_scraper.py's upload_file(..., verify=True)).

Like terminal_slot_scraper.py, each carrier is processed independently -
one carrier failing to scrape doesn't stop the others, and doesn't stop the
upload of whatever WAS successfully mapped. The run exits 1 only if the
whole thing produced nothing (e.g. Drive auth failed, or every carrier
failed) - see main() for the exact conditions.

Required environment variables (same pattern as terminal_slot_scraper.py):
  GDRIVE_FILE_ID          - the master workbook's Drive file ID (not a
                             folder - this script overwrites one specific
                             file in place)
  GDRIVE_CREDENTIALS_JSON - service account credentials JSON (or a JSON
                             string, same as terminal_slot_scraper.py)

Optional environment variables:
  CMA_CGM_MANUAL_FOLDER_ID - Drive folder ID for CMA CGM's manual PDF drop
                             (see MANUAL_DROP_CARRIERS below). If unset,
                             CMA CGM is skipped for the run with a clear
                             message instead of attempting a live scrape
                             that's known to be blocked.

Why CMA CGM isn't scraped live: its site (and Hapag-Lloyd's and MSC's) sits
behind Cloudflare/WAF bot protection that returns a "Just a moment..." JS
challenge or a hard "Access Denied" specifically to automated traffic from
GitHub Actions' datacenter IPs - confirmed directly (status 403 on both a
raw request and a full headless-browser page render, same result either
way). No fetch-mechanism change in the scraper fixes that; it's a network-
level block, not a parsing bug. Per your call, CMA CGM instead runs off a
manual monthly drop: someone at BuyCo downloads that month's PDFs from CMA
CGM's site in a normal browser and uploads them to a fixed Drive folder
(CMA_CGM_MANUAL_FOLDER_ID); this script lists whatever's in that folder,
downloads each PDF, and runs it through the exact same parse_cma_cgm()
parser a live-scraped PDF would go through - the parser only reads the
PDF's own text, so it doesn't care how the bytes got here. Hapag-Lloyd and
MSC are left on the (currently failing) live path for now; they can be
switched to the same manual-drop pattern later by adding another folder-id
secret and one more entry in MANUAL_DROP_CARRIERS.
"""
import io
import json
import os
import ssl
import sys
import time
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

import buyco_mapper
from dnd_tariff_scraper import CARRIERS, TariffDoc, fetch_and_parse_pdf, parse_pdf_bytes, close_browser

WORK_DIR = Path(__file__).parent / "output"
DOWNLOADED_MASTER = WORK_DIR / "master_downloaded.xlsx"
UPDATED_MASTER = WORK_DIR / "master_updated.xlsx"

# Carriers whose PDFs come from a human-maintained Drive folder instead of a
# live scrape, because their site blocks GitHub Actions' automated traffic.
# Maps carrier name -> the environment variable holding that carrier's
# Drive folder ID. Add an entry here (and set the matching secret) to move
# another blocked carrier onto the same manual-drop pattern.
MANUAL_DROP_CARRIERS: dict[str, str] = {
    "cma_cgm": "CMA_CGM_MANUAL_FOLDER_ID",
}


def get_drive_service():
    info = json.loads(os.environ["GDRIVE_CREDENTIALS_JSON"])
    if isinstance(info, str):
        info = json.loads(info)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


# Transient network failures talking to Google's API - a dropped TLS
# connection mid-request, a reset socket, a brief timeout - happen
# occasionally from GitHub Actions runners, especially on a multi-MB
# upload. None of these mean anything is wrong with the request itself;
# retrying the exact same call a few seconds later normally just works.
# Caught a real one of these in production: ssl.SSLEOFError partway through
# svc.files().update(), which crashed the whole run even though every
# carrier had already scraped and mapped successfully (172 rows ready to
# go) - losing that work to a one-off network blip isn't worth it.
_TRANSIENT_EXCEPTIONS = (
    ssl.SSLError, ssl.SSLEOFError, ConnectionError, TimeoutError, OSError,
)


def _with_retries(fn, *, attempts: int = 4, base_delay_s: float = 5.0, label: str = "request"):
    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            return fn()
        except _TRANSIENT_EXCEPTIONS as exc:
            last_exc = exc
            if attempt == attempts:
                break
            delay = base_delay_s * (2 ** (attempt - 1))
            print(f"  {label} failed (attempt {attempt}/{attempts}): {exc!r} - retrying in {delay:.0f}s")
            time.sleep(delay)
    raise last_exc


def download_master(svc, file_id: str, dest: Path) -> None:
    def _do():
        request = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return buf.getvalue()

    data = _with_retries(_do, label="download_master")
    dest.write_bytes(data)
    print(f"  Downloaded master workbook -> {dest} ({dest.stat().st_size:,} bytes)")


def upload_master(svc, file_id: str, path: Path, verify: bool = True) -> None:
    data = path.read_bytes()

    def _do_upload():
        media = MediaIoBaseUpload(io.BytesIO(data), mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ), resumable=False)
        svc.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()

    _with_retries(_do_upload, label="upload_master")
    print(f"  Uploaded updated workbook ({len(data):,} bytes)")

    if verify:
        def _do_verify():
            return svc.files().get_media(fileId=file_id, supportsAllDrives=True).execute()

        stored = _with_retries(_do_verify, label="upload_master verify")
        if stored != data:
            raise RuntimeError("Verification failed - Drive copy differs from what we uploaded.")
        print("  Verified: Drive copy is byte-identical to the upload.")


def discover_manual_pdfs(svc, folder_id: str, carrier: str) -> list[TariffDoc]:
    """Lists every PDF a human has dropped into a fixed Drive folder for a
    carrier whose site blocks automated access, downloads each one, and
    parses it through the same PDF-parsing path a live-scraped document
    would go through (parse_pdf_bytes -> the carrier's structured parser in
    buyco_mapper). The filename (minus extension) becomes the doc's title/
    country label purely for traceability in the output - the actual
    port/direction/penalty data is derived from the PDF's own text by
    buyco_mapper.parse_cma_cgm(), not from the filename, so there's no
    naming convention the person dropping files needs to follow."""
    docs: list[TariffDoc] = []
    query = f"'{folder_id}' in parents and mimeType='application/pdf' and trashed=false"
    resp = svc.files().list(
        q=query, fields="files(id,name)",
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    for f in files:
        doc = TariffDoc(
            carrier="CMA CGM" if carrier == "cma_cgm" else carrier,
            region="Manual upload",
            country=Path(f["name"]).stem,
            title=f["name"],
            pdf_url=f"drive:{f['id']}",  # not fetchable - just traceability
        )
        try:
            request = svc.files().get_media(fileId=f["id"], supportsAllDrives=True)
            buf = io.BytesIO()
            downloader = MediaIoBaseDownload(buf, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            parse_pdf_bytes(doc, buf.getvalue())
        except Exception as exc:  # noqa: BLE001
            doc.status = "failed"
            doc.error = f"drive download/parse failed: {exc}"
        docs.append(doc)
    return docs


def scrape_all_carriers(svc) -> tuple[list, list[str]]:
    all_docs = []
    failures = []
    try:
        for carrier, discover_fn in CARRIERS.items():
            print(f"\n── {carrier} ──")

            if carrier in MANUAL_DROP_CARRIERS:
                folder_id = os.environ.get(MANUAL_DROP_CARRIERS[carrier])
                if not folder_id:
                    print(f"  Skipped - {carrier}'s site blocks automated access and "
                          f"{MANUAL_DROP_CARRIERS[carrier]} isn't set, so there's no "
                          f"manual-drop folder to read from this month.")
                    failures.append(f"{carrier}: skipped - no manual-drop folder configured")
                    continue
                try:
                    docs = discover_manual_pdfs(svc, folder_id, carrier)
                except Exception as exc:  # noqa: BLE001
                    print(f"  ✗ manual-drop folder read failed: {exc}")
                    failures.append(f"{carrier}: manual-drop folder read failed: {exc}")
                    continue
                print(f"  Found {len(docs)} PDF(s) in the manual-drop folder")
                for i, doc in enumerate(docs, 1):
                    status = "✓" if doc.status != "failed" else "✗"
                    print(f"    [{i}/{len(docs)}] {status} {doc.title}")
                all_docs.extend(docs)
                if not docs:
                    failures.append(f"{carrier}: 0 PDFs found in the manual-drop folder this run")
                continue

            try:
                docs = discover_fn(None)
            except Exception as exc:  # noqa: BLE001
                print(f"  ✗ discovery failed: {exc}")
                failures.append(f"{carrier}: discovery failed: {exc}")
                continue
            print(f"  Discovered {len(docs)} document(s)")
            for i, doc in enumerate(docs, 1):
                fetch_and_parse_pdf(doc)
                status = "✓" if doc.status != "failed" else "✗"
                print(f"    [{i}/{len(docs)}] {status} {doc.country} — {doc.title}")
            all_docs.extend(docs)
            if not docs:
                failures.append(f"{carrier}: 0 documents discovered this run")
    finally:
        close_browser()  # always shut the headless browser down, even on failure
    return all_docs, failures


def main() -> int:
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    file_id = os.environ["GDRIVE_FILE_ID"]

    svc = get_drive_service()
    download_master(svc, file_id, DOWNLOADED_MASTER)

    docs, failures = scrape_all_carriers(svc)
    if not docs:
        print("\nNo documents discovered from any carrier - not touching the Drive "
              "file, since uploading now would just be a no-op with a scary-looking "
              "empty diff.")
        return 1

    summary = buyco_mapper.upsert_template(DOWNLOADED_MASTER, docs, UPDATED_MASTER)
    print(f"\n── Mapping summary ──\n{summary}")

    upload_master(svc, file_id, UPDATED_MASTER)

    print("\n── Run summary ──")
    print(f"  Documents scraped        : {len(docs)}")
    print(f"  Rows appended            : {summary['rows_appended']}")
    print(f"  Rows with end date update: {summary['rows_end_date_updated']}")
    print(f"  Rows skipped (manual)    : {summary['rows_skipped_manually_reviewed']}")
    if failures:
        print("\nCarrier-level issues (did not block the run):")
        for f in failures:
            print(f"  • {f}")

    print("\nDone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

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
"""
import io
import json
import os
import sys
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

import buyco_mapper
from dnd_tariff_scraper import CARRIERS, fetch_and_parse_pdf

WORK_DIR = Path(__file__).parent / "output"
DOWNLOADED_MASTER = WORK_DIR / "master_downloaded.xlsx"
UPDATED_MASTER = WORK_DIR / "master_updated.xlsx"


def get_drive_service():
    info = json.loads(os.environ["GDRIVE_CREDENTIALS_JSON"])
    if isinstance(info, str):
        info = json.loads(info)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


def download_master(svc, file_id: str, dest: Path) -> None:
    request = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    dest.write_bytes(buf.getvalue())
    print(f"  Downloaded master workbook -> {dest} ({dest.stat().st_size:,} bytes)")


def upload_master(svc, file_id: str, path: Path, verify: bool = True) -> None:
    data = path.read_bytes()
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), resumable=False)
    svc.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
    print(f"  Uploaded updated workbook ({len(data):,} bytes)")
    if verify:
        stored = svc.files().get_media(fileId=file_id, supportsAllDrives=True).execute()
        if stored != data:
            raise RuntimeError("Verification failed - Drive copy differs from what we uploaded.")
        print("  Verified: Drive copy is byte-identical to the upload.")


def scrape_all_carriers() -> tuple[list, list[str]]:
    all_docs = []
    failures = []
    for carrier, discover_fn in CARRIERS.items():
        print(f"\n── {carrier} ──")
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
    return all_docs, failures


def main() -> int:
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    file_id = os.environ["GDRIVE_FILE_ID"]

    svc = get_drive_service()
    download_master(svc, file_id, DOWNLOADED_MASTER)

    docs, failures = scrape_all_carriers()
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

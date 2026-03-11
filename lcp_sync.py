import json
import re
import os
import sys
import time
import traceback
import zipfile
import signal
from datetime import datetime, timezone
from io import BytesIO

import msal
import boto3
import requests

# =========================================================
# CONFIG
# =========================================================

STATE_FILE = "lcp_sync_state.json"
GRAPH_TIMEOUT = 120
MAX_RETRIES = 5
THROTTLE_SECONDS = 0.2       # Delay between S3 uploads (API-friendly)
GRAPH_THROTTLE = 0.5          # Delay between Graph API calls (avoid 429)
RUNNING = True                # Graceful shutdown flag

# =========================================================
# LOGGING
# =========================================================

def log(msg):
    ts = datetime.now(timezone.utc).isoformat()
    print(f"{ts} | {msg}", flush=True)

# =========================================================
# LOAD CONFIG & SECRETS
# =========================================================

log(f"Loading config from: {os.path.abspath('config.json')}")
log(f"Loading secrets from: {os.path.abspath('secrets.json')}")

with open("config.json") as f:
    config = json.load(f)

with open("secrets.json") as f:
    secrets = json.load(f)

SITE_ID = config["site_id"]
LIBRARY = config["library"]
MONITOR_ROOT = config["monitor_root"]
POLL_INTERVAL = config["poll_interval_seconds"]
BUCKET = config["s3_bucket"]
CASE_ID_PATTERN = config["case_id_pattern"]
DRY_RUN = config.get("dry_run", False)
EXCLUDED_FOLDERS = set(config.get("excluded_folders", []))

if DRY_RUN:
    log("*** DRY-RUN MODE ENABLED — no files will be uploaded ***")

# =========================================================
# STATE MANAGEMENT
# =========================================================

if os.path.exists(STATE_FILE):
    with open(STATE_FILE) as f:
        state = json.load(f)
else:
    state = {"delta_link": None}

FIRST_RUN = state.get("delta_link") is None

def save_state():
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w") as f:
        json.dump(state, f)
    os.replace(tmp, STATE_FILE)

# =========================================================
# GRAPH AUTH (MSAL)
# =========================================================

log("Authenticating to Microsoft Graph")

app = msal.ConfidentialClientApplication(
    secrets["client_id"],
    authority=f"https://login.microsoftonline.com/{secrets['tenant_id']}",
    client_credential=secrets["client_secret"]
)

def get_graph_headers():
    token = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in token:
        raise Exception(f"Graph auth failed: {token}")
    return {
        "Authorization": f"Bearer {token['access_token']}",
        "Content-Type": "application/json"
    }

# =========================================================
# S3 CLIENT (separate creds for finallcpreports bucket)
# =========================================================

s3 = boto3.client(
    "s3",
    aws_access_key_id=secrets["aws"]["access_key"],
    aws_secret_access_key=secrets["aws"]["secret_key"],
    region_name=secrets["aws"].get("region", "us-east-1")
)

# =========================================================
# HELPERS
# =========================================================

def retry(fn, *args, **kwargs):
    """Retry with exponential backoff. Handles 429 (throttled) responses."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            result = fn(*args, **kwargs)
            # Handle HTTP 429 (Too Many Requests) from Graph API
            if hasattr(result, "status_code") and result.status_code == 429:
                retry_after = int(result.headers.get("Retry-After", 30))
                log(f"Graph API throttled (429). Waiting {retry_after}s...")
                time.sleep(retry_after)
                if attempt == MAX_RETRIES:
                    result.raise_for_status()
                continue
            # Handle HTTP 503/504 (service unavailable / gateway timeout)
            if hasattr(result, "status_code") and result.status_code in (503, 504):
                sleep = min(2 ** attempt, 60)
                log(f"Graph API returned {result.status_code}. Retrying in {sleep}s...")
                time.sleep(sleep)
                if attempt == MAX_RETRIES:
                    result.raise_for_status()
                continue
            return result
        except requests.exceptions.Timeout:
            sleep = min(2 ** attempt, 60)
            log(f"Request timeout. Retry {attempt}/{MAX_RETRIES} after {sleep}s")
            if attempt == MAX_RETRIES:
                raise
            time.sleep(sleep)
        except requests.exceptions.ConnectionError:
            sleep = min(2 ** attempt, 60)
            log(f"Connection error. Retry {attempt}/{MAX_RETRIES} after {sleep}s")
            if attempt == MAX_RETRIES:
                raise
            time.sleep(sleep)
        except Exception as e:
            if attempt == MAX_RETRIES:
                raise
            sleep = min(2 ** attempt, 60)
            log(f"Retry {attempt}/{MAX_RETRIES} after {sleep}s -> {e}")
            time.sleep(sleep)

def graph_get(url, stream=False):
    """Make a Graph API GET request with retry, throttle, and timeout handling."""
    time.sleep(GRAPH_THROTTLE)  # Polite delay between Graph calls
    resp = retry(
        requests.get,
        url,
        headers=get_graph_headers(),
        stream=stream,
        timeout=GRAPH_TIMEOUT
    )
    if hasattr(resp, "raise_for_status"):
        resp.raise_for_status()
    return resp

# =========================================================
# CASE ID EXTRACTION
# =========================================================

# Date pattern: MM-DD-YYYY or DD-MM-YYYY (e.g., "03-05-2026")
_DATE_PATTERN = re.compile(r'^\d{2}-\d{2}-\d{4}$')

# Month names to skip
_MONTH_NAMES = {
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december"
}

def _is_organizational_folder(segment):
    """
    Returns True if the folder is a year, date, or month name
    (organizational folders, NOT patient case folders).
    """
    # Year folder: exactly 4 digits in range 2000-2099
    if re.match(r'^\d{4}$', segment):
        try:
            year = int(segment)
            if 2000 <= year <= 2099:
                return True
        except ValueError:
            pass

    # Date folder: MM-DD-YYYY pattern (e.g., 03-05-2026)
    if _DATE_PATTERN.match(segment):
        return True

    # Month name (case-insensitive)
    # Handles: "March", "January_2025", "April_2025", etc.
    first_word = segment.split("_")[0].split(" ")[0].lower()
    if first_word in _MONTH_NAMES:
        return True

    return False


def extract_case_info(parent_path):
    """
    Extract case ID and relative path from a SharePoint parent path.

    Supported folder name formats:
      "4846_LCP Adan Andrade"       → case_id = "4846"
      "0009"                        → case_id = "0009"
      "4843-Sheila Marie Tan Nadres"→ case_id = "4843"
      "4828_Ross Kip Hyams"         → case_id = "4828"
      "483_Ofelia Moreno"           → case_id = "483"

    Auto-skips organizational folders:
      "2026"        → year (skipped)
      "03-05-2026"  → date (skipped)
      "March"       → month name (skipped)
      "April_2025"  → month name (skipped)

    Scans path segments from shallowest to deepest, returns FIRST match.
    Everything after the case ID folder is preserved as relative path.
    """
    path_after_root = parent_path.split("root:")[-1] if "root:" in parent_path else parent_path
    segments = path_after_root.strip("/").split("/")

    for i, segment in enumerate(segments):
        if segment in EXCLUDED_FOLDERS:
            continue

        # Skip organizational folders (years, dates, months)
        if _is_organizational_folder(segment):
            continue

        match = re.match(CASE_ID_PATTERN, segment)
        if match:
            case_id = match.group(1)
            rel_parts = segments[i + 1:]
            rel_path = "/".join(rel_parts) if rel_parts else ""
            return case_id, rel_path

    return None, None

# =========================================================
# S3 FOLDER MANAGEMENT
# =========================================================

_created_case_folders = set()  # Cache to avoid repeated S3 checks

def ensure_s3_folders(case_id):
    """
    Create Input/, Output/, GroundTruth/ inside the case ID folder.
    Only runs once per case_id per daemon lifecycle.
    """
    if case_id in _created_case_folders:
        return

    for folder in ["Input/", "Output/", "GroundTruth/"]:
        key = f"{case_id}/{folder}"
        try:
            s3.head_object(Bucket=BUCKET, Key=key)
        except Exception:
            s3.put_object(Bucket=BUCKET, Key=key, Body=b"")
            log(f"Created S3 folder: s3://{BUCKET}/{key}")

    _created_case_folders.add(case_id)

# =========================================================
# S3 UPLOAD
# =========================================================

def stream_to_s3(download_url, s3_key):
    """Stream a file from SharePoint directly to S3 (no local disk)."""
    with graph_get(download_url, stream=True) as r:
        r.raise_for_status()
        s3.upload_fileobj(r.raw, BUCKET, s3_key)
    time.sleep(THROTTLE_SECONDS)

def upload_bytes_to_s3(data_bytes, s3_key):
    """Upload in-memory bytes to S3."""
    s3.upload_fileobj(BytesIO(data_bytes), BUCKET, s3_key)
    time.sleep(THROTTLE_SECONDS)

# =========================================================
# ZIP FILE HANDLING
# =========================================================

def handle_zip_file(download_url, case_id, rel_path, zip_filename):
    """
    Download ZIP from SharePoint into memory, extract all files,
    upload each to s3://finallcpreports/{case_id}/Input/{rel_path}/{extracted_path}.
    Preserves internal ZIP folder structure.
    Skips directories and __MACOSX artifacts.
    """
    resp = graph_get(download_url)
    zip_buffer = BytesIO(resp.content)
    uploaded = []

    base = f"{case_id}/Input"
    if rel_path:
        base = f"{base}/{rel_path}"

    try:
        with zipfile.ZipFile(zip_buffer) as zf:
            for member in zf.namelist():
                # Skip directories and macOS artifacts
                if member.endswith("/") or "__MACOSX" in member:
                    continue
                filename = os.path.basename(member)
                if not filename:
                    continue

                # Preserve ZIP internal folder structure
                s3_key = f"{base}/{member}"
                file_data = zf.read(member)
                upload_bytes_to_s3(file_data, s3_key)
                uploaded.append(s3_key)
                log(f"  ZIP extracted: {zip_filename}/{member} -> {s3_key}")
    except zipfile.BadZipFile:
        log(f"  WARNING: {zip_filename} is not a valid ZIP. Uploading as-is.")
        s3_key = f"{base}/{zip_filename}"
        upload_bytes_to_s3(resp.content, s3_key)
        uploaded.append(s3_key)

    return uploaded

# =========================================================
# EMAIL
# =========================================================

def send_email(subject, body):
    """Send email via Microsoft Graph API."""
    email_cfg = secrets["email"]

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [
                {"emailAddress": {"address": r}}
                for r in email_cfg["recipients"]
            ]
        },
        "saveToSentItems": "true"
    }

    r = retry(
        requests.post,
        f"https://graph.microsoft.com/v1.0/users/{email_cfg['sender']}/sendMail",
        headers=get_graph_headers(),
        json=payload,
        timeout=60
    )

    if hasattr(r, "status_code") and r.status_code in (200, 202):
        log("Email sent successfully")
        return True

    log(f"Email send failed: {r.status_code if hasattr(r, 'status_code') else r}")
    return False

# =========================================================
# MAIN POLL CYCLE
# =========================================================

def get_drive_id():
    """Get the drive ID for the Documents library."""
    resp = graph_get(
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives"
    ).json()

    drives = resp.get("value", [])
    if not drives:
        raise Exception(f"No drives found: {resp}")

    drive = next((d for d in drives if d["name"] == LIBRARY), None)
    if not drive:
        raise Exception(f"Library '{LIBRARY}' not found. Available: {[d['name'] for d in drives]}")

    return drive["id"]


def get_folder_delta_url(drive_id):
    """
    Build the initial delta URL scoped to the monitored folder only.
    Uses: GET /drives/{drive_id}/root:/{folder_path}:/delta
    This ensures delta queries only track changes inside our folder,
    NOT the entire drive. Keeps state file small and polls fast.
    """
    # MONITOR_ROOT starts with / e.g. "/General/MSP Team/MSP Team Data/LCP Report Data"
    folder_path = MONITOR_ROOT.lstrip("/")
    return f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}:/delta"


def poll_and_sync(drive_id):
    """
    Run one delta query cycle scoped to the monitored folder:
    - Fetch all changes since last delta token (folder-level, not drive-level)
    - For each new/modified file in a case ID folder -> upload to S3
    - Save delta state after processing
    """
    global FIRST_RUN

    delta_url = (
        get_folder_delta_url(drive_id)
        if FIRST_RUN else state["delta_link"]
    )

    synced_files = []
    skipped_count = 0
    error_count = 0
    page_count = 0

    while delta_url:
        page_count += 1
        log(f"Fetching delta page {page_count}...")

        resp = graph_get(delta_url).json()

        for item in resp.get("value", []):
            # Only process files (skip folders)
            if "file" not in item:
                continue

            parent_path = item.get("parentReference", {}).get("path", "")

            # First run: consume delta baseline without uploading
            if FIRST_RUN:
                continue

            # Extract case ID and relative path
            case_id, rel_path = extract_case_info(parent_path)
            filename = item.get("name", "")
            file_size = item.get("size", 0)
            size_mb = round(file_size / (1024 * 1024), 2)

            if not case_id:
                skipped_count += 1
                log(f"  SKIP (no case ID): {parent_path}/{filename}")
                continue

            # Build S3 key preserving subfolder structure
            s3_base = f"{case_id}/Input"
            if rel_path:
                s3_base = f"{s3_base}/{rel_path}"
            s3_key = f"{s3_base}/{filename}"

            # Dry-run mode: log only
            if DRY_RUN:
                log(f"  DRY-RUN: case_id={case_id} | {parent_path}/{filename} -> s3://{BUCKET}/{s3_key}")
                synced_files.append(f"{s3_key} ({size_mb} MB) [DRY-RUN]")
                continue

            # Real upload
            try:
                log(f"  SYNC: case_id={case_id} | {filename} ({size_mb} MB)")

                # Ensure Input/Output/GroundTruth folders exist
                ensure_s3_folders(case_id)

                download_url = (
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
                    f"/items/{item['id']}/content"
                )

                # Check if ZIP file
                if filename.lower().endswith(".zip"):
                    log(f"  ZIP detected: {filename} — extracting...")
                    extracted = handle_zip_file(download_url, case_id, rel_path, filename)
                    for ek in extracted:
                        synced_files.append(f"{ek} [from {filename}]")
                    log(f"  ZIP done: {len(extracted)} files extracted from {filename}")
                else:
                    # Stream regular file to S3
                    stream_to_s3(download_url, s3_key)
                    synced_files.append(f"{s3_key} ({size_mb} MB)")
                    log(f"  -> s3://{BUCKET}/{s3_key}")

            except Exception as e:
                error_count += 1
                log(f"  ERROR: {filename} -> {e}")
                log(traceback.format_exc())

        # Save delta link from this page
        if "@odata.deltaLink" in resp:
            state["delta_link"] = resp["@odata.deltaLink"]

        delta_url = resp.get("@odata.nextLink")
        save_state()

    # First run complete — next run will process files
    if FIRST_RUN:
        FIRST_RUN = False
        save_state()
        log("First run complete — delta baseline consumed. No files uploaded.")
        log("Next poll cycle will process new files.")
        return

    # Summary
    log(f"Poll complete: {len(synced_files)} synced, {skipped_count} skipped, {error_count} errors")

    # Only send email when files are actually synced (not on empty polls)
    if synced_files:
        # Group files by case ID folder for a clean summary
        folder_counts = {}
        for f in synced_files:
            case_id_part = f.split("/")[0]
            folder_counts[case_id_part] = folder_counts.get(case_id_part, 0) + 1

        mode = "[DRY-RUN] " if DRY_RUN else ""
        subject = f"{mode}LCP Sync — {len(synced_files)} files synced"

        folder_lines = "\n".join(
            f"  {cid}/ — {cnt} file(s)" for cid, cnt in sorted(folder_counts.items())
        )

        body = "\n".join([
            f"{mode}LCP Sync Summary",
            f"Time (UTC): {datetime.now(timezone.utc).isoformat()}",
            f"Total files synced: {len(synced_files)}",
            "",
            "Folders:",
            folder_lines
        ])

        try:
            send_email(subject, body)
        except Exception as e:
            log(f"Failed to send summary email: {e}")

    # Send error email only if errors occurred (separate from sync email)
    if error_count > 0:
        try:
            send_email(
                "LCP Sync — Error Alert",
                f"Errors during sync: {error_count}\n"
                f"Time (UTC): {datetime.now(timezone.utc).isoformat()}\n"
                f"Check logs for details."
            )
        except Exception as e:
            log(f"Failed to send error email: {e}")

# =========================================================
# GRACEFUL SHUTDOWN
# =========================================================

def handle_signal(signum, frame):
    global RUNNING
    log(f"Received signal {signum}. Shutting down gracefully...")
    RUNNING = False

signal.signal(signal.SIGTERM, handle_signal)
signal.signal(signal.SIGINT, handle_signal)

# =========================================================
# ENTRYPOINT — DAEMON LOOP
# =========================================================

if __name__ == "__main__":
    log("=" * 60)
    log("LCP Sync Daemon Starting")
    log(f"Monitoring: {MONITOR_ROOT}")
    log(f"S3 Bucket: {BUCKET}")
    log(f"Poll Interval: {POLL_INTERVAL}s")
    log(f"Dry Run: {DRY_RUN}")
    log(f"First Run: {FIRST_RUN}")
    log("=" * 60)

    try:
        drive_id = get_drive_id()
        log(f"Drive ID resolved: {drive_id}")
        log(f"Delta scoped to folder: {MONITOR_ROOT}")
        log(f"State file: {os.path.abspath(STATE_FILE)}")
    except Exception as e:
        log(f"FATAL: Could not resolve drive ID: {e}")
        log(traceback.format_exc())
        send_email(
            "LCP Sync — FATAL: Cannot Start",
            f"Failed to resolve SharePoint drive.\n\n{e}\n\n{traceback.format_exc()}"
        )
        sys.exit(1)

    while RUNNING:
        try:
            poll_and_sync(drive_id)
        except Exception as e:
            log(f"Poll cycle crashed: {e}")
            log(traceback.format_exc())
            try:
                send_email(
                    "LCP Sync — Poll Cycle Error",
                    f"Error: {e}\n\n{traceback.format_exc()}"
                )
            except Exception:
                log("Could not send error notification email")

        if RUNNING:
            log(f"Sleeping {POLL_INTERVAL}s until next poll...")
            # Sleep in small intervals for responsive shutdown
            for _ in range(POLL_INTERVAL):
                if not RUNNING:
                    break
                time.sleep(1)

    log("LCP Sync Daemon stopped.")

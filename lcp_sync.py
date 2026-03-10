import json
import re
import os
import sys
import time
import traceback
import zipfile
import signal
from datetime import datetime, timezone, timedelta
from io import BytesIO

import msal
import boto3
import requests
import PyPDF2
import tempfile
import subprocess
import glob
import redshift_connector

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
# PATIENT NAME LOOKUP (Redshift)
# =========================================================

_patient_name_cache = {}  # Cache: case_id -> patient name

def lookup_patient_name(case_id):
    """Look up patient full_name from Redshift by case_id. Returns 'Unknown' if not found."""
    if case_id in _patient_name_cache:
        return _patient_name_cache[case_id]

    name = "Unknown"
    try:
        rs_cfg = secrets["redshift"]
        conn = redshift_connector.connect(
            host=rs_cfg["host"],
            port=rs_cfg["port"],
            database=rs_cfg["database"],
            user=rs_cfg["user"],
            password=rs_cfg["password"]
        )
        cursor = conn.cursor()
        cursor.execute(
            "SELECT DISTINCT full_name FROM prod_pi_injury.stg_pi_injury.stg_lcp_pifirm_case_deals_info WHERE Case_ID = %s",
            (int(case_id),)
        )
        row = cursor.fetchone()
        if row and row[0]:
            name = row[0].strip()
        cursor.close()
        conn.close()
    except Exception as e:
        log(f"  Redshift lookup failed for case_id={case_id}: {e}")

    _patient_name_cache[case_id] = name
    return name

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
# PAGE COUNTING
# =========================================================

_IMAGE_EXTENSIONS = {"jpg", "jpeg", "png", "tiff", "tif", "bmp", "gif"}

def count_pages(file_bytes, filename):
    """Count pages in a file. Returns 0 if unknown type."""
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    try:
        if ext == "pdf":
            reader = PyPDF2.PdfReader(BytesIO(file_bytes))
            return len(reader.pages)
        if ext in ("docx", "doc"):
            with tempfile.TemporaryDirectory() as tmpdir:
                docx_path = os.path.join(tmpdir, filename)
                with open(docx_path, "wb") as f:
                    f.write(file_bytes)
                # Use isolated user profile to prevent lock conflicts between conversions
                profile_dir = os.path.join(tmpdir, "profile")
                subprocess.run(
                    ["/opt/libreoffice26.2/program/soffice", "--headless",
                     f"-env:UserInstallation=file://{profile_dir}",
                     "--convert-to", "pdf", "--outdir", tmpdir, docx_path],
                    capture_output=True, timeout=120
                )
                pdf_files = glob.glob(os.path.join(tmpdir, "*.pdf"))
                if pdf_files:
                    with open(pdf_files[0], "rb") as pf:
                        reader = PyPDF2.PdfReader(pf)
                        return len(reader.pages)
            return 1
        if ext in _IMAGE_EXTENSIONS:
            return 1
    except Exception as e:
        log(f"  Page count failed for {filename}: {e}")
    return 0


def download_and_upload(download_url, s3_key, filename):
    """Download file to memory, count pages, upload to S3."""
    resp = graph_get(download_url)
    file_bytes = resp.content
    pages = count_pages(file_bytes, filename)
    upload_bytes_to_s3(file_bytes, s3_key)
    return pages

# =========================================================
# ZIP FILE HANDLING
# =========================================================

def handle_zip_file(download_url, case_id, rel_path, zip_filename):
    """
    Download ZIP from SharePoint into memory, extract all files,
    upload each to s3://finallcpreports/{case_id}/Input/{rel_path}/{extracted_path}.
    Preserves internal ZIP folder structure.
    Skips directories and __MACOSX artifacts.
    Returns (uploaded_keys, total_pages).
    """
    resp = graph_get(download_url)
    zip_buffer = BytesIO(resp.content)
    uploaded = []
    total_pages = 0

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
                total_pages += count_pages(file_data, filename)
                upload_bytes_to_s3(file_data, s3_key)
                uploaded.append(s3_key)
                log(f"  ZIP extracted: {zip_filename}/{member} -> {s3_key}")
    except zipfile.BadZipFile:
        log(f"  WARNING: {zip_filename} is not a valid ZIP. Uploading as-is.")
        s3_key = f"{base}/{zip_filename}"
        upload_bytes_to_s3(resp.content, s3_key)
        uploaded.append(s3_key)

    return uploaded, total_pages

# =========================================================
# EMAIL
# =========================================================

def send_email(subject, body, html=False):
    """Send email via Microsoft Graph API. Set html=True for HTML body."""
    email_cfg = secrets["email"]

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML" if html else "Text",
                "content": body
            },
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
                synced_files.append({"s3_key": s3_key, "case_id": case_id, "pages": 0})
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
                    extracted, zip_pages = handle_zip_file(download_url, case_id, rel_path, filename)
                    for ek in extracted:
                        synced_files.append({"s3_key": ek, "case_id": case_id, "pages": 0})
                    # Assign total zip pages to the first entry
                    if extracted:
                        synced_files[-len(extracted)]["pages"] = zip_pages
                    log(f"  ZIP done: {len(extracted)} files extracted, {zip_pages} pages from {filename}")
                else:
                    # Download, count pages, upload to S3
                    pages = download_and_upload(download_url, s3_key, filename)
                    synced_files.append({"s3_key": s3_key, "case_id": case_id, "pages": pages})
                    log(f"  -> s3://{BUCKET}/{s3_key} ({pages} pages)")

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
        return [], 0

    # Summary
    log(f"Poll complete: {len(synced_files)} synced, {skipped_count} skipped, {error_count} errors")

    return synced_files, error_count


DASHBOARD_URL = "http://100.24.25.37:3005/msp/ocr-reports"


def _build_sync_html(synced_files, error_count):
    """Build HTML email body for sync summary."""
    # Aggregate per case folder
    folder_stats = {}
    for f in synced_files:
        cid = f["case_id"]
        if cid not in folder_stats:
            folder_stats[cid] = {"files": 0, "pages": 0}
        folder_stats[cid]["files"] += 1
        folder_stats[cid]["pages"] += f["pages"]

    total_files = len(synced_files)
    total_pages = sum(s["pages"] for s in folder_stats.values())
    num_cases = len(folder_stats)
    est = timezone(timedelta(hours=-5))
    ts = datetime.now(est).strftime("%Y-%m-%d %I:%M %p EST")
    mode = "[DRY-RUN] " if DRY_RUN else ""

    # Build folder rows
    PAGE_WARN_THRESHOLD = 900
    folder_rows = ""
    has_any_ready = False
    for cid, stats in sorted(folder_stats.items()):
        over_limit = stats["pages"] > PAGE_WARN_THRESHOLD
        if over_limit:
            row_style = 'style="background:#fff9e6;"'
            status_badge = (
                '<span style="color:#b26a00;background:#fff4db;padding:4px 10px;'
                'border-radius:4px;font-size:12px;font-weight:600;">'
                '&#9888; Limit exceeded</span>'
            )
        else:
            has_any_ready = True
            row_style = ''
            status_badge = (
                '<span style="color:#1f7a3e;background:#e6f6ec;padding:4px 10px;'
                'border-radius:4px;font-size:12px;font-weight:600;">'
                'Ready</span>'
            )
        patient_name = lookup_patient_name(cid)
        td = 'style="padding:10px 0;border-bottom:1px solid #f1f1f1;"'
        folder_rows += (
            f'<tr {row_style}>'
            f'<td {td}>{cid}</td>'
            f'<td {td}>{patient_name}</td>'
            f'<td {td}>{stats["files"]}</td>'
            f'<td {td}>{stats["pages"]}</td>'
            f'<td {td}>{status_badge}</td></tr>'
        )

    error_block = ""
    if error_count > 0:
        error_block = (
            '<div style="background:#fff4db;border-left:4px solid #b26a00;'
            'padding:12px 16px;border-radius:4px;margin-bottom:20px;color:#b26a00;">'
            f'&#9888; {error_count} error(s) occurred during sync. Check logs for details.</div>'
        )

    button_block = ""
    if has_any_ready:
        button_block = (
            f'<div style="text-align:center;margin:25px 0 10px;">'
            f'<a href="{DASHBOARD_URL}" '
            f'style="background-color:#2f6fed;color:white;padding:12px 28px;'
            f'text-decoration:none;border-radius:5px;font-size:15px;'
            f'font-weight:600;display:inline-block;">'
            f'Go to Dashboard</a></div>'
        )

    html = f"""
    <div style="width:100%;padding:30px 0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;">
      <div style="width:600px;margin:auto;background:#ffffff;border-radius:8px;border:1px solid #e3e6ea;overflow:hidden;">

        <div style="background:#2f6fed;color:white;padding:18px 25px;font-size:20px;font-weight:600;">
          {mode}Case Files Ready for Processing
        </div>

        <div style="padding:25px;color:#333;">

          <div style="background:#f1f4f8;padding:12px 16px;border-left:4px solid #2f6fed;border-radius:4px;margin-bottom:20px;color:#444;">
            {num_cases} case folder(s) uploaded and ready for processing
          </div>

          <div style="font-weight:600;margin-bottom:10px;color:#222;">Summary</div>
          <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">
            <tr>
              <td style="padding:10px 0;border-bottom:1px solid #f1f1f1;">Total Files</td>
              <td style="padding:10px 0;border-bottom:1px solid #f1f1f1;">{total_files}</td>
            </tr>
            <tr>
              <td style="padding:10px 0;border-bottom:1px solid #f1f1f1;">Total Pages</td>
              <td style="padding:10px 0;border-bottom:1px solid #f1f1f1;">{total_pages}</td>
            </tr>
            <tr>
              <td style="padding:10px 0;border-bottom:1px solid #f1f1f1;">Completed</td>
              <td style="padding:10px 0;border-bottom:1px solid #f1f1f1;">{ts}</td>
            </tr>
          </table>

          <div style="font-weight:600;margin-bottom:10px;color:#222;">Folders</div>
          <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">
            <tr>
              <th style="text-align:left;padding:10px 0;border-bottom:2px solid #e5e5e5;color:#444;">Case ID</th>
              <th style="text-align:left;padding:10px 0;border-bottom:2px solid #e5e5e5;color:#444;">Patient Name</th>
              <th style="text-align:left;padding:10px 0;border-bottom:2px solid #e5e5e5;color:#444;">Files</th>
              <th style="text-align:left;padding:10px 0;border-bottom:2px solid #e5e5e5;color:#444;">Pages</th>
              <th style="text-align:left;padding:10px 0;border-bottom:2px solid #e5e5e5;color:#444;">Status</th>
            </tr>
            {folder_rows}
          </table>

          {error_block}
          {button_block}

        </div>

        <div style="padding:15px 25px;font-size:12px;color:#888;background:#fafafa;">
          Automated notification from AI Tech Processing System
        </div>

      </div>
    </div>"""
    return html


def _send_sync_email(synced_files, error_count):
    """Send a consolidated HTML email summarizing all accumulated sync results."""
    if synced_files:
        mode = "[DRY-RUN] " if DRY_RUN else ""
        total_files = len(synced_files)
        subject = f"{mode}Case Files Ready — {total_files} files processed"
        html = _build_sync_html(synced_files, error_count)
        try:
            send_email(subject, html, html=True)
        except Exception as e:
            log(f"Failed to send summary email: {e}")

    elif error_count > 0:
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

    pending_files = []    # Accumulated sync results across poll cycles
    pending_errors = 0    # Accumulated error count

    while RUNNING:
        try:
            synced, errors = poll_and_sync(drive_id)
            pending_errors += errors

            if synced:
                # Files found — accumulate, deduplicate by s3_key
                existing_keys = {f["s3_key"] for f in pending_files}
                new_count = 0
                for f in synced:
                    if f["s3_key"] not in existing_keys:
                        pending_files.append(f)
                        existing_keys.add(f["s3_key"])
                        new_count += 1
                log(f"Accumulated {new_count} new files (total pending: {len(pending_files)})")
            elif pending_files:
                # No new files AND we have accumulated files — send consolidated email now
                log(f"Quiet poll detected. Sending consolidated email for {len(pending_files)} files...")
                _send_sync_email(pending_files, pending_errors)
                pending_files = []
                pending_errors = 0

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

    # Send any pending email before shutdown (don't lose sync results)
    if pending_files:
        log(f"Sending pending email before shutdown ({len(pending_files)} files)...")
        _send_sync_email(pending_files, pending_errors)

    log("LCP Sync Daemon stopped.")

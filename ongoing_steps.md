# Ongoing Steps — LCP Sync Changes

## Change 1: Case ID Regex Updated (supports all folder formats)

**Problem**: Old regex `^\d+_` only matched folders with underscore after digits. Real folders use multiple formats.

**Formats now supported**:

| SharePoint Folder Name | Extracted Case ID | Format |
|------------------------|-------------------|--------|
| `4846_LCP Adan Andrade` | `4846` | digits + underscore |
| `0009` | `0009` | digits only |
| `4843-Sheila Marie Tan Nadres` | `4843` | digits + hyphen |
| `4828_Ross Kip Hyams` | `4828` | digits + underscore |
| `483_Ofelia Moreno` | `483` | digits + underscore |

**Auto-skipped organizational folders** (won't be treated as case IDs):

| Folder Name | Why Skipped |
|-------------|-------------|
| `2026` | Year (4 digits in 2000-2099 range) |
| `03-05-2026` | Date pattern (MM-DD-YYYY) |
| `March` | Month name |
| `April_2025` | Month name prefix |

**Files changed**:
- `config.json` line 8: `"case_id_pattern"` changed from `"^(\\d+)_"` to `"^(\\d+)"`
- `lcp_sync.py` lines 175-240: `extract_case_info()` rewritten with `_is_organizational_folder()` helper

**To deploy**:
```bash
scp config.json lcp_sync.py ec2-user@<EC2_IP>:/home/ec2-user/lcp_report/
ssh ec2-user@<EC2_IP> "sudo systemctl restart lcp-sync"
```

---

## Change 2: Email only on actual file syncs (no spam)

**Problem**: Email was sent on every poll cycle including 0-file polls.

**Fix**: Email only fires when `synced_files` count > 0. Clean format shows folder name + count:

```
Subject: LCP Sync — 67 files synced

LCP Sync Summary
Time (UTC): 2026-03-05T06:55:00+00:00
Total files synced: 67

Folders:
  4846/ — 65 file(s)
  5317/ — 2 file(s)
```

**File changed**: `lcp_sync.py` lines 464-503

**Already applied** in current lcp_sync.py.

---

## Change 3: Batched email notifications (cooldown window)

**Problem**: Each 60-second poll cycle that finds files sends its own email. Uploading 1000 files over 10 minutes = ~10 separate emails flooding the inbox.

**Fix**: Accumulate sync results across poll cycles. Only send email when a "quiet" poll cycle occurs (0 new files after previous cycles had files). This means the email is sent ~60 seconds after the last file is detected.

| Scenario | Emails Before | Emails After |
|----------|--------------|-------------|
| 1000 files over 10 min | ~10 | 1 |
| 5 files in one drop | 1 | 1 |
| Continuous trickle all day | 1 per poll with files | 1 per quiet gap |
| Service restart mid-batch | email lost | email sent before shutdown |
| 0 files (empty poll) | 0 | 0 |

**How it works**:
1. `poll_and_sync()` now **returns** `(synced_files, error_count)` instead of sending email
2. Main loop accumulates results in `pending_files` list across poll cycles
3. When a poll finds 0 new files AND `pending_files` is non-empty → sends ONE consolidated email
4. On graceful shutdown (SIGTERM), sends pending email before exiting (no lost results)
5. New helper function `_send_sync_email()` handles the email formatting

**File changed**: `lcp_sync.py` — `poll_and_sync()` return value, new `_send_sync_email()`, main loop rewritten

---

## Change 4: HTML Email Redesign + Page Count Feature

**Problem**: Email was plain text with no visual styling. No page counts for uploaded files. No quick link to dashboard.

**Fix**: Three enhancements:

### 4a. HTML Email Template
- Styled like the "Report Generated Successfully" email (green header, gray boxes, tables)
- Title changed from "LCP Sync Summary" to **"Case Files Ready for Processing"**
- Subject line: `Case Files Ready — N files processed`
- Per-folder table with Case ID, file count, and page count
- Green "View Reports in Dashboard" button → `http://100.24.25.37:3005/msp/ocr-reports`
- `send_email()` now accepts `html=True` param; error/crash emails remain plain text

### 4b. Page Counting
- **PDF**: `PyPDF2.PdfReader` → exact page count
- **Word .docx**: `python-docx` core_properties.pages → best-effort (falls back to 1 if metadata missing)
- **Images** (.jpg, .png, .tiff, etc.): 1 page each
- **Other files**: 0 pages (counted as files only)
- Page counts shown per-folder in the email table and as a total

### 4c. Architecture Change
- Regular files now downloaded to memory first (like ZIPs), counted, then uploaded to S3
- Old `stream_to_s3()` still exists but replaced by `download_and_upload()` for new files
- `handle_zip_file()` now returns `(uploaded_keys, total_pages)`
- `synced_files` changed from list of strings to list of dicts: `{"s3_key", "case_id", "pages"}`

**Files changed**:
- `lcp_sync.py` — imports, `count_pages()`, `download_and_upload()`, `_build_sync_html()`, `_send_sync_email()`, `send_email()`, `poll_and_sync()`, `handle_zip_file()`
- `requirements.txt` — added `PyPDF2==3.0.1`, `python-docx==1.1.2`

**To deploy**:
```bash
scp lcp_sync.py requirements.txt ec2-user@<EC2_IP>:/home/ec2-user/lcp_report/
ssh ec2-user@<EC2_IP> "cd /home/ec2-user/lcp_report && source venv/bin/activate && pip install PyPDF2 python-docx"
ssh ec2-user@<EC2_IP> "sudo systemctl restart lcp-sync"
```

---

## Change 5: Fix DOCX Page Count + Deduplicate Files

**Problem 1**: `'CoreProperties' object has no attribute 'pages'` — `python-docx` doesn't have a `.pages` property.

**Fix 1**: Convert DOCX to PDF via LibreOffice headless (`/opt/libreoffice26.2/program/soffice --headless --convert-to pdf`), then count PDF pages with PyPDF2. This works 100% regardless of how the DOCX was created. Previous approaches failed: `python-docx` core_properties (no `.pages` attribute), `lastRenderedPageBreak` XML markers (most files don't have them), `docProps/app.xml` `<Pages>` (not present in non-Word files). Removed `python-docx` dependency — no longer needed. Also handles `.doc` files.

**Problem 2**: Same file appearing twice in email (e.g., "2 files" instead of 1). SharePoint delta returns the same file on consecutive polls due to metadata finalization after upload.

**Fix 2**: Deduplicate `pending_files` by `s3_key` when accumulating across poll cycles. If the same key appears again, skip it.

**File changed**: `lcp_sync.py` — `count_pages()` DOCX branch, main loop dedup logic

---

## How to deploy all changes at once

```bash
# From your local LCP_Reports folder:
scp lcp_sync.py config.json requirements.txt ec2-user@<EC2_IP>:/home/ec2-user/lcp_report/

# Install new dependencies
ssh ec2-user@<EC2_IP> "cd /home/ec2-user/lcp_report && source venv/bin/activate && pip install PyPDF2 python-docx"

# Restart service
ssh ec2-user@<EC2_IP> "sudo systemctl restart lcp-sync"

# Verify running
ssh ec2-user@<EC2_IP> "sudo systemctl status lcp-sync"

# Watch logs
ssh ec2-user@<EC2_IP> "tail -f /home/ec2-user/lcp_report/logs/lcp_sync.log"
```

## Test cases after deploy

1. Upload folder `4846_LCP Adan Andrade` with files → should create `4846/Input/` on S3
2. Upload folder `0009` with files → should create `0009/Input/` on S3
3. Upload folder `4843-Sheila Marie Tan Nadres` with files → should create `4843/Input/` on S3
4. Upload a ZIP file inside a case folder → should extract and upload individual files to Input/
5. Verify year folders (`2026`) and date folders (`03-05-2026`) are NOT treated as case IDs
6. Upload a batch of files over 2-3 minutes → should get exactly 1 email after uploads stop
7. Restart service mid-upload (`sudo systemctl restart lcp-sync`) → pending email should send before shutdown
8. Check HTML email has green header "Case Files Ready for Processing"
9. Check email shows per-folder table with file count + page count
10. Check "View Reports in Dashboard" button links to `http://100.24.25.37:3005/msp/ocr-reports`
11. Upload a PDF → verify correct page count in email
12. Upload an image → verify 1 page counted
13. Upload a .docx → verify page count (best-effort from metadata)

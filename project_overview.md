# LCP Reports — SharePoint Real-Time Sync to S3

## What This Project Does

This system monitors a specific SharePoint folder for new patient case files and automatically syncs them to an S3 bucket (`finallcpreports`) in near real-time (~1 minute).

When a new file is uploaded to SharePoint under a patient case folder (e.g., `3424_LCP_Tigran Grigoryan`), the system:
1. Detects the new file via Microsoft Graph delta queries (polling every 60 seconds)
2. Extracts the numeric case ID from the folder name (e.g., `3424`)
3. Creates the S3 folder structure: `3424/Input/`, `3424/Output/`, `3424/GroundTruth/`
4. Uploads the file to `s3://finallcpreports/3424/Input/{preserved_subfolder_path}/{filename}`
5. If the file is a ZIP, extracts it first and uploads individual files
6. Sends an email summary when files are synced

## Architecture

```
SharePoint Site: SpecialistMD2
  └── Documents / General / MSP Team / MSP Team Data / LCP Report Data
       └── {Year} / {Month} / {CaseID}_{PatientName} / files...

         ↓ (Microsoft Graph Delta Query — every 60 seconds)

EC2 Instance: /home/ec2-user/lcp_report/
  └── lcp_sync.py (systemd daemon)

         ↓ (Streaming upload — no local disk)

S3 Bucket: finallcpreports
  └── {case_id}/
       ├── Input/        ← patient files go here
       │    └── {subfolder_structure}/{filename}
       ├── Output/       ← empty (for downstream processing)
       └── GroundTruth/  ← empty (for downstream processing)
```

## SharePoint Folder Naming Patterns

All of these are valid case ID folder names:

| SharePoint Folder Name | Extracted Case ID | Format |
|------------------------|-------------------|--------|
| `4846_LCP Adan Andrade` | 4846 | digits + underscore |
| `0009` | 0009 | digits only |
| `4843-Sheila Marie Tan Nadres` | 4843 | digits + hyphen |
| `4828_Ross Kip Hyams` | 4828 | digits + underscore |
| `483_Ofelia Moreno` | 483 | digits + underscore |
| `1161_MCP_Maria Ortiz De Vargas` | 1161 | digits + underscore |

The regex pattern `^\d+` extracts all leading digits from the folder name.

Organizational folders are auto-skipped (NOT treated as case IDs):
- Year folders: `2026` (4 digits in 2000-2099)
- Date folders: `03-05-2026` (MM-DD-YYYY pattern)
- Month folders: `March`, `January_2025`, `April_2025`

## S3 Folder Structure

For each case ID, three folders are auto-created:
- `{case_id}/Input/` — files from SharePoint are uploaded here
- `{case_id}/Output/` — empty (created for downstream processing pipelines)
- `{case_id}/GroundTruth/` — empty (created for downstream processing pipelines)

Files inside the case folder on SharePoint **preserve their subfolder structure** in S3:

```
SharePoint: .../3424_LCP_Tigran/Medical Records/report.pdf
S3:         finallcpreports/3424/Input/Medical Records/report.pdf

SharePoint: .../3424_LCP_Tigran/Images/CT/scan1.jpg
S3:         finallcpreports/3424/Input/Images/CT/scan1.jpg
```

## Key Features

- **Real-time sync**: Files detected within ~60 seconds of upload
- **No files missed**: Delta queries guarantee every change is captured across poll cycles
- **ZIP extraction**: ZIP files are automatically extracted; individual files uploaded to Input/
- **Subfolder preservation**: Internal folder structure within case folders is maintained
- **API throttling protection**: Exponential backoff + Retry-After header handling for Graph API 429s
- **Graceful shutdown**: Handles SIGTERM/SIGINT for clean systemd stop/restart
- **Dry-run mode**: Test without uploading by setting `"dry_run": true` in config.json
- **Exclusion list**: Block specific folder names from being treated as case IDs
- **Batched HTML email alerts**: Results accumulate across polls; single styled email sent after uploads stop (~60s delay)
- **Page counting**: PDF (exact), DOCX (best-effort), images (1 each) — shown per-folder in email
- **Dashboard link**: "View Reports in Dashboard" button in email → LCP report generation tab
- **Atomic state saves**: Delta state written via tmp+rename to prevent corruption on crash

## What This Does NOT Do

- Does NOT upload anything to Output/ or GroundTruth/ folders
- Does NOT create a `pages/` folder inside Input/
- Does NOT interfere with the existing daily SharePoint backup system
- Does NOT store files on local disk (streams directly from SharePoint to S3)

## Credentials

| Service | Credential Source | Shared With Daily Backup? |
|---------|-------------------|---------------------------|
| Microsoft Graph API | Same Azure app registration (tenant_id, client_id, client_secret) | Yes |
| AWS S3 (finallcpreports) | **Separate** AWS access key / secret key | No — different bucket, different creds |
| Email (Graph sendMail) | Same sender & recipients | Yes |

## Files in This Project

| File | Purpose |
|------|---------|
| `lcp_sync.py` | Main daemon script (~300 lines) |
| `config.json` | Configuration: site ID, monitor path, bucket, polling interval, dry_run flag |
| `secrets.json` | Credentials: Azure, AWS (for finallcpreports), email |
| `requirements.txt` | Python dependencies (msal, boto3, requests) |
| `run_lcp_sync.sh` | Shell wrapper for manual execution with logging |
| `lcp-sync.service` | systemd unit file for running as a service |
| `lcp_sync_state.json` | Auto-created: stores delta token for incremental sync |
| `logs/` | Log files directory |

## Delta Query Isolation

This project uses **folder-scoped delta queries** — the delta API call targets only the `LCP Report Data` folder, NOT the entire SharePoint drive. This means:
- The `lcp_sync_state.json` file is small (only tracks files inside LCP Report Data)
- Polls are fast (fewer items to process per cycle)
- Completely independent from the daily backup's `delta_state.json` (which tracks 13,000+ files across the full drive)

## EC2 Projects

```
/home/ec2-user/
├── mapping_table_ui/              # Existing project
├── sharepoint_s3_backup_final/    # Daily SharePoint backup (5 AM UTC cron)
└── lcp_report/                    # THIS PROJECT (systemd daemon, 60s polling)
```

## Related Systems

- **Daily SharePoint Backup**: `/home/ec2-user/sharepoint_s3_backup_final/` — runs at 5 AM UTC daily, backs up to `medxprts-report-data-backup` bucket. Completely independent. Uses its own `delta_state.json` (full drive scope).

## Contact

- Email alerts go to: Tech_Alert@QGUCMSO.com, raj@QGUCMSO.com
- Sender: AI_Tech_Team@specialistMDLCP.com

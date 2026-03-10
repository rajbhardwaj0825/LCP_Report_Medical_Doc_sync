# Deployment & Testing Steps

## IMPORTANT: Isolation Guarantee

This project is 100% self-contained in `/home/ec2-user/lcp_report/`.
It does NOT touch, modify, or depend on the other EC2 projects:
- `/home/ec2-user/mapping_table_ui/` — untouched
- `/home/ec2-user/sharepoint_s3_backup_final/` — untouched

All commands below operate ONLY inside `/home/ec2-user/lcp_report/`.
The systemd service, venv, state file, logs, and config are all isolated.

## Prerequisites

- EC2 instance (same one running the daily SharePoint backup)
- Python 3.9+ installed
- AWS credentials for the `finallcpreports` S3 bucket (different from daily backup creds)
- SSH access to the EC2 instance

---

## Step 1: Upload Project Files to EC2

From your local machine, SCP the project files:

```bash
# Create the project directory on EC2
ssh ec2-user@<EC2_IP> "mkdir -p /home/ec2-user/lcp_report/logs"

# Upload all files
scp lcp_sync.py config.json secrets.json requirements.txt run_lcp_sync.sh lcp-sync.service \
    ec2-user@<EC2_IP>:/home/ec2-user/lcp_report/
```

---

## Step 2: Update secrets.json with Real AWS Credentials

SSH into EC2 and edit secrets.json:

```bash
ssh ec2-user@<EC2_IP>
cd /home/ec2-user/lcp_report
nano secrets.json
```

Replace the placeholder values:
```json
{
  "aws": {
    "access_key": "YOUR_ACTUAL_ACCESS_KEY_FOR_FINALLCPREPORTS",
    "secret_key": "YOUR_ACTUAL_SECRET_KEY_FOR_FINALLCPREPORTS",
    "region": "us-east-1"
  }
}
```

Save and exit.

---

## Step 3: Set Up Python Virtual Environment

```bash
cd /home/ec2-user/lcp_report
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---

## Step 4: First Run — Establish Delta Baseline

The first run consumes the full SharePoint delta WITHOUT uploading anything.
This is necessary to establish the baseline so future polls only catch new files.

```bash
cd /home/ec2-user/lcp_report
source venv/bin/activate
python lcp_sync.py
```

You should see output like:
```
... | LCP Sync Daemon Starting
... | Monitoring: /General/MSP Team/MSP Team Data/LCP Report Data
... | S3 Bucket: finallcpreports
... | Dry Run: True
... | First Run: True
... | Drive ID resolved: b!y57...
... | Fetching delta page 1...
... | Fetching delta page 2...
... | First run complete — delta baseline consumed. No files uploaded.
... | Sleeping 60s until next poll...
```

Press `Ctrl+C` after you see "First run complete". The delta baseline is now saved in `lcp_sync_state.json`.

---

## Step 5: Test in Dry-Run Mode

Config has `"dry_run": true` by default. Start the daemon:

```bash
python lcp_sync.py
```

Now go to SharePoint and upload a test file:
1. Navigate to: Documents > General > MSP Team > MSP Team Data > LCP Report Data
2. Open any existing case folder (e.g., `3424_LCP_Tigran`) or create a new one
3. Upload a small test file (PDF, image, or document)

Within ~60 seconds you should see in the logs:
```
... | DRY-RUN: case_id=3424 | .../3424_LCP_Tigran/test.pdf -> s3://finallcpreports/3424/Input/test.pdf
```

This confirms the system detects the file and knows where to put it, but does NOT actually upload.

Press `Ctrl+C` to stop.

---

## Step 6: Go Live — Disable Dry Run

Edit config.json:
```bash
nano config.json
```

Change `"dry_run": true` to `"dry_run": false`. Save.

---

## Step 7: Test Real Upload

```bash
python lcp_sync.py
```

Upload another test file to SharePoint. Within ~60 seconds:
1. Check the logs — should show `SYNC: case_id=XXXX | filename (X.XX MB)`
2. Check S3 — go to `finallcpreports` bucket in AWS Console
3. Navigate to `{case_id}/Input/` — your file should be there
4. Verify `Output/` and `GroundTruth/` folders were also created (empty)
5. Check your email — summary email should arrive

Press `Ctrl+C` to stop.

---

## Step 8: Test ZIP File Upload

1. Create a small ZIP file with 2-3 test files inside
2. Upload the ZIP to a case folder on SharePoint
3. Within ~60 seconds, check S3 — individual extracted files should appear in Input/ (not the ZIP itself)

---

## Step 9: Test Multi-File Upload (No Files Missed)

1. Upload 5-10 files to a case folder on SharePoint in quick succession
2. Wait 2-3 minutes (2-3 poll cycles)
3. Count files in S3 under `{case_id}/Input/` — should match exactly what you uploaded
4. No files should be missing

---

## Step 10: Install as systemd Service

```bash
# Copy service file
sudo cp /home/ec2-user/lcp_report/lcp-sync.service /etc/systemd/system/

# Reload systemd
sudo systemctl daemon-reload

# Enable auto-start on boot
sudo systemctl enable lcp-sync

# Start the service
sudo systemctl start lcp-sync

# Check status
sudo systemctl status lcp-sync
```

---

## Step 11: Verify Service is Running

```bash
# Check service status
sudo systemctl status lcp-sync

# Watch logs in real-time
tail -f /home/ec2-user/lcp_report/logs/lcp_sync.log

# Check recent logs
tail -100 /home/ec2-user/lcp_report/logs/lcp_sync.log
```

---

## Common Operations

### Stop the service
```bash
sudo systemctl stop lcp-sync
```

### Restart the service
```bash
sudo systemctl restart lcp-sync
```

### View logs
```bash
tail -f /home/ec2-user/lcp_report/logs/lcp_sync.log
# Or via journald:
sudo journalctl -u lcp-sync -f
```

### Add a folder to the exclusion list
Edit config.json and add the folder name:
```json
"excluded_folders": ["2025_Budget", "100_Templates"]
```
Then restart the service:
```bash
sudo systemctl restart lcp-sync
```

### Reset delta state (re-baseline)
If you need to start fresh:
```bash
sudo systemctl stop lcp-sync
rm /home/ec2-user/lcp_report/lcp_sync_state.json
sudo systemctl start lcp-sync
# First poll will consume baseline again without uploading
```

---

## Troubleshooting

### "Graph auth failed" error
- Check that `client_secret` in secrets.json hasn't expired
- Verify `tenant_id` and `client_id` are correct
- Azure app registration may need renewed credentials

### "No drives found" or "Library not found" error
- Verify `site_id` in config.json matches the SharePoint site
- Check that the Azure app has Sites.Read.All permission

### Graph API 429 (throttled) errors
- The script automatically handles these with Retry-After headers
- If persistent, increase `GRAPH_THROTTLE` in lcp_sync.py (default: 0.5s)

### Files not appearing in S3
- Check logs for SKIP/ERROR messages
- Verify the file is inside a folder matching `^\d+_` pattern
- Verify the path contains `/General/MSP Team/MSP Team Data/LCP Report Data`
- Ensure `dry_run` is `false` in config.json

### Service crashes and restarts
- The systemd service auto-restarts after 30 seconds
- Check logs: `sudo journalctl -u lcp-sync --since "1 hour ago"`
- Delta state is saved after each page, so no data is lost on crash

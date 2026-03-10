#!/bin/bash

PROJECT_DIR="/home/ec2-user/lcp_report"
VENV_DIR="$PROJECT_DIR/venv"
LOG_DIR="$PROJECT_DIR/logs"

mkdir -p "$LOG_DIR"

TIMESTAMP=$(date -u +"%Y-%m-%d_%H-%M-%S")
LOG_FILE="$LOG_DIR/lcp_sync_$TIMESTAMP.log"

cd "$PROJECT_DIR" || exit 1
source "$VENV_DIR/bin/activate"

python lcp_sync.py >> "$LOG_FILE" 2>&1

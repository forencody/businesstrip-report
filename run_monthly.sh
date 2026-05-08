#!/bin/bash
# 每月自動執行差旅報表產生器
# 由 LaunchAgent 在每月 1 日呼叫

SCRIPT_DIR="/Users/mac/Library/Mobile Documents/com~apple~CloudDocs/Documents/Claude/Projects/出差申報進化"
LOG_FILE="$SCRIPT_DIR/logs/run.log"

mkdir -p "$SCRIPT_DIR/logs"

echo "==============================" >> "$LOG_FILE"
echo "執行時間: $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG_FILE"

/usr/bin/python3 "$SCRIPT_DIR/出差申報自動化.py" >> "$LOG_FILE" 2>&1

echo "結束時間: $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG_FILE"
echo "" >> "$LOG_FILE"

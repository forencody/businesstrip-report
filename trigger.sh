#!/bin/bash
# 觸發 GitHub Actions 產生差旅報表
# 用法：
#   ./trigger.sh          → 跑上個月
#   ./trigger.sh 2026 4   → 跑指定年月

REPO="forencody/businesstrip-report"

if [ "$#" -eq 2 ]; then
    YEAR=$1
    MONTH=$2
    echo "▶ 觸發 ${YEAR}年${MONTH}月 差旅報表..."
    gh workflow run monthly-report.yml --repo "$REPO" -f year="$YEAR" -f month="$MONTH"
else
    echo "▶ 觸發上個月差旅報表..."
    gh workflow run monthly-report.yml --repo "$REPO"
fi

echo ""
echo "⏳ 等待執行（約 30 秒）..."
sleep 5

RUN_ID=$(gh run list --workflow="monthly-report.yml" --repo "$REPO" --limit 1 --json databaseId -q '.[0].databaseId')
echo "📋 Run ID: $RUN_ID"
echo ""

gh run watch "$RUN_ID" --repo "$REPO"

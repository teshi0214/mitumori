#!/bin/bash
# ============================================================
# Mitumori Agent - Agent Engine デプロイスクリプト
# 使い方: BQ_remote_Ver2/ 直下で ./deploy_mitumori.sh を実行
# ============================================================
set -e

export PATH="$PATH:/opt/homebrew/bin"

# ===== 設定 =====
PROJECT_ID="geminni-dev"
REGION="us-central1"
DISPLAY_NAME="Mitumori_Agent"
STAGING_BUCKET="gs://geminni-dev-adk-staging"
# ================

AGENT_DIR="$(cd "$(dirname "$0")/mitumori" && pwd)"

# adk コマンドを探す
ADK="$(which adk 2>/dev/null || echo '')"
if [ -z "$ADK" ]; then
    for candidate in \
        "$HOME/.pyenv/versions/3.11.9/bin/adk" \
        "$HOME/.local/bin/adk" \
        "/opt/homebrew/bin/adk" \
        "$(dirname "$0")/.venv/bin/adk"; do
        if [ -f "$candidate" ]; then
            ADK="$candidate"
            break
        fi
    done
fi

if [ -z "$ADK" ]; then
    echo "❌ adk コマンドが見つかりません。pip install google-adk を実行してください。"
    exit 1
fi

echo "===================================="
echo " Mitumori Agent - Agent Engine デプロイ"
echo "===================================="
echo "プロジェクト : $PROJECT_ID"
echo "リージョン   : $REGION"
echo "エージェント : $AGENT_DIR"
echo "ADK          : $ADK"
echo ""

# ステージングバケット確認・作成（PMC Agentと共用）
echo "🪣 ステージングバケット確認中..."
if ! gcloud storage buckets describe $STAGING_BUCKET \
        --project=$PROJECT_ID > /dev/null 2>&1; then
    echo "   バケットを作成します..."
    gcloud storage buckets create $STAGING_BUCKET \
        --project=$PROJECT_ID \
        --location=$REGION
    echo "   ✅ バケット作成完了"
else
    echo "   ✅ バケット既存（共用）"
fi

# デプロイ実行
echo ""
echo "🚀 デプロイ開始（5〜10分かかります）..."
$ADK deploy agent_engine \
    --project=$PROJECT_ID \
    --region=$REGION \
    --display_name=$DISPLAY_NAME \
    --staging_bucket=$STAGING_BUCKET \
    $AGENT_DIR

echo ""
echo "===================================="
echo "✅ デプロイ完了！"
echo "===================================="
echo ""
echo "▼ Engine ID 確認:"
echo "  gcloud ai reasoning-engines list --project=$PROJECT_ID --region=$REGION"
echo ""
echo "▼ Agent Engine SAへの権限付与（初回デプロイ後に必要）:"
echo "  プロジェクト番号を確認 →"
echo "  gcloud projects describe $PROJECT_ID --format='value(projectNumber)'"
echo "  SA: service-{番号}@gcp-sa-aiplatform-re.iam.gserviceaccount.com"
echo "  付与ロール: roles/aiplatform.user / roles/serviceusage.serviceUsageConsumer"
echo ""
echo "▼ ログ確認（エラー時）:"
echo "  gcloud logging read 'resource.type=\"aiplatform.googleapis.com/ReasoningEngine\" AND severity>=ERROR' \\"
echo "    --project=$PROJECT_ID --limit=10 --format='value(textPayload)'"

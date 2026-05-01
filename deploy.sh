#!/usr/bin/env bash
# deploy.sh — Build and deploy the Gavel Generator web app to Cloud Run
#
# Prerequisites:
#   gcloud CLI installed and authenticated
#   Google Cloud project with these APIs enabled:
#     Cloud Run, Cloud Build, Firestore, Secret Manager (optional)
#
# Usage:
#   chmod +x deploy.sh
#   ./deploy.sh

set -euo pipefail

# Load secrets from .env if it exists
if [ -f .env ]; then
  set -a
  source .env
  set +a
fi

# ── Config — edit these ────────────────────────────────────────────────────────
PROJECT_ID="maindb-79403"
REGION="us-central1"
SERVICE_NAME="gavel-generator"
IMAGE="gcr.io/${PROJECT_ID}/${SERVICE_NAME}"

# Secrets — store in .env or pass as env vars before running
: "${SHIPSTATION_API_KEY:?Need SHIPSTATION_API_KEY}"
: "${SHIPSTATION_API_SECRET:?Need SHIPSTATION_API_SECRET}"
: "${TRELLO_API_KEY:?Need TRELLO_API_KEY}"
: "${TRELLO_TOKEN:?Need TRELLO_TOKEN}"
: "${CRON_TOKEN:?Need CRON_TOKEN}"
# ──────────────────────────────────────────────────────────────────────────────

echo "▶ Building container image…"
gcloud builds submit \
  --project="${PROJECT_ID}" \
  --tag="${IMAGE}" \
  .

echo "▶ Deploying to Cloud Run…"
gcloud run deploy "${SERVICE_NAME}" \
  --project="${PROJECT_ID}" \
  --image="${IMAGE}" \
  --region="${REGION}" \
  --platform=managed \
  --allow-unauthenticated \
  --min-instances=1 \
  --max-instances=1 \
  --memory=512Mi \
  --cpu=1 \
  --timeout=3600 \
  --set-env-vars="\
SHIPSTATION_API_KEY=${SHIPSTATION_API_KEY},\
SHIPSTATION_API_SECRET=${SHIPSTATION_API_SECRET},\
TRELLO_API_KEY=${TRELLO_API_KEY},\
TRELLO_TOKEN=${TRELLO_TOKEN},\
CRON_TOKEN=${CRON_TOKEN},\
GOOGLE_CLOUD_PROJECT=${PROJECT_ID}"

SERVICE_URL=$(gcloud run services describe "${SERVICE_NAME}" \
  --project="${PROJECT_ID}" \
  --region="${REGION}" \
  --format='value(status.url)')

echo ""
echo "✓ Deployed: ${SERVICE_URL}"
echo ""
echo "── Next steps ────────────────────────────────────────────────────────────"
echo "1. (Optional) Add a Cloud Scheduler job for an extra reliability layer:"
echo "   gcloud scheduler jobs create http gavel-cron \\"
echo "     --location=${REGION} \\"
echo "     --schedule='0 8 * * 1-5' \\"
echo "     --uri=${SERVICE_URL}/cron \\"
echo "     --message-body='{}' \\"
echo "     --headers='X-Cron-Token=${CRON_TOKEN},Content-Type=application/json'"
echo ""
echo "3. Open the dashboard: ${SERVICE_URL}"

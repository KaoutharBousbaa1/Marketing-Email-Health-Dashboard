#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT_DIR"

echo "[1/11] analysis_lead_magnet.py"
python3 analysis_lead_magnet.py

echo "[2/11] analysis_group_ab_monthly.py"
python3 analysis_group_ab_monthly.py

echo "[3/11] analysis_lead_magnet_origin_post.py"
python3 analysis_lead_magnet_origin_post.py

echo "[4/11] analysis_signup_cohort_evolution.py"
python3 analysis_signup_cohort_evolution.py

echo "[5/11] analysis_bootcamp_buyers.py"
python3 analysis_bootcamp_buyers.py

echo "[6/11] fetch_cold_engagement.py"
python3 fetch_cold_engagement.py

echo "[7/11] analysis_cold_engagement.py"
python3 analysis_cold_engagement.py

echo "[8/11] analysis_cold_subscribers.py"
python3 analysis_cold_subscribers.py

echo "[9/11] analysis_sales_vs_value.py"
python3 analysis_sales_vs_value.py

echo "[10/11] analysis_phase_lifespan.py"
python3 analysis_phase_lifespan.py

echo "[11/11] analysis_campaign_math_check.py"
python3 analysis_campaign_math_check.py

echo "Done. Outputs are ready in charts/ and data/generated/."

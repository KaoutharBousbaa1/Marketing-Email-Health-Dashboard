# Email Health Analysis

This repository contains the analysis scripts and CSV data used for the LO email health audit.

## Can my teammate run this after cloning?
Yes. If they have Python 3.10+ and internet access, they can run everything from this folder.

## 1) Setup (one time)

```bash
git clone <your-private-repo-url>
cd "Email Health"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 2) Run everything

### Option A (easiest)

```bash
./run_all.sh
```

This runs the full analysis pipeline in the correct order.

### Option B (step-by-step)

```bash
python3 analysis_lead_magnet.py
python3 analysis_group_ab_monthly.py
python3 analysis_lead_magnet_origin_post.py
python3 analysis_signup_cohort_evolution.py
python3 analysis_bootcamp_buyers.py
python3 fetch_cold_engagement.py
python3 analysis_cold_engagement.py
python3 analysis_cold_subscribers.py
python3 analysis_sales_vs_value.py
python3 analysis_phase_lifespan.py
python3 analysis_campaign_math_check.py
```

This team workflow intentionally **does not** generate a Word report.

## 3) Input files (must stay in this folder)
- `Confirmed Subscribers.csv`
- `Cold Subscribers.csv`
- `Emails Broadcasting - broadcasts_categorised.csv`
- `AI Sprint Roadmap - Opened subscribers.csv`
- `AI Sprint Roadmap - Clicked.csv`
- `Workshops Pre-bootcamps - Sheet1.csv`
- `export (22).csv`

Do not rename these files, because scripts read them by exact name.

## 4) Important outputs generated
- Lead magnet outputs:
  - `data/generated/lead_magnet_stats.csv`
  - `data/generated/lead_magnet_post_survey_events.csv`
  - `data/generated/lead_magnet_post_survey_summary.csv`
  - `data/generated/lead_magnet_origin_post_survey_summary.csv`
  - `data/generated/lead_magnet_prepost_window_summary.csv`
  - `data/generated/lead_magnet_postwindow_nonlead_summary.csv`
- Cohort/buyer outputs:
  - `data/generated/lead_magnet_group_ab_monthly.csv`
  - `data/generated/signup_cohort_monthly_evolution.csv`
  - `data/generated/bootcamp_buyers_monthly_evolution.csv`
  - `data/generated/bootcamp_buyers_summary.csv`
- Campaign math checks:
  - `data/generated/campaign_math_check_summary.csv`
- Cold subscriber outputs:
  - `data/generated/cold_engagement.csv`
- Charts:
  - `charts/*.png`

Your teammate will get:
- Visual outputs in `charts/`
- Analysis tables/metrics in `data/generated/`
- No `.docx` report is produced in the standard run

## 5) Kit API key
API scripts already include the Kit API key in code for direct internal team use (private repo).
If the key changes, update `API_KEY` in:
- `analysis_lead_magnet.py`
- `analysis_group_ab_monthly.py`
- `analysis_lead_magnet_origin_post.py`
- `analysis_signup_cohort_evolution.py`
- `analysis_bootcamp_buyers.py`
- `fetch_cold_engagement.py`

## 6) Troubleshooting
- If a script fails with API/network errors, rerun the same script (Kit rate limits or temporary API errors can happen).
- If a script says a CSV is missing, confirm the file is present in this folder with the exact same name.
- If matplotlib font/cache warnings appear, they are usually non-blocking.

## 7) Environment notes
- Scripts use relative paths and can run from any machine.
- Recommended Python version: 3.10 or newer.

## 8) Python file catalog

| Python file | What it does | Main input CSVs | Main outputs |
|---|---|---|---|
| `analysis_lead_magnet.py` | Lead magnet core analysis (pre/post opens + Group A vs B engagement) | `Emails Broadcasting - broadcasts_categorised.csv`, `AI Sprint Roadmap - Opened subscribers.csv`, `AI Sprint Roadmap - Clicked.csv` | `data/generated/lead_magnet_stats.csv`, `data/generated/lead_magnet_post_survey_events.csv`, `data/generated/lead_magnet_post_survey_summary.csv`, lead-magnet charts |
| `analysis_group_ab_monthly.py` | Monthly evolution of Group A vs B (Value OR, Sales OR, Sales CTOR) | `AI Sprint Roadmap - Opened subscribers.csv`, `AI Sprint Roadmap - Clicked.csv`, `Emails Broadcasting - broadcasts_categorised.csv` | `data/generated/lead_magnet_broadcast_mapping.csv`, `data/generated/lead_magnet_group_ab_monthly.csv`, charts `AC/AD/AE` |
| `analysis_lead_magnet_origin_post.py` | Origin analysis: strict lead-magnet signups vs other cohorts | `Confirmed Subscribers.csv`, `data/generated/lead_magnet_post_survey_events.csv`, `data/generated/lead_magnet_broadcast_mapping.csv` | `data/generated/lead_magnet_origin_post_survey_events.csv`, `data/generated/lead_magnet_origin_post_survey_summary.csv`, `data/generated/lead_magnet_prepost_window_summary.csv`, `data/generated/lead_magnet_postwindow_nonlead_summary.csv`, chart `AD2` |
| `analysis_signup_cohort_evolution.py` | Open/CTOR evolution by signup-age cohorts (old vs new) | `data/generated/all_subscribers_created_at.csv` (cache), `data/generated/lead_magnet_broadcast_mapping.csv` | `data/generated/signup_cohort_monthly_evolution.csv`, `data/generated/signup_cohort_failed_ids.csv`, charts `AL/AM/AN` |
| `analysis_bootcamp_buyers.py` | Buyers vs non-buyers over time (OR/CTOR + age profile) | `Confirmed Subscribers.csv`, `data/generated/lead_magnet_broadcast_mapping.csv` | `data/generated/bootcamp_buyers_monthly_evolution.csv`, `data/generated/bootcamp_buyers_summary.csv`, `data/generated/bootcamp_buyers_age_buckets.csv`, `data/generated/bootcamp_buyers_failed_ids.csv`, charts `AO/AP` |
| `fetch_cold_engagement.py` | Calls Kit API to enrich cold subscribers with engagement stats | `Cold Subscribers.csv` | `data/generated/cold_engagement.csv` |
| `analysis_cold_engagement.py` | Analysis/charts of cold-subscriber engagement behavior | `data/generated/cold_engagement.csv` | cold engagement charts |
| `analysis_cold_subscribers.py` | Analysis/charts of cold-subscriber age/source composition | `Cold Subscribers.csv` | cold subscriber charts |
| `analysis_sales_vs_value.py` | Sales vs Value category trend analysis | `Emails Broadcasting - broadcasts_categorised.csv` | category trend charts |
| `analysis_phase_lifespan.py` | Unsubscribe lifespan by phase analysis | `export (22).csv` | phase-lifespan charts |
| `analysis_campaign_math_check.py` | Verifies denominator/numerator math for campaign claims (ex: recipients vs roadmap-tagged subset) | `Emails Broadcasting - broadcasts_categorised.csv`, `Confirmed Subscribers.csv` | `data/generated/campaign_math_check_summary.csv` |

## 9) CSV file catalog

### Raw input CSVs

| CSV file | What it includes | Used by script(s) / task |
|---|---|---|
| `AI Sprint Roadmap - Clicked.csv` | Subscribers who clicked the lead-magnet email | `analysis_lead_magnet.py`, `analysis_group_ab_monthly.py` |
| `AI Sprint Roadmap - Opened subscribers.csv` | Subscribers who opened the lead-magnet email | `analysis_lead_magnet.py`, `analysis_group_ab_monthly.py` |
| `Cold Subscribers.csv` | Subscribers identified as cold/inactive | `fetch_cold_engagement.py`, `analysis_cold_subscribers.py` |
| `Confirmed Subscribers.csv` | Full confirmed subscriber list (email, tags, created date, referrer, etc.) | `analysis_bootcamp_buyers.py`, `analysis_lead_magnet_origin_post.py` |
| `Emails Broadcasting - broadcasts_categorised.csv` | Broadcast-level email performance with Value/Sales category labels | `analysis_lead_magnet.py`, `analysis_group_ab_monthly.py`, `analysis_sales_vs_value.py` |
| `Workshops Pre-bootcamps - Sheet1.csv` | Workshop-to-bootcamp mapping tags by cohort | Not used in the default team pipeline (kept as reference data) |
| `export (22).csv` | Subscriber lifecycle export used for churn/lifespan phase analysis | `analysis_phase_lifespan.py` |

### Generated / intermediate CSVs

| CSV file | What it includes | Produced by | Used by |
|---|---|---|---|
| `data/generated/all_subscribers_created_at.csv` | Subscriber cache from Kit API (id, email, status, created date) | `analysis_signup_cohort_evolution.py` | `analysis_signup_cohort_evolution.py` |
| `data/generated/bootcamp_buyers_age_buckets.csv` | Buyer/non-buyer age-bucket counts and shares | `analysis_bootcamp_buyers.py` | team review / downstream analysis |
| `data/generated/bootcamp_buyers_failed_ids.csv` | Broadcast IDs that failed API filtering in buyer analysis | `analysis_bootcamp_buyers.py` | team review / QA |
| `data/generated/bootcamp_buyers_monthly_evolution.csv` | Monthly OR/CTOR by buyer vs non-buyer segment | `analysis_bootcamp_buyers.py` | team review / downstream analysis |
| `data/generated/bootcamp_buyers_summary.csv` | Overall buyer vs non-buyer benchmark metrics | `analysis_bootcamp_buyers.py` | team review / downstream analysis |
| `data/generated/bootcamp_conversion_latency.csv` | Days from signup to first bootcamp purchase per buyer | legacy report build | reference only |
| `data/generated/bootcamp_conversion_latency_by_cohort.csv` | Conversion latency aggregated by first purchase cohort | legacy report build | reference only |
| `data/generated/campaign_math_check_summary.csv` | Campaign denominator/numerator checks for the bootcamp invite vs roadmap-tag subsets | `analysis_campaign_math_check.py` | team review / QA |
| `data/generated/cold_engagement.csv` | Cold subscriber engagement stats from Kit API | `fetch_cold_engagement.py` | `analysis_cold_engagement.py` |
| `data/generated/failed_broadcasts_filter500.csv` | Diagnostic list of broadcasts with API filter 500 failures | manual/diagnostic artifact | reference only |
| `data/generated/lead_magnet_broadcast_mapping.csv` | Mapped broadcast IDs for categorized emails | `analysis_group_ab_monthly.py` | `analysis_signup_cohort_evolution.py`, `analysis_bootcamp_buyers.py`, `analysis_lead_magnet_origin_post.py` |
| `data/generated/lead_magnet_group_ab_monthly.csv` | Monthly Value OR / Sales OR / Sales CTOR for Group A vs B | `analysis_group_ab_monthly.py` | team review / downstream analysis |
| `data/generated/lead_magnet_origin_post_survey_events.csv` | Per-broadcast event table for strict lead-magnet origin analysis | `analysis_lead_magnet_origin_post.py` | reference / QA |
| `data/generated/lead_magnet_origin_post_survey_summary.csv` | Summary metrics for lead-magnet-origin vs pre-signup cohort (same post broadcasts) | `analysis_lead_magnet_origin_post.py` | team review / downstream analysis |
| `data/generated/lead_magnet_post_survey_events.csv` | Post-survey matched broadcast event table | `analysis_lead_magnet.py` | `analysis_lead_magnet_origin_post.py` |
| `data/generated/lead_magnet_post_survey_summary.csv` | Group A vs B post-survey email-type summary | `analysis_lead_magnet.py` | team review / downstream analysis |
| `data/generated/lead_magnet_postwindow_nonlead_summary.csv` | Lead-magnet signups vs non-lead signups in same post window | `analysis_lead_magnet_origin_post.py` | team review / downstream analysis |
| `data/generated/lead_magnet_prepost_window_summary.csv` | As-received comparison: pre-window cohort vs post-window lead cohort | `analysis_lead_magnet_origin_post.py` | team review / downstream analysis |
| `data/generated/lead_magnet_stats.csv` | Subscriber-level enriched stats for Group A/B analysis | `analysis_lead_magnet.py` | `analysis_lead_magnet.py` |
| `data/generated/sales_intent_outcomes_per_email.csv` | Sales intent/action diagnostics per sales broadcast | legacy report build | reference only |
| `data/generated/signup_cohort_failed_ids.csv` | Failed broadcast IDs in signup cohort analysis | `analysis_signup_cohort_evolution.py` | team review / QA |
| `data/generated/signup_cohort_monthly_evolution.csv` | Monthly OR/CTOR by signup cohort (Q1–Q4) | `analysis_signup_cohort_evolution.py` | team review / downstream analysis |

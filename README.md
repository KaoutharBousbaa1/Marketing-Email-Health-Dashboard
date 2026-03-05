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
```

Optional report generation:

```bash
python3 generate_report_v3.py
```

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
  - `lead_magnet_stats.csv`
  - `lead_magnet_post_survey_events.csv`
  - `lead_magnet_post_survey_summary.csv`
  - `lead_magnet_origin_post_survey_summary.csv`
  - `lead_magnet_prepost_window_summary.csv`
  - `lead_magnet_postwindow_nonlead_summary.csv`
- Cohort/buyer outputs:
  - `lead_magnet_group_ab_monthly.csv`
  - `signup_cohort_monthly_evolution.csv`
  - `bootcamp_buyers_monthly_evolution.csv`
  - `bootcamp_buyers_summary.csv`
- Cold subscriber outputs:
  - `cold_engagement.csv`
- Charts:
  - `charts/*.png`

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

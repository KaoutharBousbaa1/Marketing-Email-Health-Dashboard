# Streamlit Dashboard KPI Spec (Kit API Real-Time)

## 1) Goal
Build a simple dashboard for the marketing team that answers:
- Are we getting weaker or stronger this week/month?
- Which subscriber groups are still engaged?
- Which acquisition and conversion paths are working?

This spec is intentionally focused on the highest-impact KPIs only.

## 2) Core KPI Set (Most Important Only)

### A. Health Now (list-level)
1. **Value Open Rate (last 30 days)**
- Formula: `sum(value_emails_opened) / sum(value_recipients)`
- Why: best top-of-funnel attention metric for content emails.

2. **Sales Open Rate (last 30 days)**
- Formula: `sum(sales_emails_opened) / sum(sales_recipients)`
- Why: measures demand generation for commercial sends.

3. **Sales CTOR (last 30 days)**
- Formula: `sum(sales_clicks) / sum(sales_opens)`
- Why: measures click intent after open (quality of message + CTA).

4. **Cold Subscriber Share**
- Formula: `cold_subscribers / total_active_subscribers`
- Why: shows audience quality drag and deliverability risk.

### B. Segment Quality (who is engaging)
5. **Lead-Magnet Strict Signups (post window)**
- Definition: subscribers with roadmap-origin signup + valid survey response signal.
- Why: tells if lead magnet is adding qualified people.

6. **Lead vs Non-Lead Sales OR Gap (same window)**
- Formula: `sales_or_lead - sales_or_nonlead`
- Why: checks quality of lead-magnet subscribers vs regular new signups.

7. **Lead vs Non-Lead Sales CTOR Gap (same window)**
- Formula: `sales_ctor_lead - sales_ctor_nonlead`
- Why: checks if leads click better after opening.

8. **Oldest vs Newest Open Rate Gap**
- Formula: `open_rate_oldest_signup_cohort - open_rate_newest_signup_cohort`
- Why: confirms/disproves age-based disengagement.

9. **Oldest vs Newest Sales CTOR Gap**
- Formula: `sales_ctor_oldest_signup_cohort - sales_ctor_newest_signup_cohort`
- Why: confirms whether older subscribers still click less even when they open.

### C. Conversion Path (money signals)
10. **Workshop -> Bootcamp Conversion Rate**
- Formula: `# workshop buyers who bought bootcamp / # workshop buyers`
- Show baseline next to it: `# non-workshop subscribers who bought bootcamp / # non-workshop subscribers`
- Why: validates workshop as pre-bootcamp feeder.

11. **Share of Bootcamp Buyers with Prior Workshop**
- Formula: `# bootcamp buyers with workshop before purchase / # total bootcamp buyers`
- Why: indicates how much bootcamp demand is workshop-assisted.

12. **Median Days: Signup -> First Bootcamp Purchase**
- Formula: median of `(first_bootcamp_purchase_date - subscriber_created_at)`
- Why: guides cadence strategy (fast push vs mid-term nurture).

## 3) Minimal Visual Layout (keep simple)

### Page 1: Executive Health
- KPI cards: #1, #2, #3, #4
- 90-day trend lines for Value OR, Sales OR, Sales CTOR

### Page 2: Segment Quality
- KPI cards: #5, #6, #7, #8, #9
- Cohort chart (oldest vs newest trends)

### Page 3: Conversion Path
- KPI cards: #10, #11, #12
- Workshop vs non-workshop conversion comparison chart

## 4) Global Filters (must-have)
- Date range
- Email type: `All / Value / Sales`
- Window mode: `30d / 38d / 90d`
- Segment mode: `Lead / Non-lead / All`

## 5) Value vs Sales Classification (Critical)

## Current reality
Kit Broadcast APIs expose fields such as `subject`, `description`, `send_at`, `subscriber_filter`, and stats, but no native `category` field for Value/Sales.

## Recommended approach (best practical option)
Use broadcast `description` (internal note) as the source of truth and enforce a strict prefix:
- `[TYPE:VALUE]`
- `[TYPE:SALES]`

Then parse this in the dashboard ingestion layer.

## Why this works
- `description` is available in create/update/list broadcast API responses.
- Internal note/description is also editable in Kit UI before and after send.

## Fallback rules
If no type prefix exists:
1. Infer from campaign naming convention (subject keywords) and mark as `INFERRED`.
2. Display a QA warning: `Unclassified broadcasts`.
3. Do not silently force class into Value/Sales.

## 6) API Endpoints to Use
- `GET /v4/broadcasts` -> fetch broadcasts metadata (`subject`, `description`, `send_at`, filters)
- `GET /v4/broadcasts/stats` or `GET /v4/broadcasts/{id}/stats` -> recipients/opens/clicks metrics
- `POST /v4/subscribers/filter` -> per-broadcast engagement cohort slicing (opens/clicks/sent)
- `GET /v4/subscribers` -> subscriber base, created_at, state
- `GET /v4/subscribers/{id}/tags` -> segment logic from tags
- `GET /v4/subscribers/{id}/stats` -> per-subscriber engagement (June 2025+ coverage)

## 7) Data Quality Rules to Display in UI
- Always show numerators + denominators beside every rate.
- Show matched-broadcast count used in each KPI.
- Show unclassified broadcast count (missing `[TYPE:*]` description).
- Mark low-confidence metrics when coverage is low (e.g., <3 broadcasts).

## 8) Delivery Definition for V1
A successful V1 dashboard should allow a non-technical marketer to answer in under 2 minutes:
- Are we declining now?
- Is decline concentrated in older subscribers?
- Is lead magnet bringing better-quality subscribers?
- Are workshops still driving bootcamp purchases?

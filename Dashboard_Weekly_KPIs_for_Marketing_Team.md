# Weekly Marketing Dashboard KPIs (Based on Email Health Audit + LO Strategy)

## 1) Why this dashboard exists
The goal is simple: give the marketing team one weekly view that answers three questions fast:
- Are we getting healthier or weaker?
- Which subscriber groups are still responsive?
- Which path is driving purchases (especially Workshop -> Bootcamp)?

The audit showed that decline is real, but uneven:
- Click behavior (CTOR) dropped first, then opens dropped.
- Older subscribers are much less engaged than newer subscribers.
- Buyers and lead-magnet subscribers are still strong pockets.
- Workshops are a strong feeder into bootcamp.

So the dashboard should not focus on one global average. It should focus on segment-level health and conversion flow.

---

## 2) Dashboard structure (weekly)
Use **four sections** only:
1. **How big is the list now, and is it shrinking or growing?**
2. **Are people still opening and clicking our emails (by segment)?**
3. **Are we safely reactivating cold subscribers without hurting deliverability?**
4. **Are we moving people from lead magnet -> workshop -> bootcamp?**

This keeps the dashboard practical and avoids metric overload.

Use consistent time windows across all metrics:
- last 7 completed days for weekly KPI reporting
- last 30 days for more stable engagement rates
- last 90 days where conversion lag matters
- last 180 days for conversion-speed distributions
- last 4 completed calendar months for monthly trend views
- subscribers who signed up in the last 38 days for lead-magnet comparison cohorts

---

## 3) KPIs to include (with plain definitions, why, and chart type)

## A) Engagement Health

### KPI 1: Value Open Rate by segment
Track the share of delivered value emails that were opened in the last 7 completed days, with a 30-day trend for stability, across these groups: bootcamp buyers, workshop buyers without bootcamp, strict lead-magnet signups, active non-buyers, and re-warmed cold subscribers. This shows where attention is still strong, and the best view is a weekly multi-line trend plus a latest-week segment bar comparison.

### KPI 2: Sales Open Rate by segment
Track the share of delivered sales emails that were opened in the last 7 completed days, with a 30-day trend for stability, by segment. This shows where sales attention still exists and where reach is weakening, and the best view is a weekly multi-line chart by segment.

### KPI 3: Sales CTOR by segment
Track the share of sales openers who clicked in the last 7 completed days, with a 30-day trend for stability, by segment. This is the earliest warning signal for message and CTA quality, and the best view is a weekly multi-line chart by segment.

### KPI 4: Oldest vs Newest engagement gap
Track the open-rate gap and sales-CTOR gap between the oldest 25% and newest 25% of subscribers by signup date, calculated weekly over the last 7 completed days with a 4-week average beside it. This directly tests whether older subscribers are disengaging more, and the clearest view is two small bar charts (current week and 4-week average).

---

## B) Re-warm + Deliverability

### KPI 5: Cold re-permission click rate
Track the share of cold subscribers who clicked a re-permission email among all cold subscribers who received re-permission emails in the last 7 completed days. This shows whether cold subscribers are willing to raise their hand again, and the clearest view is a KPI card with a weekly line trend.

### KPI 6: Cold -> Re-warmed movement rate
Track the percentage of people who were cold at the start of the week and moved to re-warmed status by the end of the same week. This is the core list-recovery engine, and the clearest view is a funnel step chart from Cold to Re-permission Click to Re-warmed.

### KPI 7: Unsubscribe rate by segment
Track unsubscribe rate by segment for the last 7 completed days, with a 4-week average beside it. This protects list quality while scaling segmentation and re-entry tests, and the clearest view is a segment bar chart with a threshold line.

### KPI 8: Complaint rate by segment
Track complaint rate by segment for the last 7 completed days, with a 4-week average beside it. This is a key deliverability risk signal, especially for re-warmed cold cohorts, and the clearest view is a segment bar chart with a threshold line.

---

## C) Funnel and Conversion

### KPI 9: Lead magnet quality (strict)
Track strict lead-magnet signups in the last 7 completed days using roadmap-origin source plus a valid response signal, with a rolling 30-day trend for stability. This verifies whether the lead magnet is attracting qualified people, and the clearest view is a weekly count line plus a KPI card.

### KPI 10: Lead vs Non-lead performance (same signup window)
Compare lead and non-lead subscribers who signed up in the same 38-day window, then calculate the gap in sales open rate and sales CTOR using sales broadcasts from the last 30 days. This confirms whether lead-magnet subscribers are higher quality than regular new signups, and the clearest view is side-by-side bars for both gaps.

### KPI 11: Workshop purchase rate by segment
Track the share of each segment that bought a workshop in the last 30 days, using segment size at the start of that period. This shows who is entering the main on-ramp to monetization, and the clearest view is a segment bar chart.

### KPI 12: Workshop -> Bootcamp conversion rate by segment
Track the share of workshop buyers who then buy bootcamp in a rolling 90-day window, and show the latest cohort result beside it. This is one of the strongest revenue-path signals from the audit, and the clearest view is a cohort line chart plus a segment bar chart.

### KPI 13: Time to purchase from lead-magnet signup
Track median days from lead-magnet signup to first purchase, plus the share that converts within 30 days and within 90 days, using converters observed over the last 180 days. This guides nurture timing and sales cadence, and the clearest view is a histogram plus three KPI cards.

---

## D) List Size and Churn (New)

### KPI 14: Current confirmed subscribers
Track the total number of currently confirmed and sendable subscribers as of the report date. This gives a clear baseline before interpreting engagement rates, and the clearest view is a large KPI card with a small 12-week sparkline.

### KPI 15: Number of unsubscribers
Track the number of unsubscribers in the last 7 completed days, and also show monthly totals for the last 4 completed months. This keeps attention on real audience loss volume, and the clearest view is a weekly KPI card with a monthly bar chart.

### KPI 16: Monthly churn rate
Track monthly churn rate as unsubscribers divided by confirmed subscribers at the start of each month, for the last 4 completed calendar months. This normalizes churn by list size and shows trend direction, and the clearest view is a monthly line chart.

### KPI 17: Unsubscribers per month
Track total unsubscribe volume for each of the last 4 completed calendar months. This gives a simple trend view for leadership and quick diagnosis, and the clearest view is a monthly column chart.

### KPI 18: Subscriber status composition (circle chart)
Show a status snapshot as of the report date using a nested circle: outer ring for confirmed/sendable, bounced, and canceled/unsubscribed, and inner ring for engaged active vs cold inside the confirmed base based on the last 90 days of engagement behavior. This gives a fast list-quality view in one visual, and the clearest format is a nested donut chart.

### KPI 19: Last 4 months email performance snapshot
Show monthly Value Open Rate, Sales Open Rate, and Sales CTOR for the last 4 completed calendar months, with Sales CTOR treated as the primary quality signal. This gives a clear recent trend summary in one panel, and the clearest format is a combo chart with bars for open rates and a line for Sales CTOR.

---

## 4) Presentation rules (important for non-technical team use)
- Show **numerator and denominator** next to every rate.
- For each metric, show:
  - current week
  - previous week
  - 4-week rolling average
- Add a warning badge when volume is low (example: delivered < 500).
- Keep global metrics secondary; segment metrics first.

---

## 5) Recommended dashboard tabs (Streamlit)
1. **List Size & Churn**: KPIs 14–18  
2. **Email Engagement**: KPIs 1–4 and KPI 19  
3. **Cold Re-warm & Deliverability**: KPIs 5–8  
4. **Funnel Performance**: KPIs 9–13  

This gives one weekly operating view that matches both the audit findings and the March–June strategy.

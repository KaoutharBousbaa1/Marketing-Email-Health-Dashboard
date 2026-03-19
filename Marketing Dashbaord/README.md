# Marketing Dashbaord (Live Kit API)

One-page Streamlit dashboard for weekly marketing KPIs.

## Included KPIs
- KPI 3: Sales CTOR by segment
- KPI 6: Cold -> Re-warmed movement rate
- KPI 12: Workshop -> Bootcamp conversion rate by segment
- KPI 14: Current confirmed subscribers
- KPI 19: Last 4 months email performance snapshot
- KPI 16: Monthly churn rate (unsubscribe proxy)

## Data Source
- Kit API only (real-time on app load)
- No CSV files are used for KPI calculations
- Broadcast classification uses `description` (internal note):
  - `value` -> Value
  - `pre-sales` / `post-sales` / `sales` -> Sales

## Run
From this folder:

```bash
cd "Marketing Dashbaord"
python3 -m pip install -r requirements.txt
streamlit run app.py
```

Open the local URL shown by Streamlit (usually `http://localhost:8501`).

## Notes
- If needed, set API key via env var:
  - `export KIT_API_KEY="your_key_here"`
- The app includes a **Refresh now** button to force fresh API pulls.
- KPI 16 is shown as an unsubscribe-based churn proxy, because true lifecycle churn requires historical state snapshots.

## Files
- `app.py`: UI (purple modern layout + charts/cards)
- `kit_client.py`: Kit API client (pagination, retries, filter chunking)
- `kpi_service.py`: KPI computations
- `config.py`: API + KPI windows + segment/tag rules

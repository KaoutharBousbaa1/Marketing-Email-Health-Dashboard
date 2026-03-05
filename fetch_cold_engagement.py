"""
Cold Subscriber Engagement — Kit API Fetcher
=============================================
Strategy:
  1. Paginate through ALL subscribers (status=all) to build email→id map.
     ~34 API calls at 1,000/page for 33k subscribers.
  2. For each cold subscriber from Cold Subscribers.csv, call /stats.
     2,819 API calls (one per cold sub).
  Total: ~2,853 calls — well within rate limits, completes in ~5 minutes.

Output: cold_engagement.csv
"""

import pandas as pd
import requests
import time
from pathlib import Path

BASE     = Path(__file__).resolve().parent
API_KEY  = "kit_7d6b10fad06f88e1d0e47e45ef92e9cc"
API_BASE = "https://api.kit.com/v4"
HEADERS  = {"X-Kit-Api-Key": API_KEY, "Accept": "application/json"}

# ── Load cold subscribers ─────────────────────────────────────────────────────
cold = pd.read_csv(BASE / "Cold Subscribers.csv", encoding="utf-8-sig")
cold["created_at"] = pd.to_datetime(cold["created_at"], utc=True).dt.tz_localize(None)
cold_emails = set(cold["email"].str.lower().str.strip())
print(f"Cold subscribers to enrich: {len(cold):,}")

# ── Step 1: build email → id map from full subscriber list ────────────────────
print("\nStep 1: Paginating through all subscribers to build ID map …")
email_to_id = {}
cursor = None
page   = 0
req_count = 0
window_start = time.time()

while True:
    params = {"status": "all", "per_page": 1000}
    if cursor:
        params["after"] = cursor

    r = requests.get(f"{API_BASE}/subscribers", headers=HEADERS, params=params, timeout=30)
    req_count += 1

    # Rate limiting
    if req_count % 100 == 0:
        elapsed = time.time() - window_start
        if elapsed < 62:
            time.sleep(62 - elapsed)
        req_count    = 0
        window_start = time.time()

    if r.status_code != 200:
        print(f"  Error on page {page}: {r.status_code} {r.text[:200]}")
        break

    data  = r.json()
    subs  = data.get("subscribers", [])
    page += 1

    for s in subs:
        email_to_id[s["email_address"].lower().strip()] = (s["id"], s["state"])

    pagination = data.get("pagination", {})
    print(f"  Page {page}: {len(subs)} subscribers fetched (total mapped: {len(email_to_id):,})")

    if not pagination.get("has_next_page"):
        break
    cursor = pagination.get("end_cursor")

print(f"\nID map built: {len(email_to_id):,} subscribers total")

# Match cold subs
matched   = [(e, *email_to_id[e]) for e in cold_emails if e in email_to_id]
unmatched = [e for e in cold_emails if e not in email_to_id]
print(f"Matched {len(matched):,} cold subs | Unmatched: {len(unmatched)}")

# ── Step 2: fetch stats for each matched cold subscriber ──────────────────────
print(f"\nStep 2: Fetching stats for {len(matched):,} cold subscribers …")

stats_map = {}
req_count    = 0
window_start = time.time()

for i, (email, sub_id, state) in enumerate(matched):
    r = requests.get(f"{API_BASE}/subscribers/{sub_id}/stats", headers=HEADERS, timeout=15)
    req_count += 1

    # Rate limiting
    if req_count % 110 == 0:
        elapsed = time.time() - window_start
        if elapsed < 62:
            print(f"  Rate limit pause: {62-elapsed:.0f}s …")
            time.sleep(62 - elapsed)
        req_count    = 0
        window_start = time.time()

    if r.status_code == 200:
        stats = r.json().get("subscriber", {}).get("stats", {})
        stats_map[email] = {"api_id": sub_id, "api_state": state, **stats}
    else:
        stats_map[email] = {"api_id": sub_id, "api_state": state}

    if (i + 1) % 200 == 0:
        print(f"  Progress: {i+1}/{len(matched)} stats fetched …")

print(f"Stats fetched for {len(stats_map):,} subscribers")

# ── Merge back onto cold CSV ──────────────────────────────────────────────────
def enrich(row):
    key  = row["email"].lower().strip()
    data = stats_map.get(key, {})
    return pd.Series({
        "api_id":                 data.get("api_id"),
        "api_state":              data.get("api_state"),
        "sent":                   data.get("sent"),
        "opened":                 data.get("opened"),
        "clicked":                data.get("clicked"),
        "last_sent":              data.get("last_sent"),
        "last_opened":            data.get("last_opened"),
        "last_clicked":           data.get("last_clicked"),
        "sends_since_last_open":  data.get("sends_since_last_open"),
        "sends_since_last_click": data.get("sends_since_last_click"),
    })

enriched = cold.join(cold.apply(enrich, axis=1))
enriched.to_csv(BASE / "cold_engagement.csv", index=False)

print(f"\nSaved cold_engagement.csv  ({len(enriched):,} rows)")
print("\nSample stats:")
print(enriched[["email","sent","opened","last_opened","sends_since_last_open"]].head(10).to_string())

"""
Lead Magnet Survey Analysis
===========================
Q1: Did sending the survey email (AI Sprint Roadmap) boost future open rates?
    Compares broadcast open rates in the 90 days before vs after the survey send.

Q2: Group A (survey responders) vs Group B (non-responders) — subscriber engagement.
    Uses the Kit API to fetch lifetime stats (sent, opened, clicked) per subscriber.
    Groups are identified from the 'tags' column:
      Group A — clicked the survey email AND has a roadmap-choice tag ("I want…" / "I'm…")
      Group B — ALL non-responders:
                  • Clicked the survey email but NO roadmap tag  (611 subs)
                  • Opened the survey email but never clicked    (5,757 subs)

Charts produced:
  lead_magnet_segments.png — Bar chart: clicked-responded / clicked-not-responded / opened-only
  W — Q1: Pre / Post-survey broadcast open rates by category (Value vs Sales)
  X — Q2: Grouped bars — overall open rate, click rate, CTOR (Group A vs B)
  Y — Q2: Box-plot distributions of open rate and CTOR (Group A vs B)
  Z — Q2: Active vs Cancelled subscriber status (Group A vs B)
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import requests
import time
from pathlib import Path
from scipy import stats as scipy_stats
from difflib import SequenceMatcher
import warnings
warnings.filterwarnings("ignore")

# ─── Constants ───────────────────────────────────────────────────────────────
BASE        = Path(__file__).resolve().parent
OUT         = BASE / "charts"
OUT.mkdir(exist_ok=True)

API_KEY     = "kit_7d6b10fad06f88e1d0e47e45ef92e9cc"
API_BASE    = "https://api.kit.com/v4"
HEADERS     = {"X-Kit-Api-Key": API_KEY, "Accept": "application/json"}

SURVEY_DATE = pd.Timestamp("2026-01-25")
CACHE_FILE  = BASE / "lead_magnet_stats.csv"
POST_EVENTS_CACHE = BASE / "lead_magnet_post_survey_events.csv"
POST_SUMMARY_CACHE = BASE / "lead_magnet_post_survey_summary.csv"

GREEN  = "#10B981"
ORANGE = "#F59E0B"
RED    = "#EF4444"
BRAND  = "#4F46E5"
ACCENT = "#EC4899"
MUTED  = "#94A3B8"
TEAL   = "#06B6D4"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.titlesize": 14,
    "axes.labelsize": 11,
    "figure.dpi": 150,
})

# ─── Survey-response tag detection ───────────────────────────────────────────
RESPONSE_KEYWORDS = ["i want", "i'm", "i work at a company", "i own/run", "i own", "i'm building"]

def has_survey_response(tags_str):
    if pd.isna(tags_str):
        return False
    t = tags_str.lower()
    return any(k in t for k in RESPONSE_KEYWORDS)


# ══════════════════════════════════════════════════════════════════════════════
# Q1 — SURVEY EFFECT ON BROADCAST OPEN RATES
# ══════════════════════════════════════════════════════════════════════════════
print("=" * 70)
print("Q1: DID THE SURVEY BOOST FUTURE OPEN RATES?")
print("=" * 70)

bdf = pd.read_csv(BASE / "Emails Broadcasting - broadcasts_categorised.csv")
bdf["date"] = pd.to_datetime(bdf["date"])
bdf["category"] = bdf["category"].replace("Sles", "Sales")
bdf = bdf[bdf["recipients"] > 900].copy()
bdf_vs = bdf[bdf["category"].isin(["Value", "Sales"])].copy()
bdf_vs.sort_values("date", inplace=True)

PRE_START = SURVEY_DATE - pd.Timedelta(days=90)
pre  = bdf_vs[(bdf_vs["date"] >= PRE_START) & (bdf_vs["date"] < SURVEY_DATE)]
post = bdf_vs[bdf_vs["date"] >  SURVEY_DATE]

print(f"\n90-day pre-survey window:  {PRE_START.date()} → {SURVEY_DATE.date()}")
print(f"Post-survey window:        {SURVEY_DATE.date()} → {bdf_vs['date'].max().date()}")

q1_rows = []
for cat in ["Value", "Sales"]:
    pre_g  = pre[pre["category"] == cat]["open_rate"]
    post_g = post[post["category"] == cat]["open_rate"]
    delta  = post_g.mean() - pre_g.mean() if len(post_g) else float("nan")
    pct    = delta / pre_g.mean() * 100 if len(pre_g) else float("nan")
    t_stat, p_val = (float("nan"), float("nan"))
    if len(pre_g) >= 2 and len(post_g) >= 2:
        t_stat, p_val = scipy_stats.ttest_ind(pre_g, post_g, equal_var=False)
    q1_rows.append({
        "category":   cat,
        "pre_n":      len(pre_g),
        "pre_mean":   pre_g.mean(),
        "pre_std":    pre_g.std(),
        "post_n":     len(post_g),
        "post_mean":  post_g.mean() if len(post_g) else float("nan"),
        "post_std":   post_g.std() if len(post_g) else float("nan"),
        "delta_pp":   delta,
        "delta_pct":  pct,
        "p_value":    p_val,
    })
    print(f"\n  {cat}:")
    print(f"    Pre  (n={len(pre_g)}): {pre_g.mean():.1f}% ± {pre_g.std():.1f}pp")
    if len(post_g):
        sig = "✓ significant" if p_val < 0.05 else "not significant"
        print(f"    Post (n={len(post_g)}): {post_g.mean():.1f}% ± {post_g.std():.1f}pp")
        print(f"    Δ = {delta:+.1f}pp ({pct:+.1f}%)  p={p_val:.3f} [{sig}]")
    else:
        print("    Post: no data")

q1_df = pd.DataFrame(q1_rows)


# ══════════════════════════════════════════════════════════════════════════════
# LOAD LEAD MAGNET CSVs — IDENTIFY GROUPS
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("LOADING LEAD MAGNET DATA & IDENTIFYING GROUPS")
print("=" * 70)

df_clicked = pd.read_csv(BASE / "AI Sprint Roadmap - Clicked.csv")
df_clicked["email_lower"] = df_clicked["email"].str.lower().str.strip()
df_clicked["is_group_a"]  = df_clicked["tags"].apply(has_survey_response)

df_opened = pd.read_csv(BASE / "AI Sprint Roadmap - Opened subscribers.csv")
df_opened["email_lower"] = df_opened["email"].str.lower().str.strip()
df_opened["is_group_a"]  = df_opened["tags"].apply(has_survey_response)

# ── Segment counts for the breakdown chart ───────────────────────────────────
clicked_emails      = set(df_clicked["email_lower"])
n_clicked_responded = df_clicked["is_group_a"].sum()                           # 892
n_clicked_no_resp   = (~df_clicked["is_group_a"]).sum()                        # 611
# Opened-only = opened but NOT in the Clicked CSV at all
n_opened_only       = (~df_opened["email_lower"].isin(clicked_emails)).sum()   # 5,831

print(f"\nSurvey email breakdown:")
print(f"  Clicked + Responded  (Group A): {n_clicked_responded:,}")
print(f"  Clicked + No Response          : {n_clicked_no_resp:,}")
print(f"  Opened only (no click)         : {n_opened_only:,}")

# ── Define final groups ───────────────────────────────────────────────────────
# Group A: clicked + responded
group_a = df_clicked[df_clicked["is_group_a"]].copy()

# Group B: everyone who didn't respond
#   B1 = clicked but no roadmap tag
b1 = df_clicked[~df_clicked["is_group_a"]][["email_lower"]].copy()
#   B2 = opened only (not in Clicked CSV), and no roadmap tag
b2 = df_opened[
    (~df_opened["email_lower"].isin(clicked_emails)) &
    (~df_opened["is_group_a"])
][["email_lower"]].copy()

group_b = pd.concat([b1, b2], ignore_index=True).drop_duplicates("email_lower")

print(f"\nGroup A (survey responders):  {len(group_a):,}")
print(f"Group B (all non-responders): {len(group_b):,}")
print(f"  B1 — clicked, no response:  {len(b1):,}")
print(f"  B2 — opened only:           {len(b2):,}")

all_emails_needed = pd.concat([
    group_a[["email_lower"]].assign(group="A"),
    group_b[["email_lower"]].assign(group="B"),
], ignore_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# CHART: SURVEY SEGMENT BREAKDOWN (bar chart — no API needed)
# ══════════════════════════════════════════════════════════════════════════════
print("\nGenerating segment breakdown chart …")

seg_labels = [
    "Clicked &\nResponded\n(Group A)",
    "Clicked but\nDid NOT Respond",
    "Opened Only\n(Never Clicked)",
]
seg_values = [n_clicked_responded, n_clicked_no_resp, n_opened_only]
seg_colors = [GREEN, ORANGE, MUTED]

fig, ax = plt.subplots(figsize=(10, 6))
bars = ax.bar(seg_labels, seg_values, color=seg_colors, edgecolor="white",
              alpha=0.88, width=0.5)

total_clicked = n_clicked_responded + n_clicked_no_resp
total_opened  = total_clicked + n_opened_only

for bar, val in zip(bars, seg_values):
    pct_of_clicked = val / total_clicked * 100 if val <= total_clicked else None
    pct_of_opened  = val / total_opened  * 100
    label = f"{val:,}\n({pct_of_opened:.1f}% of all who opened)"
    ax.text(bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 60,
            label, ha="center", fontsize=10, fontweight="bold")

ax.set_ylabel("Number of Subscribers")
ax.set_title(
    "AI Sprint Roadmap Survey — How Subscribers Engaged\n"
    f"Total who opened: {total_opened:,}  |  Total who clicked: {total_clicked:,}",
    fontsize=13,
)
ax.set_ylim(0, max(seg_values) * 1.20)

# Response rate annotation
resp_rate = n_clicked_responded / total_clicked * 100
ax.annotate(
    f"Survey response rate\namong clickers: {resp_rate:.1f}%",
    xy=(0, n_clicked_responded),
    xytext=(0.6, n_clicked_responded + max(seg_values) * 0.08),
    fontsize=9, color=GREEN, fontweight="bold",
    arrowprops=dict(arrowstyle="->", color=GREEN, lw=1.5),
)

fig.tight_layout()
fig.savefig(OUT / "lead_magnet_segments.png")
plt.close()
print("  Segment breakdown chart saved → lead_magnet_segments.png")


# ══════════════════════════════════════════════════════════════════════════════
# KIT API — FETCH SUBSCRIBER STATS
# ══════════════════════════════════════════════════════════════════════════════

def rate_limited_get(url, params=None, timeout=20):
    """GET with simple retry on 429."""
    for attempt in range(3):
        r = requests.get(url, headers=HEADERS, params=params, timeout=timeout)
        if r.status_code == 429:
            retry_after = int(r.headers.get("Retry-After", 62))
            print(f"  Rate-limited. Waiting {retry_after}s …")
            time.sleep(retry_after)
            continue
        return r
    return r

def rate_limited_post(url, payload=None, timeout=25):
    """POST with simple retry on 429."""
    for attempt in range(3):
        r = requests.post(url, headers=HEADERS, json=payload, timeout=timeout)
        if r.status_code == 429:
            retry_after = int(r.headers.get("Retry-After", 62))
            print(f"  Rate-limited. Waiting {retry_after}s …")
            time.sleep(retry_after)
            continue
        return r
    return r


if CACHE_FILE.exists():
    print(f"\nLoading cached stats from {CACHE_FILE.name} …")
    enriched = pd.read_csv(CACHE_FILE)
    print(f"  {len(enriched):,} rows loaded from cache.")
else:
    target_emails = set(all_emails_needed["email_lower"])
    print(f"\nStep 1: Building email → subscriber-ID map (fetching all ~19k subs) …")

    email_to_id = {}
    cursor      = None
    page        = 0
    req_count   = 0
    window_start = time.time()

    while True:
        params = {"status": "all", "per_page": 1000}
        if cursor:
            params["after"] = cursor

        r = rate_limited_get(f"{API_BASE}/subscribers", params=params)
        req_count += 1

        if req_count % 100 == 0:
            elapsed = time.time() - window_start
            if elapsed < 62:
                time.sleep(62 - elapsed)
            req_count    = 0
            window_start = time.time()

        if r.status_code != 200:
            print(f"  Error on page {page}: {r.status_code}")
            break

        data  = r.json()
        subs  = data.get("subscribers", [])
        page += 1

        for s in subs:
            e = s["email_address"].lower().strip()
            if e in target_emails:
                email_to_id[e] = (s["id"], s["state"])

        pagination = data.get("pagination", {})
        if page % 5 == 0:
            print(f"  Page {page}: {len(subs)} subs fetched, matched {len(email_to_id):,} so far …")

        if not pagination.get("has_next_page"):
            break
        cursor = pagination.get("end_cursor")

    print(f"ID map complete: matched {len(email_to_id):,} of {len(target_emails):,} targets")

    # Step 2: fetch stats for each matched subscriber
    print(f"\nStep 2: Fetching subscriber stats …")
    stats_map    = {}
    req_count    = 0
    window_start = time.time()

    matched_list = [
        (row["email_lower"], row["group"])
        for _, row in all_emails_needed.iterrows()
        if row["email_lower"] in email_to_id
    ]

    for i, (email, group) in enumerate(matched_list):
        sub_id, state = email_to_id[email]
        r = rate_limited_get(f"{API_BASE}/subscribers/{sub_id}/stats")
        req_count += 1

        if req_count % 110 == 0:
            elapsed = time.time() - window_start
            if elapsed < 62:
                print(f"  Rate-limit pause: {62-elapsed:.0f}s …")
                time.sleep(62 - elapsed)
            req_count    = 0
            window_start = time.time()

        if r.status_code == 200:
            s = r.json().get("subscriber", {}).get("stats", {})
            stats_map[email] = {
                "api_id":                 sub_id,
                "api_state":              state,
                "group":                  group,
                "sent":                   s.get("sent"),
                "opened":                 s.get("opened"),
                "clicked":                s.get("clicked"),
                "sends_since_last_open":  s.get("sends_since_last_open"),
                "sends_since_last_click": s.get("sends_since_last_click"),
            }
        else:
            stats_map[email] = {"api_id": sub_id, "api_state": state, "group": group}

        if (i + 1) % 200 == 0:
            print(f"  {i+1}/{len(matched_list)} stats fetched …")

    # Merge into DataFrame
    rows = []
    for _, row in all_emails_needed.iterrows():
        e = row["email_lower"]
        s = stats_map.get(e, {})
        rows.append({"email": e, "group": row["group"], **s})

    enriched = pd.DataFrame(rows)
    enriched.to_csv(CACHE_FILE, index=False)
    print(f"Saved to {CACHE_FILE.name} ({len(enriched):,} rows)")


# ══════════════════════════════════════════════════════════════════════════════
# Q2 — COMPUTE ENGAGEMENT METRICS PER GROUP
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("Q2: GROUP A vs GROUP B ENGAGEMENT")
print("=" * 70)

# Make sure numeric
for col in ["sent", "opened", "clicked"]:
    enriched[col] = pd.to_numeric(enriched[col], errors="coerce")

# Filter: must have received at least 5 emails so rates are meaningful
enr = enriched[enriched["sent"] >= 5].copy()
enr["open_rate"]   = enr["opened"]  / enr["sent"]
enr["click_rate"]  = enr["clicked"] / enr["sent"]
enr["ctor"]        = np.where(enr["opened"] > 0, enr["clicked"] / enr["opened"], np.nan)

grp = enr.groupby("group").agg(
    n           = ("open_rate",  "count"),
    or_mean     = ("open_rate",  "mean"),
    or_median   = ("open_rate",  "median"),
    or_std      = ("open_rate",  "std"),
    cr_mean     = ("click_rate", "mean"),
    cr_std      = ("click_rate", "std"),
    ctor_mean   = ("ctor",       "mean"),
    ctor_std    = ("ctor",       "std"),
).reset_index()

print(f"\nFiltered to subscribers with ≥5 emails sent: {len(enr):,}")
print(grp.to_string(index=False))

# Statistical significance
for metric, label in [("open_rate", "Open Rate"), ("click_rate", "Click Rate"), ("ctor", "CTOR")]:
    a_vals = enr[enr["group"] == "A"][metric].dropna()
    b_vals = enr[enr["group"] == "B"][metric].dropna()
    if len(a_vals) >= 5 and len(b_vals) >= 5:
        t, p = scipy_stats.mannwhitneyu(a_vals, b_vals, alternative="two-sided")
        sig = "✓ significant" if p < 0.05 else "not significant"
        print(f"\n  {label}: A={a_vals.mean()*100:.1f}% vs B={b_vals.mean()*100:.1f}%"
              f"  MWU p={p:.4f}  [{sig}]")

# Active vs Cancelled status
if "api_state" in enriched.columns:
    status_tbl = (
        enriched[enriched["group"].isin(["A","B"])]
        .assign(is_active=lambda d: d["api_state"].str.lower().str.strip() == "active")
        .groupby("group")["is_active"]
        .agg(["sum", "count"])
        .rename(columns={"sum": "active", "count": "total"})
        .assign(active_pct=lambda d: d["active"] / d["total"] * 100)
    )
    print("\n  Subscriber Status:")
    print(status_tbl.to_string())


# ══════════════════════════════════════════════════════════════════════════════
# Q3 — POST-SURVEY GROUP A VS B BY EMAIL TYPE (KIT FILTER API)
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("Q3: POST-SURVEY GROUP A VS B — VALUE/SALES OPEN RATE + SALES CTOR")
print("=" * 70)

def norm_subject(s):
    s = str(s).lower()
    cleaned = "".join(ch if ch.isalnum() else " " for ch in s)
    return " ".join(cleaned.split())

def fetch_all_broadcasts():
    rows = []
    cursor = None
    while True:
        params = {"per_page": 1000}
        if cursor:
            params["after"] = cursor
        r = rate_limited_get(f"{API_BASE}/broadcasts", params=params)
        if r.status_code != 200:
            print(f"  Broadcast fetch error: {r.status_code}")
            break
        data = r.json()
        for b in data.get("broadcasts", []):
            rows.append({
                "api_broadcast_id": int(b["id"]),
                "api_subject": b.get("subject", ""),
                "api_date": pd.to_datetime(b.get("created_at"), utc=True, errors="coerce").tz_convert(None).date()
                if pd.notna(pd.to_datetime(b.get("created_at"), utc=True, errors="coerce")) else pd.NaT,
                "api_subject_norm": norm_subject(b.get("subject", "")),
            })
        pag = data.get("pagination", {})
        if not pag.get("has_next_page"):
            break
        cursor = pag.get("end_cursor")
    return pd.DataFrame(rows)

def map_broadcast_to_api(row, api_df):
    same_day = api_df[api_df["api_date"] == row["date"]]
    if same_day.empty:
        return np.nan
    local_norm = row["subject_norm"]
    exact = same_day[same_day["api_subject_norm"] == local_norm]
    if len(exact):
        return int(exact.iloc[0]["api_broadcast_id"])
    scored = same_day.assign(
        score=same_day["api_subject_norm"].apply(
            lambda x: SequenceMatcher(None, local_norm, x).ratio()
        )
    ).sort_values("score", ascending=False)
    best = scored.iloc[0]
    return int(best["api_broadcast_id"]) if best["score"] >= 0.80 else np.nan

def fetch_event_emails_for_broadcast(event_type, broadcast_id):
    payload_base = {
        "all": [{
            "type": event_type,
            "count_greater_than": 0,
            "any": [{"type": "broadcasts", "ids": [int(broadcast_id)]}],
        }]
    }
    emails = set()
    cursor = None
    while True:
        payload = dict(payload_base)
        if cursor:
            payload["after"] = cursor
        r = rate_limited_post(f"{API_BASE}/subscribers/filter", payload=payload)
        if r.status_code != 200:
            print(f"  Filter error ({event_type}, {broadcast_id}): {r.status_code}")
            break
        data = r.json()
        for s in data.get("subscribers", []):
            e = s.get("email_address")
            if e:
                emails.add(e.lower().strip())
        pag = data.get("pagination", {})
        if not pag.get("has_next_page"):
            break
        cursor = pag.get("end_cursor")
    return emails

group_a_emails = set(group_a["email_lower"])
group_b_emails = set(group_b["email_lower"])

if POST_EVENTS_CACHE.exists():
    print(f"Loading cached post-survey events from {POST_EVENTS_CACHE.name} …")
    post_events = pd.read_csv(POST_EVENTS_CACHE)
    if "broadcast_id" in post_events.columns:
        post_events = post_events.drop_duplicates(subset=["broadcast_id"]).copy()
else:
    api_broadcasts = fetch_all_broadcasts()
    post_local = bdf_vs[bdf_vs["date"] > SURVEY_DATE].copy()
    post_local["date"] = post_local["date"].dt.date
    post_local["subject_norm"] = post_local["subject"].apply(norm_subject)
    post_local["api_broadcast_id"] = post_local.apply(
        lambda r: map_broadcast_to_api(r, api_broadcasts), axis=1
    )
    matched = post_local[post_local["api_broadcast_id"].notna()].copy()
    matched["api_broadcast_id"] = matched["api_broadcast_id"].astype(int)
    matched = matched.sort_values("date").drop_duplicates(subset=["api_broadcast_id"], keep="last")
    unmatched = len(post_local) - len(matched)
    print(f"Post-survey broadcasts: {len(post_local)} | matched to Kit IDs: {len(matched)} | unmatched: {unmatched}")

    rows = []
    for i, r in matched.iterrows():
        bid = int(r["api_broadcast_id"])
        o = fetch_event_emails_for_broadcast("opens", bid)
        c = fetch_event_emails_for_broadcast("clicks", bid)
        a_open = len(o & group_a_emails)
        b_open = len(o & group_b_emails)
        a_click = len(c & group_a_emails)
        b_click = len(c & group_b_emails)
        rows.append({
            "broadcast_id": bid,
            "date": r["date"],
            "subject": r["subject"],
            "category": r["category"],
            "group_a_size": len(group_a_emails),
            "group_b_size": len(group_b_emails),
            "group_a_open": a_open,
            "group_b_open": b_open,
            "group_a_click": a_click,
            "group_b_click": b_click,
        })
        print(f"  {r['date']} [{r['category']}] id={bid} | A open/click {a_open}/{a_click} | B open/click {b_open}/{b_click}")

    post_events = pd.DataFrame(rows)
    post_events.to_csv(POST_EVENTS_CACHE, index=False)
    print(f"Saved {POST_EVENTS_CACHE.name} ({len(post_events)} rows)")

summary_rows = []
if len(post_events):
    a_size = int(post_events["group_a_size"].iloc[0])
    b_size = int(post_events["group_b_size"].iloc[0])
    for cat in ["Value", "Sales"]:
        sub = post_events[post_events["category"] == cat].copy()
        if len(sub) == 0:
            continue
        a_or_series = sub["group_a_open"] / a_size * 100
        b_or_series = sub["group_b_open"] / b_size * 100
        a_or = a_or_series.mean()
        b_or = b_or_series.mean()
        p_or = scipy_stats.mannwhitneyu(a_or_series, b_or_series, alternative="two-sided").pvalue if len(sub) >= 2 else np.nan
        summary_rows.append({
            "metric": f"{cat} open rate",
            "category": cat,
            "n_broadcasts": len(sub),
            "group_a_value": a_or,
            "group_b_value": b_or,
            "delta_pp": a_or - b_or,
            "p_value": p_or,
        })

    sales = post_events[post_events["category"] == "Sales"].copy()
    if len(sales):
        sales["group_a_ctor"] = np.where(sales["group_a_open"] > 0, sales["group_a_click"] / sales["group_a_open"] * 100, np.nan)
        sales["group_b_ctor"] = np.where(sales["group_b_open"] > 0, sales["group_b_click"] / sales["group_b_open"] * 100, np.nan)
        a_ctor = sales["group_a_ctor"].mean()
        b_ctor = sales["group_b_ctor"].mean()
        p_ctor = scipy_stats.mannwhitneyu(
            sales["group_a_ctor"].dropna(),
            sales["group_b_ctor"].dropna(),
            alternative="two-sided",
        ).pvalue if len(sales.dropna(subset=["group_a_ctor", "group_b_ctor"])) >= 2 else np.nan
        summary_rows.append({
            "metric": "Sales CTOR",
            "category": "Sales",
            "n_broadcasts": len(sales),
            "group_a_value": a_ctor,
            "group_b_value": b_ctor,
            "delta_pp": a_ctor - b_ctor,
            "p_value": p_ctor,
        })

post_summary = pd.DataFrame(summary_rows)
post_summary.to_csv(POST_SUMMARY_CACHE, index=False)
print("\nQ3 summary:")
if len(post_summary):
    print(post_summary.to_string(index=False, float_format=lambda v: f"{v:.2f}"))
else:
    print("  No post-survey matched broadcasts available.")

# Chart AA — Q3 summary bars
if len(post_summary):
    fig, ax = plt.subplots(figsize=(11, 6))
    x = np.arange(len(post_summary))
    w = 0.34
    a_vals = post_summary["group_a_value"].values
    b_vals = post_summary["group_b_value"].values
    labels = post_summary["metric"].tolist()

    ba = ax.bar(x - w/2, a_vals, w, color=GREEN, alpha=0.88, label="Group A (Responded)")
    bb = ax.bar(x + w/2, b_vals, w, color=ORANGE, alpha=0.88, label="Group B (Non-responded)")

    for i, (av, bv, p) in enumerate(zip(a_vals, b_vals, post_summary["p_value"].values)):
        ax.text(i - w/2, av + 0.7, f"{av:.1f}%", ha="center", fontsize=9)
        ax.text(i + w/2, bv + 0.7, f"{bv:.1f}%", ha="center", fontsize=9)
        delta = av - bv
        marker = "✓" if pd.notna(p) and p < 0.05 else "n.s."
        ax.text(i, max(av, bv) + 3.0, f"Δ {delta:+.1f}pp\n{marker}", ha="center", fontsize=9, color=MUTED)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_ylabel("Rate (%)")
    ax.set_title("Q3: Post-Survey Group A vs Group B\nValue OR, Sales OR, and Sales CTOR (Kit filter endpoint)")
    ax.legend(fontsize=9)
    ax.set_ylim(0, max(max(a_vals), max(b_vals)) * 1.35)
    fig.tight_layout()
    fig.savefig(OUT / "AA_post_survey_group_by_type.png")
    plt.close()
    print("  Chart AA saved.")


# ══════════════════════════════════════════════════════════════════════════════
# CHART W — Q1: Pre / Post survey open rates by category
# ══════════════════════════════════════════════════════════════════════════════
print("\nGenerating charts …")

fig, ax = plt.subplots(figsize=(11, 6))

categories   = ["Value", "Sales"]
x            = np.arange(len(categories))
bar_w        = 0.32
cat_colors   = [BRAND, ACCENT]

for idx, (cat, color) in enumerate(zip(categories, cat_colors)):
    row   = q1_df[q1_df["category"] == cat].iloc[0]
    bars_pre  = ax.bar(x[idx] - bar_w/2, row["pre_mean"],  bar_w,
                       color=color, alpha=0.40, edgecolor="white",
                       label=f"{cat} — Pre-survey" if idx == 0 else None)
    bars_post = ax.bar(x[idx] + bar_w/2, row["post_mean"], bar_w,
                       color=color, alpha=0.92, edgecolor="white",
                       label=f"{cat} — Post-survey" if idx == 0 else None)

    # Error bars (±1 std)
    ax.errorbar(x[idx] - bar_w/2, row["pre_mean"],  yerr=row["pre_std"],
                fmt="none", color="black", capsize=4, linewidth=1.2)
    if not np.isnan(row["post_mean"]):
        ax.errorbar(x[idx] + bar_w/2, row["post_mean"], yerr=row["post_std"],
                    fmt="none", color="black", capsize=4, linewidth=1.2)

    # Delta annotation
    if not np.isnan(row["delta_pp"]):
        c    = GREEN if row["delta_pp"] >= 0 else RED
        sign = "▲" if row["delta_pp"] >= 0 else "▼"
        mid  = x[idx]
        top  = max(row["pre_mean"], row["post_mean"]) + row["pre_std"] + 1.5
        ax.text(mid, top, f"{sign} {abs(row['delta_pp']):.1f}pp\n({row['delta_pct']:+.1f}%)",
                ha="center", fontsize=9, fontweight="bold", color=c)
        sig_marker = "(*)" if row["p_value"] < 0.05 else "(n.s.)"
        ax.text(mid, top - 3.8, sig_marker, ha="center", fontsize=8, color=MUTED)

ax.set_xticks(x)
ax.set_xticklabels(categories, fontsize=12)
ax.set_ylabel("Avg Broadcast Open Rate (%)")
ax.set_title(
    f"Q1: Did the Lead Magnet Survey Boost Open Rates?\n"
    f"90-day pre-survey vs post-survey, >900-recipient broadcasts",
    fontsize=13,
)

pre_patch  = mpatches.Patch(color=MUTED, alpha=0.55, label="Pre-survey (faded)")
post_patch = mpatches.Patch(color=MUTED, alpha=0.95, label="Post-survey (solid)")
val_patch  = mpatches.Patch(color=BRAND, label="Value emails")
sal_patch  = mpatches.Patch(color=ACCENT, label="Sales emails")
ax.legend(handles=[pre_patch, post_patch, val_patch, sal_patch], fontsize=9, ncol=2)
ax.set_ylim(0, ax.get_ylim()[1] * 1.20)

# Annotation: note on small post sample
ax.text(0.99, 0.02,
        f"Note: post-survey window is short (~{(bdf_vs['date'].max() - SURVEY_DATE).days} days).\n"
        "Interpret with caution.",
        transform=ax.transAxes, ha="right", va="bottom", fontsize=8, color=MUTED,
        style="italic")

fig.tight_layout()
fig.savefig(OUT / "W_survey_pre_post_open_rates.png")
plt.close()
print("  Chart W saved.")


# ══════════════════════════════════════════════════════════════════════════════
# CHART X — Q2: Grouped bars — Open Rate, Click Rate, CTOR (Group A vs B)
# ══════════════════════════════════════════════════════════════════════════════
metrics = [
    ("or_mean",   "or_std",   "Avg Open Rate (%)",   "Open Rate"),
    ("cr_mean",   "cr_std",   "Avg Click Rate (%)",  "Click Rate"),
    ("ctor_mean", "ctor_std", "Avg CTOR (%)",        "CTOR\n(Clicks / Opens)"),
]

a_row = grp[grp["group"] == "A"].iloc[0] if len(grp[grp["group"] == "A"]) else None
b_row = grp[grp["group"] == "B"].iloc[0] if len(grp[grp["group"] == "B"]) else None

if a_row is not None and b_row is not None:
    fig, axes = plt.subplots(1, 3, figsize=(14, 6))

    for ax, (mean_col, std_col, ylabel, title) in zip(axes, metrics):
        a_val = a_row[mean_col] * 100
        b_val = b_row[mean_col] * 100
        a_err = a_row[std_col]  * 100
        b_err = b_row[std_col]  * 100

        bars = ax.bar([0, 1], [a_val, b_val], color=[GREEN, ORANGE],
                      alpha=0.88, edgecolor="white", width=0.5)
        ax.errorbar([0, 1], [a_val, b_val], yerr=[a_err, b_err],
                    fmt="none", color="black", capsize=5, linewidth=1.3)

        # Delta label
        delta = a_val - b_val
        c     = GREEN if delta > 0 else RED
        sign  = "▲" if delta > 0 else "▼"
        ax.text(0.5, max(a_val, b_val) + a_err + 1.5,
                f"{sign} {abs(delta):.1f}pp\nA vs B",
                ha="center", fontsize=9, fontweight="bold", color=c,
                transform=ax.get_xaxis_transform())

        for bar, val, n_key in zip(bars, [a_val, b_val], ["A","B"]):
            n = int(a_row["n"] if n_key == "A" else b_row["n"])
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + b_err + 0.3,
                    f"{val:.1f}%\nn={n}", ha="center", fontsize=9)

        ax.set_xticks([0, 1])
        ax.set_xticklabels(["Group A\n(Survey Responders)", "Group B\n(Non-Responders)"], fontsize=9)
        ax.set_ylabel(ylabel)
        ax.set_title(title, fontsize=11)
        ax.set_ylim(0, ax.get_ylim()[1] * 1.25)

    fig.suptitle(
        "Q2: Are Survey Responders More Engaged?\n"
        "Group A (completed survey) vs Group B (clicked but didn't respond)",
        fontsize=13, fontweight="bold",
    )
    fig.tight_layout()
    fig.savefig(OUT / "X_group_a_vs_b_engagement.png")
    plt.close()
    print("  Chart X saved.")


# ══════════════════════════════════════════════════════════════════════════════
# CHART Y — Q2: Box-plot distributions for Open Rate and CTOR
# ══════════════════════════════════════════════════════════════════════════════
fig, axes = plt.subplots(1, 2, figsize=(13, 6))

for ax, (col, label) in zip(axes, [("open_rate", "Open Rate (%)"), ("ctor", "CTOR (%)")]):
    a_vals = enr[enr["group"] == "A"][col].dropna() * 100
    b_vals = enr[enr["group"] == "B"][col].dropna() * 100

    bp = ax.boxplot(
        [a_vals, b_vals],
        patch_artist=True,
        medianprops=dict(color="white", linewidth=2.5),
        whiskerprops=dict(linewidth=1.3),
        capprops=dict(linewidth=1.3),
        flierprops=dict(marker=".", alpha=0.3, markersize=4),
        widths=0.45,
    )
    bp["boxes"][0].set(facecolor=GREEN,  alpha=0.75)
    bp["boxes"][1].set(facecolor=ORANGE, alpha=0.75)

    ax.set_xticks([1, 2])
    ax.set_xticklabels(
        [f"Group A\n(n={len(a_vals)})\nSurvey Responders",
         f"Group B\n(n={len(b_vals)})\nNon-Responders"],
        fontsize=9,
    )
    ax.set_ylabel(label)
    ax.set_title(f"Distribution of {label}", fontsize=11)

    # Median labels
    for pos, vals, color in [(1, a_vals, GREEN), (2, b_vals, ORANGE)]:
        med = vals.median()
        ax.text(pos, med + 1, f"Med: {med:.1f}%", ha="center", fontsize=8,
                color=color, fontweight="bold")

fig.suptitle(
    "Q2: Engagement Rate Distributions — Group A vs Group B\n"
    "(per-subscriber lifetime stats from Kit API)",
    fontsize=13, fontweight="bold",
)
fig.tight_layout()
fig.savefig(OUT / "Y_group_ab_distributions.png")
plt.close()
print("  Chart Y saved.")


# ══════════════════════════════════════════════════════════════════════════════
# CHART Z — Q2: Subscriber status (Active vs Cancelled) for Group A vs B
# ══════════════════════════════════════════════════════════════════════════════
if "api_state" in enriched.columns:
    status_data = (
        enriched[enriched["group"].isin(["A", "B"])]
        .assign(state_clean=lambda d: d["api_state"].str.lower().str.strip().fillna("unknown"))
        .groupby(["group", "state_clean"])
        .size()
        .reset_index(name="count")
    )
    # Pivot
    pvt = status_data.pivot(index="group", columns="state_clean", values="count").fillna(0)
    pvt = pvt.div(pvt.sum(axis=1), axis=0) * 100   # convert to %

    fig, ax = plt.subplots(figsize=(9, 5))
    state_colors = {"active": GREEN, "cancelled": RED, "inactive": ORANGE, "unknown": MUTED}
    bottom = np.zeros(len(pvt))
    for state, color in state_colors.items():
        if state in pvt.columns:
            vals = pvt[state].values
            bars = ax.bar(pvt.index, vals, bottom=bottom,
                          color=color, label=state.capitalize(), edgecolor="white", width=0.45)
            for bar, v, bot in zip(bars, vals, bottom):
                if v > 3:
                    ax.text(bar.get_x() + bar.get_width() / 2,
                            bot + v / 2, f"{v:.1f}%",
                            ha="center", va="center", fontsize=9, color="white", fontweight="bold")
            bottom += vals

    ax.set_xticks([0, 1])
    ax.set_xticklabels(["Group A\nSurvey Responders", "Group B\nNon-Responders"], fontsize=10)
    ax.set_ylabel("Subscriber Status (%)")
    ax.set_title(
        "Q2: Subscriber Retention — Group A vs Group B\n"
        "(Current Kit subscriber status)",
        fontsize=13,
    )
    ax.legend(fontsize=9, bbox_to_anchor=(1.01, 1), loc="upper left")
    ax.set_ylim(0, 110)
    fig.tight_layout()
    fig.savefig(OUT / "Z_group_ab_status.png")
    plt.close()
print("  Chart Z saved.")

print("\nAll lead magnet charts saved: W, X, Y, Z, AA")
print(f"Intermediate API data saved to: {CACHE_FILE.name}")

"""
Lead Magnet Origin Cohort Engagement (Post-Survey Broadcasts)
==============================================================

Question answered
-----------------
Do subscribers inferred to have joined through the survey lead magnet
show higher engagement than subscribers who were already on the list
when the survey was sent?

Metrics (post-survey broadcasts only)
-------------------------------------
- Value email open rate
- Sales email open rate
- Sales CTOR (clicks / opens on sales emails)

Outputs
-------
- lead_magnet_origin_post_survey_events.csv
- lead_magnet_origin_post_survey_summary.csv
- charts/AD2_lead_magnet_origin_metrics.png
"""

from __future__ import annotations

import time
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import requests
from scipy import stats as scipy_stats


BASE = Path(__file__).resolve().parent
GENERATED = BASE / "data" / "generated"
GENERATED.mkdir(parents=True, exist_ok=True)
OUT = BASE / "charts"
OUT.mkdir(exist_ok=True)

API_KEY = "kit_7d6b10fad06f88e1d0e47e45ef92e9cc"
API_BASE = "https://api.kit.com/v4"
HEADERS = {
    "X-Kit-Api-Key": API_KEY,
    "Accept": "application/json",
    "Content-Type": "application/json",
}

SURVEY_DATE = pd.Timestamp("2026-01-25")
SURVEY_WINDOW_END = pd.Timestamp("2026-03-02")
PRE_SURVEY_WINDOW_DAYS = 38

POST_EVENTS_FILE = GENERATED / "lead_magnet_post_survey_events.csv"
OUT_EVENTS = GENERATED / "lead_magnet_origin_post_survey_events.csv"
OUT_SUMMARY = GENERATED / "lead_magnet_origin_post_survey_summary.csv"
OUT_PREPOST_SUMMARY = GENERATED / "lead_magnet_prepost_window_summary.csv"
OUT_POSTWINDOW_NONLEAD_SUMMARY = GENERATED / "lead_magnet_postwindow_nonlead_summary.csv"
OUT_CHART = OUT / "AD2_lead_magnet_origin_metrics.png"
BROADCAST_MAP_FILE = GENERATED / "lead_magnet_broadcast_mapping.csv"

RESPONSE_KEYWORDS = ["i want", "i'm", "i work at a company", "i own/run", "i own", "i'm building"]
ROADMAP_REFERRERS = {"lonelyoctopus.com/ai-sprint-roadmap"}


def has_survey_response(tags_str: str) -> bool:
    if pd.isna(tags_str):
        return False
    t = str(tags_str).lower()
    return any(k in t for k in RESPONSE_KEYWORDS)


def has_roadmap_referrer(referrer_str: str) -> bool:
    if pd.isna(referrer_str):
        return False
    ref = str(referrer_str).strip().lower()
    return ref in ROADMAP_REFERRERS


def rate_limited_post(url: str, payload: dict, timeout: int = 30):
    for attempt in range(4):
        r = requests.post(url, headers=HEADERS, json=payload, timeout=timeout)
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", 62))
            time.sleep(wait)
            continue
        if r.status_code >= 500 and attempt < 3:
            time.sleep(1 + attempt)
            continue
        return r
    return r


def fetch_event_emails_for_broadcast(event_type: str, broadcast_id: int) -> set[str]:
    payload_base = {
        "all": [{
            "type": event_type,
            "count_greater_than": 0,
            "any": [{"type": "broadcasts", "ids": [int(broadcast_id)]}],
        }]
    }
    emails: set[str] = set()
    cursor = None
    while True:
        payload = dict(payload_base)
        if cursor:
            payload["after"] = cursor
        r = rate_limited_post(f"{API_BASE}/subscribers/filter", payload=payload)
        if r.status_code != 200:
            raise RuntimeError(f"Filter API error for {event_type}/{broadcast_id}: {r.status_code} {r.text[:250]}")
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


def ztest_prop_diff(success_a: int, total_a: int, success_b: int, total_b: int) -> float:
    if min(total_a, total_b) <= 0:
        return np.nan
    p_pool = (success_a + success_b) / (total_a + total_b)
    se = np.sqrt(p_pool * (1 - p_pool) * ((1 / total_a) + (1 / total_b)))
    if se == 0:
        return np.nan
    z = (success_a / total_a - success_b / total_b) / se
    return float(2 * scipy_stats.norm.sf(abs(z)))


def load_window_broadcasts(start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    if not BROADCAST_MAP_FILE.exists():
        raise FileNotFoundError("lead_magnet_broadcast_mapping.csv is required for window mapping.")
    mp = pd.read_csv(BROADCAST_MAP_FILE)
    mp["date_dt"] = pd.to_datetime(mp["date"], errors="coerce")
    mp["category"] = mp["category"].replace("Sles", "Sales")
    mp = mp[
        (mp["date_dt"] >= start_date)
        & (mp["date_dt"] <= end_date)
        & (mp["category"].isin(["Value", "Sales"]))
        & (pd.to_numeric(mp["recipients"], errors="coerce") > 900)
        & (mp["api_broadcast_id"].notna())
    ].copy()
    if mp.empty:
        return pd.DataFrame(columns=["broadcast_id", "date_dt", "category", "subject"])
    mp["broadcast_id"] = mp["api_broadcast_id"].astype(int)
    mp = mp.sort_values("date_dt").drop_duplicates(subset=["broadcast_id"], keep="last")
    return mp[["broadcast_id", "date_dt", "category", "subject"]]


def main():
    print("=" * 76)
    print("LEAD MAGNET ORIGIN COHORTS — POST-SURVEY ENGAGEMENT")
    print("=" * 76)

    if not POST_EVENTS_FILE.exists():
        raise FileNotFoundError("lead_magnet_post_survey_events.csv is required.")

    subs = pd.read_csv(BASE / "Confirmed Subscribers.csv", usecols=["email", "created_at", "tags", "referrer"])
    subs["email_lower"] = subs["email"].astype(str).str.lower().str.strip()
    subs = subs.drop_duplicates("email_lower", keep="first")
    subs["created_at_dt"] = pd.to_datetime(subs["created_at"], utc=True, errors="coerce").dt.tz_convert(None)
    subs["created_date"] = subs["created_at_dt"].dt.normalize()
    subs = subs.dropna(subset=["created_at_dt"]).copy()
    subs["has_response_tag"] = subs["tags"].apply(has_survey_response)
    subs["has_roadmap_referrer"] = subs["referrer"].apply(has_roadmap_referrer)

    survey_cutoff = SURVEY_DATE.normalize()
    pre_window_start = (SURVEY_DATE - pd.Timedelta(days=PRE_SURVEY_WINDOW_DAYS)).normalize()
    survey_window_end = SURVEY_WINDOW_END.normalize()
    lead = subs[
        (subs["created_date"] >= survey_cutoff)
        & (subs["created_date"] <= survey_window_end)
        & (subs["has_response_tag"])
        & (subs["has_roadmap_referrer"])
    ][["email_lower", "created_at_dt"]].copy()
    post_nonlead = subs[
        (subs["created_date"] >= survey_cutoff)
        & (subs["created_date"] <= survey_window_end)
        & (~(subs["has_response_tag"] & subs["has_roadmap_referrer"]))
    ][["email_lower", "created_at_dt"]].copy()
    existing = subs[
        (subs["created_date"] >= pre_window_start)
        & (subs["created_date"] < survey_cutoff)
    ][["email_lower", "created_at_dt"]].copy()
    ambiguous = subs[
        (subs["created_date"] >= survey_cutoff)
        & (subs["created_date"] <= survey_window_end)
        & (subs["has_response_tag"])
        & (~subs["has_roadmap_referrer"])
    ][["email_lower", "created_at_dt"]].copy()
    lead_emails = set(lead["email_lower"])
    post_nonlead_emails = set(post_nonlead["email_lower"])
    existing_emails = set(existing["email_lower"])
    ambiguous_emails = set(ambiguous["email_lower"])

    print(f"Lead-magnet signups (strict inferred): {len(lead_emails):,}")
    print(f"Post-window non-lead signups: {len(post_nonlead_emails):,}")
    print(f"Ambiguous post-survey signups with response tag but non-roadmap referrer (excluded): {len(ambiguous_emails):,}")
    print(f"Existing subscribers in pre-survey {PRE_SURVEY_WINDOW_DAYS}-day window: {len(existing_emails):,}")
    print(f"Pre-survey signup window used: {pre_window_start.date()} to {(survey_cutoff - pd.Timedelta(days=1)).date()} (inclusive)")
    print(f"Post-survey signup window used: {survey_cutoff.date()} to {survey_window_end.date()} (inclusive)")

    pre_broadcasts = load_window_broadcasts(pre_window_start, survey_cutoff - pd.Timedelta(days=1))
    print(f"Pre-window mapped broadcasts: {len(pre_broadcasts):,}")

    post = pd.read_csv(POST_EVENTS_FILE).drop_duplicates("broadcast_id").copy()
    post["broadcast_id"] = post["broadcast_id"].astype(int)
    post["date_dt"] = pd.to_datetime(post["date"], errors="coerce")
    post = post.sort_values("date_dt")
    print(f"Post-survey matched broadcasts: {len(post):,}")

    opens_cache: dict[int, set[str]] = {}
    clicks_cache: dict[int, set[str]] = {}

    def _get_openers(_bid: int) -> set[str]:
        if _bid not in opens_cache:
            opens_cache[_bid] = fetch_event_emails_for_broadcast("opens", _bid)
        return opens_cache[_bid]

    def _get_clickers(_bid: int) -> set[str]:
        if _bid not in clicks_cache:
            clicks_cache[_bid] = fetch_event_emails_for_broadcast("clicks", _bid)
        return clicks_cache[_bid]

    rows = []
    for _, r in post.iterrows():
        bid = int(r["broadcast_id"])
        b_date = r["date_dt"]
        category = str(r["category"])
        subject = str(r["subject"])

        openers = _get_openers(bid)
        clickers = _get_clickers(bid) if category == "Sales" else set()

        lead_eligible = set(lead[lead["created_at_dt"] <= b_date]["email_lower"])
        post_nonlead_eligible = set(post_nonlead[post_nonlead["created_at_dt"] <= b_date]["email_lower"])
        existing_eligible = set(existing[existing["created_at_dt"] <= b_date]["email_lower"])

        lead_open = len(openers & lead_eligible)
        post_nonlead_open = len(openers & post_nonlead_eligible)
        existing_open = len(openers & existing_eligible)
        lead_click = len(clickers & lead_eligible)
        post_nonlead_click = len(clickers & post_nonlead_eligible)
        existing_click = len(clickers & existing_eligible)

        lead_or = (lead_open / len(lead_eligible) * 100) if len(lead_eligible) else np.nan
        post_nonlead_or = (post_nonlead_open / len(post_nonlead_eligible) * 100) if len(post_nonlead_eligible) else np.nan
        existing_or = (existing_open / len(existing_eligible) * 100) if len(existing_eligible) else np.nan
        lead_ctor = (lead_click / lead_open * 100) if (category == "Sales" and lead_open > 0) else np.nan
        post_nonlead_ctor = (post_nonlead_click / post_nonlead_open * 100) if (category == "Sales" and post_nonlead_open > 0) else np.nan
        existing_ctor = (existing_click / existing_open * 100) if (category == "Sales" and existing_open > 0) else np.nan

        rows.append({
            "broadcast_id": bid,
            "date": b_date.date().isoformat() if pd.notna(b_date) else "",
            "subject": subject,
            "category": category,
            "lead_eligible": len(lead_eligible),
            "post_nonlead_eligible": len(post_nonlead_eligible),
            "existing_eligible": len(existing_eligible),
            "lead_open": lead_open,
            "post_nonlead_open": post_nonlead_open,
            "existing_open": existing_open,
            "lead_click": lead_click,
            "post_nonlead_click": post_nonlead_click,
            "existing_click": existing_click,
            "lead_open_rate": lead_or,
            "post_nonlead_open_rate": post_nonlead_or,
            "existing_open_rate": existing_or,
            "lead_sales_ctor": lead_ctor,
            "post_nonlead_sales_ctor": post_nonlead_ctor,
            "existing_sales_ctor": existing_ctor,
        })
        print(
            f"{b_date.date()} | {category:<5} | id={bid} | "
            f"lead OR={lead_or:.2f}% ({lead_open}/{len(lead_eligible)}) | "
            f"existing OR={existing_or:.2f}% ({existing_open}/{len(existing_eligible)})"
        )

    pre_rows = []
    for _, r in pre_broadcasts.iterrows():
        bid = int(r["broadcast_id"])
        b_date = r["date_dt"]
        category = str(r["category"])
        subject = str(r["subject"])

        openers = _get_openers(bid)
        clickers = _get_clickers(bid) if category == "Sales" else set()
        existing_eligible = set(existing[existing["created_at_dt"] <= b_date]["email_lower"])
        existing_open = len(openers & existing_eligible)
        existing_click = len(clickers & existing_eligible)
        existing_or = (existing_open / len(existing_eligible) * 100) if len(existing_eligible) else np.nan
        existing_ctor = (existing_click / existing_open * 100) if (category == "Sales" and existing_open > 0) else np.nan

        pre_rows.append({
            "broadcast_id": bid,
            "date": b_date.date().isoformat() if pd.notna(b_date) else "",
            "subject": subject,
            "category": category,
            "existing_eligible": len(existing_eligible),
            "existing_open": existing_open,
            "existing_click": existing_click,
            "existing_open_rate": existing_or,
            "existing_sales_ctor": existing_ctor,
        })

    events = pd.DataFrame(rows)
    events.to_csv(OUT_EVENTS, index=False)
    print(f"Saved {OUT_EVENTS.name} ({len(events):,} rows)")

    out_rows = []
    for cat in ["Value", "Sales"]:
        sub = events[events["category"] == cat].copy()
        if sub.empty:
            continue

        lead_open = int(sub["lead_open"].sum())
        existing_open = int(sub["existing_open"].sum())
        lead_eligible = int(sub["lead_eligible"].sum())
        existing_eligible = int(sub["existing_eligible"].sum())
        lead_or = (lead_open / lead_eligible * 100) if lead_eligible else np.nan
        existing_or = (existing_open / existing_eligible * 100) if existing_eligible else np.nan
        p_or = ztest_prop_diff(lead_open, lead_eligible, existing_open, existing_eligible)

        out_rows.append({
            "metric": f"{cat} open rate",
            "category": cat,
            "n_broadcasts": int(sub["broadcast_id"].nunique()),
            "lead_magnet_value": lead_or,
            "existing_value": existing_or,
            "delta_pp": lead_or - existing_or if pd.notna(lead_or) and pd.notna(existing_or) else np.nan,
            "lead_success": lead_open,
            "lead_total": lead_eligible,
            "existing_success": existing_open,
            "existing_total": existing_eligible,
            "p_value": p_or,
        })

        if cat == "Sales":
            lead_click = int(sub["lead_click"].sum())
            existing_click = int(sub["existing_click"].sum())
            lead_ctor = (lead_click / lead_open * 100) if lead_open else np.nan
            existing_ctor = (existing_click / existing_open * 100) if existing_open else np.nan
            p_ctor = ztest_prop_diff(lead_click, lead_open, existing_click, existing_open)
            out_rows.append({
                "metric": "Sales CTOR",
                "category": "Sales",
                "n_broadcasts": int(sub["broadcast_id"].nunique()),
                "lead_magnet_value": lead_ctor,
                "existing_value": existing_ctor,
                "delta_pp": lead_ctor - existing_ctor if pd.notna(lead_ctor) and pd.notna(existing_ctor) else np.nan,
                "lead_success": lead_click,
                "lead_total": lead_open,
                "existing_success": existing_click,
                "existing_total": existing_open,
                "p_value": p_ctor,
            })

    summary = pd.DataFrame(out_rows)
    summary.to_csv(OUT_SUMMARY, index=False)
    print(f"Saved {OUT_SUMMARY.name} ({len(summary):,} rows)")

    # Additional diagnostic requested: pre-window (existing cohort) vs post-window (lead cohort).
    pre_events = pd.DataFrame(pre_rows)
    prepost_rows = []
    for cat in ["Value", "Sales"]:
        pre_cat = pre_events[pre_events["category"] == cat] if len(pre_events) else pd.DataFrame()
        post_cat = events[events["category"] == cat] if len(events) else pd.DataFrame()

        pre_open = int(pre_cat["existing_open"].sum()) if len(pre_cat) else 0
        pre_eligible = int(pre_cat["existing_eligible"].sum()) if len(pre_cat) else 0
        post_open = int(post_cat["lead_open"].sum()) if len(post_cat) else 0
        post_eligible = int(post_cat["lead_eligible"].sum()) if len(post_cat) else 0

        pre_or = (pre_open / pre_eligible * 100) if pre_eligible else np.nan
        post_or = (post_open / post_eligible * 100) if post_eligible else np.nan
        p_or = ztest_prop_diff(post_open, post_eligible, pre_open, pre_eligible) if (pre_eligible and post_eligible) else np.nan

        prepost_rows.append({
            "metric": f"{cat} open rate",
            "pre_window_start": pre_window_start.date().isoformat(),
            "pre_window_end": (survey_cutoff - pd.Timedelta(days=1)).date().isoformat(),
            "post_window_start": survey_cutoff.date().isoformat(),
            "post_window_end": survey_window_end.date().isoformat(),
            "pre_n_broadcasts": int(pre_cat["broadcast_id"].nunique()) if len(pre_cat) else 0,
            "post_n_broadcasts": int(post_cat["broadcast_id"].nunique()) if len(post_cat) else 0,
            "pre_value": pre_or,
            "post_value": post_or,
            "delta_pp": post_or - pre_or if pd.notna(pre_or) and pd.notna(post_or) else np.nan,
            "pre_success": pre_open,
            "pre_total": pre_eligible,
            "post_success": post_open,
            "post_total": post_eligible,
            "p_value": p_or,
        })

    pre_sales = pre_events[pre_events["category"] == "Sales"] if len(pre_events) else pd.DataFrame()
    post_sales = events[events["category"] == "Sales"] if len(events) else pd.DataFrame()
    pre_sales_click = int(pre_sales["existing_click"].sum()) if len(pre_sales) else 0
    pre_sales_open = int(pre_sales["existing_open"].sum()) if len(pre_sales) else 0
    post_sales_click = int(post_sales["lead_click"].sum()) if len(post_sales) else 0
    post_sales_open = int(post_sales["lead_open"].sum()) if len(post_sales) else 0
    pre_sales_ctor = (pre_sales_click / pre_sales_open * 100) if pre_sales_open else np.nan
    post_sales_ctor = (post_sales_click / post_sales_open * 100) if post_sales_open else np.nan
    p_sales_ctor = ztest_prop_diff(post_sales_click, post_sales_open, pre_sales_click, pre_sales_open) if (post_sales_open and pre_sales_open) else np.nan
    prepost_rows.append({
        "metric": "Sales CTOR",
        "pre_window_start": pre_window_start.date().isoformat(),
        "pre_window_end": (survey_cutoff - pd.Timedelta(days=1)).date().isoformat(),
        "post_window_start": survey_cutoff.date().isoformat(),
        "post_window_end": survey_window_end.date().isoformat(),
        "pre_n_broadcasts": int(pre_sales["broadcast_id"].nunique()) if len(pre_sales) else 0,
        "post_n_broadcasts": int(post_sales["broadcast_id"].nunique()) if len(post_sales) else 0,
        "pre_value": pre_sales_ctor,
        "post_value": post_sales_ctor,
        "delta_pp": post_sales_ctor - pre_sales_ctor if pd.notna(pre_sales_ctor) and pd.notna(post_sales_ctor) else np.nan,
        "pre_success": pre_sales_click,
        "pre_total": pre_sales_open,
        "post_success": post_sales_click,
        "post_total": post_sales_open,
        "p_value": p_sales_ctor,
    })
    prepost_summary = pd.DataFrame(prepost_rows)
    prepost_summary.to_csv(OUT_PREPOST_SUMMARY, index=False)
    print(f"Saved {OUT_PREPOST_SUMMARY.name} ({len(prepost_summary):,} rows)")

    # Additional diagnostic requested: within same post-signup window, lead-magnet vs non-lead signups.
    postwindow_rows = []
    for cat in ["Value", "Sales"]:
        sub = events[events["category"] == cat].copy()
        if sub.empty:
            continue
        lead_open = int(sub["lead_open"].sum())
        lead_eligible = int(sub["lead_eligible"].sum())
        nonlead_open = int(sub["post_nonlead_open"].sum())
        nonlead_eligible = int(sub["post_nonlead_eligible"].sum())
        lead_or = (lead_open / lead_eligible * 100) if lead_eligible else np.nan
        nonlead_or = (nonlead_open / nonlead_eligible * 100) if nonlead_eligible else np.nan
        p_or = ztest_prop_diff(lead_open, lead_eligible, nonlead_open, nonlead_eligible) if (lead_eligible and nonlead_eligible) else np.nan
        postwindow_rows.append({
            "metric": f"{cat} open rate",
            "n_broadcasts": int(sub["broadcast_id"].nunique()),
            "lead_magnet_value": lead_or,
            "non_lead_value": nonlead_or,
            "delta_pp": lead_or - nonlead_or if pd.notna(lead_or) and pd.notna(nonlead_or) else np.nan,
            "lead_success": lead_open,
            "lead_total": lead_eligible,
            "non_lead_success": nonlead_open,
            "non_lead_total": nonlead_eligible,
            "p_value": p_or,
        })
        if cat == "Sales":
            lead_click = int(sub["lead_click"].sum())
            nonlead_click = int(sub["post_nonlead_click"].sum())
            lead_ctor = (lead_click / lead_open * 100) if lead_open else np.nan
            nonlead_ctor = (nonlead_click / nonlead_open * 100) if nonlead_open else np.nan
            p_ctor = ztest_prop_diff(lead_click, lead_open, nonlead_click, nonlead_open) if (lead_open and nonlead_open) else np.nan
            postwindow_rows.append({
                "metric": "Sales CTOR",
                "n_broadcasts": int(sub["broadcast_id"].nunique()),
                "lead_magnet_value": lead_ctor,
                "non_lead_value": nonlead_ctor,
                "delta_pp": lead_ctor - nonlead_ctor if pd.notna(lead_ctor) and pd.notna(nonlead_ctor) else np.nan,
                "lead_success": lead_click,
                "lead_total": lead_open,
                "non_lead_success": nonlead_click,
                "non_lead_total": nonlead_open,
                "p_value": p_ctor,
            })
    postwindow_summary = pd.DataFrame(postwindow_rows)
    postwindow_summary.to_csv(OUT_POSTWINDOW_NONLEAD_SUMMARY, index=False)
    print(f"Saved {OUT_POSTWINDOW_NONLEAD_SUMMARY.name} ({len(postwindow_summary):,} rows)")

    plot = summary[summary["metric"].isin(["Value open rate", "Sales open rate", "Sales CTOR"])].copy()
    metric_order = ["Value open rate", "Sales open rate", "Sales CTOR"]
    plot["metric"] = pd.Categorical(plot["metric"], categories=metric_order, ordered=True)
    plot = plot.sort_values("metric")

    x = np.arange(len(plot))
    width = 0.34
    fig, ax = plt.subplots(figsize=(10.5, 5.8))
    ax.bar(x - width / 2, plot["lead_magnet_value"], width, label="Lead-magnet signups (strict inferred)", color="#10B981")
    ax.bar(x + width / 2, plot["existing_value"], width, label="Existing subscribers", color="#6366F1")

    for i, row in enumerate(plot.itertuples(index=False)):
        ax.text(i - width / 2, row.lead_magnet_value + 0.8, f"{row.lead_magnet_value:.1f}%", ha="center", va="bottom", fontsize=9)
        ax.text(i + width / 2, row.existing_value + 0.8, f"{row.existing_value:.1f}%", ha="center", va="bottom", fontsize=9)
        ax.text(i, max(row.lead_magnet_value, row.existing_value) + 4.0, f"Delta: {row.delta_pp:+.1f}pp", ha="center", va="bottom", fontsize=9, color="#374151")

    ax.set_xticks(x)
    ax.set_xticklabels(["Value OR", "Sales OR", "Sales CTOR"])
    ax.set_ylabel("Rate (%)")
    ax.set_ylim(0, max(float(plot["lead_magnet_value"].max()), float(plot["existing_value"].max())) + 12)
    ax.set_title("Post-Survey Engagement by Subscriber Origin")
    ax.grid(axis="y", alpha=0.2)
    ax.legend()
    fig.tight_layout()
    fig.savefig(OUT_CHART, dpi=180)
    plt.close()
    print(f"Saved chart {OUT_CHART.name}")

    print("\nSummary:")
    print(summary[["metric", "lead_magnet_value", "existing_value", "delta_pp", "p_value"]].to_string(index=False))
    print("\nPre-window existing cohort vs post-window lead cohort:")
    print(prepost_summary[["metric", "pre_value", "post_value", "delta_pp", "p_value", "pre_n_broadcasts", "post_n_broadcasts"]].to_string(index=False))
    print("\nWithin post window: lead-magnet vs non-lead signups:")
    print(postwindow_summary[["metric", "lead_magnet_value", "non_lead_value", "delta_pp", "p_value"]].to_string(index=False))


if __name__ == "__main__":
    main()

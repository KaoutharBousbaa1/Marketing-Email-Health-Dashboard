"""
Signup Cohort Evolution Analysis
================================

Tracks monthly Open Rate and CTOR by subscriber signup-date cohorts to test
whether older cohorts are disengaging more than newer cohorts.

Outputs:
- signup_cohort_monthly_evolution.csv
- signup_cohort_failed_ids.csv
- charts/AL_signup_cohort_openrate_trend.png
- charts/AM_signup_cohort_ctor_trend.png
- charts/AN_oldest_vs_newest_gap.png
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

MAPPING_FILE = GENERATED / "lead_magnet_broadcast_mapping.csv"
SUB_CACHE = GENERATED / "all_subscribers_created_at.csv"
OUT_MONTHLY = GENERATED / "signup_cohort_monthly_evolution.csv"
OUT_FAILED = GENERATED / "signup_cohort_failed_ids.csv"

CHART_OR = OUT / "AL_signup_cohort_openrate_trend.png"
CHART_CTOR = OUT / "AM_signup_cohort_ctor_trend.png"
CHART_GAP = OUT / "AN_oldest_vs_newest_gap.png"

COLORS = ["#2563EB", "#10B981", "#F59E0B", "#EF4444"]


def rate_limited_get(url: str, params: dict | None = None, timeout: int = 20):
    for attempt in range(6):
        try:
            r = requests.get(url, headers=HEADERS, params=params, timeout=timeout)
        except requests.RequestException:
            time.sleep(1 + attempt)
            continue
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", 62))
            print(f"GET 429. Waiting {wait}s …")
            time.sleep(wait)
            continue
        return r
    raise RuntimeError("GET failed repeatedly due network/API errors.")


def rate_limited_post(url: str, payload: dict | None = None, timeout: int = 12):
    for attempt in range(3):
        try:
            r = requests.post(url, headers=HEADERS, json=payload, timeout=timeout)
        except requests.RequestException:
            time.sleep(0.5 + attempt * 0.5)
            continue
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", 62))
            print(f"POST 429. Waiting {wait}s …")
            time.sleep(wait)
            continue
        if r.status_code >= 500 and attempt < 2:
            continue
        return r
    raise RuntimeError("POST failed repeatedly due network/API errors.")


def fetch_all_subscribers() -> pd.DataFrame:
    if SUB_CACHE.exists():
        df = pd.read_csv(SUB_CACHE)
        df["created_at_dt"] = pd.to_datetime(df["created_at"], utc=True, errors="coerce")
        return df

    rows = []
    cursor = None
    while True:
        params = {"status": "all", "per_page": 1000}
        if cursor:
            params["after"] = cursor
        r = rate_limited_get(f"{API_BASE}/subscribers", params=params)
        if r.status_code != 200:
            raise RuntimeError(f"Subscribers fetch failed: {r.status_code} {r.text[:250]}")
        j = r.json()
        subs = j.get("subscribers", [])
        for s in subs:
            rows.append({
                "id": s.get("id"),
                "email": s.get("email_address", "").lower().strip(),
                "state": s.get("state"),
                "created_at": s.get("created_at"),
            })
        p = j.get("pagination", {})
        if not p.get("has_next_page"):
            break
        cursor = p.get("end_cursor")
    df = pd.DataFrame(rows)
    df.to_csv(SUB_CACHE, index=False)
    df["created_at_dt"] = pd.to_datetime(df["created_at"], utc=True, errors="coerce")
    return df


def fetch_event_emails(event_type: str, broadcast_ids: list[int], failed_tracker: list[dict]) -> set[str]:
    if len(broadcast_ids) == 0:
        return set()

    emails = set()

    def _fetch_chunk(ids_chunk: list[int]) -> set[str]:
        out = set()
        payload_base = {
            "all": [{
                "type": event_type,
                "count_greater_than": 0,
                "any": [{"type": "broadcasts", "ids": [int(x) for x in ids_chunk]}],
            }]
        }
        cursor = None
        while True:
            payload = dict(payload_base)
            if cursor:
                payload["after"] = cursor
            r = rate_limited_post(f"{API_BASE}/subscribers/filter", payload=payload)
            if r.status_code != 200:
                if len(ids_chunk) > 1:
                    mid = len(ids_chunk) // 2
                    return _fetch_chunk(ids_chunk[:mid]) | _fetch_chunk(ids_chunk[mid:])
                failed_tracker.append({
                    "event_type": event_type,
                    "broadcast_id": int(ids_chunk[0]),
                    "status_code": r.status_code,
                    "error_preview": r.text[:180],
                })
                return set()
            j = r.json()
            for s in j.get("subscribers", []):
                e = s.get("email_address")
                if e:
                    out.add(e.lower().strip())
            p = j.get("pagination", {})
            if not p.get("has_next_page"):
                break
            cursor = p.get("end_cursor")
        return out

    chunk_size = 12
    chunks = [broadcast_ids[i:i + chunk_size] for i in range(0, len(broadcast_ids), chunk_size)]
    for ch in chunks:
        emails |= _fetch_chunk(ch)
    return emails


def main():
    print("=" * 74)
    print("SIGNUP COHORT EVOLUTION — MONTHLY OPEN RATE & CTOR")
    print("=" * 74)

    if not MAPPING_FILE.exists():
        raise FileNotFoundError(
            "lead_magnet_broadcast_mapping.csv not found. Run analysis_group_ab_monthly.py first."
        )

    # Broadcast universe (Value + Sales only)
    mp = pd.read_csv(MAPPING_FILE)
    mp["date"] = pd.to_datetime(mp["date"], errors="coerce")
    mp["category"] = mp["category"].replace("Sles", "Sales")
    mp = mp[
        mp["api_broadcast_id"].notna()
        & mp["category"].isin(["Value", "Sales"])
        & (mp["date"] >= pd.Timestamp("2025-02-01"))
    ].copy()
    mp["api_broadcast_id"] = mp["api_broadcast_id"].astype(int)
    mp = mp.sort_values("date").drop_duplicates(subset=["api_broadcast_id"], keep="last")
    mp["month"] = mp["date"].dt.to_period("M")

    print(f"Using {len(mp):,} unique matched broadcasts.")

    # Subscribers + cohort assignment by signup-date quartile
    subs = fetch_all_subscribers()
    subs = subs[subs["email"].str.len() > 3].copy()
    subs = subs.dropna(subset=["created_at_dt"]).copy()
    analysis_end = mp["date"].max()
    base_for_cut = subs[subs["created_at_dt"].dt.tz_convert(None) <= analysis_end].copy()

    q1 = base_for_cut["created_at_dt"].quantile(0.25)
    q2 = base_for_cut["created_at_dt"].quantile(0.50)
    q3 = base_for_cut["created_at_dt"].quantile(0.75)

    def assign_cohort(dt):
        if dt <= q1:
            return "Q1 Oldest Signups"
        if dt <= q2:
            return "Q2 Mid-Old Signups"
        if dt <= q3:
            return "Q3 Mid-New Signups"
        return "Q4 Newest Signups"

    subs["cohort"] = subs["created_at_dt"].apply(assign_cohort)
    cohort_order = [
        "Q1 Oldest Signups",
        "Q2 Mid-Old Signups",
        "Q3 Mid-New Signups",
        "Q4 Newest Signups",
    ]
    cohort_sets = {
        c: set(subs[subs["cohort"] == c]["email"])
        for c in cohort_order
    }
    cohort_created = {
        c: subs[subs["cohort"] == c]["created_at_dt"].dt.tz_convert(None)
        for c in cohort_order
    }
    print("Cohort sizes:")
    for c in cohort_order:
        print(f"  {c}: {len(cohort_sets[c]):,}")

    months = pd.period_range(mp["month"].min(), mp["month"].max(), freq="M")
    failed_rows: list[dict] = []
    rows = []

    for month in months:
        b_ids = mp[mp["month"] == month]["api_broadcast_id"].astype(int).tolist()
        opened = fetch_event_emails("opens", b_ids, failed_rows) if b_ids else set()
        clicked = fetch_event_emails("clicks", b_ids, failed_rows) if b_ids else set()
        month_end = (month + 1).to_timestamp(how="start") - pd.Timedelta(seconds=1)

        for c in cohort_order:
            sent_n = int((cohort_created[c] <= month_end).sum())
            open_n = len(opened & cohort_sets[c])
            click_n = len(clicked & cohort_sets[c])
            rows.append({
                "month": str(month),
                "month_dt": month.to_timestamp(how="start"),
                "cohort": c,
                "broadcasts_matched": len(b_ids),
                "sent_n": sent_n,
                "open_n": open_n,
                "click_n": click_n,
                "open_rate": (open_n / sent_n * 100) if sent_n > 0 else np.nan,
                "ctor": (click_n / open_n * 100) if open_n > 0 else np.nan,
            })

        print(
            f"{month}: broadcasts={len(b_ids):>2} | opened={len(opened):>5} | clicked={len(clicked):>5}"
        )

    out = pd.DataFrame(rows).sort_values(["month_dt", "cohort"])
    out.to_csv(OUT_MONTHLY, index=False)
    print(f"\nSaved {OUT_MONTHLY.name} ({len(out):,} rows)")

    failed_df = pd.DataFrame(failed_rows).drop_duplicates(["event_type", "broadcast_id"])
    if len(failed_df):
        failed_df = failed_df.sort_values(["broadcast_id", "event_type"])
    failed_df.to_csv(OUT_FAILED, index=False)
    print(f"Saved {OUT_FAILED.name} ({len(failed_df):,} failing id-event rows)")

    # ── Charts ───────────────────────────────────────────────────────────────
    # AL: Open-rate trend by cohort
    fig, ax = plt.subplots(figsize=(13.5, 6.5))
    for c, col in zip(cohort_order, COLORS):
        d = out[out["cohort"] == c].sort_values("month_dt")
        ax.plot(d["month_dt"], d["open_rate"], marker="o", linewidth=2.2, color=col, label=c)
    ax.set_title("Monthly Open Rate Evolution by Signup Cohort")
    ax.set_ylabel("Open Rate (%)")
    ax.set_xlabel("Month")
    ax.set_ylim(0, 100)
    ax.grid(axis="y", alpha=0.2)
    ax.legend(fontsize=8.5)
    fig.tight_layout()
    fig.savefig(CHART_OR)
    plt.close()

    # AM: CTOR trend by cohort
    fig, ax = plt.subplots(figsize=(13.5, 6.5))
    for c, col in zip(cohort_order, COLORS):
        d = out[out["cohort"] == c].sort_values("month_dt")
        ax.plot(d["month_dt"], d["ctor"], marker="o", linewidth=2.2, color=col, label=c)
    ax.set_title("Monthly CTOR Evolution by Signup Cohort")
    ax.set_ylabel("CTOR (%)")
    ax.set_xlabel("Month")
    ax.grid(axis="y", alpha=0.2)
    ax.legend(fontsize=8.5)
    fig.tight_layout()
    fig.savefig(CHART_CTOR)
    plt.close()

    # AN: Oldest vs newest cohort gap over time
    old = out[out["cohort"] == cohort_order[0]].sort_values("month_dt").copy()
    new = out[out["cohort"] == cohort_order[-1]].sort_values("month_dt").copy()
    g = old[["month_dt", "open_rate", "ctor"]].merge(
        new[["month_dt", "open_rate", "ctor"]],
        on="month_dt",
        suffixes=("_oldest", "_newest"),
    )
    g["or_gap_old_minus_new"] = g["open_rate_oldest"] - g["open_rate_newest"]
    g["ctor_gap_old_minus_new"] = g["ctor_oldest"] - g["ctor_newest"]

    fig, axes = plt.subplots(2, 1, figsize=(13.5, 8.2), sharex=True)
    axes[0].axhline(0, color="#64748B", linewidth=1.1)
    axes[0].bar(
        g["month_dt"],
        g["or_gap_old_minus_new"],
        color=g["or_gap_old_minus_new"].apply(lambda x: "#EF4444" if x < 0 else "#10B981"),
        width=22,
        alpha=0.85,
    )
    axes[0].set_title("Open-Rate Gap: Oldest Cohort - Newest Cohort (pp)")
    axes[0].set_ylabel("Gap (pp)")
    axes[0].grid(axis="y", alpha=0.18)

    axes[1].axhline(0, color="#64748B", linewidth=1.1)
    axes[1].bar(
        g["month_dt"],
        g["ctor_gap_old_minus_new"],
        color=g["ctor_gap_old_minus_new"].apply(lambda x: "#EF4444" if x < 0 else "#10B981"),
        width=22,
        alpha=0.85,
    )
    axes[1].set_title("CTOR Gap: Oldest Cohort - Newest Cohort (pp)")
    axes[1].set_ylabel("Gap (pp)")
    axes[1].set_xlabel("Month")
    axes[1].grid(axis="y", alpha=0.18)

    fig.suptitle("Disengagement Check: Are Older Subscribers Falling Behind Newer Ones?")
    fig.tight_layout()
    fig.savefig(CHART_GAP)
    plt.close()

    print("Saved charts:")
    print(f"  - {CHART_OR.name}")
    print(f"  - {CHART_CTOR.name}")
    print(f"  - {CHART_GAP.name}")


if __name__ == "__main__":
    main()

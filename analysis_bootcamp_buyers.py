"""
Bootcamp Buyers vs Non-Buyers Analysis
======================================

Goal
----
Among confirmed subscribers, compare people who bought at least one bootcamp
vs those who never bought:
1) Monthly Open Rate evolution
2) Monthly CTOR evolution
3) Subscriber age profile (old vs recent)

Data sources
------------
- Confirmed Subscribers.csv (subscriber tags + created_at + status)
- lead_magnet_broadcast_mapping.csv (matched broadcast IDs and dates)
- Kit v4 API /subscribers/filter (opens/clicks by broadcast IDs)

Outputs
-------
- bootcamp_buyers_monthly_evolution.csv
- bootcamp_buyers_summary.csv
- bootcamp_buyers_age_buckets.csv
- bootcamp_buyers_failed_ids.csv
- charts/AO_bootcamp_buyers_monthly_or_ctor.png
- charts/AP_bootcamp_buyers_age_profile.png
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


# ─── Paths / constants ────────────────────────────────────────────────────────
BASE = Path(__file__).resolve().parent
GENERATED = BASE / "data" / "generated"
GENERATED.mkdir(parents=True, exist_ok=True)
OUT = BASE / "charts"
OUT.mkdir(exist_ok=True)

CONFIRMED_FILE = BASE / "Confirmed Subscribers.csv"
MAPPING_FILE = GENERATED / "lead_magnet_broadcast_mapping.csv"

OUT_MONTHLY = GENERATED / "bootcamp_buyers_monthly_evolution.csv"
OUT_SUMMARY = GENERATED / "bootcamp_buyers_summary.csv"
OUT_AGE = GENERATED / "bootcamp_buyers_age_buckets.csv"
OUT_FAILED = GENERATED / "bootcamp_buyers_failed_ids.csv"

CHART_MONTHLY = OUT / "AO_bootcamp_buyers_monthly_or_ctor.png"
CHART_AGE = OUT / "AP_bootcamp_buyers_age_profile.png"

API_KEY = "kit_7d6b10fad06f88e1d0e47e45ef92e9cc"
API_BASE = "https://api.kit.com/v4"
HEADERS = {
    "X-Kit-Api-Key": API_KEY,
    "Accept": "application/json",
    "Content-Type": "application/json",
}

BUYER_TAGS = [
    "ai agent core [feb]",
    "ai agent core [oct]",
    "ai agent core [sept]",
    "ai agent bootcamp core [jul 2025]",
    "ai agent bootcamp core [apr 2025]",
]

GREEN = "#10B981"
ORANGE = "#F59E0B"
RED = "#EF4444"
MUTED = "#94A3B8"
BRAND = "#4F46E5"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "figure.dpi": 160,
})


def rate_limited_post(url: str, payload: dict | None = None, timeout: int = 30):
    for i in range(4):
        try:
            r = requests.post(url, headers=HEADERS, json=payload, timeout=timeout)
        except requests.RequestException:
            if i < 3:
                time.sleep(1 + i)
                continue
            raise
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", 62))
            print(f"  POST 429. Waiting {wait}s …")
            time.sleep(wait)
            continue
        if r.status_code >= 500 and i < 2:
            time.sleep(1 + i)
            continue
        return r
    return r


def has_buyer_tag(tags_str: str) -> bool:
    if pd.isna(tags_str):
        return False
    t = str(tags_str).lower()
    return any(tag in t for tag in BUYER_TAGS)


def fetch_event_emails(event_type: str, broadcast_ids: list[int], failed_rows: list[dict]) -> set[str]:
    if not broadcast_ids:
        return set()

    emails = set()

    def _fetch_chunk(ids_chunk: list[int]) -> set[str]:
        out = set()
        payload_base = {
            "all": [{
                "type": event_type,
                "count_greater_than": 0,
                "any": [{"type": "broadcasts", "ids": [int(x) for x in ids_chunk]}],
            }],
            "per_page": 1000,
        }

        cursor = None
        seen_cursors = set()
        page = 0
        while True:
            payload = dict(payload_base)
            if cursor:
                payload["after"] = cursor
            r = rate_limited_post(f"{API_BASE}/subscribers/filter", payload=payload)
            if r.status_code != 200:
                if len(ids_chunk) > 1:
                    mid = len(ids_chunk) // 2
                    return _fetch_chunk(ids_chunk[:mid]) | _fetch_chunk(ids_chunk[mid:])
                failed_rows.append({
                    "event_type": event_type,
                    "broadcast_id": int(ids_chunk[0]),
                    "status_code": r.status_code,
                    "error_preview": r.text[:200],
                })
                return set()

            j = r.json()
            page += 1
            if page % 5 == 0:
                print(f"    {event_type}: chunk {ids_chunk[0]}.. page {page} | accumulated={len(out):,}", flush=True)
            for s in j.get("subscribers", []):
                e = s.get("email_address")
                if e:
                    out.add(str(e).lower().strip())

            p = j.get("pagination", {})
            if not p.get("has_next_page"):
                break
            next_cursor = p.get("end_cursor")
            if not next_cursor or next_cursor in seen_cursors:
                # Safety break if API returns repeated/empty cursors.
                break
            seen_cursors.add(next_cursor)
            cursor = next_cursor
            if page >= 300:
                # Hard safety cap to prevent runaway loops.
                break
        return out

    # Conservative chunking to reduce API 500 on legacy IDs.
    for i in range(0, len(broadcast_ids), 12):
        ch = broadcast_ids[i:i + 12]
        print(f"  Fetching {event_type} for chunk {i//12 + 1}/{(len(broadcast_ids)-1)//12 + 1} (ids={len(ch)})", flush=True)
        emails |= _fetch_chunk(ch)
    return emails


def main():
    print("=" * 78, flush=True)
    print("BOOTCAMP BUYERS VS NON-BUYERS — MONTHLY OR/CTOR + AGE PROFILE", flush=True)
    print("=" * 78, flush=True)

    if not CONFIRMED_FILE.exists():
        raise FileNotFoundError(f"Missing file: {CONFIRMED_FILE}")
    if not MAPPING_FILE.exists():
        raise FileNotFoundError(f"Missing file: {MAPPING_FILE}")

    # ── Subscribers / segmentation ────────────────────────────────────────────
    subs = pd.read_csv(CONFIRMED_FILE)
    subs["email_lower"] = subs["email"].astype(str).str.lower().str.strip()
    subs["created_at_dt"] = pd.to_datetime(subs["created_at"], utc=True, errors="coerce")
    subs["is_active"] = subs["status"].astype(str).str.lower().eq("active")
    subs["is_buyer"] = subs["tags"].apply(has_buyer_tag)
    subs = subs[subs["is_active"]].copy()
    subs = subs.dropna(subset=["email_lower", "created_at_dt"]).copy()
    subs = subs.sort_values("created_at_dt").drop_duplicates(subset=["email_lower"], keep="last")

    buyers = subs[subs["is_buyer"]].copy()
    nonbuyers = subs[~subs["is_buyer"]].copy()
    buyer_set = set(buyers["email_lower"])
    nonbuyer_set = set(nonbuyers["email_lower"])

    print(f"\nActive confirmed subscribers: {len(subs):,}", flush=True)
    print(f"Buyers (>=1 bootcamp tag):   {len(buyers):,}", flush=True)
    print(f"Non-buyers:                  {len(nonbuyers):,}", flush=True)

    # ── Broadcast universe by month ───────────────────────────────────────────
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
    months = pd.period_range(mp["month"].min(), mp["month"].max(), freq="M")
    print(f"\nUsing {len(mp):,} matched unique broadcasts across {len(months)} months.", flush=True)

    # ── Monthly event pulls ───────────────────────────────────────────────────
    failed_rows: list[dict] = []
    rows = []

    buyer_created = buyers["created_at_dt"].dt.tz_convert(None)
    nonbuyer_created = nonbuyers["created_at_dt"].dt.tz_convert(None)

    for month in months:
        b_ids = mp[mp["month"] == month]["api_broadcast_id"].astype(int).tolist()
        month_end = (month + 1).to_timestamp(how="start") - pd.Timedelta(seconds=1)

        opened = fetch_event_emails("opens", b_ids, failed_rows) if b_ids else set()
        clicked = fetch_event_emails("clicks", b_ids, failed_rows) if b_ids else set()

        for label, created_series, group_set in [
            ("Buyer", buyer_created, buyer_set),
            ("Non-buyer", nonbuyer_created, nonbuyer_set),
        ]:
            eligible_n = int((created_series <= month_end).sum())
            open_n = len(opened & group_set)
            click_n = len(clicked & group_set)
            rows.append({
                "month": str(month),
                "month_dt": month.to_timestamp(how="start"),
                "group": label,
                "broadcasts_matched": len(b_ids),
                "eligible_n": eligible_n,
                "open_n": open_n,
                "click_n": click_n,
                "open_rate": (open_n / eligible_n * 100) if eligible_n > 0 else np.nan,
                "ctor": (click_n / open_n * 100) if open_n > 0 else np.nan,
            })

        print(
            f"{month}: broadcasts={len(b_ids):>2} | "
            f"Buyer OR={rows[-2]['open_rate']:>5.1f}% CTOR={rows[-2]['ctor']:>5.1f}% | "
            f"Non-buyer OR={rows[-1]['open_rate']:>5.1f}% CTOR={rows[-1]['ctor']:>5.1f}%"
        , flush=True)

    monthly = pd.DataFrame(rows)
    monthly.to_csv(OUT_MONTHLY, index=False)
    print(f"\nSaved monthly output: {OUT_MONTHLY.name} ({len(monthly)} rows)", flush=True)

    # ── Summary table ──────────────────────────────────────────────────────────
    piv = monthly.pivot_table(
        index=["month", "month_dt", "broadcasts_matched"],
        columns="group",
        values=["open_rate", "ctor", "eligible_n", "open_n", "click_n"],
        aggfunc="first",
    ).reset_index()
    piv.columns = [
        "_".join([str(x) for x in c if str(x) != ""]).rstrip("_")
        for c in piv.columns.to_flat_index()
    ]
    piv = piv.rename(columns={
        "open_rate_Buyer": "buyer_open_rate",
        "open_rate_Non-buyer": "nonbuyer_open_rate",
        "ctor_Buyer": "buyer_ctor",
        "ctor_Non-buyer": "nonbuyer_ctor",
        "eligible_n_Buyer": "buyer_eligible",
        "eligible_n_Non-buyer": "nonbuyer_eligible",
    })

    stable = piv[piv["broadcasts_matched"] >= 3].copy()
    stable_ctor = stable.dropna(subset=["buyer_ctor", "nonbuyer_ctor"]).copy()
    latest = piv.sort_values("month_dt").iloc[-1]

    analysis_date = pd.Timestamp.utcnow().tz_localize(None)
    subs["age_days"] = (analysis_date - subs["created_at_dt"].dt.tz_convert(None)).dt.days
    subs["age_bucket"] = pd.cut(
        subs["age_days"],
        bins=[-1, 90, 180, 365, 10000],
        labels=["<=90d", "91-180d", "181-365d", ">365d"],
    )

    summary_rows = []
    for label, d in [("Buyer", subs[subs["is_buyer"]]), ("Non-buyer", subs[~subs["is_buyer"]])]:
        summary_rows.append({
            "group": label,
            "n_subscribers": len(d),
            "mean_age_days": float(d["age_days"].mean()),
            "median_age_days": float(d["age_days"].median()),
            "recent_90_pct": float((d["age_days"] <= 90).mean() * 100),
            "old_365_pct": float((d["age_days"] > 365).mean() * 100),
        })

    summary = pd.DataFrame(summary_rows)
    summary["stable_open_rate_mean"] = np.nan
    summary["stable_ctor_mean"] = np.nan
    summary["latest_open_rate"] = np.nan
    summary["latest_ctor"] = np.nan

    summary.loc[summary["group"] == "Buyer", "stable_open_rate_mean"] = stable["buyer_open_rate"].mean()
    summary.loc[summary["group"] == "Non-buyer", "stable_open_rate_mean"] = stable["nonbuyer_open_rate"].mean()
    summary.loc[summary["group"] == "Buyer", "stable_ctor_mean"] = stable_ctor["buyer_ctor"].mean()
    summary.loc[summary["group"] == "Non-buyer", "stable_ctor_mean"] = stable_ctor["nonbuyer_ctor"].mean()
    summary.loc[summary["group"] == "Buyer", "latest_open_rate"] = latest["buyer_open_rate"]
    summary.loc[summary["group"] == "Non-buyer", "latest_open_rate"] = latest["nonbuyer_open_rate"]
    summary.loc[summary["group"] == "Buyer", "latest_ctor"] = latest["buyer_ctor"]
    summary.loc[summary["group"] == "Non-buyer", "latest_ctor"] = latest["nonbuyer_ctor"]
    summary["stable_open_gap_vs_nonbuyer_pp"] = np.nan
    summary["stable_ctor_gap_vs_nonbuyer_pp"] = np.nan
    summary.loc[summary["group"] == "Buyer", "stable_open_gap_vs_nonbuyer_pp"] = (
        stable["buyer_open_rate"].mean() - stable["nonbuyer_open_rate"].mean()
    )
    summary.loc[summary["group"] == "Buyer", "stable_ctor_gap_vs_nonbuyer_pp"] = (
        stable_ctor["buyer_ctor"].mean() - stable_ctor["nonbuyer_ctor"].mean()
    )
    summary.to_csv(OUT_SUMMARY, index=False)
    print(f"Saved summary output: {OUT_SUMMARY.name}", flush=True)

    age_tab = (
        subs.groupby([subs["is_buyer"].map({True: "Buyer", False: "Non-buyer"}), "age_bucket"])
        .size()
        .reset_index(name="n")
        .rename(columns={"is_buyer": "group"})
    )
    age_tab["pct"] = age_tab["n"] / age_tab.groupby("group")["n"].transform("sum") * 100
    age_tab.to_csv(OUT_AGE, index=False)
    print(f"Saved age-bucket output: {OUT_AGE.name}", flush=True)

    # ── Chart AO: monthly OR + CTOR ────────────────────────────────────────────
    chart_df = piv.sort_values("month_dt").copy()
    low_cov = chart_df["broadcasts_matched"] < 3

    fig, axes = plt.subplots(2, 1, figsize=(13.5, 8.6), sharex=True)

    # Open Rate
    ax = axes[0]
    ax.plot(chart_df["month_dt"], chart_df["buyer_open_rate"], color=GREEN, marker="o", linewidth=2.3, label="Buyer")
    ax.plot(chart_df["month_dt"], chart_df["nonbuyer_open_rate"], color=ORANGE, marker="o", linewidth=2.3, label="Non-buyer")
    ax.scatter(chart_df.loc[low_cov, "month_dt"], chart_df.loc[low_cov, "buyer_open_rate"], s=62, facecolors="none", edgecolors=GREEN, linewidths=1.5)
    ax.scatter(chart_df.loc[low_cov, "month_dt"], chart_df.loc[low_cov, "nonbuyer_open_rate"], s=62, facecolors="none", edgecolors=ORANGE, linewidths=1.5)
    ax.set_title("Monthly Open Rate — Bootcamp Buyers vs Non-buyers", fontsize=12)
    ax.set_ylabel("Open Rate (%)")
    ax.set_ylim(0, 100)
    ax.grid(axis="y", alpha=0.2)
    ax.legend(fontsize=9)

    # CTOR + coverage overlay
    ax = axes[1]
    chart_df["buyer_ctor_roll"] = chart_df["buyer_ctor"].rolling(3, min_periods=1).mean()
    chart_df["nonbuyer_ctor_roll"] = chart_df["nonbuyer_ctor"].rolling(3, min_periods=1).mean()
    ax.plot(chart_df["month_dt"], chart_df["buyer_ctor"], color=GREEN, marker="o", linewidth=2.0, label="Buyer (monthly)")
    ax.plot(chart_df["month_dt"], chart_df["nonbuyer_ctor"], color=ORANGE, marker="o", linewidth=2.0, label="Non-buyer (monthly)")
    ax.plot(chart_df["month_dt"], chart_df["buyer_ctor_roll"], color="#047857", linestyle="--", linewidth=1.8, label="Buyer (3-mo smooth)")
    ax.plot(chart_df["month_dt"], chart_df["nonbuyer_ctor_roll"], color="#B45309", linestyle="--", linewidth=1.8, label="Non-buyer (3-mo smooth)")
    ax.scatter(chart_df.loc[low_cov, "month_dt"], chart_df.loc[low_cov, "buyer_ctor"], s=62, facecolors="none", edgecolors=GREEN, linewidths=1.5)
    ax.scatter(chart_df.loc[low_cov, "month_dt"], chart_df.loc[low_cov, "nonbuyer_ctor"], s=62, facecolors="none", edgecolors=ORANGE, linewidths=1.5)
    ax.set_title("Monthly CTOR — Bootcamp Buyers vs Non-buyers", fontsize=12)
    ax.set_ylabel("CTOR (%)")
    ax.set_xlabel("Month")
    ax.set_ylim(0, 100)
    ax.grid(axis="y", alpha=0.2)

    ax2 = ax.twinx()
    ax2.bar(chart_df["month_dt"], chart_df["broadcasts_matched"], width=20, color=MUTED, alpha=0.18, label="Matched broadcasts")
    ax2.axhline(3, color="#64748B", linestyle=":", linewidth=1.2, alpha=0.8)
    ax2.set_ylabel("Matched broadcasts (n)")
    ax2.set_ylim(0, max(3, float(chart_df["broadcasts_matched"].max())) + 1.5)

    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, fontsize=8.5, loc="upper right")

    fig.suptitle(
        "Bootcamp Buyer vs Non-buyer Engagement Over Time (Low-coverage months shown as hollow points)",
        fontsize=14,
        fontweight="bold",
    )
    fig.tight_layout()
    fig.savefig(CHART_MONTHLY)
    plt.close()
    print(f"Saved chart: {CHART_MONTHLY.name}", flush=True)

    # ── Chart AP: age profile ──────────────────────────────────────────────────
    fig, axes = plt.subplots(1, 2, figsize=(13.8, 5.4))

    # Left: age density histogram
    for label, color in [("Buyer", GREEN), ("Non-buyer", ORANGE)]:
        d = subs[subs["is_buyer"] == (label == "Buyer")]["age_days"].dropna()
        axes[0].hist(d, bins=30, alpha=0.35, color=color, density=True, label=label)
        axes[0].axvline(d.median(), color=color, linestyle="--", linewidth=1.6)
    axes[0].set_title("Subscriber Age Distribution (days since signup)")
    axes[0].set_xlabel("Age in days")
    axes[0].set_ylabel("Density")
    axes[0].grid(axis="y", alpha=0.2)
    axes[0].legend(fontsize=9)

    # Right: stacked age buckets (%)
    bucket_order = ["<=90d", "91-180d", "181-365d", ">365d"]
    age_plot = age_tab.pivot(index="group", columns="age_bucket", values="pct").fillna(0)
    for b in bucket_order:
        if b not in age_plot.columns:
            age_plot[b] = 0.0
    age_plot = age_plot[bucket_order]
    colors = [GREEN, "#06B6D4", "#6366F1", RED]
    bottom = np.zeros(len(age_plot))
    x = np.arange(len(age_plot))
    for b, c in zip(bucket_order, colors):
        vals = age_plot[b].values
        axes[1].bar(x, vals, bottom=bottom, label=b, color=c, alpha=0.9)
        bottom += vals
    axes[1].set_xticks(x)
    axes[1].set_xticklabels(age_plot.index.tolist())
    axes[1].set_ylim(0, 100)
    axes[1].set_ylabel("Share of subscribers (%)")
    axes[1].set_title("Age Bucket Mix — Buyer vs Non-buyer")
    axes[1].grid(axis="y", alpha=0.2)
    axes[1].legend(title="Age bucket", fontsize=8.5, title_fontsize=9, loc="upper right")

    fig.suptitle("Bootcamp Buyer vs Non-buyer Age Profile", fontsize=13.5, fontweight="bold")
    fig.tight_layout()
    fig.savefig(CHART_AGE)
    plt.close()
    print(f"Saved chart: {CHART_AGE.name}", flush=True)

    # ── Failed IDs output ──────────────────────────────────────────────────────
    if failed_rows:
        pd.DataFrame(failed_rows).drop_duplicates().to_csv(OUT_FAILED, index=False)
        print(f"Saved failed IDs: {OUT_FAILED.name} ({len(pd.DataFrame(failed_rows).drop_duplicates())} rows)", flush=True)
    else:
        pd.DataFrame(columns=["event_type", "broadcast_id", "status_code", "error_preview"]).to_csv(OUT_FAILED, index=False)
        print("No failed IDs.", flush=True)

    print("\nDone.", flush=True)


if __name__ == "__main__":
    main()

"""
Monthly Group A/B Engagement Evolution (Lead Magnet Cohorts)
=============================================================

Goal
----
For Group A (survey responders) and Group B (non-responders), compute monthly trends
since Feb 2025 for:
1) Open rate on Value emails
2) Open rate on Sales emails
3) Sales CTOR only

Method
------
- Uses Kit v4 API `POST /subscribers/filter` with engagement filters and broadcast IDs.
- For each month/category:
    openers = unique subscribers who opened >=1 matching broadcasts in that month/category
    clickers = unique subscribers who clicked >=1 matching broadcasts in that month/category
- Group rates:
    open rate = openers_in_group / eligible_group_size_that_month
    sales CTOR = sales_clickers_in_group / sales_openers_in_group
- "Eligible group size" uses subscriber `created_at` <= month-end.
  (No unsubscribe-at-date field is available in these CSVs.)

Outputs
-------
- lead_magnet_group_ab_monthly.csv
- lead_magnet_broadcast_mapping.csv
- charts/AC_group_ab_monthly_openrate.png
- charts/AD_group_ab_monthly_sales_ctor.png
- charts/AE_group_value_vs_sales_openrate.png
"""

from __future__ import annotations

import time
from difflib import SequenceMatcher
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import requests


# ─── Constants ────────────────────────────────────────────────────────────────
BASE = Path(__file__).resolve().parent
OUT = BASE / "charts"
OUT.mkdir(exist_ok=True)

API_KEY = "kit_7d6b10fad06f88e1d0e47e45ef92e9cc"
API_BASE = "https://api.kit.com/v4"
HEADERS = {
    "X-Kit-Api-Key": API_KEY,
    "Accept": "application/json",
    "Content-Type": "application/json",
}

START_MONTH = pd.Period("2025-02", freq="M")

CACHE_MONTHLY = BASE / "lead_magnet_group_ab_monthly.csv"
CACHE_MAPPING = BASE / "lead_magnet_broadcast_mapping.csv"

GREEN = "#10B981"
ORANGE = "#F59E0B"
BRAND = "#4F46E5"
ACCENT = "#EC4899"
RED = "#EF4444"
MUTED = "#94A3B8"
BAD_BROADCAST_IDS: set[int] = set()

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "figure.dpi": 160,
})


RESPONSE_KEYWORDS = ["i want", "i'm", "i work at a company", "i own/run", "i own", "i'm building"]


def has_survey_response(tags_str: str) -> bool:
    if pd.isna(tags_str):
        return False
    t = str(tags_str).lower()
    return any(k in t for k in RESPONSE_KEYWORDS)


def norm_subject(s: str) -> str:
    s = str(s).lower()
    cleaned = "".join(ch if ch.isalnum() else " " for ch in s)
    return " ".join(cleaned.split())


def rate_limited_get(url: str, params: dict | None = None, timeout: int = 25):
    for _ in range(4):
        r = requests.get(url, headers=HEADERS, params=params, timeout=timeout)
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", 62))
            print(f"  GET rate-limited. Waiting {wait}s …")
            time.sleep(wait)
            continue
        return r
    return r


def rate_limited_post(url: str, payload: dict | None = None, timeout: int = 30):
    for i in range(3):
        r = requests.post(url, headers=HEADERS, json=payload, timeout=timeout)
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", 62))
            print(f"  POST rate-limited. Waiting {wait}s …")
            time.sleep(wait)
            continue
        if r.status_code >= 500:
            # Keep retries short; failing IDs are handled by chunk bisection/skip.
            wait = 1 + i
            if i < 2:
                print(f"  POST server error {r.status_code}. Retrying in {wait}s …")
                time.sleep(wait)
                continue
            return r
            continue
        return r
    return r


def fetch_all_broadcasts() -> pd.DataFrame:
    rows = []
    cursor = None
    while True:
        params = {"per_page": 1000}
        if cursor:
            params["after"] = cursor
        r = rate_limited_get(f"{API_BASE}/broadcasts", params=params)
        if r.status_code != 200:
            raise RuntimeError(f"Broadcast list error: {r.status_code} {r.text[:250]}")
        data = r.json()
        for b in data.get("broadcasts", []):
            ts = pd.to_datetime(b.get("created_at"), utc=True, errors="coerce")
            rows.append({
                "api_broadcast_id": int(b["id"]),
                "api_subject": b.get("subject", ""),
                "api_subject_norm": norm_subject(b.get("subject", "")),
                "api_dt": ts.tz_convert(None) if pd.notna(ts) else pd.NaT,
                "api_date": (ts.tz_convert(None).date() if pd.notna(ts) else pd.NaT),
            })
        pag = data.get("pagination", {})
        if not pag.get("has_next_page"):
            break
        cursor = pag.get("end_cursor")
    return pd.DataFrame(rows)


def map_local_to_api(local_df: pd.DataFrame, api_df: pd.DataFrame) -> pd.DataFrame:
    mapped_rows = []

    for _, row in local_df.iterrows():
        local_date = row["date"]
        local_norm = row["subject_norm"]
        date_min = local_date - pd.Timedelta(days=1)
        date_max = local_date + pd.Timedelta(days=1)
        cand = api_df[(api_df["api_dt"] >= date_min) & (api_df["api_dt"] <= date_max)].copy()

        if cand.empty:
            mapped_rows.append({**row.to_dict(), "api_broadcast_id": np.nan, "match_score": np.nan})
            continue

        exact = cand[cand["api_subject_norm"] == local_norm]
        if len(exact):
            best = exact.assign(day_diff=(exact["api_dt"] - local_date).abs()).sort_values("day_diff").iloc[0]
            mapped_rows.append({**row.to_dict(), "api_broadcast_id": int(best["api_broadcast_id"]), "match_score": 1.0})
            continue

        cand["score"] = cand["api_subject_norm"].apply(lambda s: SequenceMatcher(None, local_norm, s).ratio())
        cand["day_diff"] = (cand["api_dt"] - local_date).abs().dt.total_seconds()
        best = cand.sort_values(["score", "day_diff"], ascending=[False, True]).iloc[0]

        if best["score"] >= 0.78:
            mapped_rows.append({**row.to_dict(), "api_broadcast_id": int(best["api_broadcast_id"]), "match_score": float(best["score"])})
        else:
            mapped_rows.append({**row.to_dict(), "api_broadcast_id": np.nan, "match_score": float(best["score"])})

    mapped = pd.DataFrame(mapped_rows)
    return mapped


def fetch_event_emails(event_type: str, broadcast_ids: list[int]) -> set[str]:
    """Return unique subscriber emails with event_type on any of given broadcasts."""
    filtered_ids = [int(x) for x in broadcast_ids if int(x) not in BAD_BROADCAST_IDS]
    if len(filtered_ids) == 0:
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
                # Fallback: bisect problematic chunks to isolate bad IDs.
                if len(ids_chunk) > 1:
                    mid = len(ids_chunk) // 2
                    left = _fetch_chunk(ids_chunk[:mid])
                    right = _fetch_chunk(ids_chunk[mid:])
                    return left | right
                bad_id = int(ids_chunk[0])
                if bad_id not in BAD_BROADCAST_IDS:
                    print(f"  Skipping broadcast id {bad_id} for {event_type} (API {r.status_code})")
                    BAD_BROADCAST_IDS.add(bad_id)
                return set()
            data = r.json()
            for s in data.get("subscribers", []):
                e = s.get("email_address")
                if e:
                    out.add(e.lower().strip())
            pag = data.get("pagination", {})
            if not pag.get("has_next_page"):
                break
            cursor = pag.get("end_cursor")
        return out

    # Chunk IDs to reduce server-side 500s; each chunk can still be bisected if needed.
    chunk_size = 12
    chunks = [filtered_ids[i:i + chunk_size] for i in range(0, len(filtered_ids), chunk_size)]
    for ids_chunk in chunks:
        emails |= _fetch_chunk(ids_chunk)
    return emails


def main():
    print("=" * 76)
    print("MONTHLY GROUP A/B EVOLUTION — VALUE OR, SALES OR, SALES CTOR")
    print("=" * 76)

    # ── Load cohort files and define groups ──────────────────────────────────
    clicked = pd.read_csv(BASE / "AI Sprint Roadmap - Clicked.csv")
    opened = pd.read_csv(BASE / "AI Sprint Roadmap - Opened subscribers.csv")

    for df in (clicked, opened):
        df["email_lower"] = df["email"].str.lower().str.strip()
        df["is_group_a"] = df["tags"].apply(has_survey_response)
        df["created_at_dt"] = pd.to_datetime(df["created_at"], utc=True, errors="coerce")

    clicked_emails = set(clicked["email_lower"])
    group_a_emails = set(clicked[clicked["is_group_a"]]["email_lower"])
    group_b_emails = (
        set(clicked[~clicked["is_group_a"]]["email_lower"])
        | set(opened[(~opened["email_lower"].isin(clicked_emails)) & (~opened["is_group_a"])]["email_lower"])
    )

    print(f"Group A size: {len(group_a_emails):,}")
    print(f"Group B size: {len(group_b_emails):,}")

    # Created_at map for eligibility by month-end
    created_map = pd.concat([
        clicked[["email_lower", "created_at_dt"]],
        opened[["email_lower", "created_at_dt"]],
    ], ignore_index=True)
    created_map = created_map.sort_values("created_at_dt").drop_duplicates("email_lower", keep="first")
    created_map = created_map.dropna(subset=["created_at_dt"])

    a_created = created_map[created_map["email_lower"].isin(group_a_emails)][["email_lower", "created_at_dt"]].copy()
    b_created = created_map[created_map["email_lower"].isin(group_b_emails)][["email_lower", "created_at_dt"]].copy()

    # ── Load local broadcast categories and map to Kit IDs ───────────────────
    bdf = pd.read_csv(BASE / "Emails Broadcasting - broadcasts_categorised.csv")
    bdf["date"] = pd.to_datetime(bdf["date"], errors="coerce")
    bdf["category"] = bdf["category"].replace("Sles", "Sales")
    bdf = bdf[bdf["category"].isin(["Value", "Sales"])].copy()
    bdf = bdf.sort_values("date")
    bdf["subject_norm"] = bdf["subject"].apply(norm_subject)

    print("\nFetching broadcasts from Kit API …")
    api_broadcasts = fetch_all_broadcasts()
    mapped = map_local_to_api(bdf, api_broadcasts)

    # Keep one row per matched API broadcast id to avoid duplicate local rows
    matched = mapped[mapped["api_broadcast_id"].notna()].copy()
    matched["api_broadcast_id"] = matched["api_broadcast_id"].astype(int)
    matched = matched.sort_values("date").drop_duplicates(subset=["api_broadcast_id"], keep="last")
    mapped.to_csv(CACHE_MAPPING, index=False)

    print(f"Local rows: {len(bdf):,} | matched unique API broadcasts: {len(matched):,} | unmatched local rows: {(mapped['api_broadcast_id'].isna()).sum():,}")

    matched["month"] = matched["date"].dt.to_period("M")
    last_month = matched["month"].max()
    months = pd.period_range(START_MONTH, last_month, freq="M")

    rows = []
    for month in months:
        month_start = month.to_timestamp(how="start")
        month_end = (month + 1).to_timestamp(how="start") - pd.Timedelta(seconds=1)

        # eligibility by subscription date
        a_eligible = int((a_created["created_at_dt"].dt.tz_convert(None) <= month_end).sum())
        b_eligible = int((b_created["created_at_dt"].dt.tz_convert(None) <= month_end).sum())

        for cat in ["Value", "Sales"]:
            sub = matched[(matched["month"] == month) & (matched["category"] == cat)].copy()
            b_ids = sorted(sub["api_broadcast_id"].astype(int).unique().tolist())

            if len(b_ids):
                openers = fetch_event_emails("opens", b_ids)
                clickers = fetch_event_emails("clicks", b_ids) if cat == "Sales" else set()
            else:
                openers = set()
                clickers = set()

            a_open = len(openers & group_a_emails)
            b_open = len(openers & group_b_emails)
            a_click = len(clickers & group_a_emails)
            b_click = len(clickers & group_b_emails)

            a_or = (a_open / a_eligible * 100) if a_eligible > 0 else np.nan
            b_or = (b_open / b_eligible * 100) if b_eligible > 0 else np.nan

            if cat == "Sales":
                a_ctor = (a_click / a_open * 100) if a_open > 0 else np.nan
                b_ctor = (b_click / b_open * 100) if b_open > 0 else np.nan
            else:
                a_ctor = np.nan
                b_ctor = np.nan

            rows.append({
                "month": str(month),
                "month_dt": month_start,
                "category": cat,
                "broadcasts_matched": len(b_ids),
                "a_eligible": a_eligible,
                "b_eligible": b_eligible,
                "a_openers": a_open,
                "b_openers": b_open,
                "a_clickers": a_click,
                "b_clickers": b_click,
                "a_open_rate": a_or,
                "b_open_rate": b_or,
                "a_sales_ctor": a_ctor,
                "b_sales_ctor": b_ctor,
            })

            print(
                f"{month} [{cat}] broadcasts={len(b_ids):>2} | "
                f"A OR={a_or:>5.1f}% ({a_open}/{a_eligible}) | "
                f"B OR={b_or:>5.1f}% ({b_open}/{b_eligible})"
                + (f" | A CTOR={a_ctor:>5.1f}% B CTOR={b_ctor:>5.1f}%" if cat == "Sales" else "")
            )

    monthly = pd.DataFrame(rows)
    monthly.to_csv(CACHE_MONTHLY, index=False)
    print(f"\nSaved monthly metrics → {CACHE_MONTHLY.name}")
    print(f"Saved mapping details → {CACHE_MAPPING.name}")

    # ── Charts ────────────────────────────────────────────────────────────────
    # Chart AC: Open rate evolution by category (A vs B)
    fig, axes = plt.subplots(2, 1, figsize=(14, 9), sharex=True)
    for ax, cat in zip(axes, ["Value", "Sales"]):
        d = monthly[monthly["category"] == cat].sort_values("month_dt")
        ax.plot(d["month_dt"], d["a_open_rate"], color=GREEN, marker="o", linewidth=2.2, label="Group A")
        ax.plot(d["month_dt"], d["b_open_rate"], color=ORANGE, marker="o", linewidth=2.2, label="Group B")
        ax.set_ylabel("Open Rate (%)")
        ax.set_title(f"{cat} Emails — Monthly Open Rate (Group A vs Group B)")
        ax.grid(axis="y", alpha=0.2)
        for _, r in d.iterrows():
            if r["broadcasts_matched"] == 0:
                ax.axvline(r["month_dt"], color=MUTED, alpha=0.08, linewidth=1)
        ax.legend(loc="upper right")

    axes[-1].set_xlabel("Month")
    fig.suptitle("Group A vs Group B — Monthly Open Rate Evolution by Email Type", fontsize=14, fontweight="bold")
    fig.tight_layout()
    fig.savefig(OUT / "AC_group_ab_monthly_openrate.png")
    plt.close()

    # Chart AD: Sales CTOR evolution (A vs B)
    sales = monthly[monthly["category"] == "Sales"].sort_values("month_dt")
    fig, ax = plt.subplots(figsize=(14, 5.8))
    ax.plot(sales["month_dt"], sales["a_sales_ctor"], color=GREEN, marker="o", linewidth=2.4, label="Group A — Sales CTOR")
    ax.plot(sales["month_dt"], sales["b_sales_ctor"], color=ORANGE, marker="o", linewidth=2.4, label="Group B — Sales CTOR")
    ax.set_ylabel("Sales CTOR (%)")
    ax.set_xlabel("Month")
    ax.set_title("Group A vs Group B — Monthly Sales CTOR Evolution")
    ax.grid(axis="y", alpha=0.2)
    ax.legend()
    fig.tight_layout()
    fig.savefig(OUT / "AD_group_ab_monthly_sales_ctor.png")
    plt.close()

    # Chart AE: Value vs Sales open rate within each group
    fig, axes = plt.subplots(2, 1, figsize=(14, 9), sharex=True)
    value = monthly[monthly["category"] == "Value"].sort_values("month_dt")
    sales = monthly[monthly["category"] == "Sales"].sort_values("month_dt")

    # Group A panel
    axes[0].plot(value["month_dt"], value["a_open_rate"], color=BRAND, marker="o", linewidth=2.2, label="Value OR")
    axes[0].plot(sales["month_dt"], sales["a_open_rate"], color=ACCENT, marker="o", linewidth=2.2, label="Sales OR")
    axes[0].set_ylabel("Open Rate (%)")
    axes[0].set_title("Group A — Value vs Sales Monthly Open Rate")
    axes[0].grid(axis="y", alpha=0.2)
    axes[0].legend(loc="upper right")

    # Group B panel
    axes[1].plot(value["month_dt"], value["b_open_rate"], color=BRAND, marker="o", linewidth=2.2, label="Value OR")
    axes[1].plot(sales["month_dt"], sales["b_open_rate"], color=ACCENT, marker="o", linewidth=2.2, label="Sales OR")
    axes[1].set_ylabel("Open Rate (%)")
    axes[1].set_xlabel("Month")
    axes[1].set_title("Group B — Value vs Sales Monthly Open Rate")
    axes[1].grid(axis="y", alpha=0.2)
    axes[1].legend(loc="upper right")

    fig.suptitle("Monthly Open Rate by Email Type Within Each Group", fontsize=14, fontweight="bold")
    fig.tight_layout()
    fig.savefig(OUT / "AE_group_value_vs_sales_openrate.png")
    plt.close()

    print("\nSaved charts:")
    print("  - AC_group_ab_monthly_openrate.png")
    print("  - AD_group_ab_monthly_sales_ctor.png")
    print("  - AE_group_value_vs_sales_openrate.png")


if __name__ == "__main__":
    main()

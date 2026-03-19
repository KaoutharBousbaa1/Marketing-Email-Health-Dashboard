"""
Campaign Math Check (Bootcamp Invite vs Roadmap/Survey Tags)
============================================================

Purpose
-------
Sanity-check campaign math like:
  - recipients of a specific bootcamp invite email (denominator)
  - roadmap/survey-tagged subsets (possible numerators)

This script is useful when teammates report numbers like:
  "Total recipients: 8,930 | AI Roadmap subscribers: 1,566"
and we need to verify exactly which tag logic could produce that.

Default campaign checked
------------------------
Subject: "Your private invite to join the AI Agent Bootcamp"
Date:    2026-02-17

Outputs
-------
- data/generated/campaign_math_check_summary.csv
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd


BASE = Path(__file__).resolve().parent
GENERATED = BASE / "data" / "generated"
GENERATED.mkdir(parents=True, exist_ok=True)

BROADCASTS_FILE = BASE / "Emails Broadcasting - broadcasts_categorised.csv"
SUBSCRIBERS_FILE = BASE / "Confirmed Subscribers.csv"
OUT_FILE = GENERATED / "campaign_math_check_summary.csv"

DEFAULT_SUBJECT = "Your private invite to join the AI Agent Bootcamp"
DEFAULT_DATE = "2026-02-17"
DEFAULT_CLAIMED = 1566

SURVEY_GOAL_TAGS = [
    "I'm transitioning into an AI/tech career - I want to build skills and projects to land an AI role",
    "I work at a company - I want to become the AI expert at work and automate my tasks",
    "I'm an independent consultant or freelancer - I want to offer AI services to clients",
    "I own/run a business or lead a team - I want to implement AI across my organization",
    "I'm building AI products or solutions - I want to ship working AI features or products",
]


def _fmt_rate(numerator: int, denominator: int) -> float:
    if denominator <= 0:
        return float("nan")
    return round((numerator / denominator) * 100.0, 2)


def _contains_any_tags(text_series: pd.Series, tags: list[str]) -> pd.Series:
    pattern = "|".join(re.escape(tag.lower()) for tag in tags)
    return text_series.str.contains(pattern, regex=True, na=False)


def _load_campaign_row(df: pd.DataFrame, subject: str, date_str: str) -> pd.Series:
    date_target = pd.to_datetime(date_str, errors="coerce")
    if pd.isna(date_target):
        raise ValueError(f"Invalid --date value: {date_str}")

    tmp = df.copy()
    tmp["date_dt"] = pd.to_datetime(tmp["date"], errors="coerce")
    subject_norm = subject.strip().lower()
    matched = tmp[
        (tmp["subject"].astype(str).str.strip().str.lower() == subject_norm)
        & (tmp["date_dt"].dt.normalize() == date_target.normalize())
    ]
    if matched.empty:
        contains = tmp[
            tmp["subject"].astype(str).str.lower().str.contains(subject_norm, na=False)
        ]
        if contains.empty:
            raise ValueError(
                "Could not find the campaign row. "
                "Check --subject and --date values."
            )
        raise ValueError(
            "Exact subject+date match not found. "
            f"Closest subject matches found: {len(contains)} rows."
        )
    if len(matched) > 1:
        matched = matched.sort_values(["date_dt", "time"]).tail(1)
    return matched.iloc[0]


def _build_context_rows(
    context_name: str,
    subs: pd.DataFrame,
    recipients: int,
    claimed_numerator: int,
) -> list[dict]:
    tags = subs["tags_l"]
    ref = subs["ref_l"]

    waitlist = tags.str.contains("ai agent bootcamp waitlist", na=False)
    sent_survey = tags.str.contains("been sent 2025 roadmap survey", na=False)
    voted_roadmap = tags.str.contains("voted: 28-day ai sprint roadmap", na=False)
    goal_response = _contains_any_tags(tags, SURVEY_GOAL_TAGS)
    roadmap_referrer = ref.str.contains(
        r"lonelyoctopus\.com/ai-sprint-roadmap|lonelyoctopus\.com/roadmap",
        regex=True,
        na=False,
    )

    checks = [
        (
            "waitlist_total",
            int(waitlist.sum()),
            "Tag contains 'AI Agent Bootcamp Waitlist'",
        ),
        (
            "waitlist_and_sent_survey_tag",
            int((waitlist & sent_survey).sum()),
            "Waitlist + 'Been Sent 2025 Roadmap Survey'",
        ),
        (
            "waitlist_and_(sent_or_voted_roadmap)",
            int((waitlist & (sent_survey | voted_roadmap)).sum()),
            "Waitlist + ('Been Sent...' OR 'Voted: 28-Day AI Sprint Roadmap')",
        ),
        (
            "waitlist_and_goal_response_tags",
            int((waitlist & goal_response).sum()),
            "Waitlist + one of the 5 roadmap goal-response tags",
        ),
        (
            "waitlist_and_goal_response_and_roadmap_referrer",
            int((waitlist & goal_response & roadmap_referrer).sum()),
            "Waitlist + goal-response tag + roadmap referrer",
        ),
        (
            "all_subscribers_and_goal_response_tags",
            int(goal_response.sum()),
            "All confirmed subscribers with one of 5 roadmap goal-response tags",
        ),
        (
            "all_subscribers_and_roadmap_referrer",
            int(roadmap_referrer.sum()),
            "All confirmed subscribers with roadmap referrer",
        ),
        (
            "claimed_numerator_input",
            int(claimed_numerator),
            "Manual claimed numerator (default=1,566)",
        ),
    ]

    rows: list[dict] = []
    for metric, count, definition in checks:
        rows.append(
            {
                "context": context_name,
                "metric": metric,
                "count": count,
                "denominator_recipients": recipients,
                "rate_vs_recipients_pct": _fmt_rate(count, recipients),
                "delta_vs_claimed_count": count - claimed_numerator,
                "definition": definition,
            }
        )
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Check campaign denominator/numerator math for bootcamp invite analysis."
    )
    parser.add_argument("--subject", default=DEFAULT_SUBJECT, help="Exact email subject.")
    parser.add_argument("--date", default=DEFAULT_DATE, help="Campaign date (YYYY-MM-DD).")
    parser.add_argument(
        "--claimed",
        type=int,
        default=DEFAULT_CLAIMED,
        help="Claimed numerator to compare against (default: 1566).",
    )
    args = parser.parse_args()

    broadcasts = pd.read_csv(BROADCASTS_FILE)
    campaign = _load_campaign_row(broadcasts, args.subject, args.date)
    recipients = int(campaign["recipients"])
    campaign_date = pd.to_datetime(campaign["date"], errors="coerce").normalize()

    subs = pd.read_csv(SUBSCRIBERS_FILE, usecols=["email", "created_at", "tags", "referrer"])
    subs["created_at_dt"] = pd.to_datetime(subs["created_at"], errors="coerce", utc=True).dt.tz_convert(None)
    subs["created_date"] = subs["created_at_dt"].dt.normalize()
    subs["tags_l"] = subs["tags"].fillna("").astype(str).str.lower()
    subs["ref_l"] = subs["referrer"].fillna("").astype(str).str.lower()
    subs = subs.dropna(subset=["created_date"]).copy()

    as_of_send = subs[subs["created_date"] <= campaign_date].copy()
    full_snapshot = subs.copy()

    rows = []
    rows.append(
        {
            "context": "campaign_info",
            "metric": "campaign_recipients",
            "count": recipients,
            "denominator_recipients": recipients,
            "rate_vs_recipients_pct": 100.0,
            "delta_vs_claimed_count": recipients - args.claimed,
            "definition": f"Recipients for '{campaign['subject']}' on {campaign['date']} {campaign['time']}",
        }
    )
    rows.extend(_build_context_rows("as_of_send_date", as_of_send, recipients, args.claimed))
    rows.extend(_build_context_rows("full_snapshot", full_snapshot, recipients, args.claimed))

    out = pd.DataFrame(rows)
    out.to_csv(OUT_FILE, index=False)

    print("=" * 84)
    print("CAMPAIGN MATH CHECK")
    print("=" * 84)
    print(f"Campaign: {campaign['subject']}")
    print(f"Date:     {campaign['date']} {campaign['time']}")
    print(f"Recipients (denominator): {recipients:,}")
    print(f"Claimed numerator checked: {args.claimed:,}")
    print("-" * 84)
    print("Top checks (as_of_send_date):")
    top = out[out["context"] == "as_of_send_date"][
        ["metric", "count", "rate_vs_recipients_pct", "definition"]
    ]
    for _, r in top.iterrows():
        print(f"- {r['metric']}: {int(r['count']):,} ({r['rate_vs_recipients_pct']}%)")
    print("-" * 84)
    print(f"Saved: {OUT_FILE}")


if __name__ == "__main__":
    main()

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd

import config
from kit_client import KitClient


@dataclass
class SegmentSets:
    lead_magnet: set[str]
    workshop: set[str]
    bootcamp: set[str]
    active_non_buyers: set[str]
    active_all: set[str]
    active_before_window: set[str]
    workshop_by_tag: dict[str, set[str]]
    mapped_bootcamp_by_workshop: dict[str, set[str]]
    entry_sources: dict[str, set[str]]
    entry_source_tagged_at: dict[str, dict[str, str]]


class KPIService:
    def __init__(self, client: KitClient, tz_name: str = "Africa/Casablanca") -> None:
        self.client = client
        self.tz_name = tz_name

    @staticmethod
    def _contains_any(text: str, patterns: list[str]) -> bool:
        t = (text or "").lower()
        return any(p.lower() in t for p in patterns)

    @staticmethod
    def _canonical_workshop_name(tag_name_lower: str) -> str | None:
        t = (tag_name_lower or "").strip().lower()
        # Strict workshop tags used for mapped conversion:
        # - AI App Sprint
        # - Freelance Accelerator container (4 tags)
        # - Agent Breakthrough
        if t == "ai app sprint":
            return "AI App Sprint (Aug 2025)"
        freelance_container_tags = {
            "freelance accelerator [feb]",
            "freelance accelerator bundle",
            "freelance accelerator masterclass (public)",
            "freelance accelerator upsell [live masterclass]",
        }
        if t in freelance_container_tags:
            return "Freelance Accelerator (Oct 2025)"
        if t == "agent breakthrough":
            return "Agent Breakthrough (Nov 2025)"
        return None

    @staticmethod
    def _target_workshop_families() -> list[str]:
        return [
            "AI App Sprint (Aug 2025)",
            "Freelance Accelerator (Oct 2025)",
            "Agent Breakthrough (Nov 2025)",
        ]

    @staticmethod
    def _mapped_workshop_family_from_bootcamp_tag(tag_name_lower: str) -> str | None:
        t = (tag_name_lower or "").lower()
        if "ai agent core [sept" in t or "ai agent core [sep" in t:
            return "AI App Sprint (Aug 2025)"
        if "ai agent core [oct" in t:
            return "Freelance Accelerator (Oct 2025)"
        if "ai agent core [feb" in t:
            return "Agent Breakthrough (Nov 2025)"
        return None

    def _classify_broadcast(self, description: str) -> str:
        d = (description or "").strip().lower()
        d_norm = " ".join(d.replace("_", " ").replace("-", " ").split())
        if d_norm in {"pre", "post"}:
            return "Sales"
        if self._contains_any(d, config.VALUE_LABEL_KEYWORDS) or self._contains_any(
            d_norm, config.VALUE_LABEL_KEYWORDS
        ):
            return "Value"
        if self._contains_any(d, config.SALES_LABEL_KEYWORDS) or self._contains_any(
            d_norm, config.SALES_LABEL_KEYWORDS
        ):
            return "Sales"
        return "Unclassified"

    @staticmethod
    def _to_dt(value: str | None) -> datetime | None:
        if not value:
            return None
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            return None

    @staticmethod
    def _iso_key(value: str | None) -> tuple[int, datetime | None, str]:
        if not value:
            return (1, None, "")
        dt = KPIService._to_dt(value)
        if dt is None:
            return (1, None, value)
        return (0, dt, value)

    def _load_subscribers_df(self) -> pd.DataFrame:
        subs = self.client.list_subscribers(status="all")
        rows = []
        for s in subs:
            rows.append(
                {
                    "id": int(s.get("id")) if s.get("id") is not None else None,
                    "email": str(s.get("email_address", "")).strip().lower(),
                    "state": str(s.get("state", "")).strip().lower(),
                    "created_at": self._to_dt(s.get("created_at")),
                }
            )
        df = pd.DataFrame(rows)
        if len(df) == 0:
            return pd.DataFrame(columns=["id", "email", "state", "created_at"])
        df = df[df["email"].str.len() > 3].copy()
        return df

    def _load_segments(self, subscribers_df: pd.DataFrame, now_utc: datetime) -> SegmentSets:
        tags = self.client.list_tags()

        lead_tag_ids: list[int] = []
        workshop_tag_ids: list[int] = []
        bootcamp_tag_ids: list[int] = []
        source_tag_ids: dict[str, list[int]] = {
            source: [] for source in config.ENTRY_SOURCE_TAG_GROUPS.keys()
        }
        workshop_tag_families: dict[int, str] = {}
        bootcamp_tag_to_workshop_family: dict[int, str | None] = {}

        for t in tags:
            tag_id = int(t.get("id"))
            tag_name = str(t.get("name", "")).strip()
            lower_name = tag_name.lower()

            if self._contains_any(lower_name, config.LEAD_MAGNET_RESPONSE_PATTERNS):
                lead_tag_ids.append(tag_id)

            if self._contains_any(lower_name, config.WORKSHOP_TAG_PATTERNS):
                workshop_tag_ids.append(tag_id)
                family_name = self._canonical_workshop_name(lower_name)
                workshop_tag_families[tag_id] = family_name or tag_name

            if self._contains_any(lower_name, config.BOOTCAMP_TAG_PATTERNS):
                bootcamp_tag_ids.append(tag_id)
                bootcamp_tag_to_workshop_family[tag_id] = self._mapped_workshop_family_from_bootcamp_tag(
                    lower_name
                )

            for source_name, patterns in config.ENTRY_SOURCE_TAG_GROUPS.items():
                if self._contains_any(lower_name, patterns):
                    source_tag_ids[source_name].append(tag_id)

        lead_emails: set[str] = set()
        workshop_emails: set[str] = set()
        bootcamp_emails: set[str] = set()
        workshop_by_tag: dict[str, set[str]] = {}
        mapped_bootcamp_by_workshop: dict[str, set[str]] = {
            name: set() for name in self._target_workshop_families()
        }
        entry_sources: dict[str, set[str]] = {
            source: set() for source in config.ENTRY_SOURCE_TAG_GROUPS.keys()
        }
        entry_source_tagged_at: dict[str, dict[str, str]] = {
            source: {} for source in config.ENTRY_SOURCE_TAG_GROUPS.keys()
        }

        source_all_tag_ids = [tid for ids in source_tag_ids.values() for tid in ids]
        needed_tag_ids = sorted(
            set(lead_tag_ids + workshop_tag_ids + bootcamp_tag_ids + source_all_tag_ids)
        )
        tag_members: dict[int, set[str]] = {}
        if needed_tag_ids:
            with ThreadPoolExecutor(max_workers=8) as executor:
                future_map = {
                    executor.submit(self.client.list_tag_subscribers, tag_id): tag_id
                    for tag_id in needed_tag_ids
                }
                for future in as_completed(future_map):
                    tag_id = future_map[future]
                    try:
                        tag_members[tag_id] = future.result()
                    except Exception:
                        tag_members[tag_id] = set()

        for tag_id in sorted(set(lead_tag_ids)):
            lead_emails |= tag_members.get(tag_id, set())

        for tag_id in sorted(set(workshop_tag_ids)):
            members = tag_members.get(tag_id, set())
            workshop_emails |= members
            family = workshop_tag_families.get(tag_id, str(tag_id))
            workshop_by_tag[family] = workshop_by_tag.get(family, set()) | members

        for tag_id in sorted(set(bootcamp_tag_ids)):
            members = tag_members.get(tag_id, set())
            bootcamp_emails |= members
            mapped_family = bootcamp_tag_to_workshop_family.get(tag_id)
            if mapped_family in mapped_bootcamp_by_workshop:
                mapped_bootcamp_by_workshop[mapped_family] |= members

        for source_name, tag_ids in source_tag_ids.items():
            members: set[str] = set()
            for tag_id in sorted(set(tag_ids)):
                members |= tag_members.get(tag_id, set())
            entry_sources[source_name] = members

        # Pull tag timestamps for source tags to support first-touch attribution.
        source_tag_unique_ids = sorted({tid for ids in source_tag_ids.values() for tid in ids})
        source_tag_email_tagged_at: dict[int, dict[str, str]] = {}
        if source_tag_unique_ids:
            with ThreadPoolExecutor(max_workers=min(8, len(source_tag_unique_ids))) as executor:
                future_map = {
                    executor.submit(self.client.list_tag_subscribers_with_tagged_at, tag_id): tag_id
                    for tag_id in source_tag_unique_ids
                }
                for future in as_completed(future_map):
                    tag_id = future_map[future]
                    try:
                        source_tag_email_tagged_at[tag_id] = future.result()
                    except Exception:
                        source_tag_email_tagged_at[tag_id] = {}

        for source_name, tag_ids in source_tag_ids.items():
            email_to_first: dict[str, str] = {}
            for tag_id in sorted(set(tag_ids)):
                tagged_map = source_tag_email_tagged_at.get(tag_id, {})
                for email, tagged_at in tagged_map.items():
                    prev = email_to_first.get(email, "")
                    if tagged_at and (not prev or tagged_at < prev):
                        email_to_first[email] = tagged_at
                    elif email not in email_to_first:
                        email_to_first[email] = tagged_at
            entry_source_tagged_at[source_name] = email_to_first

        active_df = subscribers_df[subscribers_df["state"] == "active"].copy()
        active_all = set(active_df["email"])

        cutoff = now_utc - timedelta(days=config.REWARM_WINDOW_DAYS)
        active_before_window = set(active_df[active_df["created_at"] < cutoff]["email"])

        active_non_buyers = active_all - (workshop_emails | bootcamp_emails)

        return SegmentSets(
            lead_magnet=lead_emails,
            workshop=workshop_emails,
            bootcamp=bootcamp_emails,
            active_non_buyers=active_non_buyers,
            active_all=active_all,
            active_before_window=active_before_window,
            workshop_by_tag=workshop_by_tag,
            mapped_bootcamp_by_workshop=mapped_bootcamp_by_workshop,
            entry_sources=entry_sources,
            entry_source_tagged_at=entry_source_tagged_at,
        )

    def _load_broadcast_stats_df(self, now_utc: datetime) -> pd.DataFrame:
        broadcasts = self.client.list_broadcasts()

        # Enough range for all requested KPIs + margin.
        min_date = now_utc - timedelta(days=240)

        candidates: list[dict[str, Any]] = []
        for b in broadcasts:
            sent_dt = self._to_dt(b.get("published_at") or b.get("send_at") or b.get("created_at"))
            if not sent_dt:
                continue
            if sent_dt > now_utc:
                continue
            if sent_dt < min_date:
                continue

            candidates.append(
                {
                    "broadcast_id": int(b.get("id")),
                    "subject": str(b.get("subject") or "").strip(),
                    "description": str(b.get("description") or "").strip(),
                    "sent_at": sent_dt,
                }
            )

        rows: list[dict[str, Any]] = []
        if candidates:
            with ThreadPoolExecutor(max_workers=8) as executor:
                future_map = {
                    executor.submit(self.client.get_broadcast_stats, c["broadcast_id"]): c
                    for c in candidates
                }
                for future in as_completed(future_map):
                    c = future_map[future]
                    try:
                        stats = future.result()
                    except Exception:
                        stats = {}

                    recipients = int(stats.get("recipients") or 0)
                    opens = int(stats.get("emails_opened") or 0)
                    clicks = int(stats.get("total_clicks") or 0)
                    unsubs = int(stats.get("unsubscribes") or 0)
                    status = str(stats.get("status") or "")

                    if status in {"draft", "scheduled"}:
                        continue
                    if recipients <= 0:
                        continue

                    rows.append(
                        {
                            "broadcast_id": c["broadcast_id"],
                            "subject": c["subject"],
                            "description": c["description"],
                            "category": self._classify_broadcast(c["description"]),
                            "sent_at": c["sent_at"],
                            "recipients": recipients,
                            "opens": opens,
                            "clicks": clicks,
                            "unsubs": unsubs,
                        }
                    )

        if not rows:
            return pd.DataFrame(
                columns=[
                    "broadcast_id",
                    "subject",
                    "description",
                    "category",
                    "sent_at",
                    "recipients",
                    "opens",
                    "clicks",
                    "unsubs",
                ]
            )

        df = pd.DataFrame(rows)
        df["month"] = pd.to_datetime(df["sent_at"]).dt.tz_convert(None).dt.to_period("M")
        return df

    def _completed_months(self, n: int, now_utc: datetime) -> list[pd.Period]:
        current_p = pd.Timestamp(now_utc.date()).to_period("M")
        months = [current_p - i for i in range(1, n + 1)]
        months.reverse()
        return months

    def compute_all(self) -> dict[str, Any]:
        now_utc = datetime.now(timezone.utc)

        subscribers_df = self._load_subscribers_df()
        segments = self._load_segments(subscribers_df, now_utc)
        bdf = self._load_broadcast_stats_df(now_utc)

        # KPI 14: current confirmed (active) subscribers
        state_counts = (
            subscribers_df["state"].value_counts().to_dict() if len(subscribers_df) else {}
        )
        kpi14_current_confirmed = int(len(segments.active_all))

        # Confirmed subscribers trend (last N months) + source breakdown
        active_created_df = subscribers_df[
            (subscribers_df["state"] == "active") & subscribers_df["created_at"].notna()
        ].copy()
        if len(active_created_df):
            active_created_df["created_at"] = pd.to_datetime(
                active_created_df["created_at"], utc=True
            ).dt.tz_convert(None)

        end_month = pd.Timestamp(now_utc.date()).to_period("M")
        trend_months = [end_month - i for i in range(config.CONFIRMED_TREND_MONTHS - 1, -1, -1)]
        trend_rows = []
        for month in trend_months:
            month_start = month.to_timestamp(how="start")
            next_month_start = (month + 1).to_timestamp(how="start")
            if len(active_created_df):
                created = active_created_df["created_at"]
                new_confirmed = int(((created >= month_start) & (created < next_month_start)).sum())
                cumulative_confirmed = int((created < next_month_start).sum())
            else:
                new_confirmed = 0
                cumulative_confirmed = 0
            trend_rows.append(
                {
                    "month": str(month),
                    "new_confirmed": new_confirmed,
                    "cumulative_confirmed": cumulative_confirmed,
                }
            )
        confirmed_trend_df = pd.DataFrame(trend_rows)

        source_rows = []
        trend_start_month = trend_months[0]
        trend_start_ts = trend_start_month.to_timestamp(how="start")
        if len(active_created_df):
            active_last_6m = set(
                active_created_df[active_created_df["created_at"] >= trend_start_ts]["email"].tolist()
            )
        else:
            active_last_6m = set()

        active_total = len(active_last_6m)

        # Primary-source attribution: each subscriber is assigned to one source container
        # based on earliest source-tag timestamp (first touch).
        source_order = list(config.ENTRY_SOURCE_TAG_GROUPS.keys())
        assigned_by_source: dict[str, set[str]] = {name: set() for name in source_order}
        assigned_emails: set[str] = set()

        for email in active_last_6m:
            candidates: list[tuple[tuple[int, datetime | None, str], str]] = []
            for source_name in source_order:
                tagged_map = segments.entry_source_tagged_at.get(source_name, {})
                if email in tagged_map:
                    candidates.append((self._iso_key(tagged_map.get(email)), source_name))
                elif email in segments.entry_sources.get(source_name, set()):
                    # Fallback when tagged_at is missing: still allow source assignment,
                    # but rank after real timestamps.
                    candidates.append(((2, None, ""), source_name))

            if not candidates:
                continue

            # Sort by (has_timestamp, timestamp, original value) then source order.
            candidates.sort(
                key=lambda item: (item[0], source_order.index(item[1]))
            )
            chosen_source = candidates[0][1]
            assigned_by_source[chosen_source].add(email)
            assigned_emails.add(email)

        for source_name in source_order:
            members = assigned_by_source.get(source_name, set())
            share = (len(members) / active_total * 100.0) if active_total > 0 else 0.0
            source_rows.append(
                {
                    "source": source_name,
                    "confirmed_subscribers": int(len(members)),
                    "share_pct": round(share, 2),
                }
            )
        other_members = active_last_6m - assigned_emails
        if len(other_members) > 0:
            source_rows.append(
                {
                    "source": "Other / Not tagged",
                    "confirmed_subscribers": int(len(other_members)),
                    "share_pct": round((len(other_members) / active_total * 100.0), 2)
                    if active_total > 0
                    else 0.0,
                }
            )
        source_breakdown_df = pd.DataFrame(source_rows)

        # KPI 19: last 4 completed months snapshot
        snapshot_months = self._completed_months(config.SNAPSHOT_MONTHS, now_utc)
        snapshot_rows = []
        for month in snapshot_months:
            md = bdf[bdf["month"] == month]
            vd = md[md["category"] == "Value"]
            sd = md[md["category"] == "Sales"]

            value_or = (vd["opens"].sum() / vd["recipients"].sum() * 100.0) if vd["recipients"].sum() > 0 else 0.0
            sales_or = (sd["opens"].sum() / sd["recipients"].sum() * 100.0) if sd["recipients"].sum() > 0 else 0.0
            sales_ctor = (sd["clicks"].sum() / sd["opens"].sum() * 100.0) if sd["opens"].sum() > 0 else 0.0

            snapshot_rows.append(
                {
                    "month": str(month),
                    "value_open_rate": round(value_or, 2),
                    "sales_open_rate": round(sales_or, 2),
                    "sales_ctor": round(sales_ctor, 2),
                    "value_sends": int(len(vd)),
                    "sales_sends": int(len(sd)),
                }
            )
        snapshot_df = pd.DataFrame(snapshot_rows)

        # KPI 16: monthly churn rate (unsubscribe proxy)
        churn_months = self._completed_months(config.CHURN_MONTHS, now_utc)
        churn_rows = []
        for month in churn_months:
            md = bdf[bdf["month"] == month]
            recipients = int(md["recipients"].sum())
            unsubs = int(md["unsubs"].sum())
            churn_rate = (unsubs / recipients * 100.0) if recipients > 0 else 0.0
            churn_rows.append(
                {
                    "month": str(month),
                    "churn_rate": round(churn_rate, 4),
                    "unsubs": unsubs,
                    "recipients": recipients,
                    "sends": int(len(md)),
                }
            )
        churn_df = pd.DataFrame(churn_rows)

        # KPI 3: sales CTOR by segment (rolling 30 days)
        sales_start = now_utc - timedelta(days=config.ROLLING_DAYS_SALES_CTOR)
        sales_ids = sorted(
            bdf[(bdf["category"] == "Sales") & (bdf["sent_at"] >= sales_start)]["broadcast_id"]
            .dropna()
            .astype(int)
            .unique()
            .tolist()
        )

        sales_openers = self.client.filter_subscribers_by_broadcast_event("opens", sales_ids)
        sales_clickers = self.client.filter_subscribers_by_broadcast_event("clicks", sales_ids)

        segment_map = {
            "Lead Magnet Responders": segments.lead_magnet,
            "Workshop Buyers": segments.workshop,
            "Bootcamp Buyers": segments.bootcamp,
            "Active Non-Buyers": segments.active_non_buyers,
        }

        sales_ctor_rows = []
        for seg_name, seg_emails in segment_map.items():
            open_count = len(seg_emails & sales_openers)
            click_count = len(seg_emails & sales_clickers)
            ctor = (click_count / open_count * 100.0) if open_count > 0 else 0.0
            sales_ctor_rows.append(
                {
                    "segment": seg_name,
                    "segment_size": int(len(seg_emails)),
                    "openers": int(open_count),
                    "clickers": int(click_count),
                    "sales_ctor": round(ctor, 2),
                }
            )
        sales_ctor_df = pd.DataFrame(sales_ctor_rows)

        # KPI 6: cold -> re-warmed movement rate
        rewarm_window_start = now_utc - timedelta(days=config.REWARM_WINDOW_DAYS)
        cold_pre_start = rewarm_window_start - timedelta(days=config.COLD_LOOKBACK_DAYS)

        openers_pre = self.client.filter_subscribers_by_event_date(
            "opens",
            after_iso_date=cold_pre_start.date().isoformat(),
            before_iso_date=rewarm_window_start.date().isoformat(),
        )
        openers_last30 = self.client.filter_subscribers_by_event_date(
            "opens",
            after_iso_date=rewarm_window_start.date().isoformat(),
            before_iso_date=(now_utc + timedelta(days=1)).date().isoformat(),
        )
        openers_lookback_now = self.client.filter_subscribers_by_event_date(
            "opens",
            after_iso_date=(now_utc - timedelta(days=config.COLD_LOOKBACK_DAYS)).date().isoformat(),
            before_iso_date=(now_utc + timedelta(days=1)).date().isoformat(),
        )

        cold_before_last30 = segments.active_before_window - openers_pre
        rewarmed = cold_before_last30 & openers_last30
        current_cold = segments.active_all - openers_lookback_now

        movement_rate = (len(rewarmed) / len(current_cold) * 100.0) if len(current_cold) > 0 else 0.0

        kpi6 = {
            "cold_before_last30_count": int(len(cold_before_last30)),
            "current_cold_count": int(len(current_cold)),
            "rewarmed_count": int(len(rewarmed)),
            "movement_rate": round(movement_rate, 2),
            "window_days": config.REWARM_WINDOW_DAYS,
            "cold_lookback_days": config.COLD_LOOKBACK_DAYS,
            "window_start_utc": rewarm_window_start.isoformat(),
            "window_end_utc": now_utc.isoformat(),
        }

        # KPI 12: workshop -> bootcamp conversion rate by segment
        workshop_all = segments.workshop
        workshop_lead = segments.workshop & segments.lead_magnet
        workshop_non_lead = segments.workshop - segments.lead_magnet

        def mapped_converted_in_segment(seg_emails: set[str]) -> set[str]:
            converted: set[str] = set()
            for family in self._target_workshop_families():
                workshop_members = segments.workshop_by_tag.get(family, set()) & seg_emails
                mapped_bootcamp_members = segments.mapped_bootcamp_by_workshop.get(family, set())
                converted |= workshop_members & mapped_bootcamp_members
            return converted

        conv_segments = {
            "All Workshop Buyers": workshop_all,
            "Lead-Magnet Workshop Buyers": workshop_lead,
            "Non-Lead Workshop Buyers": workshop_non_lead,
        }

        conv_rows = []
        for seg_name, seg_emails in conv_segments.items():
            converted = mapped_converted_in_segment(seg_emails)
            denom = len(seg_emails)
            rate = (len(converted) / denom * 100.0) if denom > 0 else 0.0
            conv_rows.append(
                {
                    "segment": seg_name,
                    "workshop_buyers": int(denom),
                    "also_bootcamp": int(len(converted)),
                    "conversion_rate": round(rate, 2),
                }
            )
        workshop_to_bootcamp_df = pd.DataFrame(conv_rows)

        # Optional detail: conversion by workshop program tag.
        program_rows = []
        for tag_name in self._target_workshop_families():
            members = segments.workshop_by_tag.get(tag_name, set())
            converted = members & segments.mapped_bootcamp_by_workshop.get(tag_name, set())
            denom = len(members)
            rate = (len(converted) / denom * 100.0) if denom > 0 else 0.0
            program_rows.append(
                {
                    "workshop_program": tag_name,
                    "workshop_buyers": int(denom),
                    "also_bootcamp": int(len(converted)),
                    "conversion_rate": round(rate, 2),
                }
            )
        workshop_program_df = pd.DataFrame(program_rows)

        # Label coverage diagnostics (important for API-only categorization quality)
        labeled = bdf[bdf["category"].isin(["Value", "Sales"])]
        label_coverage = (
            (len(labeled) / len(bdf) * 100.0) if len(bdf) > 0 else 0.0
        )

        return {
            "generated_at_utc": now_utc.isoformat(),
            "kpi14_current_confirmed": kpi14_current_confirmed,
            "kpi_confirmed_trend_6m": confirmed_trend_df,
            "kpi_confirmed_source_breakdown": source_breakdown_df,
            "state_counts": state_counts,
            "segment_sizes": {
                "lead_magnet": len(segments.lead_magnet),
                "workshop": len(segments.workshop),
                "bootcamp": len(segments.bootcamp),
                "active_non_buyers": len(segments.active_non_buyers),
            },
            "kpi19_snapshot": snapshot_df,
            "kpi16_churn": churn_df,
            "kpi3_sales_ctor_by_segment": sales_ctor_df,
            "kpi6_rewarm": kpi6,
            "kpi12_workshop_to_bootcamp": workshop_to_bootcamp_df,
            "kpi12_workshop_program": workshop_program_df,
            "label_coverage_pct": round(label_coverage, 2),
            "sales_window_broadcasts": len(sales_ids),
            "failed_filter_requests": self.client.failed_filter_requests,
        }

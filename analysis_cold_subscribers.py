"""
Cold Subscriber Analysis
========================
Generates charts W1–W4 for the cold subscriber section of the report.

Input:  Cold Subscribers.csv
Output: charts/W1_cold_age_distribution.png
        charts/W2_cold_tag_breakdown.png
        charts/W3_cold_phase_bar.png
        charts/W4_cold_tag_x_phase.png
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from pathlib import Path
import warnings
warnings.filterwarnings("ignore")

BASE = Path(__file__).resolve().parent
OUT  = BASE / "charts"
OUT.mkdir(exist_ok=True)

TODAY      = pd.Timestamp("2026-02-21")
PHASE_CUT1 = pd.Timestamp("2025-06-01")
PHASE_CUT2 = pd.Timestamp("2025-11-01")

GREEN  = "#10B981"
ORANGE = "#F59E0B"
RED    = "#EF4444"
BRAND  = "#4F46E5"
ACCENT = "#EC4899"
MUTED  = "#94A3B8"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.titlesize": 14,
    "axes.labelsize": 11,
    "figure.dpi": 150,
})

# ── Load ──────────────────────────────────────────────────────────────────────
df = pd.read_csv(BASE / "Cold Subscribers.csv", encoding="utf-8-sig")
df["created_at"] = pd.to_datetime(df["created_at"], utc=True).dt.tz_localize(None)
df["age_days"]   = (TODAY - df["created_at"]).dt.days

# ── Phase when they subscribed ────────────────────────────────────────────────
def assign_phase(d):
    if d >= PHASE_CUT2: return "Visible Decline\n(Nov 2025–now)"
    if d >= PHASE_CUT1: return "Silent Warning\n(Jun–Oct 2025)"
    return "Strong Start\n(pre-Jun 2025)"

df["sub_phase"] = df["created_at"].apply(assign_phase)

PHASE_ORDER = [
    "Strong Start\n(pre-Jun 2025)",
    "Silent Warning\n(Jun–Oct 2025)",
    "Visible Decline\n(Nov 2025–now)",
]
PHASE_COLORS = {
    PHASE_ORDER[0]: GREEN,
    PHASE_ORDER[1]: ORANGE,
    PHASE_ORDER[2]: RED,
}

# ── Age buckets ───────────────────────────────────────────────────────────────
AGE_BINS   = [0, 90, 180, 365, 550, 730, 9999]
AGE_LABELS = ["< 3 months", "3–6 months", "6–12 months", "12–18 months", "18–24 months", "> 2 years"]
df["age_bucket"] = pd.cut(df["age_days"], bins=AGE_BINS, labels=AGE_LABELS)

# ── Tag grouping ──────────────────────────────────────────────────────────────
# Primary tag = first tag in the comma-separated list (acquisition source)
df["primary_tag"] = df["tags"].fillna("No Tag").str.split(",").str[0].str.strip()

TAG_MAP = {
    "AI Livestream Kit":                              "Livestream / YouTube",
    "AI Livestream List (Import CSV)":                "Livestream / YouTube",
    "AI Program List (CSV Import)":                   "AI Program Import",
    "AI Program Kit":                                 "AI Program Import",
    "AI Agent Bootcamp Waitlist":                     "Bootcamp Waitlist",
    "Been Sent 2025 Roadmap Survey":                  "Survey / Other",
    "Kit x Substack":                                 "Substack Import",
    "Agent Breakthrough":                             "Agent Breakthrough",
    "AI App Sprint":                                  "AI App Sprint",
    "Freelance Accelerator Bundle":                   "Freelance Accelerator",
    "Freelance Accelerator UPSELL [LIVE Masterclass]":"Freelance Accelerator",
    "Alumni Beta to AI":                              "Alumni / Paid Program",
    "AI Agent Bootcamp Core [Jul 2025]":              "Alumni / Paid Program",
    "AI Agent Core [Sept]":                           "Alumni / Paid Program",
    "AI Agent Core [Oct]":                            "Alumni / Paid Program",
    "AI Agent Bootcamp Core [Apr 2025]":              "Alumni / Paid Program",
    "Free":                                           "Survey / Other",
    "No Tag":                                         "No Tag",
}
df["tag_group"] = df["primary_tag"].map(TAG_MAP).fillna("Other")

# Consolidate tiny groups into "Other" for cleaner charts
KEEP_TAGS = ["Livestream / YouTube", "AI Program Import", "Bootcamp Waitlist",
             "Substack Import", "Alumni / Paid Program", "No Tag / Other"]
df["tag_group_clean"] = df["tag_group"].apply(
    lambda t: t if t in ["Livestream / YouTube", "AI Program Import",
                          "Bootcamp Waitlist", "Substack Import",
                          "Alumni / Paid Program"] else "No Tag / Other"
)

# ── Print summary ─────────────────────────────────────────────────────────────
print("=" * 60)
print(f"COLD SUBSCRIBERS — TOTAL: {len(df):,}")
print("=" * 60)
print(f"\nDate range: {df['created_at'].min().date()} → {df['created_at'].max().date()}")
print(f"\nAge distribution:")
print(df["age_bucket"].value_counts().sort_index().to_string())
print(f"\nSubscription phase:")
for ph in PHASE_ORDER:
    n = (df["sub_phase"] == ph).sum()
    print(f"  {ph.replace(chr(10),' | ')}: {n:,}  ({n/len(df)*100:.1f}%)")
print(f"\nTag groups:")
print(df["tag_group_clean"].value_counts().to_string())

# ─────────────────────────────────────────────────────────────────────────────
# CHART W1 — Age distribution of cold subscribers
#            Phase-coloured bars to show when they signed up
# ─────────────────────────────────────────────────────────────────────────────
# Map age bucket → likely phase of sign-up for colour
# (based on TODAY = Feb 21 2026)
# < 3 months    → signed up Nov 2025–Feb 2026  → Visible Decline
# 3–6 months    → signed up Aug–Nov 2025        → Silent Warning / Visible Decline
# 6–12 months   → signed up Feb–Aug 2025        → Strong Start / Silent Warning
# 12–18 months  → signed up Aug 2024–Feb 2025   → pre-analysis period
# 18–24 months  → signed up Feb–Aug 2024        → pre-analysis period
AGE_COLORS = {
    "< 3 months":    RED,
    "3–6 months":    ORANGE,
    "6–12 months":   ORANGE,
    "12–18 months":  BRAND,
    "18–24 months":  BRAND,
    "> 2 years":     MUTED,
}

age_counts = df["age_bucket"].value_counts().reindex(AGE_LABELS).fillna(0)
total = len(df)

fig, ax = plt.subplots(figsize=(11, 6))
bars = ax.bar(AGE_LABELS,
              age_counts.values,
              color=[AGE_COLORS[l] for l in AGE_LABELS],
              edgecolor="white", linewidth=0.5, alpha=0.9)

for bar, val in zip(bars, age_counts.values):
    if val > 0:
        ax.text(bar.get_x() + bar.get_width()/2,
                bar.get_height() + 8,
                f"{int(val):,}\n({val/total*100:.1f}%)",
                ha="center", fontsize=10, fontweight="bold")

ax.set_ylabel("Number of Cold Subscribers")
ax.set_title(f"How Old Are the Cold Subscribers?\n"
             f"Age Since Sign-Up  (total = {total:,})")
ax.set_ylim(0, age_counts.max() * 1.18)

legend_patches = [
    mpatches.Patch(color=RED,    label="Signed up during Visible Decline (Nov 2025+)"),
    mpatches.Patch(color=ORANGE, label="Signed up during Silent Warning (Jun–Oct 2025)"),
    mpatches.Patch(color=BRAND,  label="Signed up during / before Strong Start (pre-Jun 2025)"),
]
ax.legend(handles=legend_patches, fontsize=9, loc="upper right")
ax.set_xlabel("Age of Cold Subscriber (time since sign-up)")
fig.tight_layout()
fig.savefig(OUT / "W1_cold_age_distribution.png")
plt.close()
print("\nChart W1 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART W2 — Cold subscribers by tag group (horizontal bar)
#            Shows acquisition source of cold subs
# ─────────────────────────────────────────────────────────────────────────────
tag_counts = df["tag_group_clean"].value_counts()

TAG_COLORS = {
    "Livestream / YouTube": "#3B82F6",
    "AI Program Import":    "#8B5CF6",
    "Bootcamp Waitlist":    ORANGE,
    "Substack Import":      "#14B8A6",
    "Alumni / Paid Program": GREEN,
    "No Tag / Other":       MUTED,
}

fig, ax = plt.subplots(figsize=(11, 6))
bars = ax.barh(
    tag_counts.index[::-1],
    tag_counts.values[::-1],
    color=[TAG_COLORS.get(t, MUTED) for t in tag_counts.index[::-1]],
    edgecolor="white", linewidth=0.5, alpha=0.9
)
for bar, val in zip(bars, tag_counts.values[::-1]):
    ax.text(bar.get_width() + 8, bar.get_y() + bar.get_height()/2,
            f"{int(val):,}  ({val/total*100:.1f}%)",
            va="center", fontsize=10, fontweight="bold")

ax.set_xlabel("Number of Cold Subscribers")
ax.set_title(f"Cold Subscribers by Acquisition Source (Tag)\n"
             f"(total = {total:,}  |  tag = first tag assigned at sign-up)")
ax.set_xlim(0, tag_counts.max() * 1.28)
fig.tight_layout()
fig.savefig(OUT / "W2_cold_tag_breakdown.png")
plt.close()
print("Chart W2 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART W3 — Phase bar: when did cold subscribers sign up?
#            Compares to total list composition to show if phase 3 is over-represented
# ─────────────────────────────────────────────────────────────────────────────
phase_counts = df["sub_phase"].value_counts().reindex(PHASE_ORDER).fillna(0)

fig, ax = plt.subplots(figsize=(9, 6))
bars = ax.bar(
    [ph.replace("\n", "\n") for ph in PHASE_ORDER],
    phase_counts.values,
    color=[PHASE_COLORS[ph] for ph in PHASE_ORDER],
    edgecolor="white", linewidth=0.5, alpha=0.9, width=0.55
)
for bar, val in zip(bars, phase_counts.values):
    pct = val / total * 100
    ax.text(bar.get_x() + bar.get_width()/2,
            bar.get_height() + 10,
            f"{int(val):,}\n({pct:.1f}%)",
            ha="center", fontsize=11, fontweight="bold")

ax.set_ylabel("Number of Cold Subscribers")
ax.set_title("When Did Cold Subscribers Join?\n"
             "Sign-up Phase Breakdown")
ax.set_ylim(0, phase_counts.max() * 1.2)
ax.set_xticklabels([ph.replace("\n", "\n") for ph in PHASE_ORDER], fontsize=10)
fig.tight_layout()
fig.savefig(OUT / "W3_cold_phase_bar.png")
plt.close()
print("Chart W3 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART W4 — Stacked bar: tag group × phase
#            Shows which acquisition sources feed cold subs in each phase
# ─────────────────────────────────────────────────────────────────────────────
cross = (
    df.groupby(["sub_phase", "tag_group_clean"])
      .size()
      .unstack(fill_value=0)
      .reindex(PHASE_ORDER)
)

# Order tag groups by total size
tag_order = df["tag_group_clean"].value_counts().index.tolist()
cross = cross.reindex(columns=tag_order, fill_value=0)

fig, ax = plt.subplots(figsize=(11, 7))
bottom = np.zeros(len(PHASE_ORDER))

for tag in tag_order:
    vals = cross[tag].values
    bars = ax.bar(
        range(len(PHASE_ORDER)),
        vals, bottom=bottom,
        color=TAG_COLORS.get(tag, MUTED),
        label=tag, edgecolor="white", linewidth=0.5, alpha=0.9
    )
    for i, (bar, v) in enumerate(zip(bars, vals)):
        if v >= 15:
            ax.text(i, bottom[i] + v/2,
                    f"{int(v):,}",
                    ha="center", va="center",
                    fontsize=9, fontweight="bold", color="white")
    bottom += vals

# Add total label on top
for i, tot in enumerate(cross.sum(axis=1).values):
    ax.text(i, tot + 8, f"n={int(tot):,}",
            ha="center", fontsize=10, fontweight="bold", color="#1e293b")

ax.set_xticks(range(len(PHASE_ORDER)))
ax.set_xticklabels([ph.replace("\n", "\n") for ph in PHASE_ORDER], fontsize=10)
ax.set_ylabel("Number of Cold Subscribers")
ax.set_title("Cold Subscribers by Acquisition Source × Sign-up Phase\n"
             "(stacked = tag group composition per phase)")
ax.legend(title="Acquisition Source (Tag)", bbox_to_anchor=(1.01, 1),
          loc="upper left", fontsize=9)
ax.set_ylim(0, cross.sum(axis=1).max() * 1.15)
fig.tight_layout()
fig.savefig(OUT / "W4_cold_tag_x_phase.png")
plt.close()
print("Chart W4 saved.")

# ── Key stats for report text ─────────────────────────────────────────────────
print("\n" + "=" * 60)
print("KEY STATS FOR REPORT")
print("=" * 60)
pct_phase3 = (df["sub_phase"] == PHASE_ORDER[2]).sum() / total * 100
pct_phase1 = (df["sub_phase"] == PHASE_ORDER[0]).sum() / total * 100
pct_phase2 = (df["sub_phase"] == PHASE_ORDER[1]).sum() / total * 100
oldest_group = age_counts.idxmax()
print(f"Total cold: {total:,}")
print(f"% from Strong Start (pre-Jun 2025): {pct_phase1:.1f}%")
print(f"% from Silent Warning (Jun–Oct 2025): {pct_phase2:.1f}%")
print(f"% from Visible Decline (Nov 2025+): {pct_phase3:.1f}%")
print(f"Largest age bucket: {oldest_group} ({int(age_counts[oldest_group]):,})")
print(f"Median age (days): {df['age_days'].median():.0f}")
print(f"Mean age (days): {df['age_days'].mean():.0f}")
print(f"Top tag group: {tag_counts.index[0]} ({int(tag_counts.iloc[0]):,}, {tag_counts.iloc[0]/total*100:.1f}%)")

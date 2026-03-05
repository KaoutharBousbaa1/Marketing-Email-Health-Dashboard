"""
Cold Subscriber Engagement Analysis — Kit API Data
===================================================
Input:  cold_engagement.csv  (output of fetch_cold_engagement.py)
        Cold Subscribers.csv (for sign-up metadata)
Output: charts/X1_last_open_timeline.png
        charts/X2_sends_since_last_open.png
        charts/X3_never_opened_vs_went_cold.png
        charts/X4_disengagement_cohort.png

Key questions answered:
  1. When did cold subscribers last open an email?  (timeline)
  2. How many sends have they ignored since last open?
  3. Were they always cold (never opened) or did they disengage?
  4. For those who went cold: in which phase did they disengage?
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.dates as mdates
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
df = pd.read_csv(BASE / "cold_engagement.csv")
df["created_at"] = pd.to_datetime(df["created_at"], utc=True, errors="coerce").dt.tz_localize(None)
df["last_opened"] = pd.to_datetime(df["last_opened"], utc=True, errors="coerce").dt.tz_localize(None)
df["last_clicked"] = pd.to_datetime(df["last_clicked"], utc=True, errors="coerce").dt.tz_localize(None)

total = len(df)
never_opened = df["last_opened"].isna().sum()
ever_opened  = df["last_opened"].notna().sum()

print("=" * 60)
print(f"COLD ENGAGEMENT ANALYSIS  (n={total:,})")
print("=" * 60)
print(f"  Never opened any email : {never_opened:,}  ({never_opened/total*100:.1f}%)")
print(f"  Opened at least once   : {ever_opened:,}  ({ever_opened/total*100:.1f}%)")
print(f"  Ever clicked           : {df['last_clicked'].notna().sum():,}")
print()
print(f"sends_since_last_open (non-null):")
print(df["sends_since_last_open"].describe().to_string())

# ── Classify subscribers ───────────────────────────────────────────────────────
def classify(row):
    if pd.isna(row["last_opened"]):
        return "Never Opened"
    lo = row["last_opened"]
    if lo >= PHASE_CUT2:
        return "Went Cold in\nVisible Decline"
    if lo >= PHASE_CUT1:
        return "Went Cold in\nSilent Warning"
    return "Went Cold in\nStrong Start"

df["cold_class"] = df.apply(classify, axis=1)

CLASS_ORDER = [
    "Never Opened",
    "Went Cold in\nStrong Start",
    "Went Cold in\nSilent Warning",
    "Went Cold in\nVisible Decline",
]
CLASS_COLORS = {
    "Never Opened":                MUTED,
    "Went Cold in\nStrong Start":  GREEN,
    "Went Cold in\nSilent Warning":ORANGE,
    "Went Cold in\nVisible Decline":RED,
}

class_counts = df["cold_class"].value_counts().reindex(CLASS_ORDER).fillna(0)
print("\nCold classification:")
for k, v in class_counts.items():
    print(f"  {k.replace(chr(10),' | ')}: {int(v):,}  ({v/total*100:.1f}%)")

# ── Phase assignment ───────────────────────────────────────────────────────────
PHASE_ORDER = [
    "Strong Start\n(pre-Jun 2025)",
    "Silent Warning\n(Jun–Oct 2025)",
    "Visible Decline\n(Nov 2025–now)",
]
PHASE_COLORS = {PHASE_ORDER[0]: GREEN, PHASE_ORDER[1]: ORANGE, PHASE_ORDER[2]: RED}

def assign_phase(d):
    if pd.isna(d): return "Unknown"
    if d >= PHASE_CUT2: return PHASE_ORDER[2]
    if d >= PHASE_CUT1: return PHASE_ORDER[1]
    return PHASE_ORDER[0]

df["diseng_phase"] = df["last_opened"].apply(assign_phase)

# ─────────────────────────────────────────────────────────────────────────────
# CHART X1 — Timeline of last open dates (monthly bar chart)
#             Shows *when* cold subs stopped engaging
# ─────────────────────────────────────────────────────────────────────────────
opened_df = df[df["last_opened"].notna()].copy()
opened_df["last_open_month"] = opened_df["last_opened"].dt.to_period("M")

monthly = opened_df.groupby("last_open_month").size().reset_index(name="count")
monthly["month_dt"] = monthly["last_open_month"].dt.to_timestamp()
monthly = monthly.sort_values("month_dt")

def month_color(dt):
    if dt >= PHASE_CUT2: return RED
    if dt >= PHASE_CUT1: return ORANGE
    return GREEN

colors_x1 = [month_color(row["month_dt"]) for _, row in monthly.iterrows()]

fig, ax = plt.subplots(figsize=(13, 6))
bars = ax.bar(range(len(monthly)), monthly["count"].values,
              color=colors_x1, edgecolor="white", linewidth=0.4, alpha=0.9, width=0.8)

# Phase boundary lines
phase1_idx = next((i for i, r in monthly.iterrows() if r["month_dt"] >= PHASE_CUT1), None)
phase2_idx = next((i for i, r in monthly.iterrows() if r["month_dt"] >= PHASE_CUT2), None)

if phase1_idx is not None:
    ax.axvline(phase1_idx - 0.5, color=ORANGE, linestyle="--", linewidth=1.5, alpha=0.7)
    ax.text(phase1_idx - 0.4, monthly["count"].max() * 0.95,
            "Jun 2025\n(Silent Warning starts)", color=ORANGE, fontsize=8, va="top")
if phase2_idx is not None:
    ax.axvline(phase2_idx - 0.5, color=RED, linestyle="--", linewidth=1.5, alpha=0.7)
    ax.text(phase2_idx - 0.4, monthly["count"].max() * 0.95,
            "Nov 2025\n(Visible Decline starts)", color=RED, fontsize=8, va="top")

ax.set_xticks(range(len(monthly)))
ax.set_xticklabels([r["month_dt"].strftime("%b %Y") for _, r in monthly.iterrows()],
                   rotation=45, ha="right", fontsize=9)

for bar, val in zip(bars, monthly["count"].values):
    if val >= 20:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
                str(int(val)), ha="center", fontsize=8, fontweight="bold")

legend_patches = [
    mpatches.Patch(color=GREEN,  label="Last opened during Strong Start (pre-Jun 2025)"),
    mpatches.Patch(color=ORANGE, label="Last opened during Silent Warning (Jun–Oct 2025)"),
    mpatches.Patch(color=RED,    label="Last opened during Visible Decline (Nov 2025+)"),
]
ax.legend(handles=legend_patches, fontsize=9)
ax.set_ylabel("Number of Cold Subscribers")
ax.set_title(f"When Did Cold Subscribers Last Open an Email?\n"
             f"({ever_opened:,} of {total:,} cold subscribers had at least one open — "
             f"{never_opened:,} never opened)")
fig.tight_layout()
fig.savefig(OUT / "X1_last_open_timeline.png")
plt.close()
print("\nChart X1 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART X2 — Sends since last open: how long have they been cold?
#             Histogram showing depth of disengagement
# ─────────────────────────────────────────────────────────────────────────────
sslo = df["sends_since_last_open"].dropna()
bins = [0, 20, 40, 60, 80, 100, 120, 130]
labels_x2 = ["1–20", "21–40", "41–60", "61–80", "81–100", "101–120", "121+"]
cut = pd.cut(sslo, bins=bins, labels=labels_x2)
counts_x2 = cut.value_counts().reindex(labels_x2).fillna(0)

# colour by severity: the more sends ignored, the redder
SEVERITY_COLORS = ["#22C55E","#84CC16","#EAB308","#F97316","#EF4444","#DC2626","#991B1B"]

fig, ax = plt.subplots(figsize=(11, 6))
bars = ax.bar(labels_x2, counts_x2.values,
              color=SEVERITY_COLORS, edgecolor="white", linewidth=0.5, alpha=0.9)

for bar, val in zip(bars, counts_x2.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
            f"{int(val):,}\n({val/len(sslo)*100:.1f}%)",
            ha="center", fontsize=9, fontweight="bold")

ax.set_xlabel("Number of Sends Since Last Open")
ax.set_ylabel("Number of Cold Subscribers")
ax.set_title(f"How Long Have Cold Subscribers Been Ignoring Emails?\n"
             f"Sends Since Last Open  (median = {int(sslo.median())}, mean = {sslo.mean():.0f}  |  "
             f"{int(df['sends_since_last_open'].isna().sum()):,} never opened = excluded)")
ax.set_ylim(0, counts_x2.max() * 1.2)

# Add median line
ax.axvline(sslo.median() / 20 - 0.5, color=BRAND, linestyle=":", linewidth=2, alpha=0.6)

fig.tight_layout()
fig.savefig(OUT / "X2_sends_since_last_open.png")
plt.close()
print("Chart X2 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART X3 — Never opened vs went cold: donut / stacked bar
#             Distinguishes truly cold-from-start vs disengaged over time
# ─────────────────────────────────────────────────────────────────────────────
class_vals = class_counts.values
class_labels = [c.replace("\n", " ") for c in CLASS_ORDER]
class_colors  = [CLASS_COLORS[c] for c in CLASS_ORDER]

fig, (ax_pie, ax_bar) = plt.subplots(1, 2, figsize=(14, 6))

# Left: Donut
wedges, texts, autotexts = ax_pie.pie(
    class_vals, labels=None,
    colors=class_colors,
    autopct=lambda p: f"{p:.1f}%\n({int(p/100*total):,})",
    startangle=90, pctdistance=0.75,
    wedgeprops={"edgecolor": "white", "linewidth": 1.5},
)
for at in autotexts:
    at.set_fontsize(9)
    at.set_fontweight("bold")
centre = plt.Circle((0, 0), 0.55, fc="white")
ax_pie.add_artist(centre)
ax_pie.text(0, 0, f"n={total:,}", ha="center", va="center", fontsize=12, fontweight="bold")
ax_pie.set_title("Cold Subscribers by Disengagement Type", fontsize=13)
ax_pie.legend(handles=[
    mpatches.Patch(color=CLASS_COLORS[c], label=c.replace("\n", " "))
    for c in CLASS_ORDER
], fontsize=9, loc="lower center", bbox_to_anchor=(0.5, -0.08), ncol=2)

# Right: Horizontal bar showing counts clearly
bars = ax_bar.barh(
    class_labels[::-1], class_vals[::-1],
    color=class_colors[::-1], edgecolor="white", linewidth=0.5, alpha=0.9
)
for bar, val in zip(bars, class_vals[::-1]):
    ax_bar.text(bar.get_width() + 5, bar.get_y() + bar.get_height()/2,
                f"{int(val):,}  ({val/total*100:.1f}%)",
                va="center", fontsize=10, fontweight="bold")
ax_bar.set_xlabel("Number of Cold Subscribers")
ax_bar.set_title("Breakdown of Cold Subscribers", fontsize=13)
ax_bar.set_xlim(0, class_vals.max() * 1.3)

fig.suptitle("Never Engaged vs Disengaged Over Time", fontsize=15, fontweight="bold", y=1.01)
fig.tight_layout()
fig.savefig(OUT / "X3_disengagement_type.png", bbox_inches="tight")
plt.close()
print("Chart X3 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART X4 — Cohort analysis: sign-up phase × disengagement phase
#             For those who DID open, when did they stop?
# ─────────────────────────────────────────────────────────────────────────────
# Only subs who opened at least once
engaged = df[df["last_opened"].notna()].copy()
engaged["signup_phase"] = engaged["created_at"].apply(assign_phase)

# diseng_phase already on df
cross4 = (
    engaged.groupby(["signup_phase", "diseng_phase"])
    .size()
    .unstack(fill_value=0)
    .reindex(index=PHASE_ORDER, fill_value=0)
)
# Reorder columns to phase order (only phases that appear)
col_order = [p for p in PHASE_ORDER if p in cross4.columns]
cross4 = cross4.reindex(columns=col_order, fill_value=0)

fig, ax = plt.subplots(figsize=(11, 6))
bottom = np.zeros(len(PHASE_ORDER))

for phase in col_order:
    vals = cross4[phase].values if phase in cross4.columns else np.zeros(len(PHASE_ORDER))
    bars = ax.bar(
        range(len(PHASE_ORDER)), vals, bottom=bottom,
        color=PHASE_COLORS.get(phase, MUTED),
        label=f"Last opened during {phase.replace(chr(10), ' ')}",
        edgecolor="white", linewidth=0.5, alpha=0.9
    )
    for i, (bar, v) in enumerate(zip(bars, vals)):
        if v >= 10:
            ax.text(i, bottom[i] + v/2, f"{int(v):,}",
                    ha="center", va="center", fontsize=9, fontweight="bold", color="white")
    bottom += vals

# Totals on top
totals = cross4.sum(axis=1).values
for i, tot in enumerate(totals):
    ax.text(i, tot + 5, f"n={int(tot):,}", ha="center", fontsize=10, fontweight="bold")

ax.set_xticks(range(len(PHASE_ORDER)))
ax.set_xticklabels([p.replace("\n", "\n") for p in PHASE_ORDER], fontsize=10)
ax.set_ylabel("Number of Cold Subscribers (who opened at least once)")
ax.set_title("When Did Cold Subscribers Go Cold?\n"
             "Sign-up Phase vs Phase of Last Open\n"
             f"(only the {ever_opened:,} who had at least one open)")
ax.legend(title="Phase of Last Open", bbox_to_anchor=(1.01, 1), loc="upper left", fontsize=9)
ax.set_ylim(0, cross4.sum(axis=1).max() * 1.2)
fig.tight_layout()
fig.savefig(OUT / "X4_disengagement_cohort.png", bbox_inches="tight")
plt.close()
print("Chart X4 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# KEY STATS for report
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("KEY STATS FOR REPORT — COLD ENGAGEMENT")
print("=" * 60)
print(f"Total cold subscribers: {total:,}")
print(f"Never opened (ghost subs): {never_opened:,} ({never_opened/total*100:.1f}%)")
print(f"Opened at least once: {ever_opened:,} ({ever_opened/total*100:.1f}%)")
print(f"Ever clicked: {df['last_clicked'].notna().sum():,} ({df['last_clicked'].notna().sum()/total*100:.1f}%)")
print()
print("Disengagement phase (of those who opened at least once):")
for p in PHASE_ORDER:
    n = (engaged["diseng_phase"] == p).sum()
    print(f"  {p.replace(chr(10),' | ')}: {n:,} ({n/ever_opened*100:.1f}%)")
print()
print(f"Median sends since last open: {int(sslo.median())}")
print(f"Mean sends since last open: {sslo.mean():.0f}")
print(f"% who have been cold for 80+ sends: {(sslo >= 80).sum()/len(sslo)*100:.1f}%")
print(f"% who have been cold for 100+ sends: {(sslo >= 100).sum()/len(sslo)*100:.1f}%")
print()
print("Last open month distribution (top 5):")
print(monthly.sort_values("count", ascending=False).head(5)[["month_dt","count"]].to_string(index=False))

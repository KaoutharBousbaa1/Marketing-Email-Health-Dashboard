import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.patches as mpatches
import seaborn as sns
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

BASE = Path(__file__).resolve().parent
OUT  = BASE / "charts"
OUT.mkdir(exist_ok=True)

BRAND  = "#4F46E5"
ACCENT = "#EC4899"
GREEN  = "#10B981"
ORANGE = "#F59E0B"
RED    = "#EF4444"
MUTED  = "#94A3B8"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.titlesize": 14,
    "axes.labelsize": 11,
    "figure.dpi": 150,
})

# Phase boundaries (same as shift analysis)
phase_cut1 = pd.Timestamp("2025-06-01")
phase_cut2 = pd.Timestamp("2025-11-01")

PHASE_COLORS = {
    "Phase 1\nFeb–May 2025\n(baseline)":    GREEN,
    "Phase 2\nJun–Oct 2025\n(transition)":  ORANGE,
    "Phase 3\nNov 2025–now\n(decline)":     RED,
}
PHASE_SHORT = {
    "Phase 1\nFeb–May 2025\n(baseline)":    "Phase 1\nBaseline\nFeb–May 2025",
    "Phase 2\nJun–Oct 2025\n(transition)":  "Phase 2\nTransition\nJun–Oct 2025",
    "Phase 3\nNov 2025–now\n(decline)":     "Phase 3\nDecline\nNov 2025–now",
}

# ── Load subscribers ────────────────────────────────────────────────────────
df = pd.read_csv(
    BASE / "export (22).csv",
    encoding="utf-8-sig",
    parse_dates=["Date Subscribed", "Date Last Updated"],
)
df = df[df["State"].str.lower() == "cancelled"].copy()
df.dropna(subset=["Date Subscribed", "Date Last Updated"], inplace=True)
df["lifespan_days"] = (df["Date Last Updated"] - df["Date Subscribed"]).dt.days
df = df[df["lifespan_days"] >= 0]

# Assign phase by cancellation date
def assign_phase(d):
    if d < phase_cut1:  return "Phase 1\nFeb–May 2025\n(baseline)"
    if d < phase_cut2:  return "Phase 2\nJun–Oct 2025\n(transition)"
    return "Phase 3\nNov 2025–now\n(decline)"

df["phase"] = df["Date Last Updated"].apply(assign_phase)

phase_order = list(PHASE_COLORS.keys())
df["phase"] = pd.Categorical(df["phase"], categories=phase_order, ordered=True)

# Lifespan buckets
bins   = [0, 7, 30, 90, 180, 365, 9999]
blabels = ["0–7 d\n(very new)", "8–30 d\n(new)", "31–90 d\n(early)",
           "91–180 d\n(mid)", "181–365 d\n(long)", ">1 yr\n(veteran)"]
df["bucket"] = pd.cut(df["lifespan_days"], bins=bins, labels=blabels)

# ── Summary stats per phase ─────────────────────────────────────────────────
print("=" * 65)
print("LIFESPAN OF UNSUBSCRIBERS BY PHASE")
print("=" * 65)
for ph in phase_order:
    g = df[df["phase"] == ph]
    print(f"\n{ph.replace(chr(10), ' | ')}:")
    print(f"  Count          : {len(g):,}")
    print(f"  Median lifespan: {g['lifespan_days'].median():.0f} days")
    print(f"  Mean lifespan  : {g['lifespan_days'].mean():.1f} days")
    print(f"  Std dev        : {g['lifespan_days'].std():.1f} days")
    print(f"  % left within 30d  : {(g['lifespan_days']<=30).mean()*100:.1f}%")
    print(f"  % left after 180d  : {(g['lifespan_days']>180).mean()*100:.1f}%")

# ─────────────────────────────────────────────────────────────────────────────
# CHART M — KDE curves, one per phase
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(12, 6))
cap = 400
for ph, color in PHASE_COLORS.items():
    g = df[df["phase"] == ph]["lifespan_days"].clip(upper=cap)
    g.plot.kde(ax=ax, color=color, linewidth=2.5,
               label=f"{ph.replace(chr(10),' | ')}  (n={len(g):,}, median={g.median():.0f}d)")
    ax.axvline(g.median(), color=color, linewidth=1.2, linestyle="--", alpha=0.7)

ax.set_xlabel("Lifespan (days) — capped at 400")
ax.set_ylabel("Density")
ax.set_title("Lifespan Distribution of Unsubscribers — by Phase\n(dashed lines = phase medians)")
ax.legend(fontsize=9)
ax.set_xlim(left=-5)
fig.tight_layout()
fig.savefig(OUT / "M_phase_lifespan_kde.png")
plt.close()

# ─────────────────────────────────────────────────────────────────────────────
# CHART N — Grouped bar: bucket % per phase
# ─────────────────────────────────────────────────────────────────────────────
bucket_pct = (
    df.groupby(["phase", "bucket"], observed=True)
      .size()
      .groupby(level=0)
      .transform(lambda x: x / x.sum() * 100)
      .reset_index(name="pct")
)
pivot = bucket_pct.pivot(index="bucket", columns="phase", values="pct").reindex(blabels)

x = np.arange(len(blabels))
w = 0.27
fig, ax = plt.subplots(figsize=(13, 6))
for i, (ph, color) in enumerate(PHASE_COLORS.items()):
    vals = pivot[ph].fillna(0).values
    bars = ax.bar(x + (i-1)*w, vals, w, color=color, alpha=0.85,
                  label=ph.replace("\n", " | "), edgecolor="white")
    for bar, val in zip(bars, vals):
        if val > 1:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                    f"{val:.0f}%", ha="center", fontsize=7.5)

ax.set_xticks(x)
ax.set_xticklabels(blabels, fontsize=9)
ax.set_ylabel("% of unsubscribers in each phase")
ax.set_title("Who's Leaving in Each Phase? — Lifespan Bucket Breakdown per Phase")
ax.legend(fontsize=8.5)
fig.tight_layout()
fig.savefig(OUT / "N_phase_bucket_bars.png")
plt.close()

# ─────────────────────────────────────────────────────────────────────────────
# CHART O — Box plot per phase (with stats overlay)
# ─────────────────────────────────────────────────────────────────────────────
df_cap = df[df["lifespan_days"] <= 400].copy()

fig, ax = plt.subplots(figsize=(11, 6))
sns.boxplot(
    data=df_cap, x="phase", y="lifespan_days",
    order=phase_order,
    palette={ph: c for ph, c in PHASE_COLORS.items()},
    ax=ax,
    flierprops=dict(marker=".", markersize=2, alpha=0.25),
    width=0.5,
)
# Overlay median + count labels
for i, ph in enumerate(phase_order):
    g = df_cap[df_cap["phase"] == ph]["lifespan_days"]
    med = g.median()
    n   = len(g)
    mean = g.mean()
    ax.text(i, med + 8, f"median={med:.0f}d\nmean={mean:.0f}d\nn={n:,}",
            ha="center", fontsize=9, fontweight="bold",
            color="white",
            bbox=dict(boxstyle="round,pad=0.3",
                      fc=list(PHASE_COLORS.values())[i], alpha=0.88))

ax.set_xticklabels([ph.replace("\n", " | ") for ph in phase_order], fontsize=9)
ax.set_xlabel("")
ax.set_ylabel("Lifespan (days) — capped at 400")
ax.set_title("Lifespan of Unsubscribers per Phase — Box Plot")
fig.tight_layout()
fig.savefig(OUT / "O_phase_lifespan_boxplot.png")
plt.close()

# ─────────────────────────────────────────────────────────────────────────────
# CHART P — Stacked % bar (composition view: who leaves in each phase)
# ─────────────────────────────────────────────────────────────────────────────
stack_colors = ["#6EE7B7","#A7F3D0","#FDE68A","#FCA5A5","#F87171","#B91C1C"]

fig, ax = plt.subplots(figsize=(10, 6))
bottom = np.zeros(3)
for j, (bl, sc) in enumerate(zip(blabels, stack_colors)):
    vals = [pivot.loc[bl, ph] if bl in pivot.index and ph in pivot.columns
            else 0 for ph in phase_order]
    vals = [v if not np.isnan(v) else 0 for v in vals]
    bars = ax.bar(range(3), vals, bottom=bottom, color=sc,
                  label=bl.replace("\n", " "), edgecolor="white", linewidth=0.5)
    for k, (bar, v) in enumerate(zip(bars, vals)):
        if v > 4:
            ax.text(k, bottom[k] + v/2, f"{v:.0f}%",
                    ha="center", va="center", fontsize=8.5, fontweight="bold",
                    color="#1e293b")
    bottom += np.array(vals)

ax.set_xticks(range(3))
ax.set_xticklabels(["Phase 1\nBaseline\nFeb–May 2025",
                    "Phase 2\nTransition\nJun–Oct 2025",
                    "Phase 3\nDecline\nNov 2025–now"], fontsize=10)
ax.set_ylabel("% of unsubscribers")
ax.set_title("Composition of Unsubscribers by Lifespan — Phase by Phase\n(Darker = longer-tenured subscriber leaving)")
ax.legend(bbox_to_anchor=(1.01, 1), loc="upper left", fontsize=8.5, title="Lifespan bucket")
fig.tight_layout()
fig.savefig(OUT / "P_phase_stacked_composition.png")
plt.close()

# ─────────────────────────────────────────────────────────────────────────────
# CHART Q — Dual axis: phase open rate + median lifespan side by side
# ─────────────────────────────────────────────────────────────────────────────
phase_labels = ["Phase 1\nBaseline", "Phase 2\nTransition", "Phase 3\nDecline"]
open_rates   = [40.8, 40.3, 35.9]
ctors        = [12.6, 7.9, 6.4]
med_lifespans = [df[df["phase"]==ph]["lifespan_days"].median() for ph in phase_order]
counts        = [len(df[df["phase"]==ph]) for ph in phase_order]
colors        = [GREEN, ORANGE, RED]

fig, ax1 = plt.subplots(figsize=(10, 6))
x = np.arange(3)
w = 0.3

b1 = ax1.bar(x - w, open_rates, w, color=[c+"cc" for c in ["#10B981","#F59E0B","#EF4444"]],
             label="Open Rate %", edgecolor="white")
b2 = ax1.bar(x,     ctors,      w, color=[c+"88" for c in ["#10B981","#F59E0B","#EF4444"]],
             label="CTOR %", edgecolor="white", hatch="//")

ax2 = ax1.twinx()
ax2.plot(x + w/2, med_lifespans, color=BRAND, marker="D", markersize=9,
         linewidth=2.5, label="Median lifespan of cancellers (days)", zorder=5)
for i, (xp, ml, cnt) in enumerate(zip(x + w/2, med_lifespans, counts)):
    ax2.annotate(f"{ml:.0f}d\n(n={cnt:,})",
                 xy=(xp, ml), xytext=(xp + 0.08, ml + 4),
                 fontsize=9, color=BRAND, fontweight="bold")

ax1.set_xticks(x)
ax1.set_xticklabels(phase_labels, fontsize=11)
ax1.set_ylabel("Rate (%)")
ax2.set_ylabel("Median Lifespan of Unsubscribers (days)", color=BRAND)
ax2.tick_params(axis='y', colors=BRAND)
ax1.set_title("Email Engagement vs Lifespan of Unsubscribers — by Phase\n(Are declining rates linked to longer-tenured subscribers leaving?)")

lines1, lab1 = ax1.get_legend_handles_labels()
lines2, lab2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, lab1 + lab2, fontsize=9, loc="upper right")
fig.tight_layout()
fig.savefig(OUT / "Q_phase_engagement_vs_lifespan.png")
plt.close()

print("\nAll charts saved: M, N, O, P, Q")

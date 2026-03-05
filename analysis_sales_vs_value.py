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

phase_cut1 = pd.Timestamp("2025-06-01")
phase_cut2 = pd.Timestamp("2025-11-01")

# ── Load manual categorisation CSV ──────────────────────────────────────────
bdf_raw = pd.read_csv(BASE / "Emails Broadcasting - broadcasts_categorised.csv")
bdf_raw["date"] = pd.to_datetime(bdf_raw["date"])
bdf_raw["category"] = bdf_raw["category"].replace("Sles", "Sales")

# >900 recipients filter (applied to all analyses)
bdf_all = bdf_raw[bdf_raw["recipients"] > 900].copy()
bdf_all.sort_values("date", inplace=True)
bdf_all.reset_index(drop=True, inplace=True)

# Links=Yes filter — applied only to CTOR / click-rate analyses
# Emails with no links cannot generate clicks, so including them
# in CTOR calculations artificially depresses the metric.
bdf_link = bdf_all[bdf_all["Links"] == "Yes"].copy()

# Only Value / Sales for charting (drop "Others" if present)
bdf_vs   = bdf_all[bdf_all["category"].isin(["Value", "Sales"])].copy()
bdf_vs_l = bdf_link[bdf_link["category"].isin(["Value", "Sales"])].copy()

# Phase labels
def assign_phase(d):
    if d < phase_cut1:  return "Strong Start\nFeb–May 2025"
    if d < phase_cut2:  return "Silent Warning\nJun–Oct 2025"
    return "Visible Decline\nNov 2025–now"

phase_order = [
    "Strong Start\nFeb–May 2025",
    "Silent Warning\nJun–Oct 2025",
    "Visible Decline\nNov 2025–now",
]
phase_colors = {phase_order[0]: GREEN, phase_order[1]: ORANGE, phase_order[2]: RED}

for df in [bdf_all, bdf_vs, bdf_link, bdf_vs_l]:
    df["phase"] = df["date"].apply(assign_phase)

# ── Print summary ────────────────────────────────────────────────────────────
print("=" * 70)
print("CATEGORY COUNTS  (>900 recipients)")
print("=" * 70)
print(bdf_all["category"].value_counts().to_string())
print(f"\nLinks=Yes subset: {len(bdf_link)} emails")
print(bdf_link["category"].value_counts().to_string())

print("\n" + "=" * 70)
print("OPEN RATE by CATEGORY  (all >900 emails)")
print("=" * 70)
for cat in ["Value", "Sales"]:
    g = bdf_vs[bdf_vs["category"] == cat]
    print(f"\n{cat} (n={len(g)}):  OR = {g['open_rate'].mean():.1f}%")

print("\n" + "=" * 70)
print("CTOR by CATEGORY  (Links=Yes emails only)")
print("=" * 70)
for cat in ["Value", "Sales"]:
    g = bdf_vs_l[bdf_vs_l["category"] == cat]
    print(f"\n{cat} (n={len(g)}):  CTOR = {g['click_to_open_rate'].mean():.1f}%")

print("\n" + "=" * 70)
print("BY PHASE × CATEGORY")
print("=" * 70)
for ph in phase_order:
    print(f"\n{ph.replace(chr(10), ' | ')}:")
    for cat in ["Value", "Sales"]:
        or_g = bdf_vs[(bdf_vs["phase"] == ph) & (bdf_vs["category"] == cat)]
        ct_g = bdf_vs_l[(bdf_vs_l["phase"] == ph) & (bdf_vs_l["category"] == cat)]
        if len(or_g) == 0:
            print(f"  {cat}: no data")
            continue
        print(f"  {cat}  OR n={len(or_g)}, {or_g['open_rate'].mean():.1f}%"
              f"  |  CTOR n={len(ct_g)}, {ct_g['click_to_open_rate'].mean():.1f}%")

# ── Pre-aggregate for charts ─────────────────────────────────────────────────
# Open rate — from bdf_vs (all >900, Value/Sales only)
phase_cat_or = (
    bdf_vs.groupby(["phase", "category"])
          .agg(avg_open=("open_rate", "mean"), n_or=("open_rate", "size"))
          .reset_index()
)
# CTOR — from bdf_vs_l (Links=Yes, Value/Sales only)
phase_cat_ct = (
    bdf_vs_l.groupby(["phase", "category"])
             .agg(avg_ctor=("click_to_open_rate", "mean"), n_ct=("click_to_open_rate", "size"))
             .reset_index()
)
phase_cat = phase_cat_or.merge(phase_cat_ct, on=["phase", "category"], how="left")

# ─────────────────────────────────────────────────────────────────────────────
# CHART R — Monthly avg open rate AND CTOR: Sales vs Value lines
#   Open rate: all >900 sends
#   CTOR: Links=Yes sends only
# ─────────────────────────────────────────────────────────────────────────────
bdf_vs["month"]   = bdf_vs["date"].dt.to_period("M")
bdf_vs_l["month"] = bdf_vs_l["date"].dt.to_period("M")

monthly_or = (
    bdf_vs.groupby(["month", "category"])
          .agg(avg_open=("open_rate", "mean"))
          .reset_index()
)
monthly_or["month_dt"] = monthly_or["month"].dt.to_timestamp()

monthly_ct = (
    bdf_vs_l.groupby(["month", "category"])
             .agg(avg_ctor=("click_to_open_rate", "mean"))
             .reset_index()
)
monthly_ct["month_dt"] = monthly_ct["month"].dt.to_timestamp()

fig, axes = plt.subplots(2, 1, figsize=(15, 10), sharex=True)

for cat, color, ls in [("Value", BRAND, "-"), ("Sales", ACCENT, "--")]:
    sub_or = monthly_or[monthly_or["category"] == cat].sort_values("month_dt")
    sub_ct = monthly_ct[monthly_ct["category"] == cat].sort_values("month_dt")
    axes[0].plot(sub_or["month_dt"], sub_or["avg_open"], marker="o", markersize=5,
                 color=color, linestyle=ls, linewidth=2.2, label=f"{cat} emails")
    axes[1].plot(sub_ct["month_dt"], sub_ct["avg_ctor"], marker="o", markersize=5,
                 color=color, linestyle=ls, linewidth=2.2, label=f"{cat} emails")

for ax in axes:
    ax.axvline(phase_cut1, color=ORANGE, linewidth=1.5, linestyle=":", alpha=0.8)
    ax.axvline(phase_cut2, color=RED,    linewidth=1.5, linestyle=":", alpha=0.8)
    ylim = ax.get_ylim()
    ax.text(phase_cut1 + pd.Timedelta(days=3), ylim[1] * 0.97,
            "Silent Warning\nstarts", color=ORANGE, fontsize=8)
    ax.text(phase_cut2 + pd.Timedelta(days=3), ylim[1] * 0.97,
            "Visible Decline\nstarts", color=RED, fontsize=8)
    ax.legend(fontsize=9)

axes[0].set_ylabel("Avg Open Rate (%)")
axes[0].set_title("Open Rate Over Time — Sales vs Value Emails\n(monthly averages, >900 recipients)")
axes[1].set_ylabel("Avg CTOR (%)")
axes[1].set_title("Click-to-Open Rate Over Time — Sales vs Value Emails\n(emails with links only)")
axes[1].xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
axes[1].xaxis.set_major_locator(mdates.MonthLocator())
plt.setp(axes[1].xaxis.get_majorticklabels(), rotation=45, ha="right")
fig.tight_layout()
fig.savefig(OUT / "R_sales_vs_value_monthly_trend.png")
plt.close()

# ─────────────────────────────────────────────────────────────────────────────
# CHART S — Grouped bar: open rate + CTOR per phase × category
#   Open rate bars: all sends | CTOR bars: Links=Yes only
# ─────────────────────────────────────────────────────────────────────────────
x = np.arange(len(phase_order))
w = 0.2
short_labels = ["Strong Start\nFeb–May", "Silent Warning\nJun–Oct", "Visible Decline\nNov–now"]
fig, axes = plt.subplots(1, 2, figsize=(15, 6), sharey=False)

for ax, metric, n_col, ylabel, title, subtitle in [
    (axes[0], "avg_open", "n_or", "Avg Open Rate (%)",
     "Open Rate by Phase × Category", "(all emails >900 recipients)"),
    (axes[1], "avg_ctor", "n_ct",  "Avg CTOR (%)",
     "CTOR by Phase × Category",      "(emails with links only)"),
]:
    for cat, color, offset in [("Value", BRAND, -w), ("Sales", ACCENT, w)]:
        vals, ns = [], []
        for ph in phase_order:
            row = phase_cat[(phase_cat["phase"] == ph) & (phase_cat["category"] == cat)]
            vals.append(row[metric].values[0] if len(row) and not pd.isna(row[metric].values[0]) else 0)
            ns.append(int(row[n_col].values[0]) if len(row) and not pd.isna(row[n_col].values[0]) else 0)
        bars = ax.bar(x + offset, vals, w * 1.8, color=color, alpha=0.85,
                      label=f"{cat} emails", edgecolor="white")
        for bar, v, n in zip(bars, vals, ns):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                    f"{v:.1f}%\nn={n}", ha="center", fontsize=8)

    ax.set_xticks(x)
    ax.set_xticklabels(short_labels, fontsize=9)
    ax.set_ylabel(ylabel)
    ax.set_title(f"{title}\n{subtitle}", fontsize=12)
    ax.legend(fontsize=9)

fig.suptitle("Sales vs Value Emails — Engagement by Phase", fontsize=14, fontweight="bold", y=1.01)
fig.tight_layout()
fig.savefig(OUT / "S_phase_category_bars.png")
plt.close()

# ─────────────────────────────────────────────────────────────────────────────
# CHART T — Scatter: every send coloured by category, shaped by phase
#   Uses Links=Yes only so CTOR is meaningful on all plotted points
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(13, 6))

markers = {phase_order[0]: "o", phase_order[1]: "s", phase_order[2]: "^"}
phase_short = {
    phase_order[0]: "Strong Start",
    phase_order[1]: "Silent Warning",
    phase_order[2]: "Visible Decline",
}
for cat, color in [("Value", BRAND), ("Sales", ACCENT)]:
    for ph, marker in markers.items():
        sub = bdf_vs_l[(bdf_vs_l["category"] == cat) & (bdf_vs_l["phase"] == ph)]
        ax.scatter(sub["open_rate"], sub["click_to_open_rate"],
                   color=color, marker=marker, alpha=0.6, s=45,
                   label=f"{cat} / {phase_short[ph]}")

ax.set_xlabel("Open Rate (%)")
ax.set_ylabel("Click-to-Open Rate (%)")
ax.set_title("Open Rate vs CTOR — Every Send (emails with links only)\n"
             "(colour = category, shape = phase)")
handles, labels = ax.get_legend_handles_labels()
ax.legend(handles, labels, fontsize=8, ncol=2, loc="upper right")
fig.tight_layout()
fig.savefig(OUT / "T_scatter_category_phase.png")
plt.close()

# ─────────────────────────────────────────────────────────────────────────────
# CHART U — Drop magnitude: Open Rate (all sends) & CTOR (links only)
#           Strong Start → Visible Decline, split by category
# ─────────────────────────────────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(13, 6))

for ax, (metric, n_col, mlabel, subtitle) in zip(axes, [
    ("avg_open", "n_or", "Open Rate (%)",       "all emails >900 recipients"),
    ("avg_ctor", "n_ct",  "CTOR (%)",            "emails with links only"),
]):
    categories = ["Value", "Sales"]
    p1_vals, p3_vals = [], []
    for cat in categories:
        r1 = phase_cat[(phase_cat["phase"] == phase_order[0]) & (phase_cat["category"] == cat)][metric]
        r3 = phase_cat[(phase_cat["phase"] == phase_order[2]) & (phase_cat["category"] == cat)][metric]
        p1_vals.append(r1.values[0] if len(r1) else 0)
        p3_vals.append(r3.values[0] if len(r3) else 0)

    xi = np.arange(len(categories))
    ww = 0.35
    ax.bar(xi - ww / 2, p1_vals, ww, color=[BRAND, ACCENT], alpha=0.45,
           label="Strong Start (baseline)", edgecolor="white")
    ax.bar(xi + ww / 2, p3_vals, ww, color=[BRAND, ACCENT], alpha=0.95,
           label="Visible Decline", edgecolor="white")

    for i, (v1, v3) in enumerate(zip(p1_vals, p3_vals)):
        d = v3 - v1
        pct = d / v1 * 100 if v1 else 0
        color = RED if d < 0 else GREEN
        ax.annotate("", xy=(i + ww / 2, v3), xytext=(i - ww / 2, v1),
                    arrowprops=dict(arrowstyle="->", color=color, lw=2))
        ax.text(i, max(v1, v3) + 0.5,
                f"{d:+.1f}pp\n({pct:+.0f}%)", ha="center",
                fontsize=9, fontweight="bold", color=color)

    ax.set_xticks(xi)
    ax.set_xticklabels(categories, fontsize=11)
    ax.set_ylabel(mlabel)
    ax.set_title(f"{mlabel} — Strong Start → Visible Decline\n({subtitle})", fontsize=11)
    ax.legend(fontsize=8)

fig.suptitle("Which Emails Drove the Decline?\nStrong Start vs Visible Decline — Sales vs Value",
             fontsize=14, fontweight="bold", y=1.02)
fig.tight_layout()
fig.savefig(OUT / "U_decline_by_category.png")
plt.close()

# ─────────────────────────────────────────────────────────────────────────────
# CHART V — Rolling Open Rate (all sends) & CTOR (Links=Yes) per category
# ─────────────────────────────────────────────────────────────────────────────
fig, axes = plt.subplots(2, 1, figsize=(15, 9), sharex=True)

for cat, color, ax in [("Value", BRAND, axes[0]), ("Sales", ACCENT, axes[1])]:
    or_sub = bdf_vs[bdf_vs["category"] == cat].sort_values("date").copy()
    ct_sub = bdf_vs_l[bdf_vs_l["category"] == cat].sort_values("date").copy()

    or_sub["roll_open"] = or_sub["open_rate"].rolling(6, min_periods=2).mean()
    ct_sub["roll_ctor"] = ct_sub["click_to_open_rate"].rolling(6, min_periods=2).mean()

    ax.scatter(or_sub["date"], or_sub["open_rate"], color=color, alpha=0.2, s=12)
    ax.scatter(ct_sub["date"], ct_sub["click_to_open_rate"], color=RED, alpha=0.2, s=12)
    ax.plot(or_sub["date"], or_sub["roll_open"], color=color, linewidth=2.2,
            label="Open Rate — all sends (6-send rolling avg)")
    ax.plot(ct_sub["date"], ct_sub["roll_ctor"], color=RED, linewidth=2.2,
            linestyle="--", label="CTOR — links-only sends (6-send rolling avg)")

    ax.axvline(phase_cut1, color=ORANGE, linewidth=1.2, linestyle=":")
    ax.axvline(phase_cut2, color=RED,    linewidth=1.2, linestyle=":")
    ax.set_ylabel("Rate (%)")
    ax.set_title(f"{cat} Emails — Open Rate & CTOR Rolling Trend")
    ax.legend(fontsize=9)

axes[1].xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
axes[1].xaxis.set_major_locator(mdates.MonthLocator())
plt.setp(axes[1].xaxis.get_majorticklabels(), rotation=45, ha="right")
fig.tight_layout()
fig.savefig(OUT / "V_rolling_by_category.png")
plt.close()

print("\nAll charts saved: R, S, T, U, V")
print("Note: Open Rate uses all >900 sends; CTOR uses Links=Yes sends only.")

from __future__ import annotations

from datetime import datetime
from html import escape

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import config
from kit_client import KitClient
from kpi_service import KPIService


st.set_page_config(
    page_title="Marketing Dashboard — Live Kit KPIs",
    page_icon="📈",
    layout="wide",
)

PURPLE = "#7C3AED"
PURPLE_DARK = "#5B21B6"
PURPLE_LIGHT = "#A78BFA"
BG_GRADIENT = "linear-gradient(135deg, #0E1014 0%, #141820 45%, #1A1F28 100%)"
APP_BUILD = "2026-03-19-ui-tooltips-no-kpi-labels"


def _inject_style() -> None:
    st.markdown(
        f"""
<style>
.stApp {{
  background: {BG_GRADIENT};
  color: #F8F7FF;
}}
.block-container {{
  padding-top: 1.1rem;
  padding-bottom: 2rem;
}}
[data-testid="stMetric"] {{
  background: rgba(255,255,255,0.06);
  border: 1px solid rgba(255,255,255,0.12);
  border-radius: 16px;
  padding: 14px 16px;
  backdrop-filter: blur(6px);
}}
.card {{
  background: rgba(255,255,255,0.06);
  border: 1px solid rgba(255,255,255,0.12);
  border-radius: 18px;
  padding: 18px;
}}
.card-title {{
  color: #E5DBFF;
  font-size: 0.92rem;
  margin-bottom: 0.35rem;
  letter-spacing: .02em;
}}
.hero {{
  border-radius: 0;
  padding: 0;
  margin-bottom: 10px;
  background: transparent;
  border: none;
}}
.hero h1 {{
  font-size: 1.55rem;
  margin: 0;
  line-height: 1.2;
}}
.hero p {{
  margin: 6px 0 0;
  color: #EBDFFF;
  font-size: .95rem;
}}
.metric-card {{
  background: rgba(255,255,255,0.06);
  border: 1px solid rgba(255,255,255,0.12);
  border-radius: 16px;
  padding: 14px 16px;
  min-height: 112px;
}}
.metric-label {{
  display: flex;
  align-items: center;
  gap: 8px;
  color: #E5DBFF;
  font-size: 0.92rem;
  letter-spacing: .02em;
}}
.metric-value {{
  margin-top: 8px;
  font-size: 1.8rem;
  font-weight: 700;
  color: #F8F7FF;
}}
.section-head {{
  display: flex;
  align-items: center;
  gap: 8px;
  margin: 6px 0 6px;
}}
.section-head .title {{
  font-size: 1.35rem;
  font-weight: 700;
}}
.info-wrap {{
  position: relative;
  display: inline-flex;
  align-items: center;
}}
.info-icon {{
  width: 18px;
  height: 18px;
  border-radius: 50%;
  border: 1px solid rgba(220,210,255,0.65);
  color: #E7DAFF;
  font-size: 12px;
  line-height: 16px;
  text-align: center;
  cursor: help;
  user-select: none;
}}
.tooltip-box {{
  visibility: hidden;
  opacity: 0;
  transition: opacity .15s ease;
  position: absolute;
  z-index: 9999;
  left: 22px;
  top: -6px;
  width: 360px;
  background: rgba(12,15,22,0.97);
  border: 1px solid rgba(167,139,250,0.6);
  border-radius: 10px;
  padding: 10px 12px;
  color: #F3EEFF;
  font-size: 0.82rem;
  line-height: 1.35;
  box-shadow: 0 10px 32px rgba(0,0,0,0.35);
}}
.info-wrap:hover .tooltip-box {{
  visibility: visible;
  opacity: 1;
}}
hr {{ border-color: rgba(255,255,255,0.14); }}
</style>
""",
        unsafe_allow_html=True,
    )


@st.cache_data(ttl=120, show_spinner=False)
def load_dashboard_data(cache_bust: str = "") -> dict:
    _ = cache_bust
    client = KitClient(api_key=config.KIT_API_KEY, api_base=config.KIT_API_BASE)
    service = KPIService(client=client)
    return service.compute_all()


def _fmt_int(n: int) -> str:
    return f"{n:,}"


def _fmt_pct(v: float) -> str:
    return f"{v:.2f}%"


def _chart_template(fig: go.Figure) -> go.Figure:
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,0.03)",
        legend=dict(font=dict(size=12)),
        margin=dict(l=8, r=8, t=36, b=8),
    )
    return fig


def _tooltip_html(text: str) -> str:
    safe = escape(text).replace("\n", "<br>")
    return (
        '<span class="info-wrap"><span class="info-icon">i</span>'
        f'<span class="tooltip-box">{safe}</span></span>'
    )


def _section_header(title: str, tooltip: str) -> None:
    st.markdown(
        f"""
<div class="section-head">
  <div class="title">{escape(title)}</div>
  {_tooltip_html(tooltip)}
</div>
        """,
        unsafe_allow_html=True,
    )


def _metric_card(title: str, value: str, tooltip: str) -> None:
    st.markdown(
        f"""
<div class="metric-card">
  <div class="metric-label">{escape(title)} {_tooltip_html(tooltip)}</div>
  <div class="metric-value">{escape(value)}</div>
</div>
        """,
        unsafe_allow_html=True,
    )


def _canonical_workshop_name(name: str) -> str:
    n = str(name or "").lower()
    if "ai app sprint" in n:
        return "AI App Sprint (Aug 2025)"
    if "freelance accelerator" in n:
        return "Freelance Accelerator (Oct 2025)"
    if "agent breakthrough" in n:
        return "Agent Breakthrough (Nov 2025)"
    return "Other Workshop Tags"


def main() -> None:
    _inject_style()

    st.markdown(
        """
<div class="hero">
  <h1>Marketing Dashboard</h1>
</div>
        """,
        unsafe_allow_html=True,
    )

    with st.spinner("Loading live data from Kit API..."):
        data = load_dashboard_data(cache_bust=APP_BUILD)

    generated_at = data.get("generated_at_utc")
    generated_at_txt = generated_at
    try:
        generated_at_txt = datetime.fromisoformat(generated_at).strftime("%Y-%m-%d %H:%M:%S UTC")
    except Exception:
        pass

    kpi6 = data["kpi6_rewarm"]
    trend_df = data.get("kpi_confirmed_trend_6m", pd.DataFrame()).copy()
    source_df = data.get("kpi_confirmed_source_breakdown", pd.DataFrame()).copy()

    c1, c2, c3 = st.columns(3)
    with c1:
        _metric_card(
            "Current confirmed subscribers",
            _fmt_int(data["kpi14_current_confirmed"]),
            "Meaning: total people currently active on your email list.\n"
            "This is your real audience size right now.\n"
            "Formula: count(subscribers where state = active).",
        )
    with c2:
        _metric_card(
            "Re-warmed rate (%)",
            _fmt_pct(float(kpi6["movement_rate"])),
            f"Meaning: percent of cold subscribers who came back and opened at least one email in the last {config.REWARM_WINDOW_DAYS} days.\n"
            "Shows if re-engagement efforts are working.\n"
            "Formula: (# subscribers cold before last 30 days and opened in last 30 days) / (# current cold subscribers).",
        )
    with c3:
        _metric_card(
            "Re-warmed count",
            _fmt_int(int(kpi6["rewarmed_count"])),
            f"Meaning: number of people who were cold, then opened again in the last {config.REWARM_WINDOW_DAYS} days.\n"
            "Gives the actual volume of recovered subscribers.\n"
            "Formula: count(cold_before_last_30d ∩ opened_last_30d).",
        )
    st.caption(f"Last refresh: {generated_at_txt}")

    _section_header(
        "Confirmed subscribers trend (last 6 months)",
        "Meaning: monthly trend of confirmed subscribers over the last 6 months.\n"
        "Shows whether the confirmed audience is growing or slowing.\n"
        "Formula: cumulative count of active subscribers by month-end; plus monthly new confirmed subscribers.\n"
        "Source breakdown below uses the same last-6-month confirmed cohort.",
    )
    if len(trend_df):
        fig_confirmed = go.Figure()
        fig_confirmed.add_trace(
            go.Scatter(
                x=trend_df["month"],
                y=trend_df["cumulative_confirmed"],
                mode="lines+markers",
                name="Confirmed subscribers (cumulative)",
                line=dict(color=PURPLE, width=3),
                marker=dict(size=8),
                hovertemplate="Month: %{x}<br>Confirmed subscribers: %{y:,}<extra></extra>",
            )
        )
        _chart_template(fig_confirmed)
        fig_confirmed.update_layout(
            height=340,
            yaxis_title="Confirmed subscribers",
            xaxis_title="Month",
        )
        st.plotly_chart(fig_confirmed, use_container_width=True)
    else:
        st.info("No confirmed subscriber trend data available yet.")

    _section_header(
        "Source breakdown",
        "Meaning: confirmed subscribers split by entry form source over the same last 6 months window.\n"
        "Shows which source is bringing the most confirmed subscribers.\n"
        "Formula: for each confirmed subscriber in the last 6 months, assign one primary source based on the earliest source-tag date (if multiple source tags exist, first tag wins), then count by source; share = source count / total confirmed in last 6 months.",
    )
    if len(source_df):
        fig_sources = px.bar(
            source_df,
            x="source",
            y="confirmed_subscribers",
            text="confirmed_subscribers",
            color="confirmed_subscribers",
            color_continuous_scale=[[0, PURPLE_LIGHT], [1, PURPLE_DARK]],
            labels={"source": "Source form", "confirmed_subscribers": "Confirmed subscribers"},
        )
        fig_sources.update_traces(texttemplate="%{text:,}", textposition="outside")
        fig_sources.update_coloraxes(showscale=False)
        _chart_template(fig_sources)
        fig_sources.update_layout(height=330)
        st.plotly_chart(fig_sources, use_container_width=True)

        show_sources = source_df.rename(
            columns={
                "source": "Source form",
                "confirmed_subscribers": "Confirmed subscribers",
                "share_pct": "Share of confirmed (%)",
            }
        )
        st.dataframe(show_sources, use_container_width=True, hide_index=True)
    else:
        st.info("No source breakdown data available yet.")

    st.markdown("---")

    # KPI 3
    _section_header(
        "Sales CTOR by segment",
        "Meaning: click-to-open rate for sales emails, split by audience segment.\n"
        "Shows which segments take action after opening.\n"
        f"Formula: sales clicks / sales opens (for each segment), using rolling {config.ROLLING_DAYS_SALES_CTOR} days.",
    )
    sales_df = data["kpi3_sales_ctor_by_segment"].copy()
    if len(sales_df):
        fig_sales = px.bar(
            sales_df,
            x="segment",
            y="sales_ctor",
            text="sales_ctor",
            color="sales_ctor",
            color_continuous_scale=[[0, PURPLE_LIGHT], [1, PURPLE]],
            labels={"sales_ctor": "Sales CTOR (%)", "segment": "Segment"},
        )
        fig_sales.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig_sales.update_coloraxes(showscale=False)
        _chart_template(fig_sales)
        fig_sales.update_layout(height=360)
        st.plotly_chart(fig_sales, use_container_width=True)

        show_sales = sales_df.rename(
            columns={
                "segment": "Segment",
                "segment_size": "Segment size",
                "openers": "Sales openers",
                "clickers": "Sales clickers",
                "sales_ctor": "Sales CTOR (%)",
            }
        )
        st.dataframe(show_sales, use_container_width=True, hide_index=True)
    else:
        st.info("No sales broadcasts found in the configured rolling window.")

    st.markdown("---")

    # KPI 6
    _section_header(
        "Re-warmed rate (%)",
        f"Meaning: how many cold subscribers became active openers again in the last {config.REWARM_WINDOW_DAYS} days.\n"
        "Tracks recovery of disengaged subscribers.\n"
        "Formula: (# subscribers cold before last 30 days and opened in last 30 days) / (# current cold subscribers).",
    )
    k6a, k6b = st.columns([2, 3])

    with k6a:
        gauge = go.Figure(
            go.Indicator(
                mode="gauge+number",
                value=float(kpi6["movement_rate"]),
                number={"suffix": "%", "font": {"size": 34}},
                gauge={
                    "axis": {"range": [0, 100]},
                    "bar": {"color": PURPLE},
                    "bgcolor": "rgba(255,255,255,0.08)",
                    "steps": [
                        {"range": [0, 20], "color": "rgba(167,139,250,0.22)"},
                        {"range": [20, 60], "color": "rgba(124,58,237,0.30)"},
                        {"range": [60, 100], "color": "rgba(91,33,182,0.38)"},
                    ],
                },
                title={"text": "Re-warmed rate (%)"},
            )
        )
        _chart_template(gauge)
        gauge.update_layout(height=320)
        st.plotly_chart(gauge, use_container_width=True)

    with k6b:
        st.markdown(
            f"""
<div class="card">
  <div><b>Cold before last {config.REWARM_WINDOW_DAYS} days:</b> {_fmt_int(int(kpi6['cold_before_last30_count']))}</div>
  <div><b>Current cold subscribers (denominator):</b> {_fmt_int(int(kpi6['current_cold_count']))}</div>
  <div><b>Re-warmed subscribers:</b> {_fmt_int(int(kpi6['rewarmed_count']))}</div>
  <div><b>Re-warmed rate:</b> {_fmt_pct(float(kpi6['movement_rate']))}</div>
</div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # KPI 12
    _section_header(
        "Workshop to bootcamp conversion by segment",
        "Meaning: percent of workshop buyers who also bought the related bootcamp.\n"
        "Tells how strong workshops are as a path to bootcamp sales.\n"
        "Formula: (# workshop buyers with mapped bootcamp tag) / (# workshop buyers), shown by segment and by workshop program.",
    )
    conv_df = data["kpi12_workshop_to_bootcamp"].copy()
    prog_df = data["kpi12_workshop_program"].copy()
    if len(prog_df):
        prog_df["program_3"] = prog_df["workshop_program"].apply(_canonical_workshop_name)
        prog_df = (
            prog_df.groupby("program_3", as_index=False)
            .agg(
                workshop_buyers=("workshop_buyers", "sum"),
                also_bootcamp=("also_bootcamp", "sum"),
            )
        )
        prog_df["conversion_rate"] = (
            (prog_df["also_bootcamp"] / prog_df["workshop_buyers"]) * 100.0
        ).fillna(0.0)
        prog_df["workshop_program"] = prog_df["program_3"]
        prog_df = prog_df[
            prog_df["workshop_program"].isin(
                [
                    "AI App Sprint (Aug 2025)",
                    "Freelance Accelerator (Oct 2025)",
                    "Agent Breakthrough (Nov 2025)",
                ]
            )
        ].copy()
        order = [
            "AI App Sprint (Aug 2025)",
            "Freelance Accelerator (Oct 2025)",
            "Agent Breakthrough (Nov 2025)",
        ]
        prog_df["workshop_program"] = pd.Categorical(prog_df["workshop_program"], categories=order, ordered=True)
        prog_df = prog_df.sort_values("workshop_program")

    k12a, k12b = st.columns(2)
    with k12a:
        if len(conv_df):
            fig_conv = px.bar(
                conv_df,
                x="segment",
                y="conversion_rate",
                text="conversion_rate",
                color="conversion_rate",
                color_continuous_scale=[[0, PURPLE_LIGHT], [1, PURPLE_DARK]],
                labels={"conversion_rate": "Conversion rate (%)", "segment": "Segment"},
            )
            fig_conv.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
            fig_conv.update_coloraxes(showscale=False)
            _chart_template(fig_conv)
            fig_conv.update_layout(height=360)
            st.plotly_chart(fig_conv, use_container_width=True)

    with k12b:
        if len(prog_df):
            fig_prog = px.bar(
                prog_df,
                x="workshop_program",
                y="conversion_rate",
                text="conversion_rate",
                color="conversion_rate",
                color_continuous_scale=[[0, "#C4B5FD"], [1, "#6D28D9"]],
                labels={"conversion_rate": "Conversion rate (%)", "workshop_program": "Workshop program"},
            )
            fig_prog.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
            fig_prog.update_coloraxes(showscale=False)
            _chart_template(fig_prog)
            fig_prog.update_layout(height=360)
            st.plotly_chart(fig_prog, use_container_width=True)

    show_conv = conv_df.rename(
        columns={
            "segment": "Segment",
            "workshop_buyers": "Workshop buyers",
            "also_bootcamp": "Also bootcamp buyers",
            "conversion_rate": "Conversion rate (%)",
        }
    )
    with st.expander("Show table", expanded=False):
        st.dataframe(show_conv, use_container_width=True, hide_index=True)

    st.markdown("---")

    # KPI 19
    _section_header(
        "Last 4 months email performance snapshot",
        "Meaning: monthly view of Value Open Rate, Sales Open Rate, and Sales CTOR across the last 4 completed months.\n"
        "Gives a quick trend check of recent email health.\n"
        "Formulas: Value OR = value opens/value delivered; Sales OR = sales opens/sales delivered; Sales CTOR = sales clicks/sales opens.",
    )
    snapshot_df = data["kpi19_snapshot"].copy()
    if len(snapshot_df):
        fig_snap = go.Figure()
        fig_snap.add_trace(
            go.Bar(
                x=snapshot_df["month"],
                y=snapshot_df["value_open_rate"],
                name="Value Open Rate",
                marker_color="#C4B5FD",
                hovertemplate="Month: %{x}<br>Value OR: %{y:.2f}%<extra></extra>",
            )
        )
        fig_snap.add_trace(
            go.Bar(
                x=snapshot_df["month"],
                y=snapshot_df["sales_open_rate"],
                name="Sales Open Rate",
                marker_color="#8B5CF6",
                hovertemplate="Month: %{x}<br>Sales OR: %{y:.2f}%<extra></extra>",
            )
        )
        fig_snap.add_trace(
            go.Bar(
                x=snapshot_df["month"],
                y=snapshot_df["sales_ctor"],
                name="Sales CTOR",
                marker_color="#F0ABFC",
                hovertemplate="Month: %{x}<br>Sales CTOR: %{y:.2f}%<extra></extra>",
            )
        )
        fig_snap.update_layout(
            barmode="stack",
            yaxis_title="Rate (%)",
            xaxis_title="Month",
            legend_title="Metric",
        )
        _chart_template(fig_snap)
        fig_snap.update_layout(height=360)
        st.plotly_chart(fig_snap, use_container_width=True)

        show_snapshot = snapshot_df.rename(
            columns={
                "month": "Month",
                "value_open_rate": "Value OR (%)",
                "sales_open_rate": "Sales OR (%)",
                "sales_ctor": "Sales CTOR (%)",
                "value_sends": "Value sends",
                "sales_sends": "Sales sends",
            }
        )
        st.dataframe(show_snapshot, use_container_width=True, hide_index=True)
    else:
        st.info("Not enough labeled broadcasts to compute monthly snapshot yet.")

    st.markdown("---")

    # KPI 16
    _section_header(
        "Monthly churn rate",
        "Meaning: unsubscribe rate each month.\n"
        "Helps monitor audience loss and deliverability pressure.\n"
        "Formula: monthly unsubscribes / monthly recipients.",
    )
    churn_df = data["kpi16_churn"].copy()
    if len(churn_df):
        fig_churn = go.Figure()
        fig_churn.add_trace(
            go.Scatter(
                x=churn_df["month"],
                y=churn_df["churn_rate"],
                mode="lines+markers",
                name="Churn rate",
                line=dict(color=PURPLE, width=3),
            )
        )
        fig_churn.add_trace(
            go.Bar(
                x=churn_df["month"],
                y=churn_df["unsubs"],
                name="Unsubscribes",
                marker_color="rgba(196,181,253,0.45)",
                yaxis="y2",
                hovertemplate="Month: %{x}<br>Unsubscribers: %{y:,}<extra></extra>",
            )
        )
        fig_churn.update_layout(
            yaxis=dict(title="Churn rate (%)"),
            yaxis2=dict(title="Unsubscribes", overlaying="y", side="right", showgrid=False),
            barmode="overlay",
            height=360,
        )
        _chart_template(fig_churn)
        st.plotly_chart(fig_churn, use_container_width=True)

if __name__ == "__main__":
    main()

from html import escape
from pathlib import Path
import sys

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from analysis import (
    compute_actor_metrics,
    compute_district_breakdown,
    compute_event_breakdown,
    compute_expansion_analysis,
    compute_high_impact_incidents,
    compute_intelligence_score,
    compute_summary_metrics,
    compute_tactical_shift,
    generate_executive_summary,
    prepare_map_data,
)
from data import load_data
from maps import build_map_image, build_map_view
from processing import get_time_windows
from report import build_docx_report, build_pdf_report


st.set_page_config(page_title="CNAWS Weekly Intelligence Dashboard", layout="wide")


def _inject_styles():
    st.markdown(
        """
        <style>
        .block-container {
            max-width: 1320px;
            padding-top: 24px;
            padding-right: 24px;
            padding-bottom: 40px;
            padding-left: 24px;
        }
        .dashboard-title {
            font-size: 1.95rem;
            font-weight: 700;
            color: #102a43;
            line-height: 1.1;
            margin: 0;
        }
        .dashboard-window {
            margin-top: 12px;
            padding-top: 12px;
            border-top: 1px solid #d9e2ec;
            color: #486581;
            font-size: 0.96rem;
            font-weight: 500;
        }
        .section-label {
            color: #486581;
            font-size: 0.8rem;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            margin: 0 0 10px 0;
        }
        .panel-card {
            border: 1px solid #d9e2ec;
            border-radius: 18px;
            background: #ffffff;
            padding: 18px;
            box-shadow: 0 12px 24px rgba(15, 23, 42, 0.04);
        }
        .summary-card {
            border: 1px solid #d9e2ec;
            border-radius: 18px;
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
            padding: 20px;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.04);
        }
        .summary-title {
            color: #102a43;
            font-size: 0.92rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            margin-bottom: 12px;
        }
        .summary-line {
            color: #243b53;
            font-size: 1rem;
            line-height: 1.65;
            margin: 0 0 6px 0;
        }
        .metric-card {
            height: 116px;
            border: 1px solid #d9e2ec;
            border-radius: 16px;
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
            padding: 16px 16px 14px 16px;
            box-shadow: 0 8px 18px rgba(15, 23, 42, 0.03);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }
        .metric-label {
            font-size: 0.82rem;
            color: #486581;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }
        .metric-value {
            font-size: 1.9rem;
            line-height: 1;
            font-weight: 700;
            color: #102a43;
        }
        .metric-delta {
            font-size: 0.92rem;
            font-weight: 700;
        }
        .metric-meta {
            font-size: 0.78rem;
            color: #7b8794;
        }
        .delta-positive {
            color: #0f766e;
        }
        .delta-negative {
            color: #b91c1c;
        }
        .delta-neutral {
            color: #61758a;
        }
        .score-card {
            border: 1px solid #d9e2ec;
            border-radius: 22px;
            background: linear-gradient(135deg, #102a43 0%, #243b53 100%);
            padding: 24px;
            color: #ffffff;
            box-shadow: 0 16px 32px rgba(15, 23, 42, 0.14);
        }
        .score-heading {
            font-size: 0.84rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            opacity: 0.88;
            font-weight: 700;
        }
        .score-value {
            font-size: 3.2rem;
            font-weight: 800;
            line-height: 1;
            margin: 12px 0 6px 0;
        }
        .score-risk {
            display: inline-block;
            margin-top: 10px;
            padding: 6px 12px;
            border-radius: 999px;
            font-size: 0.82rem;
            font-weight: 700;
            background: rgba(255, 255, 255, 0.12);
            border: 1px solid rgba(255, 255, 255, 0.18);
        }
        .score-components {
            margin-top: 16px;
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 10px;
        }
        .score-component {
            border-radius: 12px;
            background: rgba(255, 255, 255, 0.08);
            padding: 10px 12px;
        }
        .score-component-label {
            font-size: 0.72rem;
            text-transform: uppercase;
            opacity: 0.8;
            letter-spacing: 0.06em;
            font-weight: 700;
        }
        .score-component-value {
            font-size: 1.08rem;
            font-weight: 700;
            margin-top: 4px;
        }
        .drivers-title {
            margin-top: 16px;
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            font-weight: 700;
            opacity: 0.84;
        }
        .driver-line {
            margin-top: 4px;
            font-size: 0.95rem;
            line-height: 1.5;
        }
        .table-wrap {
            border: 1px solid #d9e2ec;
            border-radius: 16px;
            overflow: hidden;
            background: #ffffff;
        }
        .analytic-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
        }
        .analytic-table thead th {
            background: #f8fafc;
            color: #334e68;
            text-align: left;
            font-weight: 700;
            padding: 11px 12px;
            border-bottom: 1px solid #d9e2ec;
            white-space: nowrap;
        }
        .analytic-table tbody td {
            padding: 10px 12px;
            border-top: 1px solid #eef2f6;
            color: #102a43;
            vertical-align: top;
        }
        .analytic-table tbody tr.row-top td {
            font-weight: 700;
        }
        .analytic-table tbody tr.row-new td {
            background: #fff7ed;
        }
        .analytic-table tbody tr:hover td {
            background: #f8fafc;
        }
        .table-caption {
            font-size: 0.78rem;
            font-weight: 700;
            color: #61758a;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            margin-bottom: 10px;
        }
        .pill {
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 700;
            border: 1px solid transparent;
        }
        .pill-red {
            background: #fee2e2;
            color: #991b1b;
            border-color: #fecaca;
        }
        .pill-green {
            background: #dcfce7;
            color: #166534;
            border-color: #bbf7d0;
        }
        .summary-grid {
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
            margin-bottom: 14px;
        }
        .summary-chip {
            min-width: 180px;
            padding: 12px 14px;
            border-radius: 14px;
            background: #f8fafc;
            border: 1px solid #d9e2ec;
        }
        .summary-chip-label {
            color: #61758a;
            font-size: 0.74rem;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            font-weight: 700;
        }
        .summary-chip-value {
            margin-top: 4px;
            color: #102a43;
            font-size: 1.4rem;
            font-weight: 700;
        }
        .shift-text {
            color: #102a43;
            font-size: 0.98rem;
            line-height: 1.65;
        }
        .section-spacer {
            height: 32px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _format_number(value, decimals=0):
    if value is None:
        return "-"
    if decimals == 0:
        return f"{int(round(value)):,}"
    return f"{value:,.{decimals}f}"


def _format_pct(value):
    if value is None:
        return "New"
    return f"{value:+.1%}"


def _delta_class(value):
    if value > 0:
        return "delta-positive"
    if value < 0:
        return "delta-negative"
    return "delta-neutral"


def _window_span(window):
    end_inclusive = window["end"] - pd.Timedelta(days=1)
    return f"{window['start']:%d %b} - {end_inclusive:%d %b}"


def _section_title(text):
    st.markdown(f"<div class='section-label'>{escape(text)}</div>", unsafe_allow_html=True)


def _metric_card(label, value, delta_text, previous_text, delta_value=0):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{escape(label)}</div>
            <div class="metric-value">{escape(value)}</div>
            <div class="metric-delta {_delta_class(delta_value)}">{escape(delta_text)}</div>
            <div class="metric-meta">{escape(previous_text)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _score_card(score_payload):
    component_caps = {"Activity": 30, "Lethality": 25, "Expansion": 25, "Tactical": 20}
    component_html = "".join(
        f"""
        <div class="score-component">
            <div class="score-component-label">{escape(name)}</div>
            <div class="score-component-value">{_format_number(value)} / {component_caps.get(name, 25)}</div>
        </div>
        """
        for name, value in score_payload["components"].items()
    )
    drivers_html = "".join(
        f"<div class='driver-line'>- {escape(driver)}</div>" for driver in score_payload["drivers"]
    )

    st.markdown(
        f"""
        <div class="score-card">
            <div class="score-heading">Intelligence Score</div>
            <div class="score-value">{_format_number(score_payload['final_score'])} / 100</div>
            <div>Weighted weekly risk reading across activity, lethality, geographic spread, and tactical change.</div>
            <div class="score-risk" style="color:{score_payload['risk_color']}; background: rgba(255,255,255,0.92);">
                Risk Level: {escape(score_payload['risk_level'])}
            </div>
            <div class="score-components">{component_html}</div>
            <div class="drivers-title">Drivers</div>
            {drivers_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def _build_comparison_chart(dataframe, category_column, title):
    if dataframe.empty:
        return None

    chart_df = dataframe.copy()
    categories = chart_df[category_column].astype(str).tolist()
    height = max(330, len(chart_df) * 30)

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=categories,
            y=chart_df["previous"],
            mode="lines+markers",
            name="Previous",
            line=dict(color="#16a34a", width=3, shape="spline"),
            marker=dict(size=7, color="#16a34a"),
            hovertemplate="<b>%{x}</b><br>Previous: %{y}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=categories,
            y=chart_df["current"],
            mode="lines+markers",
            name="Current",
            line=dict(color="#1f3b57", width=3, shape="spline"),
            marker=dict(size=8, color="#1f3b57"),
            hovertemplate="<b>%{x}</b><br>Current: %{y}<extra></extra>",
        )
    )

    if category_column == "event_type" and "CT Ops" in categories:
        ct_row = chart_df[chart_df[category_column].astype(str).eq("CT Ops")].iloc[0]
        fig.add_trace(
            go.Scatter(
                x=["CT Ops"],
                y=[ct_row["current"]],
                mode="markers",
                name="Security ops",
                marker=dict(size=11, color="#dc2626", line=dict(width=1.5, color="#ffffff")),
                hovertemplate="<b>CT Ops</b><br>Security ops: %{y}<extra></extra>",
            )
        )

    fig.update_layout(
        title=dict(text=title, font=dict(size=14, color="#102a43"), x=0, xanchor="left"),
        height=height,
        margin=dict(l=8, r=8, t=46, b=70),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Segoe UI, sans-serif", size=12, color="#334155"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=11, color="#64748b"),
            bgcolor="rgba(0,0,0,0)",
            borderwidth=0,
        ),
        hovermode="x unified",
    )
    fig.update_xaxes(
        title=None,
        tickfont=dict(size=10, color="#64748b"),
        showgrid=False,
        zeroline=False,
        showline=False,
        ticks="",
        tickangle=-35,
    )
    fig.update_yaxes(
        title=dict(text="Incidents", font=dict(size=11, color="#94a3b8")),
        tickfont=dict(size=10, color="#64748b"),
        gridcolor="#e5e7eb",
        gridwidth=1,
        zeroline=False,
        showline=False,
        ticks="",
    )
    return fig


def _render_html_table(columns, rows, table_height=None):
    header_html = "".join(f"<th>{escape(header)}</th>" for _, header in columns)
    body_rows = []

    if not rows:
        col_span = len(columns)
        body_rows.append(f"<tr><td colspan='{col_span}'>No data available.</td></tr>")
    else:
        for row in rows:
            row_class = row.get("_row_class", "")
            cells = "".join(f"<td>{row.get(key, '')}</td>" for key, _ in columns)
            body_rows.append(f"<tr class='{row_class}'>{cells}</tr>")

    height_style = f"max-height:{table_height}px; overflow:auto;" if table_height else ""
    st.markdown(
        f"""
        <div class="table-wrap" style="{height_style}">
            <table class="analytic-table">
                <thead><tr>{header_html}</tr></thead>
                <tbody>{''.join(body_rows)}</tbody>
            </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _delta_html(value, is_percent=False):
    display = _format_pct(value) if is_percent else f"{value:+,}"
    return f"<span class='{_delta_class(value)}'>{escape(display)}</span>"


def _district_rows(district_breakdown):
    rows = []
    for _, row in district_breakdown.iterrows():
        row_classes = []
        if row["rank"] <= 3:
            row_classes.append("row-top")
        if row["is_new"]:
            row_classes.append("row-new")

        rows.append(
            {
                "_row_class": " ".join(row_classes),
                "district": escape(str(row["district"])),
                "current": _format_number(row["current"]),
                "previous": _format_number(row["previous"]),
                "delta": _delta_html(int(row["delta"])),
                "pct_change": _delta_html(row["pct_change"], is_percent=True)
                if row["pct_change"] is not None
                else "<span class='delta-positive'>New</span>",
                "share": f"{row['share'] * 100:.1f}%",
                "lethality": f"{row['lethality']:.2f}",
            }
        )
    return rows


def _event_rows(event_breakdown):
    rows = []
    for _, row in event_breakdown.iterrows():
        row_classes = ["row-top"] if row["rank"] <= 3 else []
        rows.append(
            {
                "_row_class": " ".join(row_classes),
                "event_type": escape(str(row["event_type"])),
                "current": _format_number(row["current"]),
                "previous": _format_number(row["previous"]),
                "delta": _delta_html(int(row["delta"])),
                "pct_change": _delta_html(row["pct_change"], is_percent=True)
                if row["pct_change"] is not None
                else "<span class='delta-positive'>New</span>",
                "share": f"{row['share'] * 100:.1f}%",
            }
        )
    return rows


def _simple_rows(dataframe, columns):
    rows = []
    for _, row in dataframe.iterrows():
        built = {}
        for key, formatter in columns.items():
            built[key] = formatter(row[key])
        rows.append(built)
    return rows


def _render_executive_summary(summary_text):
    lines = [line for line in summary_text.splitlines() if line.strip()]
    title = lines[0] if lines else "EXECUTIVE ASSESSMENT"
    body_lines = lines[1:]
    body_html = "".join(f"<div class='summary-line'>{escape(line)}</div>" for line in body_lines)

    st.markdown(
        f"""
        <div class="summary-card">
            <div class="summary-title">{escape(title)}</div>
            {body_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def _render_key_takeaways(summary_metrics, expansion_analysis, tactical_shift, event_breakdown, district_breakdown):
    inc_current = summary_metrics["incidents"]["current"]
    inc_prev = summary_metrics["incidents"]["previous"]
    ct_ops_change = tactical_shift.get("ct_ops_change", 0)
    new_districts_count = expansion_analysis["new_districts"]

    # Calculate % changes
    inc_pct = ((inc_current - inc_prev) / max(inc_prev, 1)) * 100
    ct_pct = ct_ops_change  # Already calculated in tactical_shift

    bullets = []

    # 1) EXPANSION BULLET
    if new_districts_count > 0:
        bullets.append(f"Geographic expansion into {new_districts_count} new district{'s' if new_districts_count > 1 else ''} signals significant widening of operational footprint")
    else:
        bullets.append("No geographic expansion observed, indicating a stable operational footprint")

    # 2) TACTICAL BULLET
    tactical_bullet = None
    if not event_breakdown.empty:
        tactical_candidates = event_breakdown.copy()
        tactical_candidates["pct_change"] = pd.to_numeric(
            tactical_candidates["pct_change"], errors="coerce"
        )

        # pct_change is stored as a ratio, so 0.20 means +20%.
        significant_events = tactical_candidates[tactical_candidates["pct_change"] > 0.20]
        if not significant_events.empty:
            top_event = significant_events.sort_values(by="pct_change", ascending=False).iloc[0]
            event_type = str(top_event["event_type"]).strip()
            pct_val = top_event["pct_change"]

            interpretation_map = {
                "Targeted Violence": "selective engagements",
                "Explosions": "stand-off attacks",
                "Clashes": "direct engagements",
            }
            interpretation = interpretation_map.get(event_type, "changing operational patterns")
            tactical_bullet = f"{event_type} increased sharply ({pct_val:+.1%}), indicating {interpretation}"

    if tactical_bullet is None:
        tactical_bullet = "Operational patterns remained broadly stable without a dominant tactical shift"

    bullets.append(tactical_bullet)

    # 3) STATE RESPONSE BULLET
    if ct_pct < 0:
        bullets.append(f"Security force operations declined ({ct_pct:.1f}%), reducing operational pressure")
    elif ct_pct > 0:
        bullets.append(f"Security force operations increased (+{ct_pct:.1f}%), indicating heightened state response")
    else:
        bullets.append("Security force operations remained unchanged")

    # 4) ACTIVITY BULLET
    if inc_pct < 0:
        bullets.append(f"Overall activity declined marginally ({inc_pct:.1f}%), suggesting a slowdown despite evolving dynamics")
    else:
        bullets.append(f"Overall activity increased ({inc_pct:.1f}%), indicating rising operational tempo")

    # Render
    with st.container(border=True):
        st.markdown("<div class='section-label'>KEY TAKEAWAYS</div>", unsafe_allow_html=True)

        bullets_html = "\n\n".join(f"- {bullet}" for bullet in bullets)
        st.markdown(bullets_html)


def _render_state_response_analysis(summary_metrics, actor_metrics):
    militant_incidents_current = actor_metrics["militant_incidents"]["current"]
    militant_incidents_prev = actor_metrics["militant_incidents"]["previous"]
    security_ops_current = actor_metrics["security_force_operations"]["current"]
    security_ops_prev = actor_metrics["security_force_operations"]["previous"]
    ct_ops_current = actor_metrics["ct_operations"]["current"]
    ct_ops_prev = actor_metrics["ct_operations"]["previous"]

    # STEP 1: Calculate metrics
    ct_ratio_current = security_ops_current / max(militant_incidents_current, 1)
    ct_ratio_prev = security_ops_prev / max(militant_incidents_prev, 1)
    security_ops_pct = ((security_ops_current - security_ops_prev) / max(security_ops_prev, 1)) * 100
    ct_pct = ((ct_ops_current - ct_ops_prev) / max(ct_ops_prev, 1)) * 100
    militant_pct = ((militant_incidents_current - militant_incidents_prev) / max(militant_incidents_prev, 1)) * 100
    ratio_pct = ((ct_ratio_current - ct_ratio_prev) / max(ct_ratio_prev, 0.01)) * 100

    # STEP 2: Determine dominance
    if ct_ratio_current < 0.3:
        dominance = "pronounced militant operational dominance"
    elif ct_ratio_current <= 0.6:
        dominance = "imbalanced engagement environment"
    else:
        dominance = "strong state operational pressure"

    # STEP 3: Determine trend
    if ct_pct < -50:
        pressure_signal = "sharp reduction in state pressure"
    elif security_ops_pct < -20:
        pressure_signal = "contracting state operational activity"
    elif security_ops_pct > 20:
        pressure_signal = "increased state operational activity"
    else:
        pressure_signal = "limited change in state operational activity"

    # Build interpretation paragraph
    interpretation_parts = [
        f"Security operations changed by {security_ops_pct:+.1f}% while militant incidents changed by {militant_pct:+.1f}%.",
        f"The CT ratio stands at {ct_ratio_current:.2f}, indicating {dominance}.",
        f"Targeted CT operations changed by {ct_pct:+.1f}%, signaling {pressure_signal}.",
    ]

    if ct_ratio_current < 0.3:
        interpretation_parts.append("Operational control is tilted toward militant initiative.")
    elif ct_ratio_current <= 0.6:
        interpretation_parts.append("Operational control remains contested but weighted against the state response.")
    else:
        interpretation_parts.append("Operational control shows material state pressure on the environment.")

    interpretation_text = " ".join(interpretation_parts)

    # Render
    with st.container(border=True):
        st.markdown("<div class='section-label'>STATE RESPONSE ANALYSIS</div>", unsafe_allow_html=True)

        # PART 1: TABLE
        import pandas as pd
        table_data = pd.DataFrame({
            "Metric": ["Militant Incidents", "Security Operations", "CT Ratio"],
            "Current": [
                _format_number(militant_incidents_current),
                _format_number(security_ops_current),
                f"{ct_ratio_current:.2f}",
            ],
            "Previous": [
                _format_number(militant_incidents_prev),
                _format_number(security_ops_prev),
                f"{ct_ratio_prev:.2f}",
            ],
            "Change": [
                f"{militant_pct:+.1f}%",
                f"{security_ops_pct:+.1f}%",
                f"{ratio_pct:+.1f}%",
            ]
        })
        st.dataframe(table_data, hide_index=True, width="stretch")

        # PART 2: INTERPRETATION
        st.markdown(f"<div class='shift-text'>{interpretation_text}</div>", unsafe_allow_html=True)


def _render_map_legend():
    st.caption("Event Types")
    event_cols = st.columns(4)
    event_items = [
        ("Clashes", "Red"),
        ("Explosions", "Orange"),
        ("CT Ops", "Blue"),
        ("Targeted Violence", "Purple"),
    ]
    for column, (label, marker) in zip(event_cols, event_items):
        with column:
            st.markdown(f"**{marker}** {label}")

    st.caption("Actor Groups")
    actor_cols = st.columns(3)
    with actor_cols[0]:
        st.markdown("Circle Militant")
    with actor_cols[1]:
        st.markdown("Square Security Forces")
    with actor_cols[2]:
        st.markdown("Triangle Other")


def _build_report_context(
    window_header,
    summary_metrics,
    expansion_analysis,
    intelligence_score,
    tactical_shift,
    district_breakdown,
    event_breakdown,
    high_impact,
    actor_metrics,
    executive_summary,
    map_image,
):
    metric_cards = [
        ("Incidents", summary_metrics["incidents"]),
        ("Casualties", summary_metrics["casualties"]),
        ("Intensity", summary_metrics["intensity"]),
        ("Districts", summary_metrics["districts"]),
        (
            "Expansion Index",
            {
                "current": expansion_analysis["expansion_index"],
                "previous": 0,
                "delta": expansion_analysis["expansion_index"],
            },
        ),
    ]

    return {
        "window_header": window_header,
        "metric_cards": metric_cards,
        "expansion_analysis": expansion_analysis,
        "intelligence_score": intelligence_score,
        "tactical_shift": tactical_shift,
        "district_breakdown": district_breakdown,
        "event_breakdown": event_breakdown,
        "high_impact_incidents": high_impact,
        "actor_metrics": actor_metrics,
        "executive_summary": executive_summary,
        "map_image": map_image,
    }


_inject_styles()

try:
    df = load_data()
except Exception as exc:
    st.error(f"Unable to load the source data: {exc}")
    st.stop()

df_current, df_prev, windows = get_time_windows(df)

summary_metrics = compute_summary_metrics(df_current, df_prev)
district_breakdown = compute_district_breakdown(df_current, df_prev)
event_breakdown = compute_event_breakdown(df_current, df_prev)
expansion_analysis = compute_expansion_analysis(df_current, df_prev)
actor_metrics = compute_actor_metrics(df_current, df_prev)
tactical_shift = compute_tactical_shift(df_current, df_prev, event_breakdown, actor_metrics)
intelligence_score = compute_intelligence_score(summary_metrics, expansion_analysis, tactical_shift, actor_metrics)
map_data = prepare_map_data(df_current)
high_impact_incidents = compute_high_impact_incidents(df_current)
map_image = build_map_image(map_data)

window_header = f"Rolling Window: {_window_span(windows['current'])} vs {_window_span(windows['previous'])}"

executive_summary = generate_executive_summary(
    {
        "inc_current": actor_metrics["militant_incidents"]["current"],
        "inc_prev": actor_metrics["militant_incidents"]["previous"],
        "cas_current": summary_metrics["casualties"]["current"],
        "cas_prev": summary_metrics["casualties"]["previous"],
        "intensity_current": summary_metrics["intensity"]["current"],
        "intensity_prev": summary_metrics["intensity"]["previous"],
        "final_score": intelligence_score["final_score"],
    },
    expansion_analysis,
    tactical_shift,
)

report_context = _build_report_context(
    window_header,
    summary_metrics,
    expansion_analysis,
    intelligence_score,
    tactical_shift,
    district_breakdown,
    event_breakdown,
    high_impact_incidents,
    actor_metrics,
    executive_summary,
    map_image,
)

pdf_bytes = build_pdf_report(report_context)
docx_bytes = build_docx_report(report_context)

# ============================================================
# SECTION 1: HEADER REDESIGN
# ============================================================
with st.container(border=False):
    header_col1, header_col2, header_col3 = st.columns([1, 3, 1], vertical_alignment="center")

    with header_col1:
        st.markdown(
            '<img src="https://cnaws.in/assets/logoblack-BZJsSCK0.png" width="120">',
            unsafe_allow_html=True,
        )

    with header_col2:
        st.markdown(
            "<h2 style='text-align:center;margin:0;color:#102a43;'>CNAWS Weekly Intelligence Dashboard</h2>",
            unsafe_allow_html=True,
        )

    with header_col3:
        export_cols = st.columns(2)
        with export_cols[0]:
            st.download_button(
                "Export PDF",
                pdf_bytes,
                file_name="cnaws-weekly-intelligence.pdf",
                mime="application/pdf",
                width="stretch",
            )
        with export_cols[1]:
            st.download_button(
                "Export DOCX",
                docx_bytes,
                file_name="cnaws-weekly-intelligence.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )

# Horizontal divider
st.markdown("""
<hr style="height:1px; border:none; background-color:#1f3b57; margin-top:5px; margin-bottom:-2px;" />
""", unsafe_allow_html=True)

st.markdown(f"<div class='dashboard-window'>{escape(window_header)}</div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# SECTION 2: TOP LAYOUT (Intelligence Score + Executive Summary)
# ============================================================
top_col1, top_col2 = st.columns([1.2, 1])

with top_col1:
    _section_title("Intelligence Score")
    _score_card(intelligence_score)

with top_col2:
    _section_title("Executive Summary")
    _render_executive_summary(executive_summary)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# SECTION 3: SECOND ROW (Key Takeaways + State Response)
# ============================================================
row2_col1, row2_col2 = st.columns(2)

with row2_col1:
    _render_key_takeaways(
        summary_metrics,
        expansion_analysis,
        tactical_shift,
        event_breakdown,
        district_breakdown,
    )

with row2_col2:
    _render_state_response_analysis(
        summary_metrics,
        actor_metrics,
    )

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# SECTION 4: KPI CARDS ROW
# ============================================================
_section_title("Key Metrics")

# SVG Icons
SVG_INCIDENTS = '''<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" stroke="#1f3b57" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>'''
SVG_CASUALTIES = '''<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" stroke="#1f3b57" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21a8 8 0 0 0-16 0"/><circle cx="12" cy="7" r="4"/></svg>'''
SVG_INTENSITY = '''<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" stroke="#1f3b57" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>'''
SVG_DISTRICTS = '''<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" stroke="#1f3b57" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 6-9 12-9 12S3 16 3 10a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>'''
SVG_NEW_DISTRICTS = '''<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" stroke="#1f3b57" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polygon points="16 8 14 14 8 16 10 10 16 8"/></svg>'''
SVG_EXPANSION = '''<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" stroke="#1f3b57" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>'''


def _kpi_card(icon_svg, title, value, change_text, subtext, delta_value=0):
    # Determine color based on delta
    if delta_value < 0:
        change_color = "#dc2626"
    elif delta_value > 0:
        change_color = "#16a34a"
    else:
        change_color = "#6b7280"

    st.markdown(
        f"""
        <div style="
            padding:15px 16px;
            border-radius:14px;
            background:#ffffff;
            border:1px solid #e5e7eb;
            box-shadow:0 8px 18px rgba(15,23,42,0.035);
            transition:background 160ms ease, box-shadow 160ms ease, transform 160ms ease;
        " onmouseover="
            this.style.background='#f9fafb';
            this.style.boxShadow='0 12px 24px rgba(15,23,42,0.06)';
            this.style.transform='translateY(-1px)';
        " onmouseout="
            this.style.background='#ffffff';
            this.style.boxShadow='0 8px 18px rgba(15,23,42,0.035)';
            this.style.transform='translateY(0)';
        ">
            <div style="
                display:flex;
                align-items:center;
                gap:10px;
                margin-bottom:8px;
            ">
                <div style="
                    width:36px;
                    height:36px;
                    min-width:36px;
                    display:flex;
                    align-items:center;
                    justify-content:center;
                    background:#eef3f8;
                    border-radius:10px;
                ">
                    {icon_svg}
                </div>
                <div style="
                    font-size:12px;
                    line-height:1.2;
                    color:#6b7280;
                    font-weight:700;
                    letter-spacing:0.05em;
                    text-transform:uppercase;
                ">{escape(title)}</div>
            </div>
            <div style="
                font-size:29px;
                line-height:1.05;
                font-weight:700;
                color:#1f2937;
                margin-bottom:5px;
            ">{escape(value)}</div>
            <div style="
                font-size:12px;
                line-height:1.25;
                color:{change_color};
                font-weight:700;
                margin-bottom:3px;
            ">
                {escape(change_text)}
            </div>
            <div style="
                font-size:11px;
                line-height:1.25;
                color:#9ca3af;
            ">
                {escape(subtext)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


kpi_cols = st.columns(6)

with kpi_cols[0]:
    _kpi_card(
        SVG_INCIDENTS,
        "Incidents",
        _format_number(summary_metrics["incidents"]["current"]),
        f"{summary_metrics['incidents']['delta']:+,} vs previous",
        f"Previous: {_format_number(summary_metrics['incidents']['previous'])}",
        summary_metrics["incidents"]["delta"],
    )
with kpi_cols[1]:
    _kpi_card(
        SVG_CASUALTIES,
        "Casualties",
        _format_number(summary_metrics["casualties"]["current"]),
        f"{summary_metrics['casualties']['delta']:+,} vs previous",
        f"Previous: {_format_number(summary_metrics['casualties']['previous'])}",
        summary_metrics["casualties"]["delta"],
    )
with kpi_cols[2]:
    _kpi_card(
        SVG_INTENSITY,
        "Intensity",
        _format_number(summary_metrics["intensity"]["current"], 2),
        f"{summary_metrics['intensity']['delta']:+.2f} vs previous",
        f"Previous: {_format_number(summary_metrics['intensity']['previous'], 2)}",
        summary_metrics["intensity"]["delta"],
    )
with kpi_cols[3]:
    _kpi_card(
        SVG_DISTRICTS,
        "Districts",
        _format_number(summary_metrics["districts"]["current"]),
        f"{summary_metrics['districts']['delta']:+,} vs previous",
        f"Previous: {_format_number(summary_metrics['districts']['previous'])}",
        summary_metrics["districts"]["delta"],
    )
with kpi_cols[4]:
    _kpi_card(
        SVG_NEW_DISTRICTS,
        "New Districts",
        _format_number(expansion_analysis["new_districts"]),
        f"{expansion_analysis['new_districts']:+,} new",
        f"Previous footprint: {expansion_analysis['previous_footprint']}",
        expansion_analysis["new_districts"],
    )
with kpi_cols[5]:
    _kpi_card(
        SVG_EXPANSION,
        "Expansion Index",
        _format_number(expansion_analysis["expansion_index"], 2),
        f"{expansion_analysis['expansion_index']:.2f}",
        "Footprint expansion rate",
        0,
    )

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("<div class='section-spacer'></div>", unsafe_allow_html=True)
_section_title("Actor Separation")
actor_cols = st.columns(3)

with actor_cols[0]:
    _metric_card(
        "Militant Incidents",
        _format_number(actor_metrics["militant_incidents"]["current"]),
        f"{actor_metrics['militant_incidents']['delta']:+,} vs previous",
        f"Previous: {_format_number(actor_metrics['militant_incidents']['previous'])}",
        actor_metrics["militant_incidents"]["delta"],
    )
with actor_cols[1]:
    _metric_card(
        "Security Force Operations",
        _format_number(actor_metrics["security_force_operations"]["current"]),
        f"{actor_metrics['security_force_operations']['delta']:+,} vs previous",
        f"Previous: {_format_number(actor_metrics['security_force_operations']['previous'])}",
        actor_metrics["security_force_operations"]["delta"],
    )
with actor_cols[2]:
    _metric_card(
        "CT Ratio",
        _format_number(actor_metrics["ct_ratio"], 2),
        actor_metrics["interpretation"],
        ">= 1 means aggressive state response",
        actor_metrics["ct_ratio"] - 1,
    )

st.markdown("<div class='section-spacer'></div>", unsafe_allow_html=True)
_section_title("District Intelligence")
district_chart_col, district_table_col = st.columns([0.60, 0.40], vertical_alignment="top")

with district_chart_col:
    district_chart = _build_comparison_chart(
        district_breakdown,
        "district",
        "District Activity by Rolling Week",
    )
    if district_chart is None:
        st.info("No district activity found in the current comparison windows.")
    else:
        st.plotly_chart(district_chart, width="stretch")

with district_table_col:
    st.markdown("<div class='table-caption'>Analytical Table</div>", unsafe_allow_html=True)
    _render_html_table(
        [
            ("district", "District"),
            ("current", "Current"),
            ("previous", "Previous"),
            ("delta", "Delta"),
            ("pct_change", "% Change"),
            ("share", "Share (%)"),
            ("lethality", "Lethality"),
        ],
        _district_rows(district_breakdown),
        table_height=420,
    )

st.markdown("<div class='section-spacer'></div>", unsafe_allow_html=True)
_section_title("Expansion Panel")
expansion_pill_class = "pill-red" if expansion_analysis["new_districts"] else "pill-green"
with st.container(border=True):
    st.markdown(
        f"""
        <div class="summary-grid">
            <div class="summary-chip">
                <div class="summary-chip-label">New Districts</div>
                <div class="summary-chip-value">{expansion_analysis['new_districts']}</div>
            </div>
            <div class="summary-chip">
                <div class="summary-chip-label">Expansion Index</div>
                <div class="summary-chip-value">{expansion_analysis['expansion_index']:.2f}</div>
            </div>
            <div class="summary-chip">
                <div class="summary-chip-label">Status</div>
                <div class="summary-chip-value"><span class="pill {expansion_pill_class}">{escape(expansion_analysis['tag'])}</span></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _render_html_table(
        [("district", "District"), ("incidents", "Incidents"), ("casualties", "Casualties")],
        _simple_rows(
            expansion_analysis["table"],
            {
                "district": lambda value: escape(str(value)),
                "incidents": lambda value: _format_number(value),
                "casualties": lambda value: _format_number(value),
            },
        ),
    )

st.markdown("<div class='section-spacer'></div>", unsafe_allow_html=True)
_section_title("Geographic View")
map_col, map_table_col = st.columns([0.65, 0.35], vertical_alignment="top")
deck, legend = build_map_view(map_data)

with map_col:
    if deck is None:
        st.info("No valid coordinates are available for the current rolling window.")
    else:
        st.pydeck_chart(deck, width="stretch")
        _render_map_legend()

with map_table_col:
    st.markdown("<div class='table-caption'>Incident Table</div>", unsafe_allow_html=True)
    mapped_incidents = map_data.loc[:, ["date", "district", "event_category", "casualties_total"]].copy()
    mapped_incidents["date"] = pd.to_datetime(mapped_incidents["date"]).dt.strftime("%d %b %Y")
    _render_html_table(
        [("date", "Date"), ("district", "District"), ("event_type", "Event Type"), ("casualties", "Casualties")],
        _simple_rows(
            mapped_incidents.rename(
                columns={"event_category": "event_type", "casualties_total": "casualties"}
            ).head(12),
            {
                "date": lambda value: escape(str(value)),
                "district": lambda value: escape(str(value)),
                "event_type": lambda value: escape(str(value)),
                "casualties": lambda value: _format_number(value),
            },
        ),
        table_height=520,
    )

st.markdown("<div class='section-spacer'></div>", unsafe_allow_html=True)
_section_title("Event And Tactical Analysis")
event_chart_col, event_table_col = st.columns([0.60, 0.40], vertical_alignment="top")

with event_chart_col:
    event_chart = _build_comparison_chart(
        event_breakdown,
        "event_type",
        "Event Type Comparison by Rolling Week",
    )
    if event_chart is None:
        st.info("No event data found in the comparison windows.")
    else:
        st.plotly_chart(event_chart, width="stretch")

with event_table_col:
    st.markdown("<div class='table-caption'>Analytical Table</div>", unsafe_allow_html=True)
    _render_html_table(
        [
            ("event_type", "Event Type"),
            ("current", "Current"),
            ("previous", "Previous"),
            ("delta", "Delta"),
            ("pct_change", "% Change"),
            ("share", "Share"),
        ],
        _event_rows(event_breakdown),
        table_height=420,
    )

st.markdown("<div class='section-spacer'></div>", unsafe_allow_html=True)
_section_title("Tactical Shift Panel")
top_increase = tactical_shift["top_increase"]
top_decrease = tactical_shift["top_decrease"]
top_increase_text = (
    f"{top_increase['event_type']} ({_format_pct(top_increase['pct_change'])})"
    if top_increase is not None
    else "No material increase"
)
top_decrease_text = (
    f"{top_decrease['event_type']} ({_format_pct(top_decrease['pct_change'])})"
    if top_decrease is not None
    else "No material decrease"
)
observations_html = "".join(
    f"<div class='summary-line'>- {escape(observation)}</div>"
    for observation in tactical_shift["observations"]
)

st.markdown(
    f"""
    <div class="panel-card">
        <div class="shift-text">
            <strong>TACTICAL SHIFT</strong><br/><br/>
            <strong>Top Increase:</strong> {escape(top_increase_text)}<br/>
            <strong>Top Decrease:</strong> {escape(top_decrease_text)}<br/><br/>
            <strong>Key Observations:</strong><br/>
            {observations_html}<br/>
            <strong>Interpretation:</strong><br/>
            {escape(tactical_shift['interpretation'])}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div class='section-spacer'></div>", unsafe_allow_html=True)
_section_title("High Impact Incidents")
high_impact_display = high_impact_incidents.copy()
if not high_impact_display.empty:
    high_impact_display["date"] = pd.to_datetime(high_impact_display["date"]).dt.strftime("%d %b %Y")
    high_impact_display = high_impact_display.rename(columns={"casualties_total": "casualties"})

_render_html_table(
    [("date", "Date"), ("district", "District"), ("event_type", "Event Type"), ("casualties", "Casualties")],
    _simple_rows(
        high_impact_display,
        {
            "date": lambda value: escape(str(value)),
            "district": lambda value: escape(str(value)),
            "event_type": lambda value: escape(str(value)),
            "casualties": lambda value: _format_number(value),
        },
    ),
    table_height=460,
)

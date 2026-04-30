import re

import pandas as pd

CLASH_KEYWORDS = ("clash", "clashes", "engagement", "engagements", "gun battle")
EXPLOSION_KEYWORDS = ("explosion", "explosions", "ied", "blast", "bomb", "detonation")
CT_OPS_KEYWORDS = ("ct operation", "counter-terror", "counter terror", "security operation")
TARGETED_VIOLENCE_KEYWORDS = ("targeted", "assassination", "abduction", "kidnap", "execution")


def _safe_divide(numerator, denominator):
    if denominator in (0, None) or pd.isna(denominator):
        return 0.0 if numerator in (0, None) or pd.isna(numerator) else None
    return numerator / denominator


def _compute_pct_change(current_value, previous_value):
    if previous_value in (0, None) or pd.isna(previous_value):
        return None if current_value else 0.0
    return (current_value - previous_value) / previous_value


def _build_metric(current_value, previous_value):
    return {
        "current": current_value,
        "previous": previous_value,
        "delta": current_value - previous_value,
        "pct_change": _compute_pct_change(current_value, previous_value),
    }


def _count_by_category(df, category_column):
    if df.empty:
        return pd.Series(dtype="int64")

    return (
        df[category_column]
        .fillna("Unknown")
        .replace("", "Unknown")
        .value_counts(dropna=False)
        .sort_values(ascending=False)
    )


def _sum_by_category(df, category_column, value_column):
    if df.empty:
        return pd.Series(dtype="float64")

    return (
        df.groupby(category_column, dropna=False)[value_column]
        .sum()
        .sort_values(ascending=False)
    )


def _pct_change_percent(current_value, previous_value):
    change = _compute_pct_change(current_value, previous_value)
    if change is None:
        return None
    return round(change * 100, 1)


def categorize_event_type(event_type):
    text = str(event_type or "").strip().lower()
    if any(keyword in text for keyword in CLASH_KEYWORDS):
        return "Clashes"
    if any(keyword in text for keyword in EXPLOSION_KEYWORDS):
        return "Explosions"
    if any(keyword in text for keyword in CT_OPS_KEYWORDS) or text == "ct operation":
        return "CT Ops"
    if any(keyword in text for keyword in TARGETED_VIOLENCE_KEYWORDS):
        return "Targeted Violence"
    return str(event_type or "Unknown").strip() or "Unknown"


def _event_series(df):
    if df.empty:
        return pd.Series(dtype="string")
    return df["event_type"].apply(categorize_event_type)


def compute_summary_metrics(df_current, df_prev):
    incidents_current = len(df_current)
    incidents_prev = len(df_prev)

    casualties_current = float(df_current["casualties_total"].sum())
    casualties_prev = float(df_prev["casualties_total"].sum())

    intensity_current = _safe_divide(casualties_current, incidents_current) or 0.0
    intensity_prev = _safe_divide(casualties_prev, incidents_prev) or 0.0

    districts_current = int(df_current["district"].nunique())
    districts_prev = int(df_prev["district"].nunique())

    return {
        "incidents": _build_metric(incidents_current, incidents_prev),
        "casualties": _build_metric(int(casualties_current), int(casualties_prev)),
        "intensity": _build_metric(round(intensity_current, 2), round(intensity_prev, 2)),
        "districts": _build_metric(districts_current, districts_prev),
    }


def compute_actor_metrics(df_current, df_prev):
    current_event_types = _event_series(df_current)
    prev_event_types = _event_series(df_prev)

    current_actor_types = df_current["actor_type"].fillna("").str.lower()
    prev_actor_types = df_prev["actor_type"].fillna("").str.lower()

    militant_current = int(current_actor_types.eq("militant").sum())
    militant_prev = int(prev_actor_types.eq("militant").sum())
    security_force_current = int(current_actor_types.str.contains("security force", regex=False).sum())
    security_force_prev = int(prev_actor_types.str.contains("security force", regex=False).sum())
    ct_current = int(current_event_types.eq("CT Ops").sum())
    ct_prev = int(prev_event_types.eq("CT Ops").sum())

    ct_ratio = round(_safe_divide(security_force_current, max(militant_current, 1)) or 0.0, 2)
    ct_change_pct = _pct_change_percent(ct_current, ct_prev)

    if ct_ratio >= 1:
        interpretation = "Aggressive state response"
    else:
        interpretation = "Militant dominance"

    return {
        "militant_incidents": _build_metric(militant_current, militant_prev),
        "security_force_operations": _build_metric(security_force_current, security_force_prev),
        "ct_operations": _build_metric(ct_current, ct_prev),
        "ct_ratio": ct_ratio,
        "ct_ops_change_pct": ct_change_pct,
        "interpretation": interpretation,
    }


def compute_district_breakdown(df_current, df_prev, top_n=12):
    current_counts = _count_by_category(df_current, "district")
    prev_counts = _count_by_category(df_prev, "district")
    current_casualties = _sum_by_category(df_current, "district", "casualties_total")

    districts = current_counts.index.union(prev_counts.index)
    breakdown = pd.DataFrame(
        {
            "district": districts,
            "current": current_counts.reindex(districts, fill_value=0).values,
            "previous": prev_counts.reindex(districts, fill_value=0).values,
            "casualties": current_casualties.reindex(districts, fill_value=0).values,
        }
    )

    current_total = breakdown["current"].sum()
    breakdown["delta"] = breakdown["current"] - breakdown["previous"]
    breakdown["pct_change"] = breakdown.apply(
        lambda row: _compute_pct_change(row["current"], row["previous"]), axis=1
    )
    breakdown["share"] = breakdown["current"].apply(
        lambda value: _safe_divide(value, current_total) or 0.0
    )
    breakdown["lethality"] = breakdown.apply(
        lambda row: round(_safe_divide(row["casualties"], row["current"]) or 0.0, 2), axis=1
    )
    breakdown["is_new"] = (breakdown["current"] > 0) & (breakdown["previous"] == 0)
    breakdown["rank"] = breakdown["current"].rank(method="first", ascending=False).astype(int)

    breakdown = breakdown.sort_values(
        by=["current", "delta", "casualties", "district"], ascending=[False, False, False, True]
    ).reset_index(drop=True)

    return breakdown.head(top_n)


def compute_event_breakdown(df_current, df_prev, top_n=12):
    current_counts = _event_series(df_current).value_counts().sort_values(ascending=False)
    prev_counts = _event_series(df_prev).value_counts().sort_values(ascending=False)

    event_types = current_counts.index.union(prev_counts.index)
    breakdown = pd.DataFrame(
        {
            "event_type": event_types,
            "current": current_counts.reindex(event_types, fill_value=0).values,
            "previous": prev_counts.reindex(event_types, fill_value=0).values,
        }
    )

    current_total = breakdown["current"].sum()
    breakdown["delta"] = breakdown["current"] - breakdown["previous"]
    breakdown["pct_change"] = breakdown.apply(
        lambda row: _compute_pct_change(row["current"], row["previous"]), axis=1
    )
    breakdown["share"] = breakdown["current"].apply(
        lambda value: _safe_divide(value, current_total) or 0.0
    )
    breakdown["rank"] = breakdown["current"].rank(method="first", ascending=False).astype(int)

    breakdown = breakdown.sort_values(
        by=["current", "delta", "event_type"], ascending=[False, False, True]
    ).reset_index(drop=True)

    return breakdown.head(top_n)


def compute_expansion_analysis(df_current, df_prev):
    current_districts = set(df_current["district"].dropna())
    previous_districts = set(df_prev["district"].dropna())
    new_districts = sorted(current_districts - previous_districts)

    rows = []
    for district in new_districts:
        district_df = df_current[df_current["district"] == district]
        rows.append(
            {
                "district": district,
                "incidents": int(len(district_df)),
                "casualties": int(district_df["casualties_total"].sum()),
            }
        )

    expansion_table = pd.DataFrame(rows)
    if expansion_table.empty:
        expansion_table = pd.DataFrame(columns=["district", "incidents", "casualties"])
    else:
        expansion_table = expansion_table.sort_values(
            by=["incidents", "casualties", "district"], ascending=[False, False, True]
        ).reset_index(drop=True)

    previous_footprint = len(previous_districts)
    expansion_index = round(_safe_divide(len(new_districts), max(previous_footprint, 1)) or 0.0, 2)

    return {
        "new_districts": len(new_districts),
        "previous_footprint": previous_footprint,
        "expansion_index": expansion_index,
        "tag": "Expansion Detected" if new_districts else "Stable Footprint",
        "table": expansion_table,
    }


def _find_event_row(event_breakdown, label):
    matches = event_breakdown[event_breakdown["event_type"] == label]
    return matches.iloc[0] if not matches.empty else None


def compute_tactical_shift(df_current, df_prev, event_breakdown, actor_metrics):
    increases = event_breakdown[event_breakdown["delta"] > 0].sort_values(
        by=["delta", "current"], ascending=[False, False]
    )
    decreases = event_breakdown[event_breakdown["delta"] < 0].sort_values(
        by=["delta", "current"], ascending=[True, False]
    )

    top_increase = increases.iloc[0] if not increases.empty else None
    top_decrease = decreases.iloc[0] if not decreases.empty else None

    clashes_row = _find_event_row(event_breakdown, "Clashes")
    explosions_row = _find_event_row(event_breakdown, "Explosions")
    ct_row = _find_event_row(event_breakdown, "CT Ops")

    clashes_increase = 0.0 if clashes_row is None else (_pct_change_percent(clashes_row["current"], clashes_row["previous"]) or 0.0)
    explosions_increase = 0.0 if explosions_row is None else (_pct_change_percent(explosions_row["current"], explosions_row["previous"]) or 0.0)
    ct_ops_change = actor_metrics["ct_ops_change_pct"] or 0.0
    ct_ops_decrease = max(0.0, -ct_ops_change)
    clashes_share = 0.0 if clashes_row is None else round((clashes_row["share"] or 0.0) * 100, 1)

    observations = []
    if ct_ops_change <= -30:
        observations.append(f"CT operations declined sharply ({ct_ops_change:+.1f}%).")
    elif ct_ops_change >= 30:
        observations.append(f"CT operations increased significantly ({ct_ops_change:+.1f}%).")
    else:
        observations.append(f"CT operations were broadly stable ({ct_ops_change:+.1f}%).")

    if explosions_row is not None:
        explosion_change = explosions_row["pct_change"]
        if explosion_change is not None:
            observations.append(f"Explosions changed by {explosion_change:+.1%}.")

    if clashes_row is not None:
        observations.append(f"Clashes remain dominant ({clashes_share:.1f}% of recorded incidents).")

    if ct_ops_change <= -30 and explosions_increase < 0:
        interpretation = (
            "Reduced state pressure alongside declining explosive activity suggests a temporary "
            "operational slowdown rather than escalation."
        )
    elif clashes_increase > 20:
        interpretation = "Higher direct engagement levels indicate elevated close-contact violence."
    elif ct_ops_change >= 30:
        interpretation = "A stronger state response suggests a reactive posture to current-week activity."
    else:
        interpretation = "Operational patterns remained mixed without a single dominant shift."

    return {
        "top_increase": None
        if top_increase is None
        else {
            "event_type": top_increase["event_type"],
            "delta": int(top_increase["delta"]),
            "pct_change": top_increase["pct_change"],
        },
        "top_decrease": None
        if top_decrease is None
        else {
            "event_type": top_decrease["event_type"],
            "delta": int(top_decrease["delta"]),
            "pct_change": top_decrease["pct_change"],
        },
        "observations": observations,
        "interpretation": interpretation,
        "ct_ops_change": ct_ops_change,
        "clashes_increase": clashes_increase,
        "explosions_increase": explosions_increase,
        "ct_ops_decrease": ct_ops_decrease,
        "clashes_share": clashes_share,
        "top_event_type": event_breakdown.iloc[0]["event_type"] if not event_breakdown.empty else "Unknown",
        "declining_event_type": top_decrease["event_type"] if top_decrease is not None else "None",
    }


def compute_intelligence_score(summary_metrics, expansion_analysis, tactical_shift, actor_metrics):
    inc_current = summary_metrics["incidents"]["current"]
    inc_prev = summary_metrics["incidents"]["previous"]
    intensity_current = summary_metrics["intensity"]["current"]
    intensity_prev = summary_metrics["intensity"]["previous"]
    new_districts = expansion_analysis["new_districts"]

    activity_ratio = inc_current / max(inc_prev, 1)
    intensity_ratio = intensity_current / max(intensity_prev, 0.1)

    if activity_ratio >= 1.5:
        activity_score = 30
    elif activity_ratio >= 1.2:
        activity_score = 22
    elif activity_ratio >= 1.0:
        activity_score = 15
    else:
        activity_score = 8

    if intensity_ratio >= 1.5:
        lethality_score = 25
    elif intensity_ratio >= 1.2:
        lethality_score = 18
    elif intensity_ratio >= 1.0:
        lethality_score = 12
    else:
        lethality_score = 6

    if new_districts >= 10:
        expansion_score = 25
    elif new_districts >= 5:
        expansion_score = 18
    elif new_districts >= 2:
        expansion_score = 10
    else:
        expansion_score = 3

    if tactical_shift["clashes_increase"] > 20:
        tactical_score = 20
    elif tactical_shift["explosions_increase"] > 20:
        tactical_score = 15
    elif tactical_shift["ct_ops_decrease"] > 40:
        tactical_score = 12
    else:
        tactical_score = 6

    final_score = activity_score + lethality_score + expansion_score + tactical_score

    if final_score > 60:
        risk_level = "HIGH"
        risk_color = "#b91c1c"
    elif final_score > 30:
        risk_level = "ELEVATED"
        risk_color = "#c2410c"
    else:
        risk_level = "STABLE"
        risk_color = "#0f766e"

    drivers = []
    if lethality_score >= 18:
        drivers.append("High lethality increase")
    if expansion_score >= 10:
        drivers.append("Significant geographic expansion")
    if tactical_shift["clashes_increase"] > 20:
        drivers.append("Elevated direct engagements")
    elif tactical_shift["ct_ops_decrease"] > 40:
        drivers.append("Reduced security force pressure")
    if activity_score >= 22 and len(drivers) < 3:
        drivers.append("Higher overall incident tempo")
    if not drivers:
        drivers.append("Indicators remain within a comparatively stable range")

    return {
        "final_score": final_score,
        "risk_level": risk_level,
        "risk_color": risk_color,
        "drivers": drivers,
        "components": {
            "Activity": activity_score,
            "Lethality": lethality_score,
            "Expansion": expansion_score,
            "Tactical": tactical_score,
        },
    }


def generate_executive_summary(metrics, expansion, tactical):
    inc_current = metrics["inc_current"]
    inc_prev = metrics["inc_prev"]
    intensity_current = metrics["intensity_current"]
    intensity_prev = metrics["intensity_prev"]
    final_score = metrics["final_score"]

    incident_change = _pct_change_percent(inc_current, inc_prev)
    abs_incident_change = abs(incident_change or 0.0)
    if inc_current > inc_prev:
        activity = (
            f"Militant activity increased by {abs_incident_change:.1f}% "
            f"({inc_current} vs {inc_prev})."
        )
    else:
        activity = (
            f"Militant activity declined by {abs_incident_change:.1f}% "
            f"({inc_current} vs {inc_prev})."
        )

    if intensity_current > intensity_prev:
        lethality = (
            "Lethality increased, with casualties per incident rising from "
            f"{intensity_prev:.2f} to {intensity_current:.2f}."
        )
    else:
        lethality = (
            "Lethality declined slightly, with casualties per incident falling from "
            f"{intensity_prev:.2f} to {intensity_current:.2f}."
        )

    if expansion["expansion_index"] > 0:
        geo = (
            "Geographic expansion was observed, with activity spreading to "
            f"{expansion['new_districts']} new districts."
        )
    else:
        geo = "No geographic expansion was observed, indicating a stable operational footprint."

    ct_ops_change = tactical["ct_ops_change"]
    tactical_change_values = [
        abs(tactical.get("clashes_increase") or 0.0),
        abs(tactical.get("explosions_increase") or 0.0),
    ]
    for key in ("top_increase", "top_decrease"):
        event_shift = tactical.get(key)
        if not event_shift:
            continue
        if event_shift.get("pct_change") is None:
            tactical_change_values.append(100.0)
        else:
            tactical_change_values.append(abs(event_shift["pct_change"] * 100))

    notable_tactical_shift = any(value > 30 for value in tactical_change_values)
    if ct_ops_change < -30:
        tactical_line = (
            "Security force operations declined sharply, indicating reduced operational pressure."
        )
    elif ct_ops_change > 30:
        tactical_line = (
            "Security force operations increased significantly, suggesting a reactive posture."
        )
    elif notable_tactical_shift:
        tactical_line = "A notable tactical shift was observed in the current reporting window."
    else:
        tactical_line = "No single tactical category materially altered the operating picture."

    if incident_change is None:
        activity_trend = "With overall activity newly recorded"
    elif incident_change < 0:
        activity_trend = "Despite a marginal decline in overall activity"
    elif incident_change > 0:
        activity_trend = "Alongside an increase in overall activity"
    else:
        activity_trend = "With overall activity unchanged"

    if ct_ops_change < -30:
        pressure_clause = "a sharp reduction in security operations"
    elif ct_ops_change > 30:
        pressure_clause = "a stronger security-force response"
    else:
        pressure_clause = "limited change in security operations"

    if expansion["new_districts"] > 0:
        expansion_clause = "expansion into new districts"
        synthesis_effect = "increasing operational space for militant actors"
    else:
        expansion_clause = "no new geographic expansion"
        synthesis_effect = "a constrained but still contested operating environment"

    synthesis = (
        f"{activity_trend}, {pressure_clause} and {expansion_clause} indicate "
        f"{synthesis_effect}."
    )

    if final_score > 60:
        outlook = "Overall risk remains high."
    elif final_score > 30:
        outlook = "Overall risk remains elevated."
    else:
        outlook = "Overall risk remains stable."

    # Forward-looking intelligence: Outlook section (bulleted, bold heading)
    outlook_bullets = [f"- {outlook}"]
    if ct_ops_change < -30 and expansion["new_districts"] > 0:
        outlook_bullets.append(
            "- If current trends persist, reduced state pressure and continued expansion may enable further militant consolidation."
        )

    summary_lines = [
        "EXECUTIVE ASSESSMENT",
        "",
        activity,
        lethality,
        geo,
        tactical_line,
        synthesis,
        "",
        "𝐎𝐮𝐭𝐥𝐨𝐨𝐤",
        *outlook_bullets,
    ]

    return "\n".join(summary_lines)


def compute_high_impact_incidents(df_current, top_n=15):
    if df_current.empty:
        return pd.DataFrame(columns=["date", "district", "event_type", "casualties_total"])

    incidents = df_current.copy()
    incidents["casualties_total"] = incidents["casualties_total"].fillna(0)
    incidents["event_type"] = incidents["event_type"].apply(categorize_event_type)

    return (
        incidents.sort_values(
            by=["casualties_total", "date", "district"], ascending=[False, False, True]
        )
        .loc[:, ["date", "district", "event_type", "casualties_total"]]
        .head(top_n)
        .reset_index(drop=True)
    )


def prepare_map_data(df_current):
    if df_current.empty:
        return pd.DataFrame(
            columns=[
                "date",
                "district",
                "event_type",
                "event_category",
                "actor_type",
                "actor_group",
                "casualties_total",
                "latitude",
                "longitude",
                "marker_size",
            ]
        )

    map_df = df_current.dropna(subset=["latitude", "longitude"]).copy()
    map_df = map_df[map_df["latitude"].between(-90, 90) & map_df["longitude"].between(-180, 180)]
    map_df["event_category"] = map_df["event_type"].apply(categorize_event_type)
    actor_text = map_df["actor_type"].fillna("").str.lower()
    map_df["actor_group"] = "Other"
    map_df.loc[actor_text.eq("militant"), "actor_group"] = "Militant"
    map_df.loc[
        actor_text.str.contains("security force", regex=False) | map_df["event_category"].eq("CT Ops"),
        "actor_group",
    ] = "Security Forces"
    map_df["actor_group"] = map_df["actor_group"].where(
        map_df["actor_group"].ne("Other"), map_df["actor_type"].fillna("Other")
    )
    map_df["marker_size"] = map_df["casualties_total"].fillna(0).clip(lower=0) * 4200 + 14000

    return map_df[
        [
            "date",
            "district",
            "event_type",
            "event_category",
            "actor_type",
            "actor_group",
            "casualties_total",
            "latitude",
            "longitude",
            "marker_size",
        ]
    ].sort_values(by=["casualties_total", "date"], ascending=[False, False])

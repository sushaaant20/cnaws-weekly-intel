from io import BytesIO

import matplotlib.pyplot as plt
import pandas as pd
import pydeck as pdk

EVENT_COLOR_MAP = {
    "Clashes": [220, 38, 38, 200],
    "Explosions": [249, 115, 22, 200],
    "CT Ops": [37, 99, 235, 200],
    "Targeted Violence": [147, 51, 234, 200],
}

EVENT_COLOR_HEX = {
    "Clashes": "#dc2626",
    "Explosions": "#f97316",
    "CT Ops": "#2563eb",
    "Targeted Violence": "#9333ea",
}

ACTOR_OUTLINE_MAP = {
    "Militant": [15, 23, 42, 220],
    "Security Forces": [8, 47, 73, 220],
}

ACTOR_MARKER_MAP = {"Militant": "o", "Security Forces": "s"}


def _color_for_event(event_category):
    return EVENT_COLOR_MAP.get(event_category, [100, 116, 139, 180])


def _hex_for_event(event_category):
    return EVENT_COLOR_HEX.get(event_category, "#64748b")


def build_map_view(map_data):
    if map_data.empty:
        return None, {"events": [], "actors": []}

    deck_df = map_data.copy()
    deck_df["date_label"] = pd.to_datetime(deck_df["date"]).dt.strftime("%d %b %Y")
    deck_df["fill_color"] = deck_df["event_category"].map(_color_for_event)
    deck_df["line_color"] = deck_df["actor_group"].map(
        lambda value: ACTOR_OUTLINE_MAP.get(value, [51, 65, 85, 220])
    )

    view_state = pdk.ViewState(
        latitude=float(deck_df["latitude"].mean()),
        longitude=float(deck_df["longitude"].mean()),
        zoom=5.4,
        pitch=0,
    )

    layers = []
    for actor_group in ("Militant", "Security Forces"):
        actor_df = deck_df[deck_df["actor_group"] == actor_group]
        if actor_df.empty:
            continue

        layers.append(
            pdk.Layer(
                "ScatterplotLayer",
                data=actor_df,
                id=f"{actor_group.lower().replace(' ', '-')}-layer",
                get_position=["longitude", "latitude"],
                get_fill_color="fill_color",
                get_line_color="line_color",
                get_radius="marker_size",
                pickable=True,
                stroked=True,
                filled=True,
                radius_min_pixels=5,
                line_width_min_pixels=2,
                opacity=0.82,
            )
        )

    tooltip = {
        "html": (
            "<b>{district}</b><br/>"
            "Event: {event_category}<br/>"
            "Actor: {actor_group}<br/>"
            "Casualties: {casualties_total}"
        )
    }

    deck = pdk.Deck(
        map_style="light",
        initial_view_state=view_state,
        layers=layers,
        tooltip=tooltip,
    )

    event_legend = [
        {"label": label, "hex": _hex_for_event(label)}
        for label in ("Clashes", "Explosions", "CT Ops", "Targeted Violence")
        if label in deck_df["event_category"].unique()
    ]
    actor_legend = [
        {"label": label, "marker": "Circle" if label == "Militant" else "Square"}
        for label in deck_df["actor_group"].unique()
    ]
    return deck, {"events": event_legend, "actors": actor_legend}


def build_map_image(map_data):
    if map_data.empty:
        return None

    fig, ax = plt.subplots(figsize=(9, 5.2), dpi=160)
    fig.patch.set_facecolor("#ffffff")
    ax.set_facecolor("#f8fafc")

    for actor_group, marker in ACTOR_MARKER_MAP.items():
        actor_df = map_data[map_data["actor_group"] == actor_group]
        if actor_df.empty:
            continue

        for event_category, group_df in actor_df.groupby("event_category"):
            ax.scatter(
                group_df["longitude"],
                group_df["latitude"],
                s=(group_df["casualties_total"].clip(lower=0) * 55) + 70,
                c=_hex_for_event(event_category),
                marker=marker,
                edgecolors="#0f172a" if actor_group == "Militant" else "#164e63",
                linewidths=0.8,
                alpha=0.85,
                label=f"{event_category} | {actor_group}",
            )

    ax.set_title("Current Window Incident Map", fontsize=14, color="#102a43", loc="left")
    ax.set_xlabel("Longitude")
    ax.set_ylabel("Latitude")
    ax.grid(alpha=0.16)

    handles, labels = ax.get_legend_handles_labels()
    if handles:
        unique_labels = {}
        for handle, label in zip(handles, labels):
            unique_labels.setdefault(label, handle)
        ax.legend(
            unique_labels.values(),
            unique_labels.keys(),
            fontsize=7.5,
            loc="upper left",
            bbox_to_anchor=(1.02, 1.0),
            frameon=False,
        )

    plt.tight_layout()
    buffer = BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    return buffer.getvalue()

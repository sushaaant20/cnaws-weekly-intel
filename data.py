import re
from dotenv import load_dotenv
import os

import pandas as pd
import streamlit as st

load_dotenv

SHEET_CSV_URL = os.getenv("SHEET_CSV_URL")

COLUMN_ALIASES = {
    "date": ["Date", "Incident Date", "Event Date"],
    "event_type": ["Event Type", "Incident Type", "Type", "Event"],
    "district": ["District", "Area", "Location District"],
    "actor_type": ["Actor Type", "Actor", "Perpetrator Type"],
    "casualties_total": ["Casualties (Total)", "Total Casualties", "Casualties"],
    "latitude": ["Latitude", "Lat"],
    "longitude": ["Longitude", "Long", "Lng", "Lon"],
    "coordinates": ["Coordinates", "Coordinate", "Lat Long", "Lat/Long"],
}

TEXT_COLUMNS = ("event_type", "district", "actor_type")
NUMERIC_COLUMNS = ("casualties_total", "latitude", "longitude")


def _normalize_column_name(name):
    return re.sub(r"[^a-z0-9]+", " ", str(name).strip().lower()).strip()


def _build_rename_map(columns):
    normalized_lookup = {_normalize_column_name(column): column for column in columns}
    rename_map = {}

    for canonical_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            actual_name = normalized_lookup.get(_normalize_column_name(alias))
            if actual_name:
                rename_map[actual_name] = canonical_name
                break

    return rename_map


def _parse_coordinates(df):
    if "coordinates" not in df.columns:
        return df

    extracted = df["coordinates"].astype("string").str.extract(
        r"(?P<latitude>-?\d+(?:\.\d+)?)\s*[,/ ]\s*(?P<longitude>-?\d+(?:\.\d+)?)"
    )

    for column in ("latitude", "longitude"):
        parsed_values = pd.to_numeric(extracted[column], errors="coerce")
        if column not in df.columns:
            df[column] = parsed_values
        else:
            df[column] = pd.to_numeric(df[column], errors="coerce").fillna(parsed_values)

    return df


def normalize_data(df):
    df = df.copy()
    df = df.rename(columns=_build_rename_map(df.columns))
    df = _parse_coordinates(df)

    if "date" not in df.columns:
        raise ValueError("The dataset must include a date column.")

    for column in COLUMN_ALIASES:
        if column not in df.columns:
            df[column] = pd.NA

    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    for column in NUMERIC_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    for column in TEXT_COLUMNS:
        df[column] = (
            df[column]
            .astype("string")
            .fillna("Unknown")
            .str.strip()
            .replace({"": "Unknown"})
        )

    df["casualties_total"] = df["casualties_total"].fillna(0)

    df.loc[~df["latitude"].between(-90, 90, inclusive="both"), "latitude"] = pd.NA
    df.loc[~df["longitude"].between(-180, 180, inclusive="both"), "longitude"] = pd.NA

    df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
    return df


@st.cache_data(show_spinner=False)
def load_data(url=SHEET_CSV_URL):
    df = pd.read_csv(url)
    return normalize_data(df)

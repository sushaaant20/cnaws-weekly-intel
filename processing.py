import pandas as pd

WINDOW_DAYS = 7


def _format_window_label(start_date, end_date):
    inclusive_end = end_date - pd.Timedelta(days=1)
    return f"{start_date:%d %b %Y} to {inclusive_end:%d %b %Y}"


def get_time_windows(df, today=None):
    df = df.copy()

    anchor_day = pd.Timestamp(today).normalize() if today is not None else pd.Timestamp.now().normalize()
    end_current = anchor_day
    start_current = end_current - pd.Timedelta(days=WINDOW_DAYS)
    start_prev = start_current - pd.Timedelta(days=WINDOW_DAYS)

    df_current = df[(df["date"] >= start_current) & (df["date"] < end_current)].copy()
    df_prev = df[(df["date"] >= start_prev) & (df["date"] < start_current)].copy()

    windows = {
        "current": {
            "start": start_current,
            "end": end_current,
            "label": _format_window_label(start_current, end_current),
        },
        "previous": {
            "start": start_prev,
            "end": start_current,
            "label": _format_window_label(start_prev, start_current),
        },
    }

    return df_current, df_prev, windows

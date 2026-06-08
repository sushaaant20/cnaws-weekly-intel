import pandas as pd

WINDOW_DAYS = 7
PERIOD_OPTIONS = ("weekly", "fortnightly")


def _format_window_label(start_date, end_date):
    inclusive_end = end_date - pd.Timedelta(days=1)
    return f"{start_date:%d %b %Y} to {inclusive_end:%d %b %Y}"


def _get_weekly_windows(anchor_day):
    end_current = anchor_day
    start_current = end_current - pd.Timedelta(days=WINDOW_DAYS)
    start_prev = start_current - pd.Timedelta(days=WINDOW_DAYS)
    return start_current, end_current, start_prev, start_current


def _get_fortnightly_windows(anchor_day):
    if anchor_day.day <= 15:
        end_current = anchor_day.replace(day=1)
    else:
        end_current = anchor_day.replace(day=16)

    start_current = end_current - pd.Timedelta(days=15)
    end_previous = start_current
    start_previous = end_previous - pd.Timedelta(days=15)
    return start_current, end_current, start_previous, end_previous


def get_time_windows(df, period="weekly", today=None):
    df = df.copy()
    anchor_day = pd.Timestamp(today).normalize() if today is not None else pd.Timestamp.now().normalize()

    period_key = str(period).strip().lower()
    if period_key not in PERIOD_OPTIONS:
        raise ValueError(f"Unsupported reporting period '{period}'.")

    if period_key == "fortnightly":
        start_current, end_current, start_prev, end_prev = _get_fortnightly_windows(anchor_day)
        comparison_label = "Fortnight Window"
        period_label = "Fortnightly"
    else:
        start_current, end_current, start_prev, end_prev = _get_weekly_windows(anchor_day)
        comparison_label = "Rolling Window"
        period_label = "Weekly"

    df_current = df[(df["date"] >= start_current) & (df["date"] < end_current)].copy()
    df_prev = df[(df["date"] >= start_prev) & (df["date"] < end_prev)].copy()

    windows = {
        "current": {
            "start": start_current,
            "end": end_current,
            "label": _format_window_label(start_current, end_current),
        },
        "previous": {
            "start": start_prev,
            "end": end_prev,
            "label": _format_window_label(start_prev, end_prev),
        },
        "comparison_label": comparison_label,
        "period_label": period_label,
        "period_key": period_key,
    }

    return df_current, df_prev, windows

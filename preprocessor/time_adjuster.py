"""출퇴근 시간을 급여 계산 규칙에 맞게 보정한다."""

from __future__ import annotations

import pandas as pd

from config.data_config import ACTUAL_END_COL, ACTUAL_START_COL, END_COL, START_COL


def _adjust_hour_by_min(
    ts: pd.Timestamp,
    mode: str = "round",
    thresh_min: int = 5,
) -> pd.Timestamp:
    """분 단위 값을 기준으로 시간을 보정한다."""
    if pd.isna(ts):
        return pd.NaT

    ts = pd.Timestamp(ts)
    if mode == "round":
        if ts.minute < thresh_min:
            return ts.floor("h")
        return ts.floor("h") + pd.Timedelta(hours=1)

    if mode == "floor":
        if ts.minute < 56:
            return ts.floor("h")
        return ts.floor("h") + pd.Timedelta(hours=1)

    raise ValueError("mode는 'round' 또는 'floor'만 가능합니다.")


def adjust_commute_time_columns(df: pd.DataFrame) -> pd.DataFrame:
    """raw datetime 컬럼을 분 단위 규칙에 맞게 보정한다."""
    if df is None or df.empty:
        return df.copy()

    for col in (START_COL, END_COL):
        if col not in df.columns:
            raise KeyError(f"{col} 컬럼이 없습니다.")

    working = df.copy()
    working[ACTUAL_START_COL] = working[START_COL].map(
        lambda x: _adjust_hour_by_min(x, mode="round")
    )
    working[ACTUAL_END_COL] = working[END_COL].map(
        lambda x: _adjust_hour_by_min(x, mode="floor")
    )
    return working

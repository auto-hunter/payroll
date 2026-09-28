"""근무 기록에 대상 월과 달력 정보를 적용한다."""

from __future__ import annotations

import re

import pandas as pd

from config.data_config import (
    COMPANY_COL,
    HOLIDAY_DATES,
    SHIFT_DATES,
    TARGET_MONTH,
    WEEKDAY_COL,
    WEEKEND_WEEKDAYS,
    WORK_DATE_COL,
)

KOREAN_WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]


def filter_work_records_by_target_month(
    df: pd.DataFrame,
    target_month: str = TARGET_MONTH,
) -> pd.DataFrame:
    """근무일자가 대상 월에 속하는 근무 기록만 남긴다."""
    if df is None or df.empty:
        return df.copy() if df is not None else pd.DataFrame()

    if WORK_DATE_COL not in df.columns:
        raise KeyError(f"필수 컬럼이 없습니다: {WORK_DATE_COL}")

    normalized_target_month = str(target_month).strip()
    if not re.fullmatch(r"\d{4}-\d{2}", normalized_target_month):
        raise ValueError("target_month는 YYYY-MM 형식이어야 합니다.")

    target_period = pd.Period(normalized_target_month, freq="M")
    working = df.copy()
    work_dates = pd.to_datetime(working[WORK_DATE_COL], errors="coerce")
    if work_dates.isna().any():
        raise ValueError(f"{WORK_DATE_COL} 컬럼에 날짜로 변환할 수 없는 값이 있습니다.")

    in_target_month = work_dates.dt.to_period("M") == target_period
    return working.loc[in_target_month].reset_index(drop=True)


def add_weekday_column(df: pd.DataFrame) -> pd.DataFrame:
    """근무일자의 요일을 새 컬럼으로 추가한다."""
    if df is None or df.empty:
        return df.copy()

    if WORK_DATE_COL not in df.columns:
        raise KeyError(f"{WORK_DATE_COL} 컬럼이 없습니다.")

    working = df.copy()

    def _to_weekday(value: object) -> object:
        if pd.isna(value):
            return pd.NA
        ts = pd.Timestamp(value)
        return KOREAN_WEEKDAYS[ts.weekday()]

    working[WEEKDAY_COL] = working[WORK_DATE_COL].map(_to_weekday)
    return working


def add_holiday_shift_columns(df: pd.DataFrame) -> pd.DataFrame:
    """휴일, 공휴일 및 교대일 컬럼을 추가한다."""
    if df is None or df.empty:
        return df.copy()

    required_columns = (WORK_DATE_COL, WEEKDAY_COL)
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise KeyError(f"필수 컬럼이 없습니다: {', '.join(missing_columns)}")

    working = df.copy()
    work_dates = pd.to_datetime(working[WORK_DATE_COL], errors="coerce").dt.normalize()
    holiday_dates = pd.to_datetime(HOLIDAY_DATES, errors="coerce").dropna().normalize()
    shift_dates = pd.to_datetime(SHIFT_DATES, errors="coerce").dropna().normalize()

    working["휴일"] = working[WEEKDAY_COL].isin(WEEKEND_WEEKDAYS).astype(int)
    working["공휴일"] = work_dates.isin(holiday_dates).astype(int)
    working["교대일"] = work_dates.isin(shift_dates).astype(int)
    return working


def filter_date(df: pd.DataFrame, target_month) -> pd.DataFrame:
    """근무일자가 target_month 기준 시점까지인 행만 남긴다."""
    if df is None or df.empty:
        return df.copy()

    if WORK_DATE_COL not in df.columns:
        raise KeyError(f"{WORK_DATE_COL} 컬럼이 없습니다.")

    target_ts = pd.to_datetime(target_month, errors="coerce")
    if pd.isna(target_ts):
        raise ValueError("target_month는 yyyy-mm 문자열 또는 pd.Timestamp로 변환 가능한 값이어야 합니다.")

    working = df.copy()
    work_dates = pd.to_datetime(working[WORK_DATE_COL], errors="coerce")

    if isinstance(target_month, str) and re.fullmatch(r"\d{4}-\d{2}", target_month.strip()):
        cutoff = target_ts + pd.offsets.MonthBegin(1)
        mask = work_dates < cutoff
    else:
        mask = work_dates <= target_ts

    return working.loc[mask].reset_index(drop=True)


def filter_before_target_month_start_week(df: pd.DataFrame, target_month) -> pd.DataFrame:
    """대상 월 1일이 속한 월~일 주간부터의 행만 남긴다."""
    if df is None or df.empty:
        return df.copy()

    if WORK_DATE_COL not in df.columns:
        raise KeyError(f"{WORK_DATE_COL} 컬럼이 없습니다.")

    target_ts = pd.to_datetime(target_month, errors="coerce")
    if pd.isna(target_ts):
        raise ValueError("target_month는 yyyy-mm 문자열 또는 pd.Timestamp로 변환 가능한 값이어야 합니다.")

    target_month_first_day = target_ts.replace(day=1).normalize()
    week_start = target_month_first_day - pd.Timedelta(days=target_month_first_day.weekday())

    working = df.copy()
    work_dates = pd.to_datetime(working[WORK_DATE_COL], errors="coerce")
    mask = work_dates >= week_start
    return working.loc[mask].reset_index(drop=True)


def fill_missing_dates(
    df: pd.DataFrame,
    target_month: str = TARGET_MONTH,
) -> pd.DataFrame:
    """대상 월을 넘지 않는 범위에서 사용자별 누락 날짜 행을 채운다."""
    if df is None or df.empty:
        return df.copy() if df is not None else pd.DataFrame()

    required_columns = ("사용자ID", "이름", WORK_DATE_COL)
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise KeyError(f"필수 컬럼이 없습니다: {', '.join(missing_columns)}")

    working = df.copy()
    work_dates = pd.to_datetime(working[WORK_DATE_COL], errors="coerce").dt.normalize()
    if work_dates.isna().any():
        raise ValueError(f"{WORK_DATE_COL} 컬럼에 날짜로 변환할 수 없는 값이 있습니다.")

    normalized_target_month = str(target_month).strip()
    if not re.fullmatch(r"\d{4}-\d{2}", normalized_target_month):
        raise ValueError("target_month는 YYYY-MM 형식이어야 합니다.")

    target_period = pd.Period(normalized_target_month, freq="M")
    target_month_end = target_period.end_time.normalize()
    fill_end = min(work_dates.max(), target_month_end)
    full_dates = pd.date_range(start=work_dates.min(), end=fill_end, freq="D")
    calendar = pd.DataFrame(
        {
            WORK_DATE_COL: full_dates,
            WEEKDAY_COL: [KOREAN_WEEKDAYS[date.weekday()] for date in full_dates],
        }
    )

    output_columns = list(working.columns)
    if WEEKDAY_COL not in output_columns:
        output_columns.append(WEEKDAY_COL)

    working[WORK_DATE_COL] = work_dates
    groups = working.groupby(["사용자ID", "이름"], dropna=False, sort=False)
    filled_groups = []

    for _, person_df in groups:
        person_df = person_df.drop(columns=[WEEKDAY_COL], errors="ignore")
        filled = calendar.merge(person_df, on=WORK_DATE_COL, how="left")

        actual_rows_after_target = person_df.loc[
            person_df[WORK_DATE_COL] > target_month_end
        ].copy()
        if not actual_rows_after_target.empty:
            actual_rows_after_target[WEEKDAY_COL] = actual_rows_after_target[
                WORK_DATE_COL
            ].map(lambda date: KOREAN_WEEKDAYS[date.weekday()])
            filled = pd.concat(
                [filled, actual_rows_after_target],
                ignore_index=True,
                sort=False,
            )

        filled["사용자ID"] = person_df["사용자ID"].iloc[0]
        filled["이름"] = person_df["이름"].iloc[0]
        if COMPANY_COL in person_df.columns:
            filled[COMPANY_COL] = person_df[COMPANY_COL].iloc[0]
        filled = filled.sort_values(WORK_DATE_COL).reset_index(drop=True)
        filled[WORK_DATE_COL] = filled[WORK_DATE_COL].dt.strftime("%Y-%m-%d")
        filled_groups.append(filled[output_columns])

    result = pd.concat(filled_groups, ignore_index=True)
    return add_holiday_shift_columns(result)

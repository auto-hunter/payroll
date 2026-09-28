"""급여 계산용 출퇴근 데이터 전처리 파이프라인."""

from __future__ import annotations

import pandas as pd

from preprocessor.calendar_enricher import (
    add_holiday_shift_columns,
    add_weekday_column,
    fill_missing_dates,
    filter_work_records_by_target_month,
)
from preprocessor.commute_parser import parse_commute_logs
from preprocessor.raw_cleaner import add_company_column, filter_rows
from preprocessor.reference_matcher import add_nearest_reference_logs
from preprocessor.time_adjuster import adjust_commute_time_columns


def preprocess_commute_logs(raw_df: pd.DataFrame) -> pd.DataFrame:
    """원본 캡스 로그에 전체 전처리 단계를 순서대로 적용한다."""
    df = add_company_column(raw_df)
    df = filter_rows(df)
    df = parse_commute_logs(df)
    df = filter_work_records_by_target_month(df)
    df = add_weekday_column(df)
    df = adjust_commute_time_columns(df)
    df = add_holiday_shift_columns(df)
    df = fill_missing_dates(df)
    return add_nearest_reference_logs(df, raw_df)

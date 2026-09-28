"""누락된 출퇴근 시간에 가장 가까운 출입 참고 로그를 연결한다."""

from __future__ import annotations

import pandas as pd

from config.data_config import (
    COMMUTE_COLUMN_ORDER,
    END_COL,
    REFERENCE_LOG_MODES,
    REFERENCE_LOG_TOLERANCE,
    REFERENCE_OUTPUT_COLUMNS,
    REFERENCE_SHIFT_DETECTION_TOLERANCE,
    REFERENCE_WORK_SCHEDULES,
    SHIFT_DATES,
    START_COL,
    USER_ID_COL,
    WEEKDAY_COL,
    WORK_DATE_COL,
)
from preprocessor.calendar_enricher import KOREAN_WEEKDAYS


def add_nearest_reference_logs(
    commute_df: pd.DataFrame,
    reference_df: pd.DataFrame,
    tolerance: str | pd.Timedelta | None = REFERENCE_LOG_TOLERANCE,
) -> pd.DataFrame:
    """한쪽만 누락된 출퇴근 기록에 가장 가까운 출입 참고시간을 추가한다."""
    if commute_df is None:
        return pd.DataFrame()

    working = commute_df.copy()
    for reference_time_col in REFERENCE_OUTPUT_COLUMNS.values():
        working[reference_time_col] = pd.NaT

    if working.empty:
        return working

    commute_required = (
        USER_ID_COL,
        WORK_DATE_COL,
        WEEKDAY_COL,
        "교대일",
        START_COL,
        END_COL,
    )
    missing_commute_columns = [
        column for column in commute_required if column not in working.columns
    ]
    if missing_commute_columns:
        raise KeyError(
            f"출퇴근 데이터에 필수 컬럼이 없습니다: {', '.join(missing_commute_columns)}"
        )

    reference_required = ("발생일자", "발생시각", USER_ID_COL, "모드")
    if reference_df is None:
        raise ValueError("reference_df는 None일 수 없습니다.")
    missing_reference_columns = [
        column for column in reference_required if column not in reference_df.columns
    ]
    if missing_reference_columns:
        raise KeyError(
            f"참고 데이터에 필수 컬럼이 없습니다: {', '.join(missing_reference_columns)}"
        )

    if tolerance is None:
        normalized_tolerance = None
    else:
        try:
            normalized_tolerance = pd.Timedelta(tolerance)
        except (TypeError, ValueError) as exc:
            raise ValueError("tolerance는 시간 간격으로 변환 가능한 값이어야 합니다.") from exc
        if normalized_tolerance < pd.Timedelta(0):
            raise ValueError("tolerance는 0 이상이어야 합니다.")

    references = reference_df.loc[
        reference_df["모드"].isin(REFERENCE_LOG_MODES),
        list(reference_required),
    ].copy()
    if references.empty:
        return working

    references["_참고일시"] = pd.to_datetime(
        references["발생일자"].astype(str).str.strip()
        + " "
        + references["발생시각"].astype(str).str.strip(),
        errors="coerce",
    )
    invalid_reference_time = references["_참고일시"].isna()
    if invalid_reference_time.any():
        raise ValueError("참고 데이터에 일시로 변환할 수 없는 값이 있습니다.")

    references = references.loc[references[USER_ID_COL].notna()].copy()
    references = references.sort_values("_참고일시", kind="stable")
    if references.empty:
        return working

    work_dates = pd.to_datetime(working[WORK_DATE_COL], errors="coerce").dt.normalize()
    if work_dates.isna().any():
        raise ValueError(f"{WORK_DATE_COL} 컬럼에 날짜로 변환할 수 없는 값이 있습니다.")

    shift_detection_tolerance = pd.Timedelta(REFERENCE_SHIFT_DETECTION_TOLERANCE)
    actual_times = {}
    for time_col in REFERENCE_OUTPUT_COLUMNS:
        actual_times[time_col] = pd.to_datetime(working[time_col], errors="coerce")
        invalid_actual_time = working[time_col].notna() & actual_times[time_col].isna()
        if invalid_actual_time.any():
            raise ValueError(f"{time_col} 컬럼에 일시로 변환할 수 없는 값이 있습니다.")

    reference_groups = {
        user_id: group.reset_index(drop=True)
        for user_id, group in references.groupby(USER_ID_COL, sort=False)
    }

    def _find_nearest_reference(user_id, target_time):
        user_references = reference_groups.get(user_id)
        if user_references is None or user_references.empty:
            return None

        differences = (user_references["_참고일시"] - target_time).abs()
        nearest_position = differences.argmin()
        difference = differences.iloc[nearest_position]
        if normalized_tolerance is not None and difference > normalized_tolerance:
            return None

        return user_references.iloc[nearest_position], difference

    shift_dates = {
        date.normalize()
        for date in pd.to_datetime(SHIFT_DATES, errors="coerce")
        if pd.notna(date)
    }

    def _get_schedules(anchor_date):
        weekday = KOREAN_WEEKDAYS[anchor_date.weekday()]
        is_shift_day = anchor_date in shift_dates
        return {
            name: rule
            for name, rule in REFERENCE_WORK_SCHEDULES.items()
            if rule["shift_day"] == is_shift_day
            and (rule["weekdays"] is None or weekday in rule["weekdays"])
        }

    for position, (_, row) in enumerate(working.iterrows()):
        user_id = row[USER_ID_COL]
        if pd.isna(user_id):
            continue

        has_start = pd.notna(actual_times[START_COL].iloc[position])
        has_end = pd.notna(actual_times[END_COL].iloc[position])
        if has_start == has_end:
            continue

        if has_start:
            known_time_col = START_COL
            missing_time_col = END_COL
            anchor_dates = (work_dates.iloc[position],)
        else:
            known_time_col = END_COL
            missing_time_col = START_COL
            anchor_dates = (
                work_dates.iloc[position],
                work_dates.iloc[position] - pd.Timedelta(days=1),
            )

        known_time = actual_times[known_time_col].iloc[position]
        schedule_candidates = []
        for anchor_date in anchor_dates:
            for schedule_name, rule in _get_schedules(anchor_date).items():
                scheduled_known_time = anchor_date + pd.Timedelta(rule[known_time_col])
                difference = abs(known_time - scheduled_known_time)
                if difference <= shift_detection_tolerance:
                    scheduled_missing_time = anchor_date + pd.Timedelta(
                        rule[missing_time_col]
                    )
                    schedule_candidates.append(
                        (difference, schedule_name, scheduled_missing_time)
                    )

        if not schedule_candidates:
            continue

        schedule_candidates.sort(key=lambda candidate: candidate[0])
        best_candidate = schedule_candidates[0]
        if (
            len(schedule_candidates) > 1
            and schedule_candidates[1][0] == best_candidate[0]
        ):
            continue

        nearest = _find_nearest_reference(user_id, best_candidate[2])
        if nearest is None:
            continue

        reference, _ = nearest
        reference_time_col = REFERENCE_OUTPUT_COLUMNS[missing_time_col]
        working.iat[position, working.columns.get_loc(reference_time_col)] = (
            reference["_참고일시"]
        )

    ordered = [column for column in COMMUTE_COLUMN_ORDER if column in working.columns]
    remaining = [column for column in working.columns if column not in ordered]
    return working[ordered + remaining]

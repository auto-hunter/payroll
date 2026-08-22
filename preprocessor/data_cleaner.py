"""데이터 전처리 모듈

캡스 출퇴근 로그 엑셀 파일을 입력받아 급여 계산에 필요한 데이터 가공을 수행합니다.
모든 함수의 입력과 출력은 pd.DataFrame입니다.

파이프라인 흐름(Flow):
    input excel ──> file validation ──> [data cleaning] ──> output dataframe

주요 함수:
    - filter_rows: 불필요한 행을 제거하고 계산에 필요한 행들만 남깁니다.
    - parse_commute_logs: 1열로 입력된 로그를 파싱하여 각 컬럼에 맞게 분리합니다.
    - add_weekday_column: 기준이 되는 출근 시간의 요일을 파악합니다.
    - add_holiday_shift_columns: 설정된 요일과 날짜를 기준으로 휴일, 공휴일 및 교대일을 표시합니다.
    - adjust_commute_time_columns: 출퇴근 시간을 정규화합니다.
"""

from __future__ import annotations

import re

import pandas as pd
import numpy as np

from config.config import (
    DATA_CLEANER_FILTER_ITEMS,
    HOLIDAY_DATES,
    WEEKEND_WEEKDAYS,

    INPUT_FILE_REQUIRED_COLUMNS,
    START_COL,
    END_COL,
    WORK_DATE_COL,
    MAX_WORK_HOURS,

    WEEKDAY_COL,

    ACTUAL_START_COL,
    ACTUAL_END_COL,
    SHIFT_DATES,
    TARGET_MONTH,
    COMPANY_COL,

    TAEIL_CABLE, TAEIL_MATERIAL
)

KOREAN_WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]

def _normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _to_normalized_set(value: object) -> set[str]:
    if value is None:
        return set()

    if isinstance(value, str):
        normalized = _normalize_text(value)
        return {normalized} if normalized else {""}

    if isinstance(value, (list, tuple, set, frozenset)):
        return {_normalize_text(item) for item in value}

    normalized = _normalize_text(value)
    return {normalized} if normalized else {""}


def _build_rule_sets(rule: object) -> tuple[set[str], set[str]]:
    """
    DATA_CLEANER_UNNECESSARY_ITEMS의 한 컬럼 규칙을 include/exclude 집합으로 정규화한다.

    지원 포맷:
    - "값" 또는 ["값1", "값2"]
    - {"include": [...], "exclude": [...]}
    """
    if isinstance(rule, dict):
        include_values = _to_normalized_set(rule.get("include"))
        exclude_values = _to_normalized_set(rule.get("exclude"))
        return include_values, exclude_values

    return set(), _to_normalized_set(rule)


def filter_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    입력 DataFrame에서 계산에 불필요한 행을 제거한다.

    제거 조건:
    - DATA_CLEANER_FILTER_ITEMS 정의된
      각 컬럼별 include/exclude 규칙에 따라 필터링

    원본은 변경하지 않고 필터링된 새 DataFrame을 반환한다.
    """
    if df is None or df.empty:
        return df.copy()

    working = df.copy()
    keep_mask = pd.Series(True, index=working.index)

    for column, rule in DATA_CLEANER_FILTER_ITEMS.items():
        if column not in working.columns:
            continue

        include_values, exclude_values = _build_rule_sets(rule)
        normalized_column = working[column].map(_normalize_text)

        if include_values:
            keep_mask &= normalized_column.isin(include_values)

        if exclude_values:
            keep_mask &= ~normalized_column.isin(exclude_values)

    return working.loc[keep_mask].reset_index(drop=True)


def filter_logs_after_target_month(
    df: pd.DataFrame,
    target_month: str = TARGET_MONTH,
) -> pd.DataFrame:
    """대상 월 이후에는 퇴근 로그만 남긴다.

    대상 월 마지막 날까지의 로그는 출근 시각과 관계없이 모두 유지한다.
    따라서 말일 야간 출근과 다음 달에 기록된 퇴근을
    :func:`parse_commute_logs`가 하나의 근무로 연결할 수 있다.

    다음 달 이후에 새로 기록된 출근 로그는 17시 이후 야간 출근을 포함해
    모두 제거하고, 이전 근무를 마감할 수 있는 퇴근 로그만 유지한다.

    Args:
        df: ``발생일자``와 ``모드`` 컬럼을 포함한 원본 출퇴근 로그.
        target_month: 급여 대상 월을 나타내는 ``YYYY-MM`` 문자열.

    Returns:
        원본 순서를 유지하면서 조건에 맞는 행만 남긴 DataFrame.

    Raises:
        KeyError: 필요한 컬럼이 없을 때.
        ValueError: 대상 월 또는 발생일자를 날짜로 변환할 수 없을 때.
    """
    if df is None or df.empty:
        return df.copy() if df is not None else pd.DataFrame()

    required_columns = ("발생일자", "모드")
    missing_columns = [
        column for column in required_columns if column not in df.columns
    ]
    if missing_columns:
        raise KeyError(f"필수 컬럼이 없습니다: {', '.join(missing_columns)}")

    normalized_target_month = str(target_month).strip()
    if not re.fullmatch(r"\d{4}-\d{2}", normalized_target_month):
        raise ValueError("target_month는 YYYY-MM 형식이어야 합니다.")

    target_period = pd.Period(normalized_target_month, freq="M")
    working = df.copy()
    event_dates = pd.to_datetime(working["발생일자"], errors="coerce")
    if event_dates.isna().any():
        raise ValueError("발생일자 컬럼에 날짜로 변환할 수 없는 값이 있습니다.")

    after_target_month = event_dates.dt.to_period("M") > target_period
    clock_out = working["모드"].map(_normalize_text).eq("퇴근")
    keep_mask = ~after_target_month | clock_out

    return working.loc[keep_mask].reset_index(drop=True)


def parse_commute_logs(df: pd.DataFrame) -> pd.DataFrame:
    """
    캡스 조회 데이터를 가공해서 출근/퇴근 시간으로 변환한다.

    - 데이터 로그
        - 변환 전:
            - 출/퇴근 이벤트가 하나의 열에 기록되어 있음
            | 발생일자     | 발생시각   | 모드 |
            --------------------------------
            | 2026-03-25 | 00:00:27 | 출근 |
            | 2026-03-25 | 06:32:49 | 퇴근 |
        - 변환 후:
            | 출근시간              | 퇴근시간              |
            ----------------------------------------------
            | 2026-03-25 00:00:27 | 2026-03-25 06:32:49  |

    - 데이터 포맷
        - 변환 전:
            - 발생일자: YYYY-MM-DD (string)
            - 발생시각: HH:MM:SS (str)
        - 변환 후:
            - 출근시간: pd.Timestamp
            - 퇴근시간: pd.Timestamp
    """
    if df is None or df.empty:
        return pd.DataFrame()

    required_columns = (*INPUT_FILE_REQUIRED_COLUMNS, COMPANY_COL)
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise KeyError(f"필수 컬럼이 없습니다: {', '.join(missing_columns)}")

    # 1. 일시를 생성할 때부터 pd.to_datetime을 사용하여 Timestamp 객체로 만듭니다.
    df = df.copy()
    df["일시"] = pd.to_datetime(df["발생일자"] + " " + df["발생시각"])
    df = df.sort_values(by=["사용자ID", "일시"]).reset_index(drop=True)

    processed_records = []

    for user_id, group in df.groupby("사용자ID"):
        group = group.reset_index(drop=True)
        i = 0
        while i < len(group):
            row = group.iloc[i]

            if row["모드"] == "출근":
                user_name = row["이름"]
                company = row[COMPANY_COL]
                work_date = row["발생일자"]
                # 출근시간에 string 대신 Timestamp 객체('일시')를 할당합니다.
                clock_in_time = row["일시"]

                # 누락 시를 대비하여 초기값을 None 대신 pd.NaT로 설정합니다.
                clock_out_time = pd.NaT
                last_valid_out_index = None

                for j in range(i + 1, len(group)):
                    next_row = group.iloc[j]

                    if next_row["모드"] == "퇴근":
                        time_diff = (
                            next_row["일시"] - clock_in_time
                        ).total_seconds() / 3600

                        if time_diff <= MAX_WORK_HOURS:
                            # 퇴근시간에 string 대신 Timestamp 객체('일시')를 그대로 할당합니다.
                            clock_out_time = next_row["일시"]
                            last_valid_out_index = j
                        else:
                            break

                    elif next_row["모드"] == "출근":
                        if last_valid_out_index is not None:
                            break

                processed_records.append(
                    {
                        "사용자ID": user_id,
                        "이름": user_name,
                        COMPANY_COL: company,
                        WORK_DATE_COL: work_date,
                        START_COL: clock_in_time,
                        END_COL: clock_out_time,
                    }
                )

                i = last_valid_out_index + 1 if last_valid_out_index is not None else i + 1

            else:
                # 출근 없이 퇴근만 찍힌 경우 처리
                processed_records.append(
                    {
                        "사용자ID": user_id,
                        "이름": row["이름"],
                        COMPANY_COL: row[COMPANY_COL],
                        WORK_DATE_COL: row["발생일자"],
                        START_COL: pd.NaT,
                        END_COL: row["일시"],  # 여기도 Timestamp 객체 할당
                    }
                )
                i += 1

    result_df = pd.DataFrame(processed_records)
    return result_df


def add_weekday_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    work_date의 요일을 새 컬럼으로 추가한다.

    예:
    - 월요일 -> "월"
    """
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
    """
    설정된 요일과 날짜를 기준으로 휴일, 공휴일 및 교대일 컬럼을 추가한다.

    - WEEKEND_WEEKDAYS에 해당하면 휴일을 1로 표시한다.
    - HOLIDAY_DATES에 해당하면 공휴일을 1로 표시한다.
    - SHIFT_DATES에 해당하면 교대일을 1로 표시한다.
    """
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


def _adjust_hour_by_min(
    ts: pd.Timestamp,
    mode: str = "round",
    thresh_min: int = 5,
) -> pd.Timestamp:
    """
    분 단위 값을 기준으로 시간을 보정한다.

    mode="round":
    - 00~(thresh_min-1)분 -> 현재 시간
    - thresh_min분 이상 -> 다음 시간

    mode="floor":
    - 00~55분 -> 현재 시간
    - 56~59분 -> 다음 시간
    """
    if pd.isna(ts):
        return pd.NaT

    ts = pd.Timestamp(ts)

    # 출근
    if mode == "round":
        if ts.minute < thresh_min:
            return ts.floor("h")
        return ts.floor("h") + pd.Timedelta(hours=1)

    # 퇴근
    if mode == "floor":
        if ts.minute < 56:
            return ts.floor("h")
        return ts.floor("h") + pd.Timedelta(hours=1)

    raise ValueError("mode는 'round' 또는 'floor'만 가능합니다.")


def adjust_commute_time_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    raw datetime 컬럼을 분 단위 규칙에 맞게 보정한다.
    """
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


def filter_date(df: pd.DataFrame, target_month) -> pd.DataFrame:
    """
    근무일자가 target_month 기준 시점까지인 행만 남긴다.

    target_month:
    - "2026-04" 같은 yyyy-mm 문자열: 해당 월의 마지막 날까지 포함
    - "2026-04-15" 같은 날짜 문자열
    - pd.Timestamp

    반환값:
    - 월 단위 문자열이면 해당 월 이하, 그 외에는 target_month 이하인 새 DataFrame
    """
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
    """
    target_month가 속한 달의 1일이 포함된 월~일 주간부터의 행만 남긴다. (주휴수당 계산을 위함)

    target_month:
    - "2026-04" 같은 yyyy-mm 문자열
    - "2026-04-15" 같은 날짜 문자열
    - pd.Timestamp

    예:
    - 2026-04-01이 수요일이면 2026-03-30 월요일부터 남긴다.
      따라서 이전 월 데이터는 2026-03-30, 2026-03-31만 유지된다.
    - 2026-06-01이 월요일이면 2026-06-01부터 남긴다.
      따라서 이전 월 데이터는 유지하지 않는다.
    - 2026-02-01이 일요일이면 2026-01-26 월요일부터 남긴다.
      따라서 이전 월 데이터는 6일 전까지 유지된다.
    """
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
    """
    대상 월을 넘지 않는 범위에서 사용자별 누락 날짜 행을 채운다.

    입력 데이터의 근무일자 최솟값부터 입력 최댓값과 대상 월 말일 중
    이른 날짜까지를 공통 범위로 사용한다. 대상 월 이후에는 결측 행을
    새로 만들지 않으며, 원본에 실제로 존재하는 행만 그대로 유지한다.
    생성된 결과의 휴일, 공휴일과 교대일은 config 설정을 기준으로 0 또는 1로 채운다.
    """
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

        # 대상 월 이후에는 실제 기록만 추가하고 중간의 빈 날짜는 생성하지 않는다.
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

def add_company_column(
    df,
    taeil_cable=TAEIL_CABLE,
    taeil_material=TAEIL_MATERIAL,
    user_col="이름"
):
    # 조건(Conditions) 설정
    conditions = [
        df[user_col].isin(taeil_cable),
        df[user_col].isin(taeil_material)
    ]

    # 조건에 맞을 때 들어갈 값(Choices) 설정
    choices = ["태일전선", "태일소재"]

    # 조건에 맞지 않는 나머지 데이터의 기본값 지정 (예: "미등록" 또는 pd.NA)
    df[COMPANY_COL] = np.select(conditions, choices, default="미등록")

    return df

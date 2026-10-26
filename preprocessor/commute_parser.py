"""개별 출입 이벤트를 출퇴근 근무 기록으로 조합한다."""

from __future__ import annotations

import pandas as pd

from config.data_config import (
    COMPANY_COL,
    END_COL,
    INPUT_FILE_REQUIRED_COLUMNS,
    MAX_WORK_HOURS,
    START_COL,
    WORK_DATE_COL,
)


def parse_commute_logs_sequential(df: pd.DataFrame) -> pd.DataFrame:
    """이벤트 순서에 따라 출근과 다음 퇴근을 하나의 근무 건으로 조합한다.

    날짜가 바뀌어도 ``MAX_WORK_HOURS`` 이내의 퇴근은 직전 출근과
    연결하며, 근무일자는 출근 이벤트의 발생일자로 정한다. 연속 출근은
    후보로 보관하고, 퇴근과 1시간 이상 차이 나는 가장 최근 출근을
    선택한다. 출근 후 1시간 미만의 퇴근은 잘못 찍힌 이벤트로 본다.

    같은 사용자와 근무일자에 여러 근무 건이 생기면 완성된 근무 건을
    우선하고, 완성된 건이 없으면 24시간 이내에서 연결 가능한 출근과
    퇴근 중 시간 차이가 가장 짧은 조합을 선택해 한 행으로 만든다.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    required_columns = (*INPUT_FILE_REQUIRED_COLUMNS, COMPANY_COL)
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise KeyError(f"필수 컬럼이 없습니다: {', '.join(missing_columns)}")

    minimum_work_duration = pd.Timedelta(hours=1)

    working = df.copy()
    working["일시"] = pd.to_datetime(
        working["발생일자"].astype(str).str.strip()
        + " "
        + working["발생시각"].astype(str).str.strip(),
        errors="coerce",
    )
    if working["일시"].isna().any():
        raise ValueError("발생일자와 발생시각에 일시로 변환할 수 없는 값이 있습니다.")

    working = working.sort_values(
        by=["사용자ID", "일시"], kind="stable"
    ).reset_index(drop=True)
    processed_records = []

    def append_record(start_row=None, end_row=None):
        source_row = start_row if start_row is not None else end_row
        start_time = start_row["일시"] if start_row is not None else pd.NaT
        end_time = end_row["일시"] if end_row is not None else pd.NaT
        work_date = (
            start_row["발생일자"]
            if start_row is not None
            else end_row["발생일자"]
        )
        processed_records.append(
            {
                "사용자ID": source_row["사용자ID"],
                "이름": source_row["이름"],
                COMPANY_COL: source_row[COMPANY_COL],
                WORK_DATE_COL: work_date,
                START_COL: start_time,
                END_COL: end_time,
            }
        )

    for _, group in working.groupby("사용자ID", sort=False):
        open_starts = []

        for _, event in group.iterrows():
            mode = event["모드"]

            if mode == "출근":
                open_starts.append(event)
                continue

            if mode != "퇴근":
                continue

            if not open_starts:
                append_record(end_row=event)
                continue

            elapsed_hours = [
                (event["일시"] - start["일시"]).total_seconds() / 3600
                for start in open_starts
            ]
            valid_start_indexes = [
                index
                for index, elapsed in enumerate(elapsed_hours)
                if 1 <= elapsed <= MAX_WORK_HOURS
            ]
            if valid_start_indexes:
                selected_index = valid_start_indexes[-1]
                for unmatched_start in open_starts[:selected_index]:
                    append_record(start_row=unmatched_start)
                append_record(
                    start_row=open_starts[selected_index],
                    end_row=event,
                )
                open_starts = []
                continue

            if any(0 <= elapsed < 1 for elapsed in elapsed_hours):
                continue

            for unmatched_start in open_starts:
                append_record(start_row=unmatched_start)
            append_record(end_row=event)
            open_starts = []

        for unmatched_start in open_starts:
            append_record(start_row=unmatched_start)

    result_df = pd.DataFrame(processed_records)
    if result_df.empty:
        return result_df

    def consolidate_same_work_date(group):
        if len(group) == 1:
            return group.iloc[0]

        complete = group.loc[
            group[START_COL].notna() & group[END_COL].notna()
        ].copy()
        if not complete.empty:
            complete["_근무시간"] = complete[END_COL] - complete[START_COL]
            valid_complete = complete.loc[
                (complete["_근무시간"] >= minimum_work_duration)
                & (
                    complete["_근무시간"]
                    <= pd.Timedelta(hours=MAX_WORK_HOURS)
                )
            ]
            if not valid_complete.empty:
                return valid_complete.sort_values(
                    "_근무시간", kind="stable"
                ).iloc[0].drop(labels="_근무시간")

        starts = group.loc[group[START_COL].notna()]
        ends = group.loc[group[END_COL].notna()]
        candidates = []
        for start_index, start_row in starts.iterrows():
            for _, end_row in ends.iterrows():
                elapsed = end_row[END_COL] - start_row[START_COL]
                if (
                    minimum_work_duration <= elapsed
                    <= pd.Timedelta(hours=MAX_WORK_HOURS)
                ):
                    candidates.append((elapsed, start_index, end_row[END_COL]))

        if candidates:
            _, start_index, end_time = min(candidates, key=lambda item: item[0])
            consolidated = group.loc[start_index].copy()
            consolidated[END_COL] = end_time
            return consolidated

        if not starts.empty:
            return starts.iloc[0]
        return ends.iloc[0]

    consolidated_rows = [
        consolidate_same_work_date(group)
        for _, group in result_df.groupby(
            ["사용자ID", WORK_DATE_COL], sort=False, dropna=False
        )
    ]
    return pd.DataFrame(consolidated_rows).reset_index(drop=True)

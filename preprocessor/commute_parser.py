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


def parse_commute_logs(df: pd.DataFrame) -> pd.DataFrame:
    """캡스 조회 데이터를 출근/퇴근 시간이 포함된 근무 기록으로 변환한다."""
    if df is None or df.empty:
        return pd.DataFrame()

    required_columns = (*INPUT_FILE_REQUIRED_COLUMNS, COMPANY_COL)
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise KeyError(f"필수 컬럼이 없습니다: {', '.join(missing_columns)}")

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
                clock_in_time = row["일시"]
                clock_out_time = pd.NaT
                last_valid_out_index = None

                for j in range(i + 1, len(group)):
                    next_row = group.iloc[j]

                    if next_row["모드"] == "퇴근":
                        time_diff = (
                            next_row["일시"] - clock_in_time
                        ).total_seconds() / 3600

                        if time_diff <= MAX_WORK_HOURS:
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
                processed_records.append(
                    {
                        "사용자ID": user_id,
                        "이름": row["이름"],
                        COMPANY_COL: row[COMPANY_COL],
                        WORK_DATE_COL: row["발생일자"],
                        START_COL: pd.NaT,
                        END_COL: row["일시"],
                    }
                )
                i += 1

    result_df = pd.DataFrame(processed_records)
    return result_df

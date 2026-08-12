"""데이터 검증 모듈.

캡스 출퇴근 로그 엑셀 파일을 입력받아 급여 계산에 필요한 데이터 기준을 충족하는지 검증합니다.
데이터 처리 파이프라인의 전처리 전 단계에서 입력 파일을 검증합니다.

파이프라인 흐름(Flow):
    input excel ──> file validation ──> data cleaning ──> output dataframe

주요 함수:
    - valid_input_file: 입력 파일의 존재 여부, 빈 파일 여부, 필수 컬럼 존재 여부를 검증합니다.
"""


from __future__ import annotations

import pandas as pd

from config.config import INPUT_FILE_REQUIRED_COLUMNS


def _is_blank_series(series: pd.Series) -> bool:
    return series.isna().all() or series.astype(str).str.strip().eq("").all()


def valid_input_file(df: pd.DataFrame) -> tuple[bool, list[str]]:
    """
    입력 DataFrame이 최소 요건을 만족하는지 검증한다.

    검증 항목:
    1. 파일이 비어있지 않은지
    2. 필수 컬럼이 있는지: ["발생일자", "발생시각", "단말기ID", "사용자ID", "이름", "사원번호", "모드"]

    반환값:
        (ok, errors)
        ok: 모든 검증을 통과하면 True
        errors: 실패한 항목들의 메시지 목록
    """
    errors: list[str] = []

    if df is None or df.empty:
        return False, ["입력 파일이 비어 있습니다."]

    missing_columns = [
        col for col in INPUT_FILE_REQUIRED_COLUMNS if col not in df.columns
    ]
    if missing_columns:
        errors.append(f"필수 컬럼이 없습니다: {', '.join(missing_columns)}")

    if errors:
        return False, errors

    if all(_is_blank_series(df[col]) for col in INPUT_FILE_REQUIRED_COLUMNS):
        return False, ["필수 컬럼에는 값이 하나도 없습니다."]

    return len(errors) == 0, errors

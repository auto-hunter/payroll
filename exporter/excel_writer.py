"""openpyxl을 이용한 엑셀 파일 출력 모듈.

DataFrame을 이름별 시트로 나누어 기록한다.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.workbook.properties import CalcProperties

from config.config import SHEET_SPLIT_COL
from exporter.excel_formulas import (
    FormulaColumn,
    SummaryFormula,
    apply_formula_columns,
    apply_summary_formulas,
)


def _make_sheet_title(value: object, used_titles: set[str]) -> str:
    """그룹 값을 엑셀 규칙에 맞는 고유한 시트명으로 변환한다.

    엑셀 시트명에는 일부 문자를 사용할 수 없고 최대 길이는 31자이다.
    또한 대소문자만 다른 이름도 같은 시트명으로 취급되므로, 이미 사용한
    이름과 겹치면 ``_2``, ``_3``과 같은 번호를 붙인다.

    Args:
        value: 시트명의 원본이 되는 그룹 값(기본적으로 직원 이름).
        used_titles: 현재 워크북에서 이미 사용한 시트명의 소문자 집합.

    Returns:
        엑셀에서 사용할 수 있으며 워크북 안에서 중복되지 않는 시트명.
    """
    # 이름이 비어 있거나 NaN인 그룹도 시트를 만들 수 있게 기본명을 준다.
    if pd.isna(value):
        base_title = "이름없음"
    else:
        base_title = str(value).strip() or "이름없음"

    # 엑셀이 금지하는 문자와 제어 문자를 '_'로 바꾸고 31자로 자른다.
    base_title = re.sub(r"[\\/*?:\[\]\x00-\x1f]", "_", base_title)[:31]
    title = base_title
    sequence = 2

    # casefold()를 사용해 영문 대소문자가 다른 제목도 중복으로 판단한다.
    while title.casefold() in used_titles:
        suffix = f"_{sequence}"
        # 번호까지 포함한 최종 제목도 31자를 넘지 않도록 원래 이름을 줄인다.
        title = f"{base_title[:31 - len(suffix)]}{suffix}"
        sequence += 1

    used_titles.add(title.casefold())
    return title


def _to_excel_value(value: object) -> object:
    """pandas/numpy 값을 openpyxl이 셀에 기록할 수 있는 값으로 변환한다.

    결측값은 빈 셀을 뜻하는 ``None``으로, pandas 날짜는 파이썬 datetime으로,
    numpy 스칼라는 파이썬 기본 스칼라로 변환한다. 그 밖의 값은 그대로 둔다.

    Args:
        value: DataFrame의 셀 값.

    Returns:
        openpyxl의 ``Worksheet.append``에 전달할 수 있는 값.
    """
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    # numpy.int64 같은 numpy 스칼라는 item()으로 int 등의 기본형이 된다.
    if hasattr(value, "item"):
        return value.item()
    return value


def write_dataframe_by_name(
    df: pd.DataFrame,
    output_path: str | Path,
    sheet_split_col: str = SHEET_SPLIT_COL,
    personal_formula_columns: Iterable[FormulaColumn] = (),
    personal_summary_formulas: Iterable[SummaryFormula] = (),
) -> Path:
    """DataFrame을 이름별 시트로 나누어 새 엑셀 파일에 저장한다.

    Args:
        df: 저장할 전체 데이터.
        output_path: 생성할 xlsx 파일의 경로.
        sheet_split_col: 데이터를 나누는 기준 컬럼.
        personal_formula_columns: 개인별 시트에 추가할 수식 열 규칙.
        personal_summary_formulas: 개인별 시트의 데이터 우측에 추가할 요약 수식 규칙.

    Returns:
        저장을 마친 파일의 ``Path`` 객체.

    Raises:
        ValueError: DataFrame이 없거나 비어 있을 때.
        KeyError: 그룹 기준 컬럼이 DataFrame에 없을 때.
    """
    # 빈 데이터로는 시트를 만들 수 없으므로 파일 생성 전에 명확히 실패시킨다.
    if df is None or df.empty:
        raise ValueError("저장할 데이터가 없습니다.")
    if sheet_split_col not in df.columns:
        raise KeyError(f"{sheet_split_col} 컬럼이 없습니다.")

    workbook = Workbook()

    # openpyxl은 수식을 기록할 수는 있지만 계산 엔진은 제공하지 않는다.
    # 따라서 Excel에서 파일을 열 때 모든 수식을 자동으로 다시 계산하도록
    # 통합문서 계산 속성을 명시한다. calcId=0은 현재 Excel 버전의 계산
    # 엔진으로 수식 캐시를 새로 만들도록 유도한다.
    workbook.calculation = CalcProperties(
        calcMode="auto",
        calcId=0,
        fullCalcOnLoad=True,
        forceFullCalc=True,
        calcOnSave=True,
    )

    # Workbook이 자동 생성하는 빈 기본 시트를 제거한다.
    workbook.remove(workbook.active)
    used_titles: set[str] = set()
    # generator로 전달된 규칙도 모든 시트에서 재사용할 수 있게 고정한다.
    personal_formula_rules = tuple(personal_formula_columns)
    personal_summary_rules = tuple(personal_summary_formulas)

    # sort=False는 원본에 등장한 이름 순서를 유지하고, dropna=False는
    # 이름이 비어 있는 행도 누락하지 않고 '이름없음' 시트에 포함한다.
    for name, df_name in df.groupby(sheet_split_col, dropna=False, sort=False):
        worksheet = workbook.create_sheet(_make_sheet_title(name, used_titles))

        # 첫 행에는 모든 컬럼명을 문자열로 변환하여 헤더로 기록한다.
        worksheet.append([str(column) for column in df_name.columns])

        # itertuples는 iterrows보다 가볍게 각 행의 값만 순서대로 제공한다.
        for row in df_name.itertuples(index=False, name=None):
            worksheet.append([_to_excel_value(value) for value in row])

        # 데이터 오른쪽에 호출자가 정의한 수식 열을 동일하게 추가한다.
        apply_formula_columns(worksheet, personal_formula_rules)
        # 빈 열 하나를 사이에 두고 제목/값 형태의 요약 영역을 추가한다.
        apply_summary_formulas(worksheet, personal_summary_rules)

    destination = Path(output_path)
    # 중간 폴더가 없어도 저장할 수 있도록 상위 경로를 먼저 생성한다.
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    return destination

"""설정으로 전달된 Excel 조건부서식을 컬럼명 기준으로 적용하는 모듈.

실제 서식 규칙은 ``config.excel_config``에서
``ColumnConditionalFormat`` 객체로 정의한다. 이 모듈은 사람이 이해하기
쉬운 DataFrame 컬럼명을 Excel 열 주소와 데이터 범위로 변환하고,
openpyxl의 조건부서식 규칙을 워크시트에 등록한다.

하나의 열뿐 아니라 ``column``부터 ``end_column``까지의 연속 범위에도
같은 규칙을 적용할 수 있다. ``excel_writer``가 개인 시트와 급여대장을
만드는 과정에서 이 모듈을 호출한다.
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from typing import Iterable

from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True)
class ColumnConditionalFormat:
    """컬럼명과 openpyxl Rule을 묶은 조건부서식 적용 규칙.

    ``end_column``을 생략하면 ``column`` 하나에만 적용하고, 지정하면 두
    컬럼 사이의 연속된 데이터 범위 전체에 적용한다.
    """

    column: str
    rule: Rule
    end_column: str | None = None


def apply_column_conditional_formats(
    worksheet: Worksheet,
    conditional_formats: Iterable[ColumnConditionalFormat],
    *,
    header_row: int = 1,
) -> None:
    """열 이름으로 데이터 범위를 찾아 조건부서식을 적용한다."""
    rules = tuple(conditional_formats)
    if not rules or worksheet.max_row <= header_row:
        return

    # 헤더명을 실제 Excel 열 번호로 변환하기 위한 색인을 만든다.
    column_indexes = {
        str(worksheet.cell(header_row, column).value): column
        for column in range(1, worksheet.max_column + 1)
        if worksheet.cell(header_row, column).value is not None
    }

    for conditional_format in rules:
        column_name = str(conditional_format.column).strip()
        if not column_name:
            raise ValueError("조건부서식을 적용할 열 이름은 비어 있을 수 없습니다.")
        if column_name not in column_indexes:
            raise KeyError(f"{column_name} 컬럼을 조건부서식에서 찾을 수 없습니다.")

        end_column_name = (
            str(conditional_format.end_column).strip()
            if conditional_format.end_column is not None
            else column_name
        )
        if not end_column_name:
            raise ValueError("조건부서식 범위의 끝 열 이름은 비어 있을 수 없습니다.")
        if end_column_name not in column_indexes:
            raise KeyError(
                f"{end_column_name} 컬럼을 조건부서식 범위에서 찾을 수 없습니다."
            )

        start_column_index = column_indexes[column_name]
        end_column_index = column_indexes[end_column_name]
        if start_column_index > end_column_index:
            raise ValueError(
                "조건부서식 범위의 끝 열은 시작 열보다 앞에 있을 수 없습니다."
            )

        start_column_letter = get_column_letter(start_column_index)
        end_column_letter = get_column_letter(end_column_index)
        cell_range = (
            f"{start_column_letter}{header_row + 1}:"
            f"{end_column_letter}{worksheet.max_row}"
        )
        # 동일한 설정을 여러 워크시트에 적용하므로 Rule 객체를 복사한다.
        worksheet.conditional_formatting.add(
            cell_range,
            copy(conditional_format.rule),
        )

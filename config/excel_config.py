"""급여 Excel 출력에 사용할 수식과 조건부서식 규칙을 정의한다.

이 파일은 Excel 파일을 직접 생성하지 않는다. 출력할 파일명과 시트 분리
기준, 개인별 시트의 계산 열, 개인 요약 영역, 급여대장의 조회 열 및
조건부서식 규칙을 선언한다. 실제 적용과 저장은 ``exporter`` 패키지가
담당한다.

설정이 적용되는 순서::

    전처리 데이터
        -> personal_formula_columns       # 개인별 일자 단위 근무시간 계산
        -> personal_summary_formulas      # 개인별 시간 합계·수당·공제 계산
        -> overall_formula_columns        # 개인 요약값을 급여대장으로 조회
        -> personal_conditional_formats   # 이상값과 휴일 행 시각화

규칙 객체의 역할:
    - ``FormulaColumn``: 개인 시트의 각 데이터 행에 수식 열을 추가한다.
    - ``SummaryFormula``: 개인 시트 우측에 집계 또는 급여 수식을 배치한다.
    - ``OverallFormula``: 사용자ID로 개인 요약값을 급여대장에 조회한다.
    - ``ColumnConditionalFormat``: 컬럼명으로 조건부서식 범위를 지정한다.

수식 문자열은 Excel에서 파일을 열 때 계산된다. 따라서 여기서 사용하는
컬럼명과 요약 항목명은 전처리 결과 및 다른 규칙의 이름과 일치해야 한다.
"""

from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill

from exporter.excel_conditional_formats import ColumnConditionalFormat
from exporter.excel_formulas import FormulaColumn, SummaryFormula, OverallFormula

from config.data_config import NAME_COL, TARGET_MONTH


# 직원 이름별로 개인 워크시트를 만들고 대상 월을 출력 파일명에 사용한다.
SHEET_SPLIT_COL = NAME_COL
OUTPUT_FILE_NAME = f"output_{TARGET_MONTH.replace('-', '')}.xlsx"


def payroll_lookup_summary(label, start_row, default=0):
    """급여대장 PayrollTable의 값을 개인 요약 영역에서 조회하는 규칙을 만든다.

    개인 시트의 첫 번째 이름을 기준으로 급여대장 테이블의 동일한 이름을
    찾아 ``label`` 컬럼 값을 가져온다. 통상시급과 각종 공제 항목처럼
    급여대장에 이미 들어 있는 값을 개인 명세에서 재사용할 때 사용한다.

    ``default``는 호출부의 의도를 표현하기 위해 유지된 매개변수이며, 현재
    생성되는 INDEX/MATCH 수식에는 직접 사용되지 않는다.
    """
    def build_formula(ctx):
        return (
            f"=INDEX(PayrollTable[{label}], "
            f"MATCH(INDEX({ctx.column_range('이름')},1), "
            f"PayrollTable[이름], 0))"
        )

    return SummaryFormula(
        label=label,
        formula=build_formula,
        number_format="#,##0",
        start_row=start_row,
    )

# ---------------------------------------------------------------------------
# 개인 시트: 일자별 근무시간 계산 열
#
# 위에서 아래 순서대로 워크시트 오른쪽에 추가된다. 뒤쪽 규칙은 앞서 생성된
# 열을 이름으로 참조할 수 있으므로 선언 순서가 계산 의존성을 나타낸다.
# ---------------------------------------------------------------------------
personal_formula_columns = [
    FormulaColumn(
        header="총근무시간",
        formula=lambda ctx: (
            f"=INT(ROUND(({ctx.cell('실퇴근시간')}-{ctx.cell('실출근시간')})*24, 6))"
        ),
        number_format="0"
    ),
    FormulaColumn(
        header="실근무시간",
        formula=lambda ctx: (
            f"=MAX({ctx.cell('총근무시간')} - 1, 0)"
        ),
        number_format="0"
    ),
    FormulaColumn(
        header="소정근무시간",
        formula=lambda ctx: (
            f"=IF({ctx.cell('휴일')} = 1, 0, MIN(MAX({ctx.cell('실근무시간')}, 0), 8))"
        ),
        number_format="0"
    ),
    FormulaColumn(
        header="평일연장근무시간",
        formula=lambda ctx: (
            f"=IF({ctx.cell('휴일')} = 1, 0, MAX({ctx.cell('실근무시간')} - {ctx.cell('소정근무시간')}, 0))"
        ),
        number_format="0"
    ),
    FormulaColumn(
        header="휴일근무시간",
        formula=lambda ctx: (
            f"=IF({ctx.cell('휴일')} = 1, MIN(MAX({ctx.cell('실근무시간')}, 0), 8), 0)"
        ),
        number_format="0"
    ),
    FormulaColumn(
        header="휴일연장근무시간",
        formula=lambda ctx: (
            f"=IF({ctx.cell('휴일')} = 1, MAX({ctx.cell('실근무시간')} - {ctx.cell('휴일근무시간')}, 0), 0)"
        ),
        number_format="0"
    ),
    FormulaColumn(
        header="야간근무시간",
        # 실제 근무 구간과 22:00~익일 06:00 구간이 겹치는 시간을 계산한다.
        formula=lambda ctx: (
            f'=IF('
            f'OR({ctx.cell("실출근시간")}="", {ctx.cell("실퇴근시간")}=""),'
            f'0,'
            f'MAX('
            f'0,'
            f'MIN('
            f'{ctx.cell("실퇴근시간")},'
            f'INT({ctx.cell("실출근시간")}) - (MOD({ctx.cell("실출근시간")}, 1) < 6/24) + 30/24'
            f') - MAX('
            f'{ctx.cell("실출근시간")},'
            f'INT({ctx.cell("실출근시간")}) - (MOD({ctx.cell("실출근시간")}, 1) < 6/24) + 22/24'
            f')'
            f') * 24'
            f')'
        ),
        number_format="0"
    ),
    FormulaColumn(
        header="주휴인정시간",
        # 일요일 행에서 직전 평일 5일 모두 소정근무가 있으면 8시간을 인정한다.
        formula=lambda ctx: (
            f'=IFERROR(IF(AND('
            f'{ctx.cell("요일")}="일",'
            f'COUNTIF(OFFSET({ctx.cell("소정근무시간")},-6,0,5,1),">0")=5'
            f'),8,0),0)'
        ),
        number_format="0",
    ),
    FormulaColumn(
        header="공휴일인정시간",
        formula=lambda ctx: f"=IF({ctx.cell("공휴일")} = 1, 8, 0)",
        number_format="0",
    ),
    FormulaColumn(
        header="연차인정시간",
        formula=lambda ctx: "=0",
        number_format="0",
    )
]

# ---------------------------------------------------------------------------
# 개인 시트: 우측 요약 영역
#
# 2~12행은 신원·시급·근무시간 집계, 14~21행은 지급 항목, 23~37행은
# 공제 항목, 39행은 최종 차인지급액이다. start_row는 명세서의 배치를
# 고정하므로 변경할 때 다른 요약 셀 참조도 함께 확인해야 한다.
# ---------------------------------------------------------------------------
personal_summary_formulas = [
    SummaryFormula(
        label="이름",
        formula=lambda ctx: (
            f"=INDEX({ctx.column_range('이름')},1)"
        ),
        number_format="0",
        start_row=2
    ),
    payroll_lookup_summary("통상시급", start_row=3, default=10320),
    SummaryFormula(
        label="소정근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('소정근무시간')})"
        ),
        number_format="0",
        start_row=5
    ),
    SummaryFormula(
        label="주휴인정시간",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('주휴인정시간')})"
        ),
        number_format="0",
        start_row=6
    ),
    SummaryFormula(
        label="평일연장근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('평일연장근무시간')})"
        ),
        number_format="0",
        start_row=7
    ),
    SummaryFormula(
        label="휴일근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('휴일근무시간')})"
        ),
        number_format="0",
        start_row=8
    ),
    SummaryFormula(
        label="휴일연장근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('휴일연장근무시간')})"
        ),
        number_format="0",
        start_row=9
    ),
    SummaryFormula(
        label="야간근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('야간근무시간')})"
        ),
        number_format="0",
        start_row=10
    ),
    SummaryFormula(
        label="공휴일인정시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('공휴일인정시간')})"
        ),
        number_format="0",
        start_row=11
    ),
    SummaryFormula(
        label="연차인정시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('연차인정시간')})"
        ),
        number_format="0",
        start_row=12
    ),

    SummaryFormula(
        label="기본급",
        formula=lambda ctx: (
            f"=( {ctx.summary_cell('소정근무시간합계')} + {ctx.summary_cell('주휴인정시간')} ) * {ctx.summary_cell('통상시급')}"
        ),
        number_format="#,##0",
        start_row=14,
    ),
    SummaryFormula(
        label="평일연장수당",
        formula=lambda ctx: (
            f"={ctx.summary_cell('평일연장근무시간합계')} * {ctx.summary_cell('통상시급')} * 1.5"
        ),
        number_format="#,##0",
        start_row=15,
    ),
    SummaryFormula(
        label="야간수당",
        formula=lambda ctx: (
            f"={ctx.summary_cell('야간근무시간합계')} * {ctx.summary_cell('통상시급')} * 0.5"
        ),
        number_format="#,##0",
        start_row=16,
    ),
    SummaryFormula(
        label="휴일수당",
        formula=lambda ctx: (
            f"=({ctx.summary_cell('휴일근무시간합계')} * {ctx.summary_cell('통상시급')} * 1.5) + ({ctx.summary_cell('휴일연장근무시간합계')} * {ctx.summary_cell('통상시급')} * 2.0)"
        ),
        number_format="#,##0",
        start_row=17,
    ),
    SummaryFormula(
        label="연차수당",
        formula=lambda ctx: (
            f"={ctx.summary_cell('연차인정시간합계')} * {ctx.summary_cell('통상시급')}"
        ),
        number_format="#,##0",
        start_row=18,
    ),
    SummaryFormula(
        label="공휴일수당",
        formula=lambda ctx: (
            f"={ctx.summary_cell('공휴일인정시간합계')} * {ctx.summary_cell('통상시급')}"
        ),
        number_format="#,##0",
        start_row=19,
    ),
    SummaryFormula(
        label="기타수당",
        formula=lambda ctx: (
            f"=0"
        ),
        number_format="#,##0",
        start_row=20,
    ),
    SummaryFormula(
        label="지급합계",
        formula=lambda ctx: (
            f"=SUM({ctx.summary_cell('기본급')},{ctx.summary_cell('평일연장수당')},{ctx.summary_cell('휴일수당')},{ctx.summary_cell('연차수당')},{ctx.summary_cell('공휴일수당')},{ctx.summary_cell('야간수당')},{ctx.summary_cell('기타수당')})"
        ),
        number_format="#,##0",
        start_row=21,
    ),
    *[
        # 급여대장에 입력된 공제 정보를 개인 명세의 공제 영역으로 가져온다.
        payroll_lookup_summary(label, start_row)
        for label, start_row in [
            ("고용보험", 23),
            ("고용보험정산", 24),
            ("국민연금", 25),
            ("건강보험", 26),
            ("건강보험정산", 27),
            ("장기요양", 28),
            ("장기요양정산", 29),
            ("환급금이자", 30),
            ("관리비", 31),
            ("식대비", 32),
            ("소득세", 33),
            ("지방소득세", 34),
            ("지방소득세정산", 35),
            ("기타공제", 36),
        ]
    ],
    SummaryFormula(
        label="공제합계",
        formula=lambda ctx: (
            f"=SUM({ctx.summary_cell('고용보험')},{ctx.summary_cell('고용보험정산')},{ctx.summary_cell('국민연금')},{ctx.summary_cell('건강보험')},{ctx.summary_cell('건강보험정산')},{ctx.summary_cell('장기요양')},{ctx.summary_cell('장기요양정산')},{ctx.summary_cell('환급금이자')},{ctx.summary_cell('관리비')},{ctx.summary_cell('식대비')},{ctx.summary_cell('소득세')},{ctx.summary_cell('지방소득세')},{ctx.summary_cell('지방소득세정산')},{ctx.summary_cell('기타공제')})"
        ),
        number_format="#,##0",
        start_row=37,
    ),
    SummaryFormula(
        label="차인지급액",
        formula=lambda ctx: (
            f"={ctx.summary_cell('지급합계')}-{ctx.summary_cell('공제합계')}"
        ),
        number_format="#,##0",
        start_row=39,
    )
]

# ---------------------------------------------------------------------------
# 급여대장: 개인 요약값 조회 열
#
# OverallFormula 자체는 값을 계산하는 모델이 아니라 급여대장 열의 정의다.
# 여기서는 사용자ID를 키로 개인 시트의 SummaryFormula 결과를 VLOOKUP한다.
# ---------------------------------------------------------------------------
overall_formula_columns = [
    OverallFormula(
        header="소정근무",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "소정근무시간합계",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="주휴",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "주휴인정시간",
            default=0,
        ),
        number_format="0",
    ),

    OverallFormula(
        header="평일연장",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "평일연장근무시간합계",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="휴일근무",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "휴일근무시간합계",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="휴일연장근무",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "휴일연장근무시간합계",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="야간근무",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "야간근무시간합계",
            default=0,
        ),
        number_format="0",
    ),

    OverallFormula(
        header="기본급",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "기본급",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="평일연장수당",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "평일연장수당",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="야간수당",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "야간수당",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="휴일수당",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "휴일수당",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="연차수당",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "연차수당",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="공휴일수당",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "공휴일수당",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="기타수당",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "기타수당",
            default=0,
        ),
        number_format="0",
    ),

    OverallFormula(
        header="지급합계",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "지급합계",
            default=0,
        ),
        number_format="0",
    ),
]


# ---------------------------------------------------------------------------
# 개인 시트: 검토가 필요한 행과 휴일 범위를 강조하는 조건부서식
# ---------------------------------------------------------------------------
# 이상값은 빨간 글씨, 공휴일·주말 행은 연한 빨간 배경으로 표시한다.
red_font = Font(
    color="FFFF0000",  # 순수 빨간색 (또는 엑셀 기본 진한 빨강: "FF9C0006")
    bold=True          # (선택) 굵게 표시하고 싶을 경우
)
red_fill = PatternFill(
    start_color="FFFFC7CE",
    end_color="FFFFC7CE",
    fill_type="solid"
)

personal_conditional_formats = [
    ColumnConditionalFormat(
        column="근무일자",
        # 같은 개인 시트 안에 근무일자가 중복된 경우 표시한다.
        rule=FormulaRule(
            formula=[
                'COUNTIF('
                'INDEX($2:$1000, 0, MATCH("근무일자", $1:$1, 0)),'
                'INDEX(2:2, 1, MATCH("근무일자", $1:$1, 0))'
                ') > 1'
            ],
            font=red_font
        )
    ),
    ColumnConditionalFormat(
        column="총근무시간",
        # 하루 총근무시간이 비정상적으로 긴 경우 표시한다.
        rule=CellIsRule(
            operator="greaterThan",
            formula=["22"],
            font=red_font
        ),
    ),
    ColumnConditionalFormat(
        column="총근무시간",
        # 출퇴근 역전 등으로 음수 시간이 계산된 경우 표시한다.
        rule=CellIsRule(
            operator="lessThan",
            formula=["-1"],
            font=red_font
        ),
    ),
    ColumnConditionalFormat(
        column="실근무시간",
        # 평일 야간근무가 허용 범위를 넘는 경우 표시한다.
        rule=FormulaRule(
            formula=[
                'AND('
                'INDEX(2:2,1,MATCH("휴일",$1:$1,0))=0,'
                'INDEX(2:2,1,MATCH("교대일",$1:$1,0))=0,'
                'INDEX(2:2,1,MATCH("야간근무시간",$1:$1,0))>=8,'
                'INDEX(2:2,1,MATCH("실근무시간",$1:$1,0))>=13'
                ')'
            ],
            font=red_font,
        ),
    ),
    ColumnConditionalFormat(
        column="실근무시간",
        # 평일 주간근무가 허용 범위를 넘는 경우 표시한다.
        rule=FormulaRule(
            formula=[
                'AND('
                'INDEX(2:2,1,MATCH("휴일",$1:$1,0))=0,'
                'INDEX(2:2,1,MATCH("교대일",$1:$1,0))=0,'
                'INDEX(2:2,1,MATCH("야간근무시간",$1:$1,0))=0,'
                'INDEX(2:2,1,MATCH("실근무시간",$1:$1,0))>=11'
                ')'
            ],
            font=red_font,
        ),
    ),
    ColumnConditionalFormat(
        column="주휴인정시간",
        # 일요일이 존재하지만 주휴시간이 인정되지 않은 주를 검토 대상으로 삼는다.
        rule=FormulaRule(
            formula=[
                'AND('
                'INDEX(2:2,1,MATCH("주휴인정시간",$1:$1,0))=0,'
                'INDEX(2:2,1,MATCH("요일",$1:$1,0))="일",'
                'COUNTIFS(OFFSET(INDEX(1:1,1,MATCH("요일",$1:$1,0)),1,0,ROW()-1,1),"일")=1'
                ')'
            ],
            font=red_font,
        ),
    ),
    ColumnConditionalFormat(
        column="근무일자",
        end_column="연차인정시간",
        # 공휴일과 토·일요일 행 전체를 배경색으로 구분한다.
        rule=FormulaRule(
            formula=[
                'OR('
                'INDEX(2:2, 1, MATCH("공휴일", $1:$1, 0))=1, '
                'INDEX(2:2, 1, MATCH("요일", $1:$1, 0))="토", '
                'INDEX(2:2, 1, MATCH("요일", $1:$1, 0))="일"'
                ')'
            ],
            fill=red_fill,
        ),
    )
]

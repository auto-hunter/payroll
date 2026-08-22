"""데이터 검증 및 전처리 설정."""

from exporter.excel_formulas import FormulaColumn, SummaryFormula, OverallFormula

TARGET_MONTH = "2026-07"

# 입력 엑셀 파일에 요구되는 컬럼명
INPUT_FILE_REQUIRED_COLUMNS = (
    "발생일자",
    "발생시각",
    "단말기ID",
    "사용자ID",
    "이름",
    "사원번호",
    "모드",
)

# 출력 컬럼명
WORK_DATE_COL = "근무일자"
WEEKDAY_COL = "요일"
START_COL = "출근시간"
END_COL = "퇴근시간"
ACTUAL_START_COL = "실출근시간"
ACTUAL_END_COL = "실퇴근시간"
NAME_COL = "이름"
COMPANY_COL = "등록사업장"

# 출근 시간 전처리 규칙
# 출근 기록과 퇴근 기록을 같은 근무 건으로 묶을 수 있는 최대 시간 간격
# 출근 시간과 퇴근시간이 24시간 이상 차이나면 같은 근무 건으로 묶지 않고, 출근 기록만 있는 근무 건으로 처리
MAX_WORK_HOURS = 24

# 컬럼명별 필터 규칙
# - exclude: 해당 값이면 제거
# - include: 해당 값일 때만 유지
# 단일 값/리스트/세트/튜플도 호환으로 허용
DATA_CLEANER_FILTER_ITEMS = {
    "이름": {"exclude": [""]},
    "모드": {"include": ["출근", "퇴근"]}
}

# 주말, 공휴일 및 교대일 설정
WEEKEND_WEEKDAYS = ["토", "일"]
HOLIDAY_DATES = ["2026-07-17"]
SHIFT_DATES = ["2026-07-05"]

SHEET_SPLIT_COL = NAME_COL

TAEIL_CABLE = [
    "마빈","나낭","김위","알존","이영련","우마르","노닐","아왈","리오","젤리빈","홍용표",
    "노베르토,""무하마드","렌델","김용범","도니","이라완","아리아","안드린","에르딘","누리스"
]
TAEIL_MATERIAL = ["심상복","최영일","지노","하디","다낭","존 폴","와유디","이반","아디"]

# excel 수식

def payroll_lookup_summary(label, start_row, default=0):
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
        formula=lambda ctx: (
            f"=MAX(0, MIN({ctx.cell('실퇴근시간')}, INT({ctx.cell('실출근시간')}) - (MOD({ctx.cell('실출근시간')}, 1) < 6/24) + 30/24) - MAX({ctx.cell('실출근시간')}, INT({ctx.cell('실출근시간')}) - (MOD({ctx.cell('실출근시간')}, 1) < 6/24) + 22/24)) * 24"
        ),
        number_format="0"
    ),
    FormulaColumn(
        header="주휴인정시간",
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
        formula=lambda ctx: f'=IF({ctx.cell("공휴일")} = 1, 8, 0)',
        number_format="0",
    )
]

personal_summary_formulas = [
    payroll_lookup_summary("통상시급", start_row=2, default=10320),
    SummaryFormula(
        label="소정근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('소정근무시간')})"
        ),
        number_format="0",
        start_row=4
    ),
    SummaryFormula(
        label="주휴인정시간",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('주휴인정시간')})"
        ),
        number_format="0",
        start_row=5
    ),
    SummaryFormula(
        label="평일연장근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('평일연장근무시간')})"
        ),
        number_format="0",
        start_row=6
    ),
    SummaryFormula(
        label="휴일근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('휴일근무시간')})"
        ),
        number_format="0",
        start_row=7
    ),
    SummaryFormula(
        label="휴일연장근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('휴일연장근무시간')})"
        ),
        number_format="0",
        start_row=8
    ),
    SummaryFormula(
        label="야간근무시간합계",
        formula=lambda ctx: (
            f"=SUM({ctx.column_range('야간근무시간')})"
        ),
        number_format="0",
        start_row=9
    ),
    SummaryFormula(
        label="기본급",
        formula=lambda ctx: (
            f"=( {ctx.summary_cell('소정근무시간합계')} + {ctx.summary_cell('주휴인정시간')} ) * {ctx.summary_cell('통상시급')}"
        ),
        number_format="#,##0",
        start_row=12,
    ),
    SummaryFormula(
        label="평일연장수당",
        formula=lambda ctx: (
            f"={ctx.summary_cell('평일연장근무시간합계')} * {ctx.summary_cell('통상시급')} * 1.5"
        ),
        number_format="#,##0",
        start_row=13,
    ),
    SummaryFormula(
        label="휴일근무수당",
        formula=lambda ctx: (
            f"={ctx.summary_cell('휴일근무시간합계')} * {ctx.summary_cell('통상시급')} * 1.5"
        ),
        number_format="#,##0",
        start_row=14,
    ),
    SummaryFormula(
        label="휴일연장수당",
        formula=lambda ctx: (
            f"={ctx.summary_cell('휴일연장근무시간합계')} * {ctx.summary_cell('통상시급')} * 2.0"
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
        label="기타수당",
        formula=lambda ctx: (
            f"=0"
        ),
        number_format="#,##0",
        start_row=17,
    ),
    SummaryFormula(
        label="지급합계",
        formula=lambda ctx: (
            f"=SUM({ctx.summary_cell('기본급')},{ctx.summary_cell('평일연장수당')},{ctx.summary_cell('휴일근무수당')},{ctx.summary_cell('휴일연장수당')},{ctx.summary_cell('야간수당')},{ctx.summary_cell('기타수당')})"
        ),
        number_format="#,##0",
        start_row=18,
    ),
    *[
        payroll_lookup_summary(label, start_row)
        for label, start_row in [
            ("고용보험", 20),
            ("고용보험정산", 21),
            ("국민연금", 22),
            ("건강보험", 23),
            ("건강보험정산", 24),
            ("장기요양", 25),
            ("장기요양정산", 26),
            ("환급금이자", 27),
            ("관리비", 28),
            ("식대비", 29),
            ("소득세", 30),
            ("지방소득세", 31),
            ("지방소득세정산", 32),
            ("기타공제", 33),
        ]
    ],
    SummaryFormula(
        label="공제합계",
        formula=lambda ctx: (
            f"=SUM({ctx.summary_cell('고용보험')},{ctx.summary_cell('고용보험정산')},{ctx.summary_cell('국민연금')},{ctx.summary_cell('건강보험')},{ctx.summary_cell('건강보험정산')},{ctx.summary_cell('장기요양')},{ctx.summary_cell('장기요양정산')},{ctx.summary_cell('환급금이자')},{ctx.summary_cell('관리비')},{ctx.summary_cell('식대비')},{ctx.summary_cell('소득세')},{ctx.summary_cell('지방소득세')},{ctx.summary_cell('지방소득세정산')},{ctx.summary_cell('기타공제')})"
        ),
        number_format="#,##0",
        start_row=34,
    ),
]

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
        header="휴일근무수당",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "휴일근무수당",
            default=0,
        ),
        number_format="0",
    ),
    OverallFormula(
        header="휴일연장수당",
        formula=lambda ctx: ctx.vlookup_summary(
            "사용자ID",
            "휴일연장수당",
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

import pandas as pd
df = pd.read_excel('caps_rawdata_2607_260805.xlsx')

from validator.input_validator import valid_input_file

from preprocessor.data_cleaner import (
    filter_rows, parse_commute_logs,
    add_weekday_column, add_holiday_shift_columns,
    adjust_commute_time_columns, fill_missing_dates,
    filter_work_records_by_target_month,
)

def payroll_lookup_summary(label, start_row, default=0):
    def build_formula(ctx):
        return (
            f"=IFERROR("
            f"INDEX('급여대장'!$C:$AD,"
            f"MATCH(INDEX({ctx.column_range('이름')},1),'급여대장'!$C:$C,0),"
            f"MATCH(\"{label}\","
            f"INDEX('급여대장'!$C:$AD,MATCH(\"이름\",'급여대장'!$C:$C,0),0),0)),"
            f"{default})"
        )

    return SummaryFormula(
        label=label,
        formula=build_formula,
        number_format="#,##0",
        start_row=start_row,
    )

from exporter.excel_formulas import FormulaColumn, SummaryFormula
from exporter.excel_writer import write_dataframe_by_name

if valid_input_file(df)[0]:
    print('검증성공')
else:
    print('검증실패')


df = filter_rows(df) # 불필요행 삭제
df = parse_commute_logs(df) # 출퇴근 로그 파싱
df = filter_work_records_by_target_month(df) # 대상 월 근무만 유지
df = add_weekday_column(df) # 출근시간 기준 요일 추가
df = adjust_commute_time_columns(df) # 실출근/퇴근 시간 계산
df = add_holiday_shift_columns(df) # 휴일, 교대일 추가
df = fill_missing_dates(df) # 결측 근무일 채우기

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

write_dataframe_by_name(
    df,
    "temp_output.xlsx",
    personal_formula_columns=personal_formula_columns,
    personal_summary_formulas=personal_summary_formulas,
)

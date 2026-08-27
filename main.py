import pandas as pd

from config.data_config import (
    INPUT_FILE_NAME,
    DEDUCT_FILE_NAME,
    DEDUCTION_COL,
    USER_ID_COL
)
from config.excel_config import (
    personal_conditional_formats,
    personal_formula_columns,
    personal_summary_formulas,
    overall_formula_columns,
    OUTPUT_FILE_NAME
)

from validator.input_validator import valid_input_file

from preprocessor.data_cleaner import (
    filter_rows, parse_commute_logs,
    add_weekday_column, add_holiday_shift_columns,
    adjust_commute_time_columns, fill_missing_dates,
    filter_work_records_by_target_month, add_company_column
)

from exporter.excel_writer import write_dataframe_by_name


# 입력 파일 읽기 (캡스)
df = pd.read_excel(INPUT_FILE_NAME)

if valid_input_file(df)[0]:
    print('검증성공')
    print(df.shape)
else:
    print('검증실패')

df = add_company_column(df) # 등록사업장(전선 or 소재) 열 추가
df = filter_rows(df) # 불필요행 삭제 (이름이 없는 행, 모드가 출근/퇴근이 아닌 행 등)
df = parse_commute_logs(df) # 출퇴근 로그 파싱
df = filter_work_records_by_target_month(df) # 급여계산에 필요한 대상 월 근무만 유지, 나머지는 제거
df = add_weekday_column(df) # 근무일자 기준 요일 열 추가
df = adjust_commute_time_columns(df) # 실출근/퇴근 시간 계산
df = add_holiday_shift_columns(df) # 휴일, 교대일 열 추가
df = fill_missing_dates(df) # 결측 근무일 채우기
print('전처리 완료')
print(df.shape)


# 공제 정보
df_deductions = pd.read_excel(DEDUCT_FILE_NAME)
df_deductions = df_deductions[[USER_ID_COL] + DEDUCTION_COL] # 필요한 컬럼만 유지
df_deductions[USER_ID_COL] = df_deductions[USER_ID_COL].astype(str).str.zfill(4) # 사용자ID를 4자리 문자열로 변환

# 엑셀 출력
write_dataframe_by_name(
    df,
    output_path=OUTPUT_FILE_NAME,
    personal_formula_columns=personal_formula_columns,
    personal_conditional_formats=personal_conditional_formats,
    personal_summary_formulas=personal_summary_formulas,
    overall_formula_columns=overall_formula_columns,
    user_id_col=USER_ID_COL,
    df_deductions = df_deductions,
)
print("엑셀 출력 완료")
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

from preprocessor.pipeline import preprocess_commute_logs

from exporter.excel_writer import write_dataframe_by_name


# 입력 파일 읽기 (캡스)
df_raw = pd.read_excel(INPUT_FILE_NAME)

if valid_input_file(df_raw)[0]:
    print('검증성공')
    print(df_raw.shape)
else:
    print('검증실패')

df_commute = preprocess_commute_logs(df_raw)
print('전처리 완료')
print(df_commute.shape)


# 공제 정보
df_deductions = pd.read_excel(DEDUCT_FILE_NAME)
df_deductions = df_deductions[[USER_ID_COL] + DEDUCTION_COL] # 필요한 컬럼만 유지
df_deductions[USER_ID_COL] = df_deductions[USER_ID_COL].astype(str).str.zfill(4) # 사용자ID를 4자리 문자열로 변환

# 엑셀 출력
write_dataframe_by_name(
    df_commute,
    output_path=OUTPUT_FILE_NAME,
    personal_formula_columns=personal_formula_columns,
    personal_conditional_formats=personal_conditional_formats,
    personal_summary_formulas=personal_summary_formulas,
    overall_formula_columns=overall_formula_columns,
    user_id_col=USER_ID_COL,
    df_deductions = df_deductions,
)
print("엑셀 출력 완료")

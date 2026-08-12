"""데이터 검증 및 전처리 설정."""

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

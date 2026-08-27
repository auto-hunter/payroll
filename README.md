# 태일전선 급여계산 자동화 프로그램

> 태일전선 급여계산 자동화 프로그램은 급여 계산을 자동화하여 효율성을 높이고 오류를 최소화하는 솔루션입니다.  
> 이 프로그램은 수동 작업을 자동화하여 급여 계산 프로세스를 간소화합니다.

------

# 주요 기능
1. **Human error 탐지**: 누락되거나 잘못 입력된 데이터를 자동으로 검증 
2. **출퇴근 로그 파싱 자동화**: 근태 기록 데이터를 자동으로 추출 및 정형화
3. **엑셀 수식 및 레이아웃 자동화**: 계산 수식 반영 및 지정 레이아웃으로 엑셀 파일 출력

# 사용 방법
```bash
...
```

------

# 운영/개발 참고 사항
![Version](https://img.shields.io/badge/version-v26.08.01-blue.svg)
![Python](https://img.shields.io/badge/Python-3.12+-3776AB?style=flat&logo=python&logoColor=white)

## Versioning
- **날짜 기반 버저닝 (Calendar Versioning)**
- 형식: `vYY.MM.Patch`
- 개발자 판단하에 2개 이상의 PR 이후에 버전을 올린다.
- 버전 업데이트 기준
  - 새로운 기능 추가
  - 에러 수정
- 버전 업데이트 하지 않는 케이스
  - 단순 주석, README 추가 변경
  - 그 외 매우 사소한 변경
  
## Branch Rule
- `feature/YYMMDD-NN/` 새로운 기능이나 자동화 로직 추가
- `fix/YYMMDD-NN/` 기존 기능의 버그나 계산 수식 오류 수정
- `docs/YYMMDD-NN/` README, 명세서, 주석 등 문서 수정
- `refactor/YYMMDD-NN/` 기능 변경 없는 코드 구조/성능 개선
- `YYMMDD-NN` 작업날짜, 일련번호 (예: 20260801-01)

## PR Rule
- 제목: Branch의 prefix를 대괄호로 감싸고 업데이트 명을 적는다. (예: `[feature] 출퇴근 로그 파싱 자동화`)
- 본문: 변경 사항 요약, 변경 이유, 테스트 방법 등 작성한다.
- 머지 전 스스로 코드리뷰 후 Merge 진행한다.

## Release Process
- **작업 시작**: master에서 새로운 작업 브랜치 생성 (예: `feature/log-parser` 또는 `fix/excel-bug`)
- **개발 및 커밋**: 개발, 커밋 진행 후 해당 브랜치 푸시
- **Self PR**: GitHub 웹으로 이동하여 master 방향으로 PR을 생성하고, 코드 변경점을 최종 자가 검수 후 Merge
- **태그/릴리스 생성** (GitHub Web):
  - 배포할 시점이 되면 GitHub 웹의 Releases 메뉴로 이동
  - `Draft a new release`를 누르고 새 버전에 맞는 태그(예: `v26.08.01`) 생성 및 변경 사항 요약 작성 후 Publish

  
## Architecture (WIP)

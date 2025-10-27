# 웹 계약서 비교 도구

이 도구는 https://harim.business.lawform.io 사이트에 로그인하여 체결 계약서를 조회하고, 엑셀 파일의 데이터와 비교하는 Python 스크립트입니다.

## 기능

- 웹사이트 자동 로그인
- 체결 계약서 조회 메뉴 자동 탐색
- 계약서 목록 및 상세 내용 자동 추출
- 엑셀 파일 데이터와 유사성 비교
- 비교 결과를 엑셀 파일로 저장

## 설치 방법

1. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

2. Chrome 브라우저가 설치되어 있어야 합니다.

## 사용 방법

1. 스크립트 실행:
```bash
python web_contract_comparator.py
```

2. 스크립트가 자동으로 다음 작업을 수행합니다:
   - https://harim.business.lawform.io 사이트에 로그인
   - 체결 계약서 조회 메뉴로 이동
   - 계약서 목록 추출
   - 각 계약서의 상세 내용 추출
   - 로아이_통합파일.xlsx와 비교
   - 결과를 contract_comparison_results_YYYYMMDD_HHMMSS.xlsx 파일로 저장

## 설정 변경

`main()` 함수에서 다음 설정을 변경할 수 있습니다:

- `username`: 로그인 ID
- `password`: 로그인 비밀번호  
- `excel_file_path`: 비교할 엑셀 파일 경로

## 주의사항

- 테스트를 위해 처음 5개 계약서만 처리합니다.
- 모든 계약서를 처리하려면 코드에서 `[:5]` 부분을 제거하세요.
- 웹사이트 구조가 변경되면 셀렉터를 수정해야 할 수 있습니다.
- 네트워크 상태에 따라 처리 시간이 달라질 수 있습니다.

## 출력 파일

비교 결과는 다음 형식으로 저장됩니다:
- 파일명: `contract_comparison_results_YYYYMMDD_HHMMSS.xlsx`
- 시트명: `비교결과`
- 컬럼: 웹 데이터, 엑셀 데이터, 유사성 점수, 매치 개수



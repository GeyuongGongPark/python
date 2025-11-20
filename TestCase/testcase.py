import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account
import re

# Google Sheets API 설정
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1kDt4dwKX3O6cggiYIlju8PoPwnSh8IphLtGg9Dy9qhg'
SHEET_NAME = '[삼성전자] 개인정보 문서 자동화 솔루션'

def get_google_sheets_service(credentials_file):
    """Google Sheets API 서비스 객체 생성"""
    creds = service_account.Credentials.from_service_account_file(
        credentials_file, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    return service

def read_sheet_data(service, spreadsheet_id, range_name):
    """시트에서 데이터 읽기"""
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    values = result.get('values', [])
    return values

def split_test_cases(data):
    """
    복합 기대결과를 가진 TC를 개별 TC로 분리
    
    Args:
        data: Google Sheets에서 읽은 2D 리스트
        
    Returns:
        modified_data: 수정된 원본 데이터
        new_test_cases: 새로 생성된 TC 리스트
    """
    # 헤더 찾기 (TC No. 행)
    header_row_idx = None
    for idx, row in enumerate(data):
        if row and 'TC No.' in row[0]:
            header_row_idx = idx
            break
    
    if header_row_idx is None:
        raise ValueError("헤더 행을 찾을 수 없습니다.")
    
    # 컬럼 인덱스 매핑
    headers = data[header_row_idx]
    col_mapping = {header: idx for idx, header in enumerate(headers)}
    
    # 필요한 컬럼 인덱스
    tc_no_idx = col_mapping.get('TC No.', 0)
    expected_result_idx = col_mapping.get('기대결과', 6)  # 기본값 6 (G열)
    
    modified_data = data.copy()
    new_test_cases = []
    
    # TC 데이터 시작 행부터 처리
    tc_start_idx = header_row_idx + 1
    
    # 마지막 TC 번호 추적
    last_tc_num = 0
    
    for idx in range(tc_start_idx, len(data)):
        row = data[idx]
        
        if not row or len(row) <= tc_no_idx:
            continue
            
        tc_no = row[tc_no_idx]
        
        # TC 번호 형식 확인 (TC_XXX)
        if not tc_no.startswith('TC_'):
            continue
        
        # 현재 TC 번호 추출
        try:
            tc_num = int(tc_no.split('_')[1])
            last_tc_num = max(last_tc_num, tc_num)
        except:
            continue
        
        # 기대결과 확인
        if len(row) > expected_result_idx:
            expected_result = row[expected_result_idx]
            
            # 쉼표로 구분된 여러 기대결과 확인
            if ',' in expected_result:
                # 기대결과 분리
                results = [r.strip() for r in expected_result.split(',')]
                
                # 첫 번째 결과는 원본 TC에 유지
                modified_data[idx][expected_result_idx] = results[0]
                
                # 나머지 결과들은 새로운 TC로 생성
                for additional_result in results[1:]:
                    last_tc_num += 1
                    new_tc_no = f"TC_{last_tc_num:03d}"
                    
                    # 새로운 TC 행 생성 (원본 TC의 모든 필드 복사)
                    new_row = row.copy()
                    new_row[tc_no_idx] = new_tc_no
                    new_row[expected_result_idx] = additional_result
                    
                    new_test_cases.append(new_row)
    
    return modified_data, new_test_cases

def write_to_sheet(service, spreadsheet_id, range_name, values):
    """시트에 데이터 쓰기"""
    body = {'values': values}
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    return result

def append_to_sheet(service, spreadsheet_id, range_name, values):
    """시트에 데이터 추가"""
    body = {'values': values}
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()
    return result

def main(credentials_file):
    """메인 실행 함수"""
    # 1. Google Sheets 서비스 초기화
    service = get_google_sheets_service(credentials_file)
    
    # 2. 현재 데이터 읽기
    print("데이터를 읽는 중...")
    range_name = f"{SHEET_NAME}!A1:I200"  # 충분히 큰 범위로 읽기
    data = read_sheet_data(service, SPREADSHEET_ID, range_name)
    
    print(f"총 {len(data)}개 행을 읽었습니다.")
    
    # 3. TC 분리 처리
    print("\n복합 기대결과를 분리하는 중...")
    modified_data, new_test_cases = split_test_cases(data)
    
    print(f"총 {len(new_test_cases)}개의 새로운 TC가 생성되었습니다.")
    
    # 4. 수정된 데이터를 시트에 업데이트
    print("\n시트를 업데이트하는 중...")
    write_to_sheet(service, SPREADSHEET_ID, range_name, modified_data)
    
    # 5. 새로운 TC들을 시트에 추가
    if new_test_cases:
        append_range = f"{SHEET_NAME}!A:I"
        append_to_sheet(service, SPREADSHEET_ID, append_range, new_test_cases)
    
    print("\n✅ 작업이 완료되었습니다!")
    print(f"   - 수정된 기존 TC: {len([tc for tc in modified_data if tc and tc[0].startswith('TC_')])}개")
    print(f"   - 새로 추가된 TC: {len(new_test_cases)}개")

# CSV 파일로 작업하는 경우 (Google Sheets API 없이)
def main_csv(input_csv, output_csv):
    """CSV 파일로 작업"""
    # CSV 읽기
    df = pd.read_csv(input_csv)
    
    new_rows = []
    last_tc_num = df['TC No.'].str.extract(r'TC_(\d+)')[0].astype(int).max()
    
    # 각 행 처리
    for idx, row in df.iterrows():
        expected_result = str(row['기대결과'])
        
        if ',' in expected_result:
            results = [r.strip() for r in expected_result.split(',')]
            
            # 첫 번째 결과로 원본 업데이트
            df.at[idx, '기대결과'] = results[0]
            
            # 나머지는 새 행으로 추가
            for additional_result in results[1:]:
                last_tc_num += 1
                new_row = row.copy()
                new_row['TC No.'] = f"TC_{last_tc_num:03d}"
                new_row['기대결과'] = additional_result
                new_rows.append(new_row)
    
    # 새로운 행 추가
    if new_rows:
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
    
    # CSV로 저장
    df.to_csv(output_csv, index=False, encoding='utf-8-sig')
    print(f"✅ 작업 완료! 결과가 {output_csv}에 저장되었습니다.")
    print(f"   총 TC 개수: {len(df)}")

if __name__ == "__main__":
    # Google Sheets API 사용
    # credentials_file = "path/to/your/credentials.json"
    # main(credentials_file)
    
    # CSV 파일 사용 (더 간단한 방법)
    main_csv('input.csv', 'output.csv')

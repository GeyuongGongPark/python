"""Excel의 모든 시트를 데이터(JSON)로 직렬화하는 스크립트.

개요
- 입력: Excel 파일 1개
- 처리: 각 시트를 B2 기준으로 헤더 재구성 후 행 단위 딕셔너리 리스트로 변환
- 출력: {"시트명": [{col: value, ...}, ...]} 형태의 JSON 파일

사용 예시
    python extract_sheets_to_json.py /path/to/input.xlsx
    python extract_sheets_to_json.py /path/to/input.xlsx /path/to/output.json
"""

import sys
import json
from pathlib import Path
from typing import Dict, Any

import pandas as pd


def frame_from_B2_as_header(df_raw: pd.DataFrame) -> pd.DataFrame:
    """주어진 원본 DataFrame(header=None 기반)을 B2를 헤더로 사용하도록 변환한다.
    - B2(행 index 1, 열 index 1)를 기준으로 잘라낸 뒤, 잘린 프레임의 첫 행을 컬럼으로 설정
    - 결과 데이터는 B3부터 시작
    - 공백/문자열 정리 적용
    비어있거나 충분한 영역이 없으면 빈 DataFrame 반환
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()
    # 최소 2행 2열 필요
    if df_raw.shape[0] < 2 or df_raw.shape[1] < 2:
        return pd.DataFrame()

    trimmed = df_raw.iloc[1:, 1:].reset_index(drop=True)
    if trimmed.empty:
        return pd.DataFrame()

    # 첫 행을 헤더로 사용
    new_header = trimmed.iloc[0].astype(str).map(lambda x: str(x).strip())
    df = trimmed.iloc[1:].copy()
    df.columns = [str(c).strip() for c in new_header]
    return df


def dataframe_to_dict_list(df: pd.DataFrame) -> list[Dict[str, Any]]:
    """DataFrame을 딕셔너리 리스트로 변환. NaN은 None으로 변환"""
    if df is None or df.empty:
        return []
    
    # NaN을 None으로 변환 (JSON 직렬화를 위해)
    df_cleaned = df.where(pd.notnull(df), None)
    
    # 각 행을 딕셔너리로 변환
    records = df_cleaned.to_dict('records')
    
    # 컬럼 이름은 문자열로 보장
    result = []
    for record in records:
        clean_record = {}
        for key, value in record.items():
            # 키를 문자열로 변환
            str_key = str(key).strip()
            # 값 처리: pandas의 NaN, NaT 등을 None으로 변환
            if pd.isna(value):
                clean_record[str_key] = None
            else:
                # 숫자 타입이면 적절히 변환
                if isinstance(value, (int, float)):
                    clean_record[str_key] = value
                else:
                    clean_record[str_key] = str(value)
        result.append(clean_record)
    
    return result


def extract_sheets_to_json(input_excel: Path, output_json: Path) -> None:
    """Excel 파일의 모든 시트를 읽어서 JSON으로 저장"""
    print(f"[시작] Excel 파일 읽는 중: {input_excel.name}")
    
    # 모든 시트를 헤더 없이 로드하여 B2를 헤더로 변환
    xls = pd.ExcelFile(input_excel)
    raw_sheets: Dict[str, pd.DataFrame] = {
        name: xls.parse(name, header=None) 
        for name in xls.sheet_names
    }
    
    sheets: Dict[str, pd.DataFrame] = {
        name: frame_from_B2_as_header(df) 
        for name, df in raw_sheets.items()
    }
    
    # 시트별 데이터를 JSON 형태로 변환
    result: Dict[str, list[Dict[str, Any]]] = {}
    
    for sheet_name, df in sheets.items():
        print(f"[처리] 시트 '{sheet_name}' 처리 중...")
        
        if df is None or df.empty:
            result[sheet_name] = []
            print(f"  - '{sheet_name}': 빈 시트 (데이터 없음)")
        else:
            data_list = dataframe_to_dict_list(df)
            result[sheet_name] = data_list
            print(f"  - '{sheet_name}': {len(data_list)}개 행, {len(df.columns)}개 컬럼")
            # 컬럼 목록 출력
            cols = [str(c) for c in df.columns]
            print(f"    컬럼: {', '.join(cols[:10])}" + (f" ... (총 {len(cols)}개)" if len(cols) > 10 else ""))
    
    # JSON 파일로 저장
    print(f"[저장] JSON 파일 저장 중: {output_json.name}")
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    # 통계 출력
    total_sheets = len(result)
    total_rows = sum(len(data) for data in result.values())
    print(f"[완료] {input_excel.name} → {output_json.name}")
    print(f"  - 총 {total_sheets}개 시트")
    print(f"  - 총 {total_rows:,}개 행")


def main(argv: list[str]) -> None:
    if len(argv) < 2:
        print("사용법: python extract_sheets_to_json.py \"/절대/경로/원본.xlsx\"")
        print("또는:   python extract_sheets_to_json.py \"/절대/경로/원본.xlsx\" \"/절대/경로/출력.json\"")
        sys.exit(1)
    
    input_path = Path(argv[1]).expanduser().resolve()
    if not input_path.exists():
        print(f"입력 파일을 찾을 수 없습니다: {input_path}")
        sys.exit(1)
    
    # 출력 경로: 두 번째 인자가 있으면 사용, 없으면 입력 파일 이름 기반으로 생성
    if len(argv) >= 3:
        output_path = Path(argv[2]).expanduser().resolve()
    else:
        output_path = input_path.with_suffix('.json')
    
    extract_sheets_to_json(input_path, output_path)


if __name__ == "__main__":
    main(sys.argv)


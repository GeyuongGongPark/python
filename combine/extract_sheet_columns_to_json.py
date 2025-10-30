import sys
import json
from pathlib import Path
from typing import Dict, List

import pandas as pd


def frame_from_B2_as_header(df_raw: pd.DataFrame) -> pd.DataFrame:
    """원본 DataFrame(header=None)을 B2 기준으로 잘라 첫 행을 컬럼으로 설정.
    B3부터 데이터가 시작하는 형태로 변환. 조건 미충족 시 빈 DataFrame.
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()
    if df_raw.shape[0] < 2 or df_raw.shape[1] < 2:
        return pd.DataFrame()

    trimmed = df_raw.iloc[1:, 1:].reset_index(drop=True)
    if trimmed.empty:
        return pd.DataFrame()

    new_header = trimmed.iloc[0].astype(str).map(lambda x: str(x).strip())
    df = trimmed.iloc[1:].copy()
    df.columns = [str(c).strip() for c in new_header]
    return df


def extract_sheet_columns_to_json(input_excel: Path, output_json: Path) -> None:
    """엑셀의 모든 시트에서 컬럼명만 추출하여 JSON으로 저장한다.
    결과 형식: { "시트명": ["컬럼1", "컬럼2", ...], ... }
    """
    print(f"[시작] Excel 파일 읽는 중: {input_excel.name}")

    xls = pd.ExcelFile(input_excel)
    # 각 시트를 header=None으로 읽고 B2 헤더 변환
    raw_sheets: Dict[str, pd.DataFrame] = {
        name: xls.parse(name, header=None) for name in xls.sheet_names
    }
    sheets: Dict[str, pd.DataFrame] = {
        name: frame_from_B2_as_header(df) for name, df in raw_sheets.items()
    }

    result: Dict[str, List[str]] = {}

    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            result[sheet_name] = []
            print(f" - {sheet_name}: (빈 시트)")
        else:
            cols = [str(c).strip() for c in df.columns]
            result[sheet_name] = cols
            preview = ", ".join(cols[:10])
            suffix = f" ... (총 {len(cols)}개)" if len(cols) > 10 else ""
            print(f" - {sheet_name}: {preview}{suffix}")

    print(f"[저장] JSON 파일 저장 중: {output_json.name}")
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"[완료] {input_excel.name} → {output_json.name}")
    print(f"  - 총 {len(result)}개 시트의 컬럼 정보를 저장했습니다.")


def main(argv: List[str]) -> None:
    if len(argv) < 2:
        print("사용법: python extract_sheet_columns_to_json.py \"/절대/경로/원본.xlsx\"")
        print("또는:   python extract_sheet_columns_to_json.py \"/절대/경로/원본.xlsx\" \"/절대/경로/출력.json\"")
        sys.exit(1)

    input_path = Path(argv[1]).expanduser().resolve()
    if not input_path.exists():
        print(f"입력 파일을 찾을 수 없습니다: {input_path}")
        sys.exit(1)

    if len(argv) >= 3:
        output_path = Path(argv[2]).expanduser().resolve()
    else:
        # 원본 파일명 기반 기본 출력 이름
        output_path = input_path.with_suffix(".columns.json")

    extract_sheet_columns_to_json(input_path, output_path)


if __name__ == "__main__":
    main(sys.argv)

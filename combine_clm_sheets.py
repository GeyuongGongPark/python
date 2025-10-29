import sys
from pathlib import Path
from typing import Dict, List

import pandas as pd


def _normalize_colname(name: str) -> str:
    return (
        str(name)
        .strip()
        .lower()
        .replace(" ", "")
        .replace(".", "")
        .replace("-", "")
        .replace("_", "")
    )


def resolve_column(df: pd.DataFrame, candidates: List[str]) -> str | None:
    # 정확 매칭
    for c in candidates:
        if c in df.columns:
            return c
    # 정규화 매칭
    norm_map = {_normalize_colname(c): c for c in df.columns}
    for c in candidates:
        key = _normalize_colname(c)
        if key in norm_map:
            return norm_map[key]
    return None


def resolve_sheet(xls: pd.ExcelFile, candidates: List[str]) -> str | None:
    # 정확 매칭
    for name in xls.sheet_names:
        if name in candidates:
            return name
    # 정규화 매칭
    norm_map = {_normalize_colname(n): n for n in xls.sheet_names}
    for c in candidates:
        key = _normalize_colname(c)
        if key in norm_map:
            return norm_map[key]
    return None


def read_all_sheets(input_path: Path) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(input_path)
    return {name: xls.parse(name) for name in xls.sheet_names}


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


def consolidate_by_no(input_excel: Path, output_excel: Path) -> None:
    xls = pd.ExcelFile(input_excel)

    # 베이스 시트: CLM등록(변형 표기 허용)
    base_sheet = resolve_sheet(xls, [
        "CLM등록",
        "CLM 등록",
        "CLM_REG",
        "등록",
    ])
    if base_sheet is None:
        raise ValueError("베이스 시트(예: 'CLM등록')를 찾지 못했습니다.")

    # 모든 시트를 미리 읽되, 각 시트를 헤더 없이 로드하여 B2를 헤더로 변환
    xls_all = pd.ExcelFile(input_excel)
    raw_sheets: Dict[str, pd.DataFrame] = {name: xls_all.parse(name, header=None) for name in xls_all.sheet_names}
    sheets: Dict[str, pd.DataFrame] = {name: frame_from_B2_as_header(df) for name, df in raw_sheets.items()}

    # 시트별 컬럼 출력 (정규화 후)
    print("[정보] 정규화된 각 시트의 컬럼 목록")
    for name, df_norm in sheets.items():
        try:
            if df_norm is None or df_norm.empty:
                print(f" - {name}: (빈 시트)")
            else:
                cols_preview = ", ".join([str(c) for c in df_norm.columns])
                print(f" - {name}: {cols_preview}")
        except Exception:
            print(f" - {name}: (컬럼 출력 중 오류)")

    # 베이스 시트 프레임
    df_base = sheets.get(base_sheet, pd.DataFrame())

    # print(f"[정보] 베이스 시트: {base_sheet}")
    # try:
    #     # 컬럼 출력: 삭제 적용된 베이스 데이터의 컬럼 출력
    #     col_list = [str(c) for c in df_base.columns]
    #     print("[컬럼] " + ", ".join(col_list))
    # except Exception:
    #     # 컬럼 출력이 실패하더라도 통합은 계속 진행
    #     pass
    
    # 베이스 키: NO.
    base_key = resolve_column(df_base, ["NO.", "No.", "NO", "No", "no", "번호"])  # 폭넓게 허용
    if base_key is None:
        raise KeyError("베이스 시트에서 키 컬럼(예: 'NO.')을 찾지 못했습니다.")

    # 요청: CLM등록(베이스 시트)의 컬럼 목록 출력

    # 병합 대상 시트: 베이스 제외 전체
    other_sheet_names = [n for n in sheets.keys() if n != base_sheet]

    merged = df_base.copy()
    merged[base_key] = merged[base_key].astype(str).str.strip()

    for sheet_name in other_sheet_names:
        df = sheets[sheet_name]
        if df is None or df.empty:
            continue

        # 대상 시트 키: NO.
        target_key = resolve_column(df, ["NO.", "No.", "NO", "No", "no", "번호"])  # 폭넓게 허용
        if target_key is None:
            # 키가 없으면 스킵 (경고성 로그)
            # print(f"스킵: '{sheet_name}' 시트에서 키 컬럼('NO.')을 찾지 못했습니다.")
            continue

        df = df.copy()
        df[target_key] = df[target_key].astype(str).str.strip()

        # 컬럼 프리픽스 (키 제외)
        non_key_cols = [c for c in df.columns if c != target_key]
        prefixed = df[[target_key] + non_key_cols].copy()
        rename_map = {c: f"{sheet_name}.{c}" for c in non_key_cols}
        prefixed = prefixed.rename(columns=rename_map)

        # 중복 열 이름 충돌 방지: 이미 존재하면 .1, .2 식으로 증가(간단 처리)
        for col in list(rename_map.values()):
            if col in merged.columns:
                i = 1
                new_col = f"{col}.{i}"
                while new_col in merged.columns:
                    i += 1
                    new_col = f"{col}.{i}"
                prefixed = prefixed.rename(columns={col: new_col})

        # Left join: 오른쪽 키를 인덱스로 사용해 왼쪽 키 컬럼이 _x로 바뀌지 않도록 한다
        prefixed_indexed = prefixed.set_index(target_key)
        merged = merged.merge(
            prefixed_indexed,
            left_on=base_key,
            right_index=True,
            how="left",
        )

    # 저장
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="통합", index=False)

    print(f"완료: {input_excel.name} → {output_excel.name} | 통합 행수: {len(merged):,}")


def main(argv: List[str]) -> None:
    if len(argv) < 2:
        print("사용법: python combine_clm_sheets.py \"/절대/경로/원본.xlsx\"")
        sys.exit(1)

    input_path = Path(argv[1]).expanduser().resolve()
    if not input_path.exists():
        print(f"입력 파일을 찾을 수 없습니다: {input_path}")
        sys.exit(1)

    output_path = input_path.with_name("CLM_통합.xlsx")
    consolidate_by_no(input_path, output_path)


if __name__ == "__main__":
    main(sys.argv)



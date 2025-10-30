import sys
from pathlib import Path
from typing import List, Optional

import pandas as pd


def read_excel_sheets(excel_path: Path) -> dict:
    """
    Read all required sheets into dataframes.
    Expected sheet names (Korean):
      - "CLM등록"
      - "CLM카테고리"
      - "상대계약자"
      - "인적정보등록"
      - "CLM계약처첨부파일"
    """
    xls = pd.ExcelFile(excel_path)

    required = [
        "CLM등록",
        "CLM카테고리",
        "상대계약자",
        "인적정보등록",
        "CLM계약처첨부파일",
    ]

    missing = [s for s in required if s not in xls.sheet_names]
    if missing:
        raise ValueError(f"필수 시트 누락: {', '.join(missing)}")

    sheets = {name: xls.parse(name) for name in required}
    return sheets


def _normalize_colname(name: str) -> str:
    """간단 정규화: 공백/점/하이픈 제거, 소문자화."""
    return (
        str(name)
        .strip()
        .lower()
        .replace(" ", "")
        .replace(".", "")
        .replace("-", "")
    )


def resolve_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """주어진 후보명들 중 실제 존재하는 컬럼을 찾는다. 정규화 매칭 포함."""
    # 1) 정확 일치
    for c in candidates:
        if c in df.columns:
            return c
    # 2) 정규화 일치
    norm_map = {_normalize_colname(c): c for c in df.columns}
    for c in candidates:
        key = _normalize_colname(c)
        if key in norm_map:
            return norm_map[key]
    return None


def duplicate_register_sheet(df_register: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of the CLM등록 sheet for processing (안전 복제)."""
    return df_register.copy(deep=True)


def map_category_codes(
    df_register: pd.DataFrame,
    df_category: pd.DataFrame,
    large_col: str = "대분류",
    mid_col: str = "분류",
    large_code_col: str = "대분류코드",
    mid_code_col: str = "분류코드",
) -> pd.DataFrame:
    """
    Step 2-3: 대분류/분류 코드값을 기준으로 CLM카테고리에서 표시명(혹은 추가 정보) 매핑.

    동작:
      - df_register의 대분류(코드), 분류(코드)를 기준으로
      - df_category에서 코드→이름(또는 상세) 매핑
      - 대분류, 분류 컬럼의 우측에 각각 1열씩 신규 컬럼 추가 (예: 대분류_명, 분류_명)

    Assumptions:
      - df_register has columns: 대분류, 분류 (값은 코드라고 가정)
      - df_category has columns: 대분류코드, 대분류, 분류코드, 분류
    """
    df = df_register.copy()

    # 준비: 코드→이름 매핑 딕셔너리 생성 (가능한 경우에 한해)
    large_map = None
    if {large_code_col, large_col}.issubset(df_category.columns):
        large_map = (
            df_category[[large_code_col, large_col]]
            .dropna()
            .drop_duplicates(subset=[large_code_col])
            .set_index(large_code_col)[large_col]
            .to_dict()
        )

    mid_map = None
    if {mid_code_col, mid_col}.issubset(df_category.columns):
        mid_map = (
            df_category[[mid_code_col, mid_col]]
            .dropna()
            .drop_duplicates(subset=[mid_code_col])
            .set_index(mid_code_col)[mid_col]
            .to_dict()
        )

    # 신규 컬럼명
    large_name_col = f"{large_col}_명"
    mid_name_col = f"{mid_col}_명"

    # 값 매핑 (df_register의 대분류/분류 값이 코드라고 가정)
    if large_map is not None and large_col in df.columns:
        df[large_name_col] = df[large_col].map(lambda x: large_map.get(str(x).strip(), None))
    else:
        df[large_name_col] = None

    if mid_map is not None and mid_col in df.columns:
        df[mid_name_col] = df[mid_col].map(lambda x: mid_map.get(str(x).strip(), None))
    else:
        df[mid_name_col] = None

    # 컬럼 위치: 각 원본 컬럼 우측에 삽입
    def insert_right_of(df_in: pd.DataFrame, anchor: str, new_col: str) -> pd.DataFrame:
        if anchor in df_in.columns:
            cols = list(df_in.columns)
            cols.remove(new_col)
            idx = cols.index(anchor) + 1
            cols.insert(idx, new_col)
            return df_in[cols]
        return df_in

    df = insert_right_of(df, large_col, large_name_col)
    df = insert_right_of(df, mid_col, mid_name_col)

    return df


def map_counterparty_person_info(
    df_register: pd.DataFrame,
    df_counterparty: pd.DataFrame,
    df_person: pd.DataFrame,
    clm_col: str = "CLM  NO.",
    person_key: str = "인적정보  NO.",
    explode_multiple: bool = False,
    person_prefix: str = "인적정보.",
) -> pd.DataFrame:
    """
    Steps 3 and 4:
      - Join counterparties by CLM  NO. to get person info keys
      - Join with person info table to fetch person details
      - If multiple person info numbers per CLM  NO., either aggregate (default) or explode rows

    Returns a new dataframe with person info columns added.
    """
    df = df_register.copy()

    # Normalize key columns
    for d in (df, df_counterparty, df_person):
        if clm_col in d.columns:
            d[clm_col] = d[clm_col].astype(str).str.strip()
        if person_key in d.columns:
            d[person_key] = d[person_key].astype(str).str.strip()

    # 상대계약자에서 (CLM  NO. -> 인적정보  NO.)를 연결
    # 하나의 CLM  NO.에 여러 인적정보  NO.가 있을 수 있음
    if not {clm_col, person_key}.issubset(df_counterparty.columns):
        # 필수 컬럼이 없으면 그대로 반환
        return df

    # 집계: CLM  NO. 별 인적정보  NO. 리스트
    agg = (
        df_counterparty[[clm_col, person_key]]
        .dropna(subset=[person_key])
        .groupby(clm_col)[person_key]
        .apply(lambda s: [v for v in s.astype(str) if v != ""])
        .reset_index(name="인적정보No_list")
    )

    df = df.merge(agg, on=clm_col, how="left")

    if explode_multiple:
        # 핵심: 동일 CLM  NO.에 여러 인적정보  NO.가 있는 경우 행 확장(explode)
        df = df.explode("인적정보No_list", ignore_index=True)
        df[person_key] = df["인적정보No_list"].fillna("")
        df.drop(columns=["인적정보No_list"], inplace=True)
    else:
        # 대안: 콤마로 합치기
        df[person_key] = df["인적정보No_list"].apply(
            lambda lst: ",".join(lst) if isinstance(lst, list) else ""
        )
        df.drop(columns=["인적정보No_list"], inplace=True)

    # 인적정보 상세 병합 (prefix 부여)
    if person_key in df_person.columns:
        person_cols = [c for c in df_person.columns if c != person_key]
        df_person_prefixed = df_person.copy()
        df_person_prefixed.columns = [
            person_key if c == person_key else f"{person_prefix}{c}"
            for c in df_person.columns
        ]
        df = df.merge(df_person_prefixed, on=person_key, how="left")

    # 계약 시작일 좌측에 인적정보  NO. 컬럼을 위치시키기
    start_col = "계약 시작일"
    if start_col in df.columns and person_key in df.columns:
        cols = list(df.columns)
        cols.remove(person_key)
        insert_idx = cols.index(start_col)
        cols.insert(insert_idx, person_key)
        df = df[cols]

    return df


def enrich_counterparty_with_person(
    df_counterparty: pd.DataFrame,
    df_person: pd.DataFrame,
    counterparty_person_key: str = "인적정보  NO.", # 상대계약자 시트의 인적정보 키 컬럼
    person_sheet_key: str = " NO.", # 인적정보등록 시트의 인적정보 키 컬럼  
    person_prefix: str = "인적정보.",
) -> pd.DataFrame:
    """
    상대계약자 시트의 `인적정보  NO.` 우측에 인적정보등록 시트의 매핑된 데이터를 붙여넣는다.

    매핑 기준: 인적정보  NO.(상대계약자) =  NO.(인적정보등록)
    결과: 인적정보등록의 컬럼들에 `인적정보.` prefix를 붙여서 병합하고,
         이 컬럼들을 상대계약자 내에서 `인적정보  NO.` 우측에 연속 삽입한다.
    """
    cnt = df_counterparty.copy()
    per = df_person.copy()

    # 키 컬럼 유연 탐지
    counterparty_key = resolve_column(
        cnt,
        [counterparty_person_key, "인적정보  NO.", "인적정보번호", "인적정보ID", "인적정보id"],
    )
    person_key = resolve_column(per, [person_sheet_key, "No", "no", "번호", "고유번호"])

    if counterparty_key is None:
        raise KeyError("상대계약자 시트에서 인적정보 키 컬럼을 찾지 못했습니다. (예: '인적정보  NO.')")
    if person_key is None:
        raise KeyError("인적정보등록 시트에서 키 컬럼을 찾지 못했습니다. (예: ' NO.')")

    cnt[counterparty_key] = cnt[counterparty_key].astype(str).str.strip()
    per[person_key] = per[person_key].astype(str).str.strip()

    # 인적정보 컬럼들 prefix 적용
    per_prefixed = per.copy()
    person_cols = [c for c in per_prefixed.columns if c != person_sheet_key]
    per_prefixed.columns = [
        counterparty_key if c == person_key else f"{person_prefix}{c}"
        for c in per_prefixed.columns
    ]

    merged = cnt.merge(per_prefixed, on=counterparty_key, how="left")

    # 컬럼 재배치: 인적정보.* 컬럼을 인적정보  NO. 우측에 연속 배치
    info_cols = [c for c in merged.columns if c.startswith(person_prefix)]
    if counterparty_key in merged.columns and info_cols:
        cols = list(merged.columns)
        # 먼저 인적정보.* 제거
        for c in info_cols:
            cols.remove(c)
        anchor_idx = cols.index(counterparty_key) + 1
        for offset, c in enumerate(info_cols):
            cols.insert(anchor_idx + offset, c)
        merged = merged[cols]

    return merged


def merge_person_into_register_left_of_start(
    df_register: pd.DataFrame,
    df_counterparty_enriched: pd.DataFrame,
    clm_col: str = "CLM  NO.",
    start_date_col: str = "계약 시작일",
) -> pd.DataFrame:
    """
    인적정보가 붙은 상대계약자 데이터를 CLM  NO. 기준으로 집계하여 CLM등록의 계약 시작일 좌측에 삽입.

    - 동일 CLM  NO.에 다수 인적정보가 있으면 각 인적정보 컬럼을 콤마로 병합하여 한 셀에 요약
    - 병합 대상 컬럼: 상대계약자 내 `인적정보.` prefix로 시작하는 모든 컬럼
    """
    df_reg = df_register.copy()
    dfe = df_counterparty_enriched.copy()

    # 키 컬럼 유연 탐지
    clm_key_reg = resolve_column(df_reg, [clm_col, "CLM NO.", "CLM No", "CLM번호", "CLMno"])
    clm_key_cnt = resolve_column(dfe, [clm_col, "CLM NO.", "CLM No", "CLM번호", "CLMno"])
    start_col = resolve_column(df_reg, [start_date_col, "계약시작일", "시작일", "계약 시작", "startdate"])

    if clm_key_reg is None or clm_key_cnt is None:
        return df_reg

    df_reg[clm_key_reg] = df_reg[clm_key_reg].astype(str).str.strip()
    dfe[clm_key_cnt] = dfe[clm_key_cnt].astype(str).str.strip()

    info_cols = [c for c in dfe.columns if c.startswith("인적정보.")]
    if not info_cols or clm_key_cnt not in dfe.columns:
        return df_reg

    # CLM  NO.별 인적정보 컬럼 콤마 집계
    grouped = (
        dfe[[clm_key_cnt] + info_cols]
        .groupby(clm_key_cnt)
        .agg(lambda s: ",".join(sorted({str(v).strip() for v in s.dropna().astype(str) if str(v).strip() != ""})))
        .reset_index()
    )

    df_merged = df_reg.merge(grouped, left_on=clm_key_reg, right_on=clm_key_cnt, how="left")

    # 계약 시작일 좌측으로 인적정보.* 컬럼 이동
    if start_col in df_merged.columns:
        cols = list(df_merged.columns)
        # 꺼내서 재삽입
        extracting = [c for c in info_cols if c in cols]
        for c in extracting:
            cols.remove(c)
        insert_idx = cols.index(start_col)
        for offset, c in enumerate(extracting):
            cols.insert(insert_idx + offset, c)
        df_merged = df_merged[cols]

    return df_merged


def map_attachments(
    df_register: pd.DataFrame,
    df_attachments: pd.DataFrame,
    clm_col: str = "CLM  NO.",
    attach_count_col: str = "첨부파일개수",
    first_file_col_name: str = "첫번째첨부파일명",
    possible_filename_cols: Optional[List[str]] = None,
) -> pd.DataFrame:
    """
    Step 5: CLM계약처첨부파일 시트의 문서 이름 목록을 CLM등록 최우측 열에 추가.
    기본 동작: 파일명(혹은 유사 컬럼)을 집계하여 콤마로 결합한 문자열을 '첨부문서목록' 컬럼으로 추가.
    """
    df = df_register.copy()

    if clm_col in df_attachments.columns:
        df_attachments = df_attachments.copy()
        df_attachments[clm_col] = df_attachments[clm_col].astype(str).str.strip()
        # 후보 파일명 컬럼들 자동 탐색
        if possible_filename_cols is None:
            candidates = [
                "파일명",
                "첨부파일명",
                "FileName",
                "file_name",
                "FILE_NAME",
            ]
        else:
            candidates = possible_filename_cols

        filename_col = next((c for c in candidates if c in df_attachments.columns), None)

        grp = df_attachments.groupby(clm_col)
        if filename_col is not None:
            attach_list = (
                grp[filename_col]
                .apply(lambda s: ",".join([str(v) for v in s.dropna().astype(str) if str(v).strip() != ""]))
                .rename("첨부문서목록")
                .reset_index()
            )
        else:
            # 파일명 컬럼이 없으면 개수를 문자열로 대체
            attach_list = grp.size().astype(str).rename("첨부문서목록").reset_index()

        df = df.merge(attach_list, on=clm_col, how="left")
    else:
        # 첨부파일 시트에 CLM  NO.가 없으면 빈 문자열
        df["첨부문서목록"] = ""

    # '첨부문서목록'을 최우측으로 이동
    if "첨부문서목록" in df.columns:
        cols = [c for c in df.columns if c != "첨부문서목록"] + ["첨부문서목록"]
        df = df[cols]

    return df


def process(
    input_excel: Path,
    output_excel: Path,
    explode_multiple_persons: bool = False,
) -> None:
    sheets = read_excel_sheets(input_excel)

    df_reg_raw = sheets["CLM등록"]
    df_cat = sheets["CLM카테고리"]
    df_cnt = sheets["상대계약자"]
    df_person = sheets["인적정보등록"]
    df_attach = sheets["CLM계약처첨부파일"]

    # 1) 복제
    df_reg = duplicate_register_sheet(df_reg_raw)

    # 2) 카테고리 코드 매핑
    df_reg = map_category_codes(df_reg, df_cat)

    # 3) 상대계약자에 인적정보등록 매핑 (인적정보  NO. 우측에 추가)
    df_cnt_enriched = enrich_counterparty_with_person(
        df_counterparty=df_cnt,
        df_person=df_person,
        counterparty_person_key="인적정보  NO.",
        person_sheet_key=" NO.",
        person_prefix="인적정보.",
    )

    # 4) 매핑된 인적정보를 CLM등록의 계약 시작일 좌측에 추가 (CLM  NO. 기준 집계)
    df_reg = merge_person_into_register_left_of_start(
        df_register=df_reg,
        df_counterparty_enriched=df_cnt_enriched,
        clm_col="CLM  NO.",
        start_date_col="계약 시작일",
    )

    # 5) 첨부파일 매핑 (마지막 열들로 추가됨)
    df_reg = map_attachments(df_reg, df_attach)

    # 저장: 새로운 워크북에 처리된 CLM등록만 포함
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df_reg.to_excel(writer, sheet_name="CLM등록(처리)", index=False)

    # 간단한 로그
    print(
        f"완료: {input_excel.name} → {output_excel.name}  | 결과 행수: {len(df_reg):,}"
    )


def main(argv: List[str]) -> None:
    if len(argv) < 2:
        print(
            "사용법: python clm_process.py \"/절대/경로/CLM 이관_하림_마이그레이션_애그리보텍_완.xlsx\" [--explode]"
        )
        print("  기본: 병합 모드 (동일 CLM  NO.의 여러 인적정보  NO.를 콤마로 합침)")
        print("  --explode 옵션: 행 확장 모드 (각 인적정보  NO.마다 별도 행 생성)")
        sys.exit(1)

    input_path = Path(argv[1]).expanduser().resolve()
    if not input_path.exists():
        print(f"입력 파일을 찾을 수 없습니다: {input_path}")
        sys.exit(1)

    explode = False
    if len(argv) >= 3 and argv[2] == "--explode":
        explode = True

    output_path = input_path.with_name("CLM_등록_처리결과.xlsx")
    process(input_path, output_path, explode_multiple_persons=explode)


if __name__ == "__main__":
    main(sys.argv)



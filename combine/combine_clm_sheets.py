"""CLM 다중 시트 병합 도구.

개요
- 입력: 개별/다중 Excel 파일(디렉터리 지정 시 내부 *.xlsx 일괄 처리)
- 처리:
1) 모든 시트를 헤더 없이 로드한 뒤 A1을 헤더로 재구성
2) 베이스 시트(예: CLM등록)를 기준으로 기타 시트의 컬럼을 프리픽스 붙여 병합
3) 특수 시트 처리(예: 첨부파일 → 파일명만 키 기준으로 콤마 결합)
4) 인적정보등록에서 기업명(법인명) 매핑하여 상대계약자에 요약 컬럼 추가
5) 충돌/중복 열 이름 정리 후 결과 저장
- 출력: "통합" 시트를 가진 결과 워크북(파일명 규칙에 따라 done 폴더에 저장)

사용 예시
    # 디렉터리 내 모든 .xlsx 처리
    python combine_clm_sheets.py /path/to/folder

    # 단일 파일 처리
    python combine_clm_sheets.py /path/to/file.xlsx
"""

import sys
import re
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


def resolve_key_with_fallback(df: pd.DataFrame) -> str | None:
    """기본 키 'NO.' 우선, 없으면 'CLM NO.' 류로 대체."""
    key = resolve_column(df, [
        "NO.", "No.", "NO", "No", "no", "번호",
    ])
    if key is not None:
        return key
    return resolve_column(df, [
        "CLM NO.", "CLM NO", "CLMNO.", "CLMNO", "CLM no", "CLM번호",
        "clm no", "clmno",
    ])


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


def frame_from_A1_as_header(df_raw: pd.DataFrame) -> pd.DataFrame:
    """주어진 원본 DataFrame(header=None 기반)을 A1을 헤더로 사용하도록 변환한다.
    - A1(행 index 0, 열 index 0)을 헤더로 사용
    - 결과 데이터는 A2(행 index 1)부터 시작
    - 공백/문자열 정리 적용
    비어있거나 충분한 영역이 없으면 빈 DataFrame 반환
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()
    # 최소 2행 필요
    if df_raw.shape[0] < 2:
        return pd.DataFrame()

    # 첫 행을 헤더로 사용
    new_header = df_raw.iloc[0].astype(str).map(lambda x: str(x).strip())
    df = df_raw.iloc[1:].copy().reset_index(drop=True)
    df.columns = [str(c).strip() for c in new_header]
    return df


def consolidate_by_no(input_excel: Path, output_excel: Path) -> None:
    xls = pd.ExcelFile(input_excel)

    # 베이스 시트: CLM
    base_sheet = resolve_sheet(xls, [
        "CLM",
    ])
    if base_sheet is None:
        raise ValueError("베이스 시트(예: 'CLM')를 찾지 못했습니다.")

    # 모든 시트를 미리 읽되, 각 시트를 헤더 없이 로드하여 A1을 헤더로 변환
    xls_all = pd.ExcelFile(input_excel)
    raw_sheets: Dict[str, pd.DataFrame] = {name: xls_all.parse(name, header=None) for name in xls_all.sheet_names}
    sheets: Dict[str, pd.DataFrame] = {name: frame_from_A1_as_header(df) for name, df in raw_sheets.items()}

    # 특수 처리: 상대계약자 시트에 인적정보등록의 "기업명(법인명)" 매핑 및 CLM NO. 기준 콤마 병합
    rel_sheet_name = resolve_sheet(xls, [
        "CLM_CUSTOMER",
    ])
    person_sheet_name = resolve_sheet(xls, [
        "CLM_USER_CONTACT",
    ])
    if rel_sheet_name and person_sheet_name:
        df_rel = sheets.get(rel_sheet_name)
        df_person = sheets.get(person_sheet_name)
        try:
            if df_rel is not None and not df_rel.empty and df_person is not None and not df_person.empty:
                person_no_col = resolve_column(df_rel, [
                    "인적정보 NO.", "인적정보 NO", "인적정보NO.", "인적정보NO", "인적정보 no", "인적정보번호",
                ])
                rel_clm_no_col = resolve_column(df_rel, [
                    "CLM NO.", "CLM NO", "CLMNO.", "CLMNO", "CLM no", "CLM번호",
                ])
                person_key_col = resolve_column(df_person, [
                    "NO.", "No.", "NO", "No", "no", "번호",
                ])
                company_col = resolve_column(df_person, [
                    "기업명(법인명)", "기업명", "법인명", "기업명(법인 명)",
                ])

                if person_no_col and rel_clm_no_col and person_key_col and company_col:
                    # 매핑: 인적정보 NO. -> 기업명(법인명)
                    person_map = (
                        df_person[[person_key_col, company_col]]
                        .dropna(subset=[person_key_col])
                        .assign(**{person_key_col: lambda d: d[person_key_col].astype(str).str.strip()})
                        .set_index(person_key_col)[company_col]
                        .to_dict()
                    )

                    df_rel_local = df_rel.copy()
                    df_rel_local[person_no_col] = df_rel_local[person_no_col].astype(str).str.strip()
                    df_rel_local[rel_clm_no_col] = df_rel_local[rel_clm_no_col].astype(str).str.strip()

                    comp_col_name = "기업명(법인명)"
                    df_rel_local[comp_col_name] = df_rel_local[person_no_col].map(person_map)

                    # CLM NO. 기준으로 기업명(법인명)만 콤마 병합하여 추가 컬럼으로 생성
                    def _join_unique(series: pd.Series) -> str:
                        vals = [str(v).strip() for v in series.dropna().tolist() if str(v).strip() and str(v).lower() != "nan"]
                        return ", ".join(sorted(set(vals))) if vals else ""

                    # CLM NO. 기준으로 기업명(법인명) 병합
                    rel_comp_agg = (
                        df_rel_local.groupby(rel_clm_no_col, dropna=False)[comp_col_name]
                        .apply(_join_unique)
                        .reset_index()
                        .rename(columns={comp_col_name: f"{comp_col_name}_병합"})
                    )
                    
                    # 원본 상대계약자 시트와 병합된 기업명(법인명)을 합침
                    df_rel_final = df_rel_local.merge(
                        rel_comp_agg,
                        left_on=rel_clm_no_col,
                        right_on=rel_clm_no_col,
                        how="left"
                    )
                    
                    # 기업명(법인명)_병합을 기업명(법인명)으로 덮어쓰기 (요구사항 2-3: 한 셀에 콤마로 구분)
                    df_rel_final[comp_col_name] = df_rel_final[f"{comp_col_name}_병합"]
                    df_rel_final = df_rel_final.drop(columns=[f"{comp_col_name}_병합"])

                    # 상대계약자 시트 교체: 모든 컬럼 유지하되 기업명(법인명) 추가됨
                    sheets[rel_sheet_name] = df_rel_final
        except Exception:
            # 전처리 실패 시 기본 흐름으로 진행
            pass

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

    # 병합 전 단계에서 CLM등록의 NO. 기준 중복 제거(요청): 선제적으로 깨끗한 베이스 확보
    try:
        df_base = df_base.copy()
        df_base[base_key] = df_base[base_key].astype(str).str.strip()
        df_base = df_base.drop_duplicates(subset=[base_key], keep="first")
    except Exception:
        pass

    # CLM카테고리 매핑: 대분류/분류 컬럼에 카테고리이름 매핑 (요구사항 1)
    category_sheet_name = resolve_sheet(xls, [
        "CLM_CATEGORY",
    ])
    if category_sheet_name:
        df_category = sheets.get(category_sheet_name)
        if df_category is not None and not df_category.empty:
            try:
                # CLM카테고리 시트의 키: 보통 카테고리 코드나 NO. 컬럼
                category_key_col = resolve_column(df_category, [
                    "NO.", "No.", "NO", "No", "no", "번호", "코드", "카테고리코드",
                ])
                category_name_col = resolve_column(df_category, [
                    "카테고리이름", "카테고리 이름", "이름", "카테고리명",
                ])

                if category_key_col and category_name_col:
                    # 매핑 딕셔너리 생성
                    category_map = (
                        df_category[[category_key_col, category_name_col]]
                        .dropna(subset=[category_key_col])
                        .assign(**{category_key_col: lambda d: d[category_key_col].astype(str).str.strip()})
                        .set_index(category_key_col)[category_name_col]
                        .to_dict()
                    )

                    # 대분류 컬럼 확인 및 매핑
                    major_category_col = resolve_column(df_base, [
                        "대분류", "대 분류",
                    ])
                    if major_category_col:
                        df_base[major_category_col] = df_base[major_category_col].astype(str).str.strip()
                        # 대분류 우측에 카테고리이름 추가
                        major_cols = list(df_base.columns)
                        major_idx = major_cols.index(major_category_col)
                        major_category_name_col = f"{major_category_col}_카테고리이름"
                        df_base.insert(
                            major_idx + 1,
                            major_category_name_col,
                            df_base[major_category_col].map(category_map)
                        )

                    # 분류 컬럼 확인 및 매핑
                    sub_category_col = resolve_column(df_base, [
                        "분류", "소분류",
                    ])
                    if sub_category_col:
                        df_base[sub_category_col] = df_base[sub_category_col].astype(str).str.strip()
                        # 분류 우측에 카테고리이름 추가
                        sub_cols = list(df_base.columns)
                        sub_idx = sub_cols.index(sub_category_col)
                        sub_category_name_col = f"{sub_category_col}_카테고리이름"
                        df_base.insert(
                            sub_idx + 1,
                            sub_category_name_col,
                            df_base[sub_category_col].map(category_map)
                        )
            except Exception:
                # 매핑 실패 시 경고 없이 진행
                pass

    # 요청: CLM등록(베이스 시트)의 컬럼 목록 출력

    # 병합 대상 시트: 베이스 시트와 _기존이 붙은 모든 시트 제외 (요구사항 4)
    exclude_sheet_names = [base_sheet]
    # _기존이 포함된 시트 찾기
    for sheet_name in sheets.keys():
        if "_기존" in sheet_name or " 기존" in sheet_name:
            exclude_sheet_names.append(sheet_name)
    
    # CLM카테고리 시트 제외: 매핑용으로만 사용, 직접 병합하지 않음 (요구사항 1-3)
    if category_sheet_name:
        exclude_sheet_names.append(category_sheet_name)
    
    # 인적정보등록 시트 제외: 상대계약자를 통해서만 사용, 직접 병합하지 않음 (요구사항 2-2)
    if person_sheet_name:
        exclude_sheet_names.append(person_sheet_name)
    
    # 정규화 매칭으로도 제외 시트 찾기
    exclude_normalized = {_normalize_colname(n) for n in exclude_sheet_names}
    other_sheet_names = [
        n for n in sheets.keys()
        if n not in exclude_sheet_names
        and _normalize_colname(n) not in exclude_normalized
    ]

    merged = df_base.copy()
    merged[base_key] = merged[base_key].astype(str).str.strip()

    for sheet_name in other_sheet_names:
        df = sheets[sheet_name]
        if df is None or df.empty:
            continue

        # 대상 시트 키: 기본 NO. → 대체 'CLM NO.' 허용(요구사항 1-1)
        target_key = resolve_key_with_fallback(df)
        if target_key is None:
            # 키가 없으면 스킵
            continue

        df = df.copy()
        df[target_key] = df[target_key].astype(str).str.strip()

        # 특수 처리: 첨부파일 시트는 파일명 컬럼만 가져오기
        # 새 명칭만 허용: CLM_FILE, CLMATTACHMENT
        if (
            _normalize_colname(sheet_name) in {"clmfile", "clmattachment"}
            or sheet_name in {"CLM_FILE", "CLMATTACHMENT"}
        ):
            filename_col = resolve_column(df, [
                "파일명", "파일 이름", "파일명칭",
            ])
            if filename_col:
                # 키와 파일명만 유지
                df = df[[target_key, filename_col]].copy()
                # 같은 키 내 다건 파일을 콤마로 합쳐 1행으로 축약해 베이스 중복 방지
                def _join_files(series: pd.Series) -> str:
                    vals = [str(v).strip() for v in series.dropna().tolist() if str(v).strip() and str(v).lower() != "nan"]
                    return ", ".join(sorted(set(vals))) if vals else ""
                df = (
                    df.groupby(target_key, dropna=False)[filename_col]
                    .apply(_join_files)
                    .reset_index()
                )
            else:
                # 파일명 컬럼이 없으면 스킵
                continue

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

    # 병합 후 키 컬럼(base_key)의 중복 변형 정리: NO._x, NO._y, NO..1 등 제거
    try:
        cols_to_drop: List[str] = []
        for col in list(merged.columns):
            if col == base_key:
                continue
            # pandas merge 잔재(_x/_y) 제거
            if col == f"{base_key}_x" or col == f"{base_key}_y":
                cols_to_drop.append(col)
                continue
            # 충돌 회피로 생성된 숫자 접미사 형태(NO..1, NO..2 등) 제거
            if col.startswith(f"{base_key}."):
                suffix = col[len(base_key) + 1:]
                if suffix.isdigit():
                    cols_to_drop.append(col)
                    continue
        if cols_to_drop:
            merged = merged.drop(columns=cols_to_drop)
    except Exception:
        # 정리 실패 시에도 저장은 진행
        pass

    # 최종 단계에서 베이스 행 중복 제거는 수행하지 않음(상세 정보 보존)

    # 저장
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="통합", index=False)

    print(f"완료: {input_excel.name} → {output_excel.name} | 통합 행수: {len(merged):,}")


def _make_output_name_from_input(input_file: Path) -> str:
    stem = input_file.stem
    pattern = r"[_\s]*이관\s*_?\s*하림\s*_?\s*마이그레이션[_\s]*"
    cleaned = re.sub(pattern, " ", stem)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = re.sub(r"\s*_\s*", "_", cleaned)
    cleaned = re.sub(r"_+", "_", cleaned)
    cleaned = cleaned.strip(" _")
    final_name = (cleaned if cleaned else "CLM") + "_통합.xlsx"
    return final_name


def main(argv: List[str]) -> None:
    # 인자가 없으면 현재 파일 위치(combine 폴더)를 기본 대상으로 처리
    target_arg = Path(argv[1]).expanduser().resolve() if len(argv) >= 2 else Path.cwd()
    if not target_arg.exists():
        print(f"입력 경로를 찾을 수 없습니다: {target_arg}")
        sys.exit(1)

    # 입력 경로가 파일이면 그 파일이 있는 폴더를, 디렉터리면 그대로 사용
    if target_arg.is_file():
        input_dir = target_arg.parent
    else:
        input_dir = target_arg

    # 동일 폴더 내 모든 .xlsx 파일 처리 (done 하위 폴더에 저장)
    done_dir = input_dir / "done"
    done_dir.mkdir(parents=True, exist_ok=True)

    excel_files = sorted([p for p in input_dir.glob("*.xlsx") if p.is_file() and p.parent == input_dir])
    if not excel_files:
        print(f"📁 '{input_dir}'에서 .xlsx 파일을 찾을 수 없습니다.")
        sys.exit(1)

    print(f"🚀 다중 처리 시작: {len(excel_files)}개 파일")
    for f in excel_files:
        try:
            out_name = _make_output_name_from_input(f)
            output_path = done_dir / out_name
            consolidate_by_no(f, output_path)
        except Exception as e:
            print(f"❌ 실패: {f.name} → {e}")
    print(f"✅ 완료: 결과는 '{done_dir}' 폴더에 저장되었습니다.")


if __name__ == "__main__":
    main(sys.argv)



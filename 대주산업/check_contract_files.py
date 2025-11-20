import os
import pandas as pd
from pathlib import Path
from collections import defaultdict

def get_actual_files(base_path):
    """모두싸인 폴더 내의 모든 파일명을 수집합니다 (감사 추적 인증서 제외)."""
    base_dir = Path(base_path)
    
    영업관리팀_path = base_dir / "모두싸인_ 8월 4일 기준 전자계약서" / "영업관리팀"
    인사팀_path = base_dir / "모두싸인_ 8월 4일 기준 전자계약서" / "인사팀"
    
    actual_files = set()
    
    # 영업관리팀 파일 수집
    if 영업관리팀_path.exists():
        for root, dirs, files in os.walk(영업관리팀_path):
            for file in files:
                # 감사 추적 인증서 제외
                if "감사 추적 인증서" in file or "감사 추적 인증서.pdf" in file:
                    continue
                file_path = Path(root) / file
                relative_path = file_path.relative_to(영업관리팀_path)
                actual_files.add(str(relative_path))
    
    # 인사팀 파일 수집
    if 인사팀_path.exists():
        for root, dirs, files in os.walk(인사팀_path):
            for file in files:
                # 감사 추적 인증서 제외
                if "감사 추적 인증서" in file or "감사 추적 인증서.pdf" in file:
                    continue
                file_path = Path(root) / file
                relative_path = file_path.relative_to(인사팀_path)
                actual_files.add(str(relative_path))
    
    return actual_files

def get_excel_contract_files(excel_path):
    """Excel 파일에서 계약서 파일명을 추출합니다."""
    try:
        df = pd.read_excel(excel_path, header=1)
        
        # 계약서 파일 명 컬럼 찾기
        contract_file_col = '계약서 파일 명'
        
        if contract_file_col not in df.columns:
            print(f"경고: '{contract_file_col}' 컬럼을 찾을 수 없습니다.")
            print(f"사용 가능한 컬럼: {df.columns.tolist()}")
            return set()
        
        # NaN이 아닌 계약서 파일명 추출
        contract_files = set()
        for idx, row in df.iterrows():
            file_name = row[contract_file_col]
            if pd.notna(file_name) and str(file_name).strip():
                file_name = str(file_name).strip()
                # .pdf 확장자가 없으면 추가
                if not file_name.endswith('.pdf'):
                    file_name += '.pdf'
                contract_files.add(file_name)
        
        return contract_files
    except Exception as e:
        print(f"Excel 파일 읽기 오류: {e}")
        return set()

def normalize_filename(filename):
    """파일명을 정규화합니다 (공백, 경로 구분자 등 처리)."""
    # 경로 구분자를 통일
    filename = filename.replace('\\', '/')
    # 앞뒤 공백 제거
    filename = filename.strip()
    return filename

def extract_filename_only(filepath):
    """경로에서 파일명만 추출합니다."""
    return Path(filepath).name

def compare_files(excel_files, actual_files):
    """Excel의 파일명과 실제 폴더의 파일명을 비교합니다."""
    # Excel 파일명 정규화 (파일명만)
    excel_files_normalized = {normalize_filename(f) for f in excel_files}
    
    # 실제 폴더 파일명 정규화 (경로에서 파일명만 추출)
    actual_files_normalized = {normalize_filename(extract_filename_only(f)) for f in actual_files}
    
    # Excel에만 있는 파일 (실제 폴더에 없는 파일)
    only_in_excel = excel_files_normalized - actual_files_normalized
    
    # 실제 폴더에만 있는 파일 (Excel에 없는 파일)
    # 실제 경로 정보를 유지하기 위해 원본 경로를 반환
    only_in_actual_paths = []
    actual_filenames = {normalize_filename(extract_filename_only(f)): f for f in actual_files}
    for filename in actual_files_normalized:
        if filename not in excel_files_normalized:
            only_in_actual_paths.append(actual_filenames[filename])
    
    # 일치하는 파일
    matched = excel_files_normalized & actual_files_normalized
    
    return {
        'matched': matched,
        'only_in_excel': only_in_excel,
        'only_in_actual': only_in_actual_paths,
        'excel_count': len(excel_files_normalized),
        'actual_count': len(actual_files_normalized)
    }

def print_section(title, items):
    """공통 출력 형식"""
    print(f"\n{'='*80}")
    print(f"【{title}】: {len(items)}개")
    print(f"{'='*80}")
    for i, filename in enumerate(sorted(items), 1):
        print(f"{i:4d}. {filename}")

def print_comparison_result(result):
    """비교 결과를 출력합니다."""
    print(f"\n{'='*80}")
    print("【계약서 파일명 비교 결과】")
    print(f"{'='*80}\n")
    
    print(f"Excel 계약서리스트 파일 수: {result['excel_count']}개")
    print(f"실제 폴더 파일 수 (감사 추적 인증서 제외): {result['actual_count']}개")
    print(f"일치하는 파일 수: {len(result['matched'])}개\n")
    
    if result['only_in_excel']:
        print_section("Excel에만 있는 파일 (실제 폴더에 없음)", result['only_in_excel'])
    
    if result['only_in_actual']:
        print_section("실제 폴더에만 있는 파일 (Excel에 없음)", result['only_in_actual'])
    
    if not result['only_in_excel'] and not result['only_in_actual']:
        print("\n✅ 모든 파일이 일치합니다!")

def save_comparison_result(result, output_path):
    """비교 결과를 파일로 저장합니다."""
    lines = []
    lines.append("=" * 80)
    lines.append("【계약서 파일명 비교 결과】")
    lines.append("=" * 80)
    lines.append("")
    lines.append(f"Excel 계약서리스트 파일 수: {result['excel_count']}개")
    lines.append(f"실제 폴더 파일 수 (감사 추적 인증서 제외): {result['actual_count']}개")
    lines.append(f"일치하는 파일 수: {len(result['matched'])}개")
    lines.append("")
    
    if result['only_in_excel']:
        lines.append("=" * 80)
        lines.append(f"【Excel에만 있는 파일 (실제 폴더에 없음)】: {len(result['only_in_excel'])}개")
        lines.append("=" * 80)
        for i, filename in enumerate(sorted(result['only_in_excel']), 1):
            lines.append(f"{i:4d}. {filename}")
        lines.append("")
    
    if result['only_in_actual']:
        lines.append("=" * 80)
        lines.append(f"【실제 폴더에만 있는 파일 (Excel에 없음)】: {len(result['only_in_actual'])}개")
        lines.append("=" * 80)
        for i, filename in enumerate(sorted(result['only_in_actual']), 1):
            lines.append(f"{i:4d}. {filename}")
        lines.append("")
    
    if not result['only_in_excel'] and not result['only_in_actual']:
        lines.append("✅ 모든 파일이 일치합니다!")
    
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines), encoding="utf-8")

def save_comparison_excel(result, output_path):
    """비교 결과를 Excel 파일로 저장합니다."""
    summary_df = pd.DataFrame(
        [
            ("Excel 계약서리스트 파일 수", result["excel_count"]),
            ("실제 폴더 파일 수 (감사 추적 인증서 제외)", result["actual_count"]),
            ("일치하는 파일 수", len(result["matched"])),
            ("Excel에만 있는 파일 수", len(result["only_in_excel"])),
            ("실제 폴더에만 있는 파일 수", len(result["only_in_actual"])),
        ],
        columns=["항목", "값"],
    )
    
    only_excel_df = pd.DataFrame(sorted(result["only_in_excel"]), columns=["파일명"])
    only_actual_df = pd.DataFrame(sorted(result["only_in_actual"]), columns=["파일 경로"])
    matched_df = pd.DataFrame(sorted(result["matched"]), columns=["파일명"])
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        only_excel_df.to_excel(writer, sheet_name="OnlyInExcel", index=False)
        only_actual_df.to_excel(writer, sheet_name="OnlyInActual", index=False)
        matched_df.to_excel(writer, sheet_name="Matched", index=False)

def main():
    # 현재 스크립트가 있는 디렉토리를 기준으로 설정
    base_path = Path(__file__).parent
    excel_path = base_path / "계약서리스트_양식_대주산업_모두싸인.xlsx"
    output_path = base_path / "comparison_result.txt"
    output_excel_path = base_path / "comparison_result.xlsx"
    
    if not excel_path.exists():
        print(f"오류: Excel 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    print("Excel 파일에서 계약서 파일명 추출 중...")
    excel_files = get_excel_contract_files(excel_path)
    print(f"Excel에서 추출한 계약서 파일 수: {len(excel_files)}개")
    
    print("\n실제 폴더에서 파일명 수집 중...")
    actual_files = get_actual_files(base_path)
    print(f"실제 폴더 파일 수 (감사 추적 인증서 제외): {len(actual_files)}개")
    
    print("\n파일명 비교 중...")
    result = compare_files(excel_files, actual_files)
    
    print_comparison_result(result)
    save_comparison_result(result, output_path)
    save_comparison_excel(result, output_excel_path)
    print(f"\n비교 결과 텍스트 파일: {output_path}")
    print(f"비교 결과 엑셀 파일: {output_excel_path}")

if __name__ == "__main__":
    main()


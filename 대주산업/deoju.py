import os
from pathlib import Path
from collections import defaultdict

def get_file_list(base_path):
    """모두싸인 폴더 내의 모든 파일명을 수집합니다."""
    base_dir = Path(base_path)
    
    영업관리팀_path = base_dir / "모두싸인_ 8월 4일 기준 전자계약서" / "영업관리팀"
    인사팀_path = base_dir / "모두싸인_ 8월 4일 기준 전자계약서" / "인사팀"
    
    영업관리팀_files = []
    인사팀_files = []
    
    # 영업관리팀 파일 수집
    if 영업관리팀_path.exists():
        for root, dirs, files in os.walk(영업관리팀_path):
            for file in files:
                file_path = Path(root) / file
                # 상대 경로로 저장
                relative_path = file_path.relative_to(영업관리팀_path)
                영업관리팀_files.append(str(relative_path))
    
    # 인사팀 파일 수집
    if 인사팀_path.exists():
        for root, dirs, files in os.walk(인사팀_path):
            for file in files:
                file_path = Path(root) / file
                # 상대 경로로 저장
                relative_path = file_path.relative_to(인사팀_path)
                인사팀_files.append(str(relative_path))
    
    return 영업관리팀_files, 인사팀_files

def print_file_list(team_name, files):
    """팀별 파일 리스트를 출력합니다."""
    print(f"\n{'='*80}")
    print(f"【{team_name}】")
    print(f"{'='*80}")
    print(f"총 파일 수: {len(files)}개\n")
    
    # 폴더별로 그룹화
    folder_files = defaultdict(list)
    for file in sorted(files):
        folder = str(Path(file).parent)
        if folder == '.':
            folder = '루트'
        folder_files[folder].append(Path(file).name)
    
    # 폴더별로 출력
    for folder in sorted(folder_files.keys()):
        print(f"\n📁 {folder}")
        print("-" * 80)
        for filename in sorted(folder_files[folder]):
            print(f"  • {filename}")

def main():
    # 현재 스크립트가 있는 디렉토리를 기준으로 설정
    base_path = Path(__file__).parent
    
    print("모두싸인 폴더 내 파일명 수집 중...")
    영업관리팀_files, 인사팀_files = get_file_list(base_path)
    
    # 결과 출력
    print_file_list("영업관리팀", 영업관리팀_files)
    print_file_list("인사팀", 인사팀_files)
    
    # 요약 정보
    print(f"\n{'='*80}")
    print("【요약】")
    print(f"{'='*80}")
    print(f"영업관리팀: {len(영업관리팀_files)}개 파일")
    print(f"인사팀: {len(인사팀_files)}개 파일")
    print(f"전체: {len(영업관리팀_files) + len(인사팀_files)}개 파일")

if __name__ == "__main__":
    main()


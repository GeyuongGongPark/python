# import os
from pathlib import Path

# filenames.txt 읽기
with open('filenames.txt', 'r', encoding='utf-8') as f:
    filenames_txt = [line.strip() for line in f if line.strip()]

# 무제 폴더의 파일 목록 가져오기
folder_path = Path('무제 폴더')
folder_files = []
if folder_path.exists():
    folder_files = [f.name for f in folder_path.iterdir() if f.is_file()]

# 집합으로 변환하여 비교
filenames_txt_set = set(filenames_txt)
folder_files_set = set(folder_files)

# 비교 결과
matched = filenames_txt_set & folder_files_set  # 일치하는 파일
missing_in_folder = filenames_txt_set - folder_files_set  # filenames.txt에는 있지만 폴더에는 없는 파일
extra_in_folder = folder_files_set - filenames_txt_set  # 폴더에는 있지만 filenames.txt에는 없는 파일

# 결과 출력
print("=" * 80)
print("파일 비교 결과")
print("=" * 80)
print(f"\nfilenames.txt 총 파일 수: {len(filenames_txt_set)}")
print(f"무제 폴더 총 파일 수: {len(folder_files_set)}")
print(f"일치하는 파일 수: {len(matched)}")
print(f"filenames.txt에만 있는 파일 수: {len(missing_in_folder)}")
print(f"무제 폴더에만 있는 파일 수: {len(extra_in_folder)}")

print("\n" + "=" * 80)
print("✅ 일치하는 파일 ({}개)".format(len(matched)))
print("=" * 80)
for filename in sorted(matched):
    print(f"  ✓ {filename}")

if missing_in_folder:
    print("\n" + "=" * 80)
    print("❌ filenames.txt에만 있는 파일 ({}개) - 무제 폴더에 없음".format(len(missing_in_folder)))
    print("=" * 80)
    for filename in sorted(missing_in_folder):
        print(f"  ✗ {filename}")

if extra_in_folder:
    print("\n" + "=" * 80)
    print("➕ 무제 폴더에만 있는 파일 ({}개) - filenames.txt에 없음".format(len(extra_in_folder)))
    print("=" * 80)
    for filename in sorted(extra_in_folder):
        print(f"  + {filename}")

# 결과를 파일로 저장
with open('comparison_result.txt', 'w', encoding='utf-8') as f:
    f.write("=" * 80 + "\n")
    f.write("파일 비교 결과\n")
    f.write("=" * 80 + "\n")
    f.write(f"\nfilenames.txt 총 파일 수: {len(filenames_txt_set)}\n")
    f.write(f"무제 폴더 총 파일 수: {len(folder_files_set)}\n")
    f.write(f"일치하는 파일 수: {len(matched)}\n")
    f.write(f"filenames.txt에만 있는 파일 수: {len(missing_in_folder)}\n")
    f.write(f"무제 폴더에만 있는 파일 수: {len(extra_in_folder)}\n")
    
    f.write("\n" + "=" * 80 + "\n")
    f.write("✅ 일치하는 파일 ({}개)\n".format(len(matched)))
    f.write("=" * 80 + "\n")
    for filename in sorted(matched):
        f.write(f"  ✓ {filename}\n")
    
    if missing_in_folder:
        f.write("\n" + "=" * 80 + "\n")
        f.write("❌ filenames.txt에만 있는 파일 ({}개) - 무제 폴더에 없음\n".format(len(missing_in_folder)))
        f.write("=" * 80 + "\n")
        for filename in sorted(missing_in_folder):
            f.write(f"  ✗ {filename}\n")
    
    if extra_in_folder:
        f.write("\n" + "=" * 80 + "\n")
        f.write("➕ 무제 폴더에만 있는 파일 ({}개) - filenames.txt에 없음\n".format(len(extra_in_folder)))
        f.write("=" * 80 + "\n")
        for filename in sorted(extra_in_folder):
            f.write(f"  + {filename}\n")

print("\n" + "=" * 80)
print("결과가 'comparison_result.txt' 파일에 저장되었습니다.")
print("=" * 80)


#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 파일에서 계약명과 계약서명을 추출하고,
하위 폴더에 동일한 이름의 파일이 있는지 검수하는 스크립트
"""

import os
import pandas as pd
from pathlib import Path
from typing import List, Tuple, Dict

def read_excel_file(excel_path: str) -> pd.DataFrame:
    """Excel 파일을 읽어서 DataFrame으로 반환"""
    print(f"Excel 파일 읽는 중: {excel_path}")
    df = pd.read_excel(excel_path, skiprows=1)
    return df

def extract_contract_info(df: pd.DataFrame) -> List[Tuple[str, str]]:
    """계약명과 계약서 파일명을 추출"""
    contracts = []
    
    # 계약명과 계약서 파일 명 컬럼 확인
    contract_name_col = '계약명'
    contract_file_col = '계약서 파일 명'
    
    if contract_name_col not in df.columns or contract_file_col not in df.columns:
        print(f"컬럼을 찾을 수 없습니다.")
        print(f"사용 가능한 컬럼: {df.columns.tolist()}")
        return []
    
    for idx, row in df.iterrows():
        contract_name = str(row[contract_name_col]).strip() if pd.notna(row[contract_name_col]) else ""
        contract_file = str(row[contract_file_col]).strip() if pd.notna(row[contract_file_col]) else ""
        
        # 빈 행 건너뛰기
        if not contract_name and not contract_file:
            continue
            
        contracts.append((contract_name, contract_file))
    
    return contracts

def save_to_txt(contracts: List[Tuple[str, str]], output_path: str):
    """계약명과 계약서명을 txt 파일로 저장"""
    print(f"\n결과를 {output_path}에 저장 중...")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("계약명\t계약서 파일명\n")
        f.write("=" * 100 + "\n")
        for contract_name, contract_file in contracts:
            f.write(f"{contract_name}\t{contract_file}\n")
    print(f"총 {len(contracts)}개의 계약 정보를 저장했습니다.")

def normalize_filename(filename: str) -> str:
    """파일명을 정규화 (확장자 제거, 공백 처리 등)"""
    # 확장자 제거
    name = os.path.splitext(filename)[0]
    # 공백 정규화
    name = ' '.join(name.split())
    return name

def find_files_in_subdirs(base_dir: Path, search_filename: str) -> List[Path]:
    """하위 폴더에서 파일 검색 (부분 일치 포함)"""
    found_files = []
    normalized_search = normalize_filename(search_filename).lower()
    
    # PDF 파일만 검색
    for pdf_file in base_dir.rglob("*.pdf"):
        pdf_name = normalize_filename(pdf_file.name).lower()
        # 정확히 일치하거나 부분 일치하는 경우
        if normalized_search in pdf_name or pdf_name in normalized_search:
            found_files.append(pdf_file)
    
    return found_files

def verify_files(base_dir: Path, contracts: List[Tuple[str, str]]) -> Dict:
    """하위 폴더에서 파일 존재 여부 검수"""
    print(f"\n하위 폴더에서 파일 검색 중: {base_dir}")
    
    results = {
        'found': [],  # 파일을 찾은 경우
        'not_found': [],  # 파일을 찾지 못한 경우
        'multiple_matches': []  # 여러 파일이 매칭된 경우
    }
    
    for contract_name, contract_file in contracts:
        if not contract_file or contract_file == 'nan':
            results['not_found'].append({
                'contract_name': contract_name,
                'contract_file': contract_file,
                'reason': '계약서 파일명이 없음'
            })
            continue
        
        # 계약서 파일명에서 실제 파일명 추출 (확장자 포함 가능)
        search_name = contract_file.strip()
        
        # 찾은 파일들
        found_files = find_files_in_subdirs(base_dir, search_name)
        
        if len(found_files) == 0:
            results['not_found'].append({
                'contract_name': contract_name,
                'contract_file': contract_file,
                'found_files': []
            })
        elif len(found_files) == 1:
            results['found'].append({
                'contract_name': contract_name,
                'contract_file': contract_file,
                'found_path': str(found_files[0])
            })
        else:
            results['multiple_matches'].append({
                'contract_name': contract_name,
                'contract_file': contract_file,
                'found_files': [str(f) for f in found_files]
            })
    
    return results

def print_verification_results(results: Dict):
    """검수 결과 출력"""
    print("\n" + "=" * 100)
    print("파일 검수 결과")
    print("=" * 100)
    
    total = len(results['found']) + len(results['not_found']) + len(results['multiple_matches'])
    
    print(f"\n총 계약 수: {total}")
    print(f"✅ 파일 발견: {len(results['found'])}개")
    print(f"❌ 파일 미발견: {len(results['not_found'])}개")
    print(f"⚠️  여러 파일 매칭: {len(results['multiple_matches'])}개")
    
    if results['found']:
        print("\n" + "-" * 100)
        print("✅ 파일을 찾은 계약서:")
        print("-" * 100)
        for item in results['found'][:10]:  # 처음 10개만 출력
            print(f"계약명: {item['contract_name']}")
            print(f"  계약서 파일명: {item['contract_file']}")
            print(f"  발견 위치: {item['found_path']}")
            print()
        if len(results['found']) > 10:
            print(f"... 외 {len(results['found']) - 10}개")
    
    if results['not_found']:
        print("\n" + "-" * 100)
        print("❌ 파일을 찾지 못한 계약서:")
        print("-" * 100)
        for item in results['not_found'][:20]:  # 처음 20개만 출력
            print(f"계약명: {item['contract_name']}")
            print(f"  계약서 파일명: {item['contract_file']}")
            if 'reason' in item:
                print(f"  사유: {item['reason']}")
            print()
        if len(results['not_found']) > 20:
            print(f"... 외 {len(results['not_found']) - 20}개")
    
    if results['multiple_matches']:
        print("\n" + "-" * 100)
        print("⚠️  여러 파일이 매칭된 계약서:")
        print("-" * 100)
        for item in results['multiple_matches']:
            print(f"계약명: {item['contract_name']}")
            print(f"  계약서 파일명: {item['contract_file']}")
            print(f"  발견된 파일들:")
            for file_path in item['found_files']:
                print(f"    - {file_path}")
            print()

def save_verification_results(results: Dict, output_path: str):
    """검수 결과를 파일로 저장"""
    print(f"\n검수 결과를 {output_path}에 저장 중...")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("=" * 100 + "\n")
        f.write("파일 검수 결과\n")
        f.write("=" * 100 + "\n\n")
        
        total = len(results['found']) + len(results['not_found']) + len(results['multiple_matches'])
        f.write(f"총 계약 수: {total}\n")
        f.write(f"✅ 파일 발견: {len(results['found'])}개\n")
        f.write(f"❌ 파일 미발견: {len(results['not_found'])}개\n")
        f.write(f"⚠️  여러 파일 매칭: {len(results['multiple_matches'])}개\n\n")
        
        if results['found']:
            f.write("-" * 100 + "\n")
            f.write("✅ 파일을 찾은 계약서:\n")
            f.write("-" * 100 + "\n")
            for item in results['found']:
                f.write(f"계약명: {item['contract_name']}\n")
                f.write(f"  계약서 파일명: {item['contract_file']}\n")
                f.write(f"  발견 위치: {item['found_path']}\n\n")
        
        if results['not_found']:
            f.write("-" * 100 + "\n")
            f.write("❌ 파일을 찾지 못한 계약서:\n")
            f.write("-" * 100 + "\n")
            for item in results['not_found']:
                f.write(f"계약명: {item['contract_name']}\n")
                f.write(f"  계약서 파일명: {item['contract_file']}\n")
                if 'reason' in item:
                    f.write(f"  사유: {item['reason']}\n")
                f.write("\n")
        
        if results['multiple_matches']:
            f.write("-" * 100 + "\n")
            f.write("⚠️  여러 파일이 매칭된 계약서:\n")
            f.write("-" * 100 + "\n")
            for item in results['multiple_matches']:
                f.write(f"계약명: {item['contract_name']}\n")
                f.write(f"  계약서 파일명: {item['contract_file']}\n")
                f.write(f"  발견된 파일들:\n")
                for file_path in item['found_files']:
                    f.write(f"    - {file_path}\n")
                f.write("\n")
    
    print(f"검수 결과 저장 완료!")

def main():
    # 경로 설정
    base_dir = Path("/Users/ggpark/Desktop/python/스캔된계약서")
    excel_path = base_dir / "계약서 정리" / "계약서리스트_양식_크리에이션뮤직라이츠_2025_07_25.xlsx"
    output_txt = "계약서_추출_결과.txt"
    verification_result_txt = "계약서_검수_결과.txt"
    
    # Excel 파일 읽기
    df = read_excel_file(str(excel_path))
    
    # 계약 정보 추출
    print("\n계약 정보 추출 중...")
    contracts = extract_contract_info(df)
    print(f"총 {len(contracts)}개의 계약 정보를 추출했습니다.")
    
    # txt 파일로 저장
    save_to_txt(contracts, output_txt)
    
    # 하위 폴더에서 파일 검수
    results = verify_files(base_dir, contracts)
    
    # 결과 출력
    print_verification_results(results)
    
    # 결과 파일로 저장
    save_verification_results(results, verification_result_txt)
    
    print("\n" + "=" * 100)
    print("작업 완료!")
    print(f"- 추출 결과: {output_txt}")
    print(f"- 검수 결과: {verification_result_txt}")
    print("=" * 100)

if __name__ == "__main__":
    main()


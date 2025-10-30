import pandas as pd
import os
from pathlib import Path

def merge_excel_files():
    """
    로아이 원본 폴더의 모든 엑셀 파일을 하나의 엑셀 파일로 합치는 함수
    각 파일명을 시트 탭으로 사용
    """
    
    # 로아이 원본 폴더 경로
    source_folder = "/Users/ggpark/Desktop/python/로아이 원본"
    
    # 출력 파일 경로
    output_file = "/Users/ggpark/Desktop/python/로아이_통합파일.xlsx"
    
    # 엑셀 파일 목록 가져오기
    excel_files = []
    for file in os.listdir(source_folder):
        if file.endswith('.xlsx') and not file.startswith('~'):
            excel_files.append(file)
    
    print(f"발견된 엑셀 파일 수: {len(excel_files)}")
    
    # ExcelWriter 객체 생성
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        for i, file_name in enumerate(excel_files, 1):
            file_path = os.path.join(source_folder, file_name)
            
            try:
                print(f"[{i}/{len(excel_files)}] 처리 중: {file_name}")
                
                # 파일명에서 확장자 제거하여 시트명으로 사용
                sheet_name = file_name.replace('.xlsx', '')
                
                # 시트명이 너무 길면 31자로 제한 (엑셀 시트명 제한)
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                
                # 엑셀 파일 읽기
                df = pd.read_excel(file_path)
                
                # 데이터프레임을 시트로 저장
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"  ✓ 완료: {sheet_name} 시트에 {len(df)}행 저장")
                
            except Exception as e:
                print(f"  ✗ 오류 발생: {file_name} - {str(e)}")
                continue
    
    print(f"\n통합 완료! 파일 저장 위치: {output_file}")
    
    # 결과 요약
    try:
        # 생성된 파일의 시트 정보 확인
        result_df = pd.ExcelFile(output_file)
        print(f"\n생성된 통합 파일 정보:")
        print(f"- 총 시트 수: {len(result_df.sheet_names)}")
        print(f"- 시트 목록: {', '.join(result_df.sheet_names)}")
        
    except Exception as e:
        print(f"결과 확인 중 오류: {str(e)}")

if __name__ == "__main__":
    merge_excel_files()

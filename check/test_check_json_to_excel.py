"""선진 파일 기준 테스트 스크립트 - 매칭 실패 원인 디버깅용"""

import json
import pandas as pd
from pathlib import Path
from fuzzywuzzy import fuzz
from typing import Dict, Any, Optional
from check_json_to_excel import normalize_value, find_column, find_excel_key_column


def load_json_data(raw_data_path: Path) -> Dict[str, Dict[str, Any]]:
    """raw_data 폴더에서 선진 관련 selectDetail.json 파일만 로드"""
    json_data_map: Dict[str, Dict[str, Any]] = {}
    
    print(f"[JSON 수집] 선진 관련 selectDetail.json 파일 수집 중...")
    
    # 선진 관련 폴더 직접 찾기 (os.listdir 사용)
    import os
    seonjin_keywords = ['선진']
    seonjin_folders = []
    
    # os.listdir로 디렉토리 확인
    try:
        dir_items = os.listdir(raw_data_path)
        for item_name in dir_items:
            item_path = raw_data_path / item_name
            if item_path.is_dir() and not item_name.endswith('.zip'):
                for keyword in seonjin_keywords:
                    if keyword in item_name:
                        seonjin_folders.append(item_path)
                        print(f"  발견: {item_name}")
                        break
    except Exception as e:
        print(f"  경고: 디렉토리 읽기 실패: {e}")
    
    # 직접 경로도 시도
    direct_path = raw_data_path / "선진"
    if direct_path.exists() and direct_path not in seonjin_folders:
        seonjin_folders.append(direct_path)
        print(f"  발견 (직접 경로): 선진")
    
    # 선진 관련 다른 폴더들도 찾기
    for item_name in ['선진에프에스', '선진팜', '선진한마을', '선진에프에스', '선진EFS']:
        item_path = raw_data_path / item_name
        if item_path.exists() and item_path not in seonjin_folders:
            seonjin_folders.append(item_path)
            print(f"  발견: {item_name}")
    
    if not seonjin_folders:
        print(f"  경고: 선진 관련 폴더를 찾을 수 없습니다.")
        print(f"  디렉토리 목록 확인 중...")
        try:
            dir_items = os.listdir(raw_data_path)
            print(f"  총 {len(dir_items)}개 항목:")
            for item_name in dir_items[:15]:
                item_path = raw_data_path / item_name
                if item_path.is_dir():
                    print(f"    - {item_name}")
        except:
            pass
        return json_data_map
    
    # 각 선진 폴더에서 JSON 파일 찾기
    json_files = []
    for folder in seonjin_folders:
        found = list(folder.rglob("selectDetail.json"))
        json_files.extend(found)
        print(f"  {folder.name}: {len(found)}개 파일")
    
    print(f"  총 발견된 JSON 파일 수: {len(json_files)}")
    
    if not json_files:
        print(f"  경고: JSON 파일을 찾을 수 없습니다.")
        return json_data_map
    
    # JSON 파일 로드 (SignedContractUUID를 키로 사용)
    for json_file in json_files:
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    # SignedContractUUID를 키로 사용 (엑셀의 관리번호와 매칭)
                    uuid = str(data.get('SignedContractUUID', '')).strip()
                    if uuid:
                        json_data_map[uuid] = data
                    # ManageNo도 키로 추가 (백업용)
                    manage_no = str(data.get('ManageNo', '')).strip()
                    if manage_no and manage_no not in json_data_map:
                        json_data_map[manage_no] = data
        except Exception as e:
            print(f"  경고: {json_file} 읽기 실패: {e}")
            continue
    
    print(f"  로드된 JSON 데이터 수: {len(json_data_map)}")
    
    if json_data_map:
        # 샘플 데이터 출력
        sample_key = list(json_data_map.keys())[0]
        sample_data = json_data_map[sample_key]
        print(f"\n  [샘플 JSON 데이터]")
        print(f"    키 (SignedContractUUID): {sample_key}")
        print(f"    ManageNo: {sample_data.get('ManageNo')}")
        print(f"    SignedContractUUID: {sample_data.get('SignedContractUUID')}")
        print(f"    ContractName: {sample_data.get('ContractName')}")
        print(f"    StatusName: {sample_data.get('StatusName')}")
        print(f"    ManagerUserName: {sample_data.get('ManagerUserName')}")
    
    return json_data_map


def test_single_excel():
    """선진 통합 엑셀 파일 하나만 테스트"""
    script_dir = Path(__file__).parent
    tonghab_dir = script_dir / "통합본"
    raw_data_dir = script_dir / "raw_data"
    
    # 선진 통합 파일 찾기
    excel_file = tonghab_dir / "선진_통합.xlsx"
    if not excel_file.exists():
        print(f"오류: {excel_file} 파일을 찾을 수 없습니다.")
        return
    
    print(f"\n{'='*60}")
    print(f"테스트 대상: {excel_file.name}")
    print(f"{'='*60}\n")
    
    # JSON 데이터 로드
    json_data_map = load_json_data(raw_data_dir)
    
    if not json_data_map:
        print("오류: JSON 데이터를 찾을 수 없습니다.")
        return
    
    # 엑셀 파일 읽기
    xls = pd.ExcelFile(excel_file)
    print(f"시트 목록: {xls.sheet_names}\n")
    
    # CLM등록 시트 찾기
    sheet_name = None
    for name in xls.sheet_names:
        if '등록' in name or 'CLM' in name.upper():
            sheet_name = name
            break
    
    if sheet_name is None:
        sheet_name = xls.sheet_names[0]
    
    print(f"사용할 시트: {sheet_name}\n")
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    print(f"엑셀 행 수: {len(df)}\n")
    
    # 키 컬럼 찾기
    key_col = find_excel_key_column(df)
    if key_col is None:
        print(f"경고: 키 컬럼을 찾을 수 없습니다.")
        print(f"컬럼 목록: {list(df.columns)[:20]}")
        return
    
    print(f"키 컬럼: {key_col}\n")
    
    # 비교할 컬럼 찾기
    contract_name_col = find_column(df, [
        "계약명", "계약 명", "계약이름", "계약 이름",
        "ContractName", "contract name", "계약서명"
    ])
    status_col = find_column(df, [
        "진행 상태", "진행상태", "상태", "진행 상태 ", "Status",
        "계약단계", "계약 단계"
    ])
    manager_col = find_column(df, [
        "담당자 이름", "담당자", "담당자명", "담당자 이름 ",
        "ManagerUserName", "manager", "담당자 이메일"
    ])
    
    print(f"계약명 컬럼: {contract_name_col if contract_name_col else '없음'}")
    print(f"진행상태 컬럼: {status_col if status_col else '없음'}")
    print(f"담당자 컬럼: {manager_col if manager_col else '없음'}")
    print(f"전체 컬럼 수: {len(df.columns)}개")
    print(f"비교 대상 컬럼: {len([c for c in df.columns if c != key_col])}개 (키 컬럼 제외)\n")
    
    # 상세 비교 (처음 10개 행)
    print(f"{'='*60}")
    print(f"상세 비교 결과 (처음 10개 행)")
    print(f"{'='*60}\n")
    
    match_count = 0
    mismatch_count = 0
    not_found_count = 0
    
    for i, (index, row) in enumerate(df.iterrows()):
        if i >= 10:  # 처음 10개만
            break
        
        excel_key = str(row[key_col]).strip() if pd.notna(row[key_col]) else ""
        
        if not excel_key:
            continue
        
        print(f"\n[{i+1}] 관리번호 (엑셀): {excel_key}")
        print(f"  관리번호 길이: {len(excel_key)}")
        print(f"-" * 60)
        
        # JSON 데이터 찾기 (SignedContractUUID로 먼저 시도)
        json_data = json_data_map.get(excel_key)
        
        if json_data is None:
            # UUID 형식이 아닌 경우 ManageNo로도 시도
            print(f"  SignedContractUUID로 매칭 실패, ManageNo로 재시도...")
            # JSON 데이터에서 ManageNo로 찾기
            for key, data in json_data_map.items():
                if data.get('ManageNo') == excel_key:
                    json_data = data
                    print(f"  ManageNo로 매칭 성공!")
                    break
        
        if json_data is None:
            print(f"  ❌ JSON 데이터를 찾을 수 없음")
            print(f"  JSON 키 목록 샘플: {list(json_data_map.keys())[:5]}")
            not_found_count += 1
            continue
        
        # 엑셀과 JSON 필드 매핑 정의 (check_json_to_excel.py와 동일)
        field_mappings = {
            '계약명': ['ContractName'],
            '계약 명': ['ContractName'],
            '진행 상태': ['StatusName'],
            '진행상태': ['StatusName'],
            '상태': ['StatusName'],
            '담당자 이름': ['ManagerUserName'],
            '담당자': ['ManagerUserName'],
            '담당자명': ['ManagerUserName'],
            '검토 요청자': ['ManagerUserName'],
            '검토요청자': ['ManagerUserName'],
            '계약 시작일': ['ContractStartDate'],
            '계약 완료일': ['ContractEndDate'],
            '계약 종료': ['ContractEndDate'],
            '계약 체결일': ['SignedDate'],
            '계약예정일': ['SignedDate'],
            '계약 규모': ['ContractAmountList'],
            '통화': ['CurrencyCode'],
            '상대 계약자': ['SignedContractPartnerInfoList'],
            '상대계약자': ['SignedContractPartnerInfoList'],
            '대분류': ['MainContractTypeCode'],
            '대분류_카테고리이름': ['MainContractTypeName'],
            '분류': ['ContractClassCode'],
            '분류_카테고리이름': ['ContractClassName'],
            # 새로 확인된 매핑
            '계열사명': ['CCName'],
            '관리 번호': ['ManageNo'],  # ManageName이 아니라 ManageNo인 것으로 확인
            '관리번호': ['ManageNo'],
            '개정번호': ['Revision'],
            '수정일': ['UpdateDate'],
            '담당자 이메일': ['ManagerUserEmail'],
            '담당자 휴대폰 번호': ['ManagerUserPhoneNumber'],
            '담당자 휴대폰': ['ManagerUserPhoneNumber'],
            '담당자 휴대폰번호': ['ManagerUserPhoneNumber'],
            '담당자 퇴사여부': ['ManagerUserIsActive'],
            '담당자 퇴사': ['ManagerUserIsActive'],
        }
        
        # 모든 컬럼 비교 결과 수집
        all_columns_comparison = []
        
        # 엑셀의 모든 컬럼에 대해 비교
        for excel_col in df.columns:
            if excel_col == key_col:  # 키 컬럼은 제외
                continue
            
            excel_val_raw = row[excel_col] if pd.notna(row[excel_col]) else ''
            excel_val = normalize_value(excel_val_raw)
            
            # JSON 필드 찾기
            json_field = None
            json_val = None
            json_val_raw = None
            
            # 매핑 테이블에서 찾기
            excel_col_normalized = str(excel_col).strip()
            if excel_col_normalized in field_mappings:
                for json_field_candidate in field_mappings[excel_col_normalized]:
                    if json_field_candidate in json_data:
                        json_field = json_field_candidate
                        json_val_raw = json_data[json_field_candidate]
                        if isinstance(json_val_raw, (list, dict)):
                            json_val = str(json_val_raw)
                        else:
                            json_val = normalize_value(json_val_raw)
                        break
            
            # 매핑 테이블에 없으면 컬럼명과 유사한 JSON 필드 찾기
            if json_val is None:
                col_normalized = str(excel_col).strip().lower().replace(" ", "").replace("_", "")
                for json_key in json_data.keys():
                    json_key_normalized = str(json_key).lower().replace(" ", "").replace("_", "")
                    if col_normalized in json_key_normalized or json_key_normalized in col_normalized:
                        if len(col_normalized) >= 3:
                            json_field = json_key
                            json_val_raw = json_data[json_key]
                            if isinstance(json_val_raw, (list, dict)):
                                json_val = str(json_val_raw)
                            else:
                                json_val = normalize_value(json_val_raw)
                            break
            
            # 비교 결과 저장
            is_match = False
            similarity = 100
            mismatch_reason = ''
            
            if json_val is not None:
                if not excel_val and not json_val:
                    is_match = True
                elif excel_val and json_val:
                    if excel_val == json_val:
                        is_match = True
                    else:
                        similarity = fuzz.ratio(excel_val, json_val)
                        if similarity >= 95:
                            is_match = True
                        else:
                            mismatch_reason = '값 불일치'
                else:
                    mismatch_reason = '한쪽 값만 존재'
                    similarity = 0
            elif excel_val:
                mismatch_reason = 'JSON에 해당 필드 없음'
                similarity = 0
            elif json_val:
                mismatch_reason = '엑셀에 해당 필드 없음'
                similarity = 0
            else:
                is_match = True
            
            all_columns_comparison.append({
                '컬럼명': excel_col,
                '엑셀_값': excel_val if excel_val else '(없음)',
                '엑셀_원본값': str(excel_val_raw) if excel_val_raw else '(없음)',
                'JSON_필드명': json_field if json_field else '(매칭 실패)',
                'JSON_값': json_val if json_val else '(없음)',
                'JSON_원본값': str(json_val_raw) if json_val_raw is not None else '(없음)',
                '일치': is_match,
                '유사도': similarity,
                '비고': mismatch_reason if not is_match else '',
            })
        
        # 엑셀 값 가져오기
        excel_contract = ''
        excel_status = ''
        excel_manager = ''
        
        if contract_name_col and pd.notna(row.get(contract_name_col)):
            excel_contract_raw = row[contract_name_col]
            excel_contract = normalize_value(excel_contract_raw)
        if status_col and pd.notna(row.get(status_col)):
            excel_status_raw = row[status_col]
            excel_status = normalize_value(excel_status_raw)
        if manager_col and pd.notna(row.get(manager_col)):
            excel_manager_raw = row[manager_col]
            excel_manager = normalize_value(excel_manager_raw)
        
        # JSON 값 가져오기
        json_contract_raw = json_data.get('ContractName') or json_data.get('CCName', '')
        json_contract = normalize_value(json_contract_raw)
        
        json_status_raw = json_data.get('StatusName', '')
        json_status = normalize_value(json_status_raw)
        
        json_manager_raw = json_data.get('ManagerUserName', '')
        json_manager = normalize_value(json_manager_raw)
        
        # 계약명 비교
        print(f"\n  [계약명]")
        print(f"    엑셀 원본: {repr(excel_contract_raw) if contract_name_col and pd.notna(row.get(contract_name_col)) else '(없음)'}")
        print(f"    엑셀 정규화: {repr(excel_contract)}")
        print(f"    JSON 원본: {repr(json_contract_raw)}")
        print(f"    JSON 정규화: {repr(json_contract)}")
        if excel_contract and json_contract:
            is_match = excel_contract == json_contract
            similarity = fuzz.ratio(excel_contract, json_contract)
            print(f"    일치 여부: {'✓' if is_match else '✗'} (유사도: {similarity}%)")
            if not is_match:
                print(f"    ⚠️  불일치!")
        else:
            print(f"    ⚠️  한쪽 값만 존재")
        
        # 진행상태 비교
        print(f"\n  [진행상태]")
        print(f"    엑셀 원본: {repr(excel_status_raw) if status_col and pd.notna(row.get(status_col)) else '(없음)'}")
        print(f"    엑셀 정규화: {repr(excel_status)}")
        print(f"    JSON 원본: {repr(json_status_raw)}")
        print(f"    JSON 정규화: {repr(json_status)}")
        if excel_status and json_status:
            is_match = excel_status == json_status
            similarity = fuzz.ratio(excel_status, json_status)
            print(f"    일치 여부: {'✓' if is_match else '✗'} (유사도: {similarity}%)")
            if not is_match:
                print(f"    ⚠️  불일치!")
        else:
            print(f"    ⚠️  한쪽 값만 존재")
        
        # 담당자 비교
        print(f"\n  [담당자]")
        print(f"    엑셀 원본: {repr(excel_manager_raw) if manager_col and pd.notna(row.get(manager_col)) else '(없음)'}")
        print(f"    엑셀 정규화: {repr(excel_manager)}")
        print(f"    JSON 원본: {repr(json_manager_raw)}")
        print(f"    JSON 정규화: {repr(json_manager)}")
        if excel_manager and json_manager:
            is_match = excel_manager == json_manager
            similarity = fuzz.ratio(excel_manager, json_manager)
            print(f"    일치 여부: {'✓' if is_match else '✗'} (유사도: {similarity}%)")
            if not is_match:
                print(f"    ⚠️  불일치!")
        else:
            print(f"    ⚠️  한쪽 값만 존재")
        
        # 모든 컬럼 비교 결과 출력
        print(f"\n  [전체 컬럼 비교 결과]")
        match_count_col = sum(1 for c in all_columns_comparison if c['일치'])
        mismatch_count_col = len(all_columns_comparison) - match_count_col
        print(f"    총 비교 컬럼: {len(all_columns_comparison)}개")
        print(f"    일치: {match_count_col}개")
        print(f"    불일치: {mismatch_count_col}개")
        
        if mismatch_count_col > 0:
            print(f"\n    [불일치 컬럼 목록]")
            mismatch_items = [c for c in all_columns_comparison if not c['일치']]
            # 전체 불일치 컬럼 상세 출력
            for idx, comp in enumerate(mismatch_items):
                print(f"      • {comp['컬럼명']}")
                excel_val_display = comp['엑셀_값'][:50] if len(comp['엑셀_값']) > 50 else comp['엑셀_값']
                json_val_display = comp['JSON_값'][:50] if len(comp['JSON_값']) > 50 else comp['JSON_값']
                print(f"        엑셀: {excel_val_display}")
                print(f"        JSON: {json_val_display}")
                print(f"        유사도: {comp['유사도']}% | {comp['비고']}")
        
        # 전체 일치 여부
        all_match = (
            (excel_contract == json_contract if excel_contract and json_contract else False) or
            (not excel_contract and not json_contract)
        ) and (
            (excel_status == json_status if excel_status and json_status else False) or
            (not excel_status and not json_status)
        ) and (
            (excel_manager == json_manager if excel_manager and json_manager else False) or
            (not excel_manager and not json_manager)
        )
        
        if all_match:
            match_count += 1
        else:
            mismatch_count += 1
    
    print(f"\n{'='*60}")
    print(f"요약")
    print(f"{'='*60}")
    print(f"  매칭 성공: {match_count}개")
    print(f"  매칭 실패: {mismatch_count}개")
    print(f"  JSON 없음: {not_found_count}개")
    print(f"  총 비교: {match_count + mismatch_count + not_found_count}개")


if __name__ == "__main__":
    test_single_excel()


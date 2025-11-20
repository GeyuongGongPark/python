"""통합본 엑셀 파일 vs raw_data JSON 파일 비교 리포트 생성 스크립트.

개요
- 입력: 
- 통합본 폴더의 엑셀 파일들 (check/통합본/*.xlsx)
- raw_data 폴더의 JSON 파일들 (check/raw_data/**/selectDetail.json)
- 처리: 주요 필드의 Fuzzy 매칭으로 유사도 계산 및 불일치 분류
- 출력: 불일치 항목을 시트별로 정리한 Excel 파일
"""

import json
import pandas as pd
import sys
import os
import re
from pathlib import Path
from fuzzywuzzy import fuzz
from typing import Dict, List, Any, Optional


def load_json_data(raw_data_path: Path) -> Dict[str, Dict[str, Any]]:
    """raw_data 폴더에서 모든 selectDetail.json 파일을 로드하여 ManageNo를 키로 하는 딕셔너리 생성"""
    json_data_map: Dict[str, Dict[str, Any]] = {}
    
    print(f"[JSON 수집] {raw_data_path}에서 selectDetail.json 파일 수집 중...")
    
    json_files = list(raw_data_path.rglob("selectDetail.json"))
    print(f"  발견된 JSON 파일 수: {len(json_files)}")
    
    for json_file in json_files:
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    # SignedContractUUID를 우선 키로 사용 (엑셀의 관리번호와 매칭)
                    uuid = str(data.get('SignedContractUUID', '')).strip()
                    if uuid:
                        json_data_map[uuid] = data
                    # ManageNo도 키로 추가 (백업용)
                    manage_no = str(data.get('ManageNo', '')).strip()
                    if manage_no:
                        # UUID가 이미 키로 사용되지 않은 경우에만 ManageNo로 추가
                        if manage_no not in json_data_map:
                            json_data_map[manage_no] = data
        except Exception as e:
            print(f"  경고: {json_file} 읽기 실패: {e}")
            continue
    
    print(f"  로드된 JSON 데이터 수: {len(json_data_map)}")
    return json_data_map


def find_excel_key_column(df: pd.DataFrame) -> Optional[str]:
    """엑셀 DataFrame에서 키 컬럼 찾기 (NO., 관리 번호 등)"""
    candidates = [
        "NO.", "No.", "NO", "No", "no", "번호",
        "관리 번호", "관리번호", "관리번호 ", "CLM NO.", "CLM NO", "CLMNO."
    ]
    
    for col in df.columns:
        col_str = str(col).strip()
        if col_str in candidates:
            return col
    
    # 정규화된 매칭
    for candidate in candidates:
        for col in df.columns:
            col_normalized = str(col).strip().lower().replace(" ", "").replace(".", "")
            candidate_normalized = candidate.lower().replace(" ", "").replace(".", "")
            if col_normalized == candidate_normalized:
                return col
    
    return None


def normalize_value(val) -> str:
    """값을 정규화하여 비교 (공백 제거, 줄바꿈 제거, None 처리)"""
    if pd.isna(val) or val is None:
        return ''
    val_str = str(val)
    # 줄바꿈, 탭, 캐리지 리턴 제거
    val_str = val_str.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
    # 앞뒤 공백 제거
    val_str = val_str.strip()
    # 연속된 공백을 단일 공백으로 통일
    val_str = ' '.join(val_str.split())
    return val_str


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """후보 목록에서 컬럼 찾기 (정확 매칭 → 정규화 매칭 → 부분 매칭)"""
    # 1. 정확 매칭
    for col in df.columns:
        col_str = str(col).strip()
        if col_str in candidates:
            return col
    
    # 2. 정규화된 정확 매칭
    for candidate in candidates:
        for col in df.columns:
            col_normalized = str(col).strip().lower().replace(" ", "").replace(".", "").replace("_", "")
            candidate_normalized = candidate.lower().replace(" ", "").replace(".", "").replace("_", "")
            if col_normalized == candidate_normalized:
                return col
    
    # 3. 부분 매칭 (후보의 핵심 키워드가 컬럼명에 포함되어 있는지)
    for candidate in candidates:
        candidate_keywords = candidate.replace(" ", "").replace("_", "").lower()
        for col in df.columns:
            col_normalized = str(col).strip().replace(" ", "").replace("_", "").lower()
            # 후보 키워드가 컬럼명에 포함되어 있으면 (최소 3글자 이상)
            if len(candidate_keywords) >= 3 and candidate_keywords in col_normalized:
                return col
            # 또는 컬럼명의 핵심 부분이 후보에 포함되어 있으면
            if len(col_normalized) >= 3 and col_normalized in candidate_keywords:
                return col
    
    return None


def compare_excel_with_json(
    excel_path: Path,
    json_data_map: Dict[str, Dict[str, Any]],
    output_path: Path
) -> None:
    """엑셀 파일과 JSON 데이터를 비교하여 결과를 저장"""
    print(f"\n[처리 시작] {excel_path.name}")
    
    try:
        # 엑셀 파일 읽기
        xls = pd.ExcelFile(excel_path)
        print(f"  시트 목록: {xls.sheet_names}")
        
        # CLM등록 시트 찾기 (없으면 첫 번째 시트 사용)
        sheet_name = None
        for name in xls.sheet_names:
            if '등록' in name or 'CLM' in name.upper():
                sheet_name = name
                break
        
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        
        print(f"  사용할 시트: {sheet_name}")
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        print(f"  엑셀 행 수: {len(df)}")
        
        # 키 컬럼 찾기
        key_col = find_excel_key_column(df)
        if key_col is None:
            print(f"  경고: 키 컬럼을 찾을 수 없습니다. 컬럼: {list(df.columns)[:10]}")
            return
        
        print(f"  키 컬럼: {key_col}")
        
        # 비교할 컬럼 찾기 (더 많은 후보 추가)
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
        
        # 찾은 컬럼 출력
        print(f"  계약명 컬럼: {contract_name_col if contract_name_col else '없음'}")
        print(f"  진행상태 컬럼: {status_col if status_col else '없음'}")
        print(f"  담당자 컬럼: {manager_col if manager_col else '없음'}")
        print(f"  전체 컬럼 수: {len(df.columns)}개")
        print(f"  비교 대상 컬럼: {len([c for c in df.columns if c != key_col])}개 (키 컬럼 제외)")
        
        # 샘플 데이터 확인 (처음 3개 행)
        if len(df) > 0:
            print(f"\n  [샘플 데이터 확인]")
            for i in range(min(3, len(df))):
                sample_row = df.iloc[i]
                sample_key = str(sample_row[key_col]).strip() if pd.notna(sample_row[key_col]) else ""
                if sample_key:
                    json_data = json_data_map.get(sample_key)
                    if json_data:
                        excel_contract = normalize_value(sample_row[contract_name_col]) if contract_name_col and pd.notna(sample_row.get(contract_name_col)) else ""
                        excel_status = normalize_value(sample_row[status_col]) if status_col and pd.notna(sample_row.get(status_col)) else ""
                        excel_manager = normalize_value(sample_row[manager_col]) if manager_col and pd.notna(sample_row.get(manager_col)) else ""
                        
                        json_contract = normalize_value(json_data.get('ContractName') or json_data.get('CCName', ''))
                        json_status = normalize_value(json_data.get('StatusName', ''))
                        json_manager = normalize_value(json_data.get('ManagerUserName', ''))
                        
                        print(f"    [{i+1}] 관리번호: {sample_key}")
                        print(f"      계약명 - 엑셀: '{excel_contract}' | JSON: '{json_contract}' | 일치: {excel_contract == json_contract}")
                        print(f"      진행상태 - 엑셀: '{excel_status}' | JSON: '{json_status}' | 일치: {excel_status == json_status}")
                        print(f"      담당자 - 엑셀: '{excel_manager}' | JSON: '{json_manager}' | 일치: {excel_manager == json_manager}")
                        break
        
        # 불일치 결과 저장 리스트
        management_mismatch = []
        contract_mismatch = []
        status_mismatch = []
        manager_mismatch = []
        not_found = []
        overall_mismatch = []
        all_fields_mismatch = []  # 모든 필드 불일치
        
        # 각 행 비교
        for index, row in df.iterrows():
            excel_key = str(row[key_col]).strip() if pd.notna(row[key_col]) else ""
            
            if not excel_key:
                continue
            
            # JSON 데이터에서 매칭 찾기 (SignedContractUUID로 먼저 시도)
            json_data = json_data_map.get(excel_key)
            
            if json_data is None:
                # UUID 형식이 아닌 경우 ManageNo로도 시도
                for key, data in json_data_map.items():
                    if data.get('ManageNo') == excel_key:
                        json_data = data
                        break
            
            if json_data is None:
                # 정확 매칭 실패 시 Fuzzy 매칭으로 가장 유사한 것 찾기
                best_match = None
                best_score = 0
                best_key = None
                
                for json_key, json_val in json_data_map.items():
                    score = fuzz.ratio(excel_key, json_key)
                    if score > best_score:
                        best_score = score
                        best_match = json_val
                        best_key = json_key
                
                if best_score >= 80:  # 80% 이상 유사하면 매칭
                    json_data = best_match
                    excel_key = best_key
                else:
                    not_found.append({
                        '엑셀_관리번호': excel_key,
                        '계약명': row[contract_name_col] if contract_name_col and pd.notna(row.get(contract_name_col)) else '',
                        '진행_상태': row[status_col] if status_col and pd.notna(row.get(status_col)) else '',
                        '담당자': row[manager_col] if manager_col and pd.notna(row.get(manager_col)) else '',
                    })
                    continue
            
            # 엑셀과 JSON 필드 매핑 정의
            field_mappings = {
                # 엑셀 컬럼명: [JSON 필드명 후보들]
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
            
            # 모든 엑셀 컬럼과 JSON 필드 비교
            all_mismatches = []
            base_data = {
                '엑셀_관리번호': excel_key,
                'JSON_관리번호': json_data.get('ManageNo', ''),
                'JSON_SignedContractUUID': json_data.get('SignedContractUUID', ''),
            }
            
            # 엑셀의 모든 컬럼에 대해 비교
            for excel_col in df.columns:
                if excel_col == key_col:  # 키 컬럼은 제외
                    continue
                
                excel_val_raw = row[excel_col] if pd.notna(row[excel_col]) else ''
                excel_val = normalize_value(excel_val_raw)
                
                # JSON 필드 찾기 (매핑 테이블 사용)
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
                            # 리스트나 딕셔너리는 문자열로 변환
                            if isinstance(json_val_raw, (list, dict)):
                                json_val = str(json_val_raw)
                            else:
                                json_val = normalize_value(json_val_raw)
                            break
                
                # 매핑 테이블에 없으면 컬럼명과 유사한 JSON 필드 찾기
                if json_val is None:
                    # 컬럼명을 정규화하여 JSON 필드명과 매칭 시도
                    col_normalized = str(excel_col).strip().lower().replace(" ", "").replace("_", "")
                    for json_key in json_data.keys():
                        json_key_normalized = str(json_key).lower().replace(" ", "").replace("_", "")
                        if col_normalized in json_key_normalized or json_key_normalized in col_normalized:
                            if len(col_normalized) >= 3:  # 최소 3글자 이상
                                json_field = json_key
                                json_val_raw = json_data[json_key]
                                if isinstance(json_val_raw, (list, dict)):
                                    json_val = str(json_val_raw)
                                else:
                                    json_val = normalize_value(json_val_raw)
                                break
                
                # 값 비교 (빈 값도 포함하여 비교)
                is_match = False
                similarity = 100
                mismatch_reason = ''
                
                if json_val is not None:
                    # 둘 다 빈 값이면 일치
                    if not excel_val and not json_val:
                        is_match = True
                    # 둘 다 값이 있으면 비교
                    elif excel_val and json_val:
                        if excel_val == json_val:
                            is_match = True
                        else:
                            similarity = fuzz.ratio(excel_val, json_val)
                            if similarity >= 95:
                                is_match = True  # 95% 이상이면 일치로 간주
                            else:
                                mismatch_reason = '값 불일치'
                    # 한쪽만 값이 있으면 불일치
                    else:
                        mismatch_reason = '한쪽 값만 존재'
                        similarity = 0
                elif excel_val:  # JSON에 해당 필드가 없는 경우
                    mismatch_reason = 'JSON에 해당 필드 없음'
                    similarity = 0
                elif json_val:  # 엑셀에 값이 없지만 JSON에는 있는 경우
                    mismatch_reason = '엑셀에 해당 필드 없음'
                    similarity = 0
                else:
                    # 둘 다 없으면 일치로 간주
                    is_match = True
                
                # 불일치인 경우에만 기록
                if not is_match:
                    all_mismatches.append({
                        '엑셀_컬럼명': excel_col,
                        '엑셀_값': excel_val if excel_val else '(없음)',
                        '엑셀_원본값': str(excel_val_raw) if excel_val_raw else '(없음)',
                        'JSON_필드명': json_field if json_field else '(매칭 실패)',
                        'JSON_값': json_val if json_val else '(없음)',
                        'JSON_원본값': str(json_val_raw) if json_val_raw is not None else '(없음)',
                        '유사성_점수': similarity,
                        '비고': mismatch_reason,
                    })
            
            # 기존 방식 유지 (계약명, 진행상태, 담당자)
            json_contract_name = ''
            if json_data.get('ContractName'):
                json_contract_name = normalize_value(json_data.get('ContractName'))
            elif json_data.get('CCName'):
                json_contract_name = normalize_value(json_data.get('CCName'))
            
            json_status = normalize_value(json_data.get('StatusName', ''))
            json_manager = normalize_value(json_data.get('ManagerUserName', ''))
            
            excel_contract = ''
            excel_status = ''
            excel_manager = ''
            
            if contract_name_col and pd.notna(row.get(contract_name_col)):
                excel_contract = normalize_value(row[contract_name_col])
            if status_col and pd.notna(row.get(status_col)):
                excel_status = normalize_value(row[status_col])
            if manager_col and pd.notna(row.get(manager_col)):
                excel_manager = normalize_value(row[manager_col])
            
            # base_data에 기존 필드 추가
            base_data.update({
                '엑셀_계약명': excel_contract,
                '엑셀_진행상태': excel_status,
                '엑셀_담당자': excel_manager,
                'JSON_계약명': json_contract_name,
                'JSON_진행상태': json_status,
                'JSON_담당자': json_manager,
            })
            
            # 관리번호 비교 (이미 매칭되었으므로 100%여야 함)
            if base_data['엑셀_관리번호'] != base_data['JSON_관리번호']:
                management_mismatch.append({
                    **base_data,
                    '유사성_점수': fuzz.ratio(base_data['엑셀_관리번호'], base_data['JSON_관리번호'])
                })
            
            # 계약명 비교 (둘 다 값이 있을 때만)
            if excel_contract and json_contract_name:
                # 정확히 같으면 일치로 처리
                if excel_contract == json_contract_name:
                    pass  # 일치, 불일치 리스트에 추가하지 않음
                else:
                    similarity = fuzz.ratio(excel_contract, json_contract_name)
                    # 95% 이상이면 실질적으로 일치로 간주 (공백, 특수문자 차이 등)
                    if similarity < 95:
                        contract_mismatch.append({
                            **base_data,
                            '유사성_점수': similarity,
                            '엑셀_원본값': str(row[contract_name_col]) if contract_name_col and pd.notna(row.get(contract_name_col)) else '',
                            'JSON_원본값': str(json_data.get('ContractName') or json_data.get('CCName', ''))
                        })
            elif excel_contract or json_contract_name:
                # 한쪽만 값이 있는 경우도 불일치로 기록
                contract_mismatch.append({
                    **base_data,
                    '유사성_점수': 0,
                    '비고': '한쪽 값만 존재'
                })
            
            # 진행상태 비교 (둘 다 값이 있을 때만)
            if excel_status and json_status:
                # 정확히 같으면 일치로 처리
                if excel_status == json_status:
                    pass  # 일치, 불일치 리스트에 추가하지 않음
                else:
                    similarity = fuzz.ratio(excel_status, json_status)
                    # 95% 이상이면 실질적으로 일치로 간주 (공백 차이 등)
                    if similarity < 95:
                        status_mismatch.append({
                            **base_data,
                            '유사성_점수': similarity,
                            '엑셀_원본값': str(row[status_col]) if status_col and pd.notna(row.get(status_col)) else '',
                            'JSON_원본값': str(json_data.get('StatusName', ''))
                        })
            elif excel_status or json_status:
                # 한쪽만 값이 있는 경우도 불일치로 기록
                status_mismatch.append({
                    **base_data,
                    '유사성_점수': 0,
                    '비고': '한쪽 값만 존재'
                })
            
            # 담당자 비교 (둘 다 값이 있을 때만)
            if excel_manager and json_manager:
                # 정확히 같으면 일치로 처리
                if excel_manager == json_manager:
                    pass  # 일치, 불일치 리스트에 추가하지 않음
                else:
                    similarity = fuzz.ratio(excel_manager, json_manager)
                    # 95% 이상이면 실질적으로 일치로 간주 (공백 차이 등)
                    if similarity < 95:
                        manager_mismatch.append({
                            **base_data,
                            '유사성_점수': similarity,
                            '엑셀_원본값': str(row[manager_col]) if manager_col and pd.notna(row.get(manager_col)) else '',
                            'JSON_원본값': str(json_data.get('ManagerUserName', ''))
                        })
            elif excel_manager or json_manager:
                # 한쪽만 값이 있는 경우도 불일치로 기록
                manager_mismatch.append({
                    **base_data,
                    '유사성_점수': 0,
                    '비고': '한쪽 값만 존재'
                })
            
            # 종합 불일치 계산 (정규화된 값으로 비교)
            contract_sim = 100
            if excel_contract and json_contract_name:
                if excel_contract == json_contract_name:
                    contract_sim = 100  # 정확히 일치
                else:
                    contract_sim = fuzz.ratio(excel_contract, json_contract_name)
                    if contract_sim >= 95:
                        contract_sim = 100  # 95% 이상이면 일치로 간주
            elif excel_contract or json_contract_name:
                contract_sim = 0  # 한쪽만 값이 있으면 불일치
            
            status_sim = 100
            if excel_status and json_status:
                if excel_status == json_status:
                    status_sim = 100  # 정확히 일치
                else:
                    status_sim = fuzz.ratio(excel_status, json_status)
                    if status_sim >= 95:
                        status_sim = 100  # 95% 이상이면 일치로 간주
            elif excel_status or json_status:
                status_sim = 0  # 한쪽만 값이 있으면 불일치
            
            manager_sim = 100
            if excel_manager and json_manager:
                if excel_manager == json_manager:
                    manager_sim = 100  # 정확히 일치
                else:
                    manager_sim = fuzz.ratio(excel_manager, json_manager)
                    if manager_sim >= 95:
                        manager_sim = 100  # 95% 이상이면 일치로 간주
            elif excel_manager or json_manager:
                manager_sim = 0  # 한쪽만 값이 있으면 불일치
            
            overall_score = (contract_sim * 0.5 + status_sim * 0.3 + manager_sim * 0.2)
            
            if overall_score < 100:
                overall_mismatch.append({
                    **base_data,
                    '종합_유사성_점수': round(overall_score, 2),
                    '계약명_유사성': contract_sim,
                    '진행상태_유사성': status_sim,
                    '담당자_유사성': manager_sim,
                })
            
            # 모든 필드 불일치 추가
            if all_mismatches:
                for mismatch in all_mismatches:
                    all_fields_mismatch.append({
                        **base_data,
                        **mismatch,
                    })
            
            # 첫 번째 행에서만 디버깅 정보 출력
            if index == df.index[0]:
                print(f"\n  [디버깅] 첫 번째 행의 모든 컬럼 비교 결과:")
                print(f"    - 총 비교한 컬럼 수: {len([c for c in df.columns if c != key_col])}개")
                print(f"    - 불일치한 컬럼 수: {len(all_mismatches)}개")
                if all_mismatches:
                    print(f"    - 불일치 컬럼 목록 (전체):")
                    for mismatch in all_mismatches:
                        print(f"      • {mismatch['엑셀_컬럼명']}: {mismatch['비고']}")
        
        # 결과를 DataFrame으로 변환 (불일치 항목만 저장)
        result_dfs = {}
        
        # 불일치 항목이 있는 경우에만 시트 생성
        if management_mismatch:
            result_dfs['관리번호_불일치'] = pd.DataFrame(management_mismatch)
        if contract_mismatch:
            result_dfs['계약명_불일치'] = pd.DataFrame(contract_mismatch)
        if status_mismatch:
            result_dfs['진행상태_불일치'] = pd.DataFrame(status_mismatch)
        if manager_mismatch:
            result_dfs['담당자_불일치'] = pd.DataFrame(manager_mismatch)
        if not_found:
            result_dfs['JSON_매칭_실패'] = pd.DataFrame(not_found)
        if overall_mismatch:
            result_dfs['종합_불일치'] = pd.DataFrame(overall_mismatch)
        if all_fields_mismatch:
            result_dfs['전체필드_불일치'] = pd.DataFrame(all_fields_mismatch)
        
        # 불일치 항목이 있는 경우에만 결과 파일 저장
        if result_dfs:
            output_file = output_path / f"{excel_path.stem}_비교결과.xlsx"
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df_result in result_dfs.items():
                    # 불일치 항목만 포함된 DataFrame 저장
                    df_result.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"  [완료] 불일치 항목만 저장: {output_file.name}")
            print(f"    - 관리번호 불일치: {len(management_mismatch)}개")
            print(f"    - 계약명 불일치: {len(contract_mismatch)}개")
            print(f"    - 진행상태 불일치: {len(status_mismatch)}개")
            print(f"    - 담당자 불일치: {len(manager_mismatch)}개")
            print(f"    - JSON 매칭 실패: {len(not_found)}개")
            print(f"    - 종합 불일치: {len(overall_mismatch)}개")
            print(f"    - 전체 필드 불일치: {len(all_fields_mismatch)}개")
        else:
            print(f"  [완료] 모든 데이터가 일치합니다! 결과 파일을 생성하지 않습니다.")
    
    except Exception as e:
        print(f"  [오류] {excel_path.name} 처리 실패: {e}")
        import traceback
        traceback.print_exc()


def main(argv: List[str]) -> None:
    """메인 함수"""
    # 경로 설정
    script_dir = Path(__file__).parent
    tonghab_dir = script_dir / "통합본"
    raw_data_dir = script_dir / "raw_data"
    output_dir = script_dir / "비교결과"
    
    # 출력 디렉토리 생성
    output_dir.mkdir(exist_ok=True)
    
    # 디렉토리 확인
    if not tonghab_dir.exists():
        print(f"오류: 통합본 폴더를 찾을 수 없습니다: {tonghab_dir}")
        return
    
    if not raw_data_dir.exists():
        print(f"오류: raw_data 폴더를 찾을 수 없습니다: {raw_data_dir}")
        return
    
    # JSON 데이터 로드
    json_data_map = load_json_data(raw_data_dir)
    
    if not json_data_map:
        print("오류: JSON 데이터를 찾을 수 없습니다.")
        return
    
    # 엑셀 파일 처리
    xlsx_files = list(tonghab_dir.glob("*.xlsx"))
    print(f"\n[엑셀 파일] 발견된 파일 수: {len(xlsx_files)}")
    
    if not xlsx_files:
        print("오류: 엑셀 파일을 찾을 수 없습니다.")
        return
    
    # 각 엑셀 파일 처리
    for excel_file in xlsx_files:
        compare_excel_with_json(excel_file, json_data_map, output_dir)
    
    print(f"\n[전체 완료] 모든 비교 작업이 완료되었습니다.")
    print(f"결과 파일은 {output_dir} 폴더에 저장되었습니다.")


if __name__ == "__main__":
    main(sys.argv)

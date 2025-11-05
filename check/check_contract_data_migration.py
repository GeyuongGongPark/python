"""로폼(원본) 시트 vs 로아이(웹 수집) 시트 컬럼 비교 리포트 생성 스크립트.

개요
- 입력: 문서비교.xlsx (시트: 로폼, 로아이)
- 처리: 관리번호/계약명/진행상태/상대계약자/요청자/검토담당자에 대해 Fuzzy 매칭
- 출력: 불일치 항목을 시트별로 정리한 Excel 파일(column_comparison_results.xlsx)
"""

import pandas as pd
from fuzzywuzzy import fuzz

# 엑셀 파일 경로
file_path = '문서비교.xlsx'  # 엑셀 파일 경로

# 엑셀 파일 읽기
sheet1 = pd.read_excel(file_path, sheet_name='로폼')  # 원본 데이터
sheet2 = pd.read_excel(file_path, sheet_name='로아이')  # 웹에서 받은 데이터

# 비교할 열 이름 설정 (실제 엑셀 파일의 열 이름에 맞게 수정)
column_name1 = '관리번호'  # 시트1의 열 이름
column_name2 = '계약명'  # 시트1의 계약명 열 이름
column_name2_web = '계약명 '  # 시트2의 열 이름 (공백 포함)
column_name3 = '진행 상태'
column_name4 = '상대 계약자'
column_name5 = '요청자'
column_name6 = '검토담당자'

# 일치율 계산 함수
def calculate_similarity(row):
    return fuzz.ratio(row[column_name1], row[column_name2])  # 두 셀 간의 유사성 비율 계산

# 각 컬럼별 비교 결과를 저장할 리스트들
management_mismatch = []  # 관리번호 불일치
contract_mismatch = []   # 계약명 불일치
status_mismatch = []     # 진행상태 불일치
contractor_mismatch = [] # 상대계약자 불일치
requester_mismatch = []  # 요청자 불일치
reviewer_mismatch = []   # 검토담당자 불일치
overall_mismatch = []    # 종합 불일치

# 유사성 임계값 설정 (이 값 이하는 "없는 정보"로 판단)
similarity_threshold = 80  # 80% 이하는 매칭되지 않은 것으로 간주

# 각 원본 데이터와 웹 데이터 비교
for index, row in sheet1.iterrows():
    # 웹 데이터에서 가장 유사한 데이터 찾기 (모든 컬럼 비교)
    def calculate_overall_similarity(web_row):
        # 각 컬럼별 유사성 계산
        management_similarity = fuzz.ratio(str(row[column_name1]), str(web_row['관리 번호']))
        contract_similarity = fuzz.ratio(str(row[column_name2]), str(web_row[column_name2_web]))
        status_similarity = fuzz.ratio(str(row[column_name3]), str(web_row[column_name3]))
        contractor_similarity = fuzz.ratio(str(row[column_name4]), str(web_row[column_name4]))
        requester_similarity = fuzz.ratio(str(row[column_name5]), str(web_row[column_name5]))
        reviewer_similarity = fuzz.ratio(str(row[column_name6]), str(web_row[column_name6]))
        
        # 종합 점수 계산 (가중평균)
        overall_score = (
            management_similarity * 0.3 +  # 관리번호 30%
            contract_similarity * 0.4 +    # 계약명 40%
            status_similarity * 0.1 +      # 진행 상태 10%
            contractor_similarity * 0.1 +   # 상대 계약자 10%
            requester_similarity * 0.05 +  # 요청자 5%
            reviewer_similarity * 0.05    # 검토담당자 5%
        )
        
        return {
            'overall_score': overall_score,
            'management_similarity': management_similarity,
            'contract_similarity': contract_similarity,
            'status_similarity': status_similarity,
            'contractor_similarity': contractor_similarity,
            'requester_similarity': requester_similarity,
            'reviewer_similarity': reviewer_similarity
        }
    
    # 모든 웹 데이터와 비교하여 최고 점수 찾기
    best_scores = sheet2.apply(calculate_overall_similarity, axis=1)
    best_match_index = best_scores.apply(lambda x: x['overall_score']).idxmax()
    best_match_data = best_scores.iloc[best_match_index]
    
    # 각 컬럼별로 불일치 데이터 분류
    base_data = {
        '관리번호': row[column_name1],
        '계약명': row[column_name2],
        '진행 상태': row[column_name3],
        '상대 계약자': row[column_name4],
        '요청자': row[column_name5],
        '검토담당자': row[column_name6],
        '계약 시작일': row['계약 시작일'],
        '계약 종료': row['계약 종료']
    }
    
    # 관리번호 불일치 (100% 미만)
    if best_match_data['management_similarity'] < 100:
        management_mismatch.append({
            **base_data,
            '유사성 점수': best_match_data['management_similarity'],
            '웹 데이터 관리번호': sheet2.at[best_match_index, '관리 번호']
        })
    
    # 계약명 불일치 (100% 미만)
    if best_match_data['contract_similarity'] < 100:
        contract_mismatch.append({
            **base_data,
            '유사성 점수': best_match_data['contract_similarity'],
            '웹 데이터 계약명': sheet2.at[best_match_index, column_name2_web]
        })
    
    # 진행상태 불일치 (100% 미만)
    if best_match_data['status_similarity'] < 100:
        status_mismatch.append({
            **base_data,
            '유사성 점수': best_match_data['status_similarity'],
            '웹 데이터 진행상태': sheet2.at[best_match_index, column_name3]
        })
    
    # 상대계약자 불일치 (100% 미만)
    if best_match_data['contractor_similarity'] < 100:
        contractor_mismatch.append({
            **base_data,
            '유사성 점수': best_match_data['contractor_similarity'],
            '웹 데이터 상대계약자': sheet2.at[best_match_index, column_name4]
        })
    
    # 요청자 불일치 (100% 미만)
    if best_match_data['requester_similarity'] < 100:
        requester_mismatch.append({
            **base_data,
            '유사성 점수': best_match_data['requester_similarity'],
            '웹 데이터 요청자': sheet2.at[best_match_index, column_name5]
        })
    
    # 검토담당자 불일치 (100% 미만)
    if best_match_data['reviewer_similarity'] < 100:
        reviewer_mismatch.append({
            **base_data,
            '유사성 점수': best_match_data['reviewer_similarity'],
            '웹 데이터 검토담당자': sheet2.at[best_match_index, column_name6]
        })
    
    # 종합 불일치 (종합 점수 100% 미만)
    if best_match_data['overall_score'] < 100:
        overall_mismatch.append({
            **base_data,
            '종합 유사성 점수': round(best_match_data['overall_score'], 2),
            '관리번호 유사성': best_match_data['management_similarity'],
            '계약명 유사성': best_match_data['contract_similarity'],
            '진행상태 유사성': best_match_data['status_similarity'],
            '상대계약자 유사성': best_match_data['contractor_similarity'],
            '요청자 유사성': best_match_data['requester_similarity'],
            '검토담당자 유사성': best_match_data['reviewer_similarity']
        })

# 각 리스트를 DataFrame으로 변환
management_df = pd.DataFrame(management_mismatch)
contract_df = pd.DataFrame(contract_mismatch)
status_df = pd.DataFrame(status_mismatch)
contractor_df = pd.DataFrame(contractor_mismatch)
requester_df = pd.DataFrame(requester_mismatch)
reviewer_df = pd.DataFrame(reviewer_mismatch)
overall_df = pd.DataFrame(overall_mismatch)

# 결과를 여러 시트로 나누어 엑셀 파일로 저장
with pd.ExcelWriter('column_comparison_results.xlsx', engine='openpyxl') as writer:
    management_df.to_excel(writer, sheet_name='관리번호_불일치', index=False)
    contract_df.to_excel(writer, sheet_name='계약명_불일치', index=False)
    status_df.to_excel(writer, sheet_name='진행상태_불일치', index=False)
    contractor_df.to_excel(writer, sheet_name='상대계약자_불일치', index=False)
    requester_df.to_excel(writer, sheet_name='요청자_불일치', index=False)
    reviewer_df.to_excel(writer, sheet_name='검토담당자_불일치', index=False)
    overall_df.to_excel(writer, sheet_name='종합_불일치', index=False)

print(f"각 컬럼별 비교 결과가 'column_comparison_results.xlsx' 파일로 저장되었습니다.")
print(f"관리번호 불일치: {len(management_mismatch)}개")
print(f"계약명 불일치: {len(contract_mismatch)}개")
print(f"진행상태 불일치: {len(status_mismatch)}개")
print(f"상대계약자 불일치: {len(contractor_mismatch)}개")
print(f"요청자 불일치: {len(requester_mismatch)}개")
print(f"검토담당자 불일치: {len(reviewer_mismatch)}개")
print(f"종합 불일치: {len(overall_mismatch)}개")
print("각 시트별로 불일치 데이터를 확인할 수 있습니다.")

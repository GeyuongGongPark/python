import pandas as pd
import time
import json
import csv
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import os
from datetime import datetime

def parse_custom_env():
    """.env 파일을 직접 파싱 (KEY : VALUE 형식 지원)"""
    env_vars = {}
    
    try:
        # .env 파일 읽기
        with open('.env', 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                
                # 주석이나 빈 줄 무시
                if not line or line.startswith('#'):
                    continue
                
                # KEY : VALUE 형식 파싱
                if ':' in line:
                    key, value = line.split(':', 1)
                    key = key.strip()
                    value = value.strip()
                    
                    # 따옴표 제거
                    value = value.strip('"').strip("'")
                    
                    env_vars[key] = value
                    
        return env_vars
    except Exception as e:
        print(f"⚠ .env 파일 읽기 실패: {e}")
        return {}

# .env 파일 파싱
env_vars = parse_custom_env()

# 환경변수 가져오기 (줄 시작부터 공백 없이 매칭)
BASE_URL = {
    'PRODUCTION': env_vars.get('prod_BASE_URL', '').strip() or env_vars.get('dev_BASE_URL', '').strip(),
}

class ContractComparator:
    def __init__(self):
        self.driver = None
        self.contract_data = []
        
    def setup_driver(self):
        """Chrome 드라이버 설정"""
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        # 헤드리스 모드 비활성화 (디버깅을 위해)
        # chrome_options.add_argument("--headless")
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            print("✓ Chrome 드라이버가 성공적으로 설정되었습니다.")
            return True
        except Exception as e:
            print(f"✗ Chrome 드라이버 설정 실패: {str(e)}")
            return False
    
    def login(self, username, password):
        """웹사이트 로그인"""
        try:
            print("로그인 시도 중...")
            self.driver.get(BASE_URL['PRODUCTION'])
            
            # 페이지 로딩 대기
            time.sleep(3)
            
            # 현재 페이지 정보 출력
            print(f"현재 URL: {self.driver.current_url}")
            print(f"페이지 제목: {self.driver.title}")
            
            # 로그인 페이지 로딩 대기 (최적화: k6에서 확인된 셀렉터 우선 사용)
            wait = WebDriverWait(self.driver, 10)  # 대기 시간 단축
            
            # ID 입력 필드 셀렉터 (k6 성능 측정에서 확인된 셀렉터 우선 사용)
            # 첫 번째 셀렉터로 바로 찾기 시도
            try:
                id_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']")))
                print("✓ ID 필드 발견: input[type='email']")
            except:
                # 폴백: 다른 셀렉터 시도
                id_selectors = [
                    (By.XPATH, "//input[@type='email']"),
                    (By.NAME, "username"),
                    (By.NAME, "email"),
                    (By.ID, "username"),
                ]
                id_field = None
                for selector_type, selector_value in id_selectors:
                    try:
                        id_field = wait.until(EC.presence_of_element_located((selector_type, selector_value)))
                        print(f"✓ ID 필드 발견: {selector_type} = {selector_value}")
                        break
                    except:
                        continue
            
            if not id_field:
                print("✗ ID 입력 필드를 찾을 수 없습니다.")
                # 페이지 소스 일부 출력
                print("페이지 소스 일부:")
                print(self.driver.page_source[:1000])
                return False
            
            # ID 입력
            id_field.clear()
            id_field.send_keys(username)
            print("✓ ID 입력 완료")
            
            # 비밀번호 입력 필드 (k6 성능 측정에서 확인된 셀렉터 우선 사용)
            try:
                pw_field = self.driver.find_element(By.CSS_SELECTOR, "input[type='password']")
                print("✓ 비밀번호 필드 발견: input[type='password']")
            except:
                # 폴백: 다른 셀렉터 시도
                pw_selectors = [
                    (By.XPATH, "//input[@type='password']"),
                    (By.NAME, "password"),
                    (By.ID, "password"),
                ]
                pw_field = None
                for selector_type, selector_value in pw_selectors:
                    try:
                        pw_field = self.driver.find_element(selector_type, selector_value)
                        print(f"✓ 비밀번호 필드 발견: {selector_type} = {selector_value}")
                        break
                    except:
                        continue
            
            if not pw_field:
                print("✗ 비밀번호 입력 필드를 찾을 수 없습니다.")
                return False
            
            # 비밀번호 입력
            pw_field.clear()
            pw_field.send_keys(password)
            print("✓ 비밀번호 입력 완료")
            
            # 로그인 버튼 (k6 성능 측정에서 확인된 셀렉터 우선 사용)
            try:
                login_button = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
                print("✓ 로그인 버튼 발견: button[type='submit']")
            except:
                # 폴백: 다른 셀렉터 시도
                login_button_selectors = [
                    (By.XPATH, "//button[@type='submit']"),
                    (By.XPATH, "//input[@type='submit']"),
                    (By.CSS_SELECTOR, "input[type='submit']"),
                    (By.XPATH, "//button[contains(text(), '로그인')]"),
                ]
                login_button = None
                for selector_type, selector_value in login_button_selectors:
                    try:
                        login_button = self.driver.find_element(selector_type, selector_value)
                        print(f"✓ 로그인 버튼 발견: {selector_type} = {selector_value}")
                        break
                    except:
                        continue
            
            if not login_button:
                print("✗ 로그인 버튼을 찾을 수 없습니다.")
                return False
            
            # 로그인 버튼 클릭
            login_button.click()
            print("✓ 로그인 버튼 클릭 완료")
            
            # 로그인 성공 확인 (대기 시간 최소화)
            time.sleep(2)  # 3초 -> 2초로 단축
            
            # URL 변경 확인
            current_url = self.driver.current_url
            print(f"로그인 후 URL: {current_url}")
            
            # 대시보드나 메인 페이지 요소 확인
            success_indicators = [
                (By.CLASS_NAME, "dashboard"),
                (By.CLASS_NAME, "main"),
                (By.CLASS_NAME, "home"),
                (By.XPATH, "//a[contains(text(), '로그아웃')]"),
                (By.XPATH, "//a[contains(text(), 'Logout')]"),
                (By.XPATH, "//div[contains(@class, 'user')]"),
                (By.XPATH, "//div[contains(@class, 'profile')]")
            ]
            
            login_success = False
            for selector_type, selector_value in success_indicators:
                try:
                    element = self.driver.find_element(selector_type, selector_value)
                    print(f"✓ 로그인 성공 확인: {selector_type} = {selector_value}")
                    login_success = True
                    break
                except:
                    continue
            
            if not login_success:
                # URL이 변경되었거나 로그인 페이지가 아닌 경우 성공으로 간주
                if "login" not in current_url.lower() and "signin" not in current_url.lower():
                    print("✓ URL 변경으로 로그인 성공 확인")
                    login_success = True
            
            if login_success:
                print("✓ 로그인이 성공적으로 완료되었습니다.")
                return True
            else:
                print("✗ 로그인 실패 - 성공 지표를 찾을 수 없습니다.")
                return False
            
        except TimeoutException:
            print("✗ 로그인 시간 초과")
            return False
        except Exception as e:
            print(f"✗ 로그인 실패: {str(e)}")
            return False
    
    def navigate_to_contracts(self):
        """체결 계약서 조회 메뉴로 이동"""
        try:
            print("체결 계약서 조회 페이지로 이동 중...")
            contract_url = BASE_URL['PRODUCTION'] + "/clm/complete?page=0"
            print(f"URL: {contract_url}")
            
            self.driver.get(contract_url)
            time.sleep(3)
            
            current_url = self.driver.current_url
            print(f"현재 URL: {current_url}")
            
            if "clm/complete" in current_url:
                print("✓ 체결 계약서 조회 페이지로 이동 성공")
                return True
            else:
                print(f"✗ 체결 계약서 페이지로 이동 실패")
                return False
                
        except Exception as e:
            print(f"✗ 메뉴 이동 실패: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def extract_current_page_contracts(self):
        """현재 페이지의 계약서 추출"""
        try:
            wait = WebDriverWait(self.driver, 10)
            time.sleep(2)
            
            # 계약서 목록 테이블 찾기
            table_selectors = [
                "//table",
                "//div[contains(@class, 'table')]",
                "//div[contains(@class, 'list')]",
                "//div[contains(@class, 'contract')]"
            ]
            
            table = None
            for selector in table_selectors:
                try:
                    table = self.driver.find_element(By.XPATH, selector)
                    break
                except:
                    continue
            
            if not table:
                return []
            
            rows = table.find_elements(By.XPATH, ".//tr")
            if len(rows) <= 1:
                return []
            
            headers = []
            header_row = rows[0]
            header_cells = header_row.find_elements(By.XPATH, ".//th | .//td")
            for cell in header_cells:
                headers.append(cell.text.strip())
            
            contract_list = []
            for i, row in enumerate(rows[1:], 1):
                try:
                    cells = row.find_elements(By.XPATH, ".//td")
                    if len(cells) == 0:
                        continue
                    
                    row_data = {}
                    for j, cell in enumerate(cells):
                        if j < len(headers):
                            row_data[headers[j]] = cell.text.strip()
                    
                    try:
                        link_element = row.find_element(By.XPATH, ".//a")
                        row_data['link'] = link_element.get_attribute('href')
                    except:
                        row_data['link'] = None
                    
                    contract_list.append(row_data)
                except:
                    continue
            
            return contract_list
        except:
            return []
    
    def extract_contract_list(self):
        """모든 페이지의 계약서 추출 (page 파라미터 사용)"""
        try:
            print("\n=== 계약서 목록 추출 시작 ===")
            
            all_contracts = []
            page_num = 0
            empty_page_count = 0  # 빈 페이지 연속 카운트
            
            while True:
                print(f"\n--- page={page_num} 추출 중 ---")
                
                # 현재 페이지 URL
                current_url = f"{BASE_URL['PRODUCTION']}/clm/complete?page={page_num}"
                self.driver.get(current_url)
                time.sleep(3)
                
                print(f"URL: {current_url}")
                print(f"현재 페이지 URL: {self.driver.current_url}")
                
                # 먼저 "등록된 내용이 없습니다" 메시지 확인
                try:
                    page_text = self.driver.find_element(By.TAG_NAME, "body").text
                    if "등록된 내용이 없습니다" in page_text:
                        print(f"  ⚠ '등록된 내용이 없습니다' 메시지 발견. 추출 종료.")
                        break
                except:
                    pass
                
                # 추가 안전장치: 다양한 "데이터 없음" 메시지 확인
                try:
                    no_data_keywords = [
                        "등록된 내용이 없습니다",
                        "데이터가 없습니다",
                        "no data available",
                        "등록된 계약서가 없습니다"
                    ]
                    
                    for keyword in no_data_keywords:
                        elements = self.driver.find_elements(By.XPATH, f"//*[contains(text(), '{keyword}')]")
                        if elements:
                            for elem in elements:
                                if keyword in elem.text:
                                    print(f"  ⚠ 데이터 없음 메시지 발견: '{elem.text}'. 추출 종료.")
                                    return all_contracts
                except:
                    pass
                
                # 현재 페이지 계약서 추출
                current_contracts = self.extract_current_page_contracts()
                
                print(f"  → 추출 결과: {len(current_contracts)}개")
                
                # 빈 페이지 체크
                if not current_contracts or len(current_contracts) == 0:
                    empty_page_count += 1
                    print(f"  ⚠ 빈 페이지 감지 (연속 {empty_page_count}번)")
                    
                    # 빈 페이지가 2번 연속 나오면 종료
                    if empty_page_count >= 2:
                        print(f"  ✓ 빈 페이지 2번 연속 확인. 추출 종료.")
                        break
                else:
                    # 계약서가 있으면 빈 페이지 카운트 리셋
                    empty_page_count = 0
                    print(f"  ✓ {len(current_contracts)}개 추출")
                    all_contracts.extend(current_contracts)
                
                # 최대 100페이지 제한
                if page_num >= 100:
                    print("⚠ 최대 페이지 수 도달")
                    break
                
                page_num += 1
            
            print(f"\n✓ 총 {len(all_contracts)}개 계약서 추출 완료")
            return all_contracts
            
        except Exception as e:
            print(f"✗ 계약서 목록 추출 실패: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def _parse_contract_info(self, text):
        """계약 정보 영역 파싱"""
        data = {}
        
        # 각 필드를 키워드로 찾아서 추출
        field_patterns = {
            '관리번호': ['관리번호', '관리 번호'],
            '계약명': ['계약명', '계약 명', '계약서명'],
            '계약 분류': ['계약 분류', '분류'],
            '체결계약서 사본': ['체결계약서 사본', '사본'],
            '원본 보관 위치': ['원본 보관 위치', '원본 보관', '보관 위치'],
            '요청자': ['요청자', '요청인'],
            '계약 기간': ['계약 기간', '기간'],
            '계약 자동 연장 여부': ['자동 연장', '연장 여부'],
            '보안여부': ['보안여부', '보안 여부', '보안'],
            '서면 실태 조사': ['서면 실태 조사', '실태 조사'],
            '연관 계약': ['연관 계약', '연관'],
            '첨부/별첨': ['첨부', '별첨'],
            '상대 계약자 정보': ['상대 계약자', '계약자 정보'],
            '참조 수신자 정보': ['참조', '수신자']
        }
        
        lines = text.split('\n')
        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            
            for field_name, keywords in field_patterns.items():
                for keyword in keywords:
                    if keyword in line:
                        # 해당 라인의 값 추출 (콜론, 공백, 탭 등으로 구분)
                        parts = line.split(':', 1)
                        if len(parts) > 1:
                            value = parts[1].strip()
                            
                            # 특별 처리
                            if field_name == '계약 분류':
                                # 계약 분류 파싱 개선
                                # 예: "[중분류명](대분류명 > [중분류명])" 또는 "대분류명 > 중분류명"
                                print(f"      계약 분류 원본: {value}")
                                
                                # 1. 괄호 안의 내용 추출 시도
                                bracket_match = re.search(r'\((.*?)\)', value)
                                if bracket_match:
                                    bracket_content = bracket_match.group(1).strip()  # "대분류명 > [중분류명]"
                                    print(f"      괄호 내용: {bracket_content}")
                                    
                                    # 중괄호 제거
                                    bracket_content = bracket_content.replace('[', '').replace(']', '').strip()
                                    
                                    # ">" 또는 " > "로 분리 시도
                                    if '>' in bracket_content:
                                        parts = re.split(r'\s*>\s*', bracket_content, 1)
                                        data['계약분류_대분류'] = parts[0].strip()
                                        data['계약분류_중분류'] = parts[1].strip() if len(parts) > 1 else ''
                                    else:
                                        # 괄호는 있지만 > 없는 경우
                                        data['계약분류_대분류'] = bracket_content.strip()
                                        data['계약분류_중분류'] = ''
                                
                                # 2. 괄호 없이 ">" 로 구분된 경우
                                elif '>' in value:
                                    value_clean = value.replace('[', '').replace(']', '').replace('(', '').replace(')', '').strip()
                                    parts = re.split(r'\s*>\s*', value_clean, 1)
                                    data['계약분류_대분류'] = parts[0].strip()
                                    data['계약분류_중분류'] = parts[1].strip() if len(parts) > 1 else ''
                                
                                # 3. 그 외의 경우
                                else:
                                    value_clean = value.replace('[', '').replace(']', '').replace('(', '').replace(')', '').strip()
                                    data['계약분류_대분류'] = value_clean
                                    data['계약분류_중분류'] = ''
                                
                                print(f"      파싱 결과 - 대분류: {data.get('계약분류_대분류')}, 중분류: {data.get('계약분류_중분류')}")
                            
                            elif field_name == '요청자' and '?' in value:
                                # "팀 ? 이름" 형태 파싱
                                parts = value.split('?')
                                data['요청자_팀'] = parts[0].strip()
                                data['요청자_이름'] = parts[1].strip() if len(parts) > 1 else ''
                            
                            elif field_name == '계약 기간' and '~' in value:
                                # "시작일 ~ 종료일" 형태 파싱
                                parts = value.split('~')
                                data['계약기간_시작일'] = parts[0].strip()
                                data['계약기간_종료일'] = parts[1].strip() if len(parts) > 1 else ''
                            
                            elif field_name == '계약 자동 연장 여부':
                                if value.lower() == 'yes' or value.lower() == '예':
                                    data['자동연장_여부'] = 'Yes'
                                    if '?' in value or '/' in value:
                                        data['자동연장_코멘트'] = value.split('/')[-1].strip() if '/' in value else ''
                                else:
                                    data['자동연장_여부'] = value
                            
                            else:
                                data[field_name] = value
                        break
        
        return data
    
    def _parse_detail_info(self, text):
        """상세 정보 영역 파싱"""
        data = {}
        
        # 각 필드를 키워드로 찾아서 추출
        field_patterns = {
            '계약 체결일': ['계약 체결일', '체결일', '체결 일자'],
            '계약규모': ['계약규모', '계약 규모', '규모'],
            '지급 상세': ['지급 상세', '지급'],
            '계약 배경/목적': ['계약 배경', '배경/목적', '목적', '배경'],
            '주요 협의사항': ['주요 협의사항', '협의사항', '협의 사항']
        }
        
        lines = text.split('\n')
        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            
            for field_name, keywords in field_patterns.items():
                for keyword in keywords:
                    if keyword in line:
                        # 해당 라인의 값 추출
                        parts = line.split(':', 1)
                        if len(parts) > 1:
                            data[field_name] = parts[1].strip()
                        break
        
        return data
    
    def _parse_contract_info_special(self, data):
        """추출된 계약 정보에서 특별 파싱 수행"""
        parsed = {}
        
        # 계약 분류 파싱
        if '계약 분류' in data and data['계약 분류']:
            value = data['계약 분류']
            print(f"      계약 분류 원본: {value}")
            
            # 괄호 안의 내용 추출 시도
            bracket_match = re.search(r'\((.*?)\)', value)
            if bracket_match:
                bracket_content = bracket_match.group(1).strip()
                bracket_content = bracket_content.replace('[', '').replace(']', '').strip()
                
                if '>' in bracket_content:
                    parts = re.split(r'\s*>\s*', bracket_content, 1)
                    parsed['계약분류_대분류'] = parts[0].strip()
                    parsed['계약분류_중분류'] = parts[1].strip() if len(parts) > 1 else ''
                else:
                    parsed['계약분류_대분류'] = bracket_content.strip()
                    parsed['계약분류_중분류'] = ''
            elif '>' in value:
                value_clean = value.replace('[', '').replace(']', '').replace('(', '').replace(')', '').strip()
                parts = re.split(r'\s*>\s*', value_clean, 1)
                parsed['계약분류_대분류'] = parts[0].strip()
                parsed['계약분류_중분류'] = parts[1].strip() if len(parts) > 1 else ''
            else:
                parsed['계약분류_대분류'] = value
                parsed['계약분류_중분류'] = ''
            
            print(f"      파싱 결과 - 대분류: {parsed.get('계약분류_대분류')}, 중분류: {parsed.get('계약분류_중분류')}")
        
        # 요청자 파싱
        if '요청자' in data and data['요청자']:
            value = data['요청자']
            if '/' in value:
                parts = value.split('/')
                if len(parts) >= 2:
                    parsed['요청자_팀'] = parts[0].strip()
                    parsed['요청자_이름'] = parts[-1].strip()
        
        # 계약 기간 파싱
        if '계약 기간' in data and data['계약 기간']:
            value = data['계약 기간']
            if '~' in value:
                parts = value.split('~')
                parsed['계약기간_시작일'] = parts[0].strip()
                parsed['계약기간_종료일'] = parts[1].strip() if len(parts) > 1 else ''
        
        # 계약 자동 연장 여부
        if '계약 자동 연장 여부' in data and data['계약 자동 연장 여부']:
            value = data['계약 자동 연장 여부']
            if value.upper() in ['YES', '예', 'Y']:
                parsed['자동연장_여부'] = 'Yes'
            else:
                parsed['자동연장_여부'] = 'No'
        
        return parsed
    
    def _parse_detail_info_special(self, data):
        """추출된 상세 정보에서 특별 파싱 수행"""
        parsed = {}
        
        # 계약 체결일
        if '계약 체결일' in data:
            parsed['계약체결일'] = data['계약 체결일']
        
        # 계약 규모
        if '계약 규모' in data:
            parsed['계약규모'] = data['계약 규모']
        
        # 지급 상세
        if '지급 상세' in data:
            parsed['지급상세'] = data['지급 상세']
        
        # 계약 배경/목적
        if '계약 배경/목적' in data:
            parsed['계약배경_목적'] = data['계약 배경/목적']
        
        # 주요 협의사항
        if '주요 협의사항' in data:
            parsed['주요협의사항'] = data['주요 협의사항']
        
        return parsed
    
    def _extract_table_key_values(self, table_element):
        """테이블에서 각 행의 th → td 1:1 매핑으로 키-값을 안전 추출"""
        result = {}
        try:
            rows = table_element.find_elements(By.XPATH, ".//tr")
            for row in rows:
                try:
                    th_elements = row.find_elements(By.XPATH, ".//th")
                    td_elements = row.find_elements(By.XPATH, ".//td")
                    if len(th_elements) >= 1 and len(td_elements) >= 1:
                        key = th_elements[0].text.strip()
                        value = td_elements[0].text.strip()
                        if key:
                            result[key] = value
                except Exception:
                    continue
        except Exception:
            pass
        return result
    
    def _map_to_template_format(self, data):
        """추출된 데이터를 양식 파일 구조에 맞게 매핑"""
        if not data:
            return {}
        
        mapped = {}
        
        # 관리 번호
        mapped['관리 번호'] = data.get('관리번호', '')
        
        # 계약명
        mapped['계약명 '] = data.get('계약명', '')
        
        # 대분류, 분류 (계약 분류 파싱 결과)
        mapped['대분류'] = data.get('계약분류_대분류', '')
        mapped['분류'] = data.get('계약분류_중분류', '')
        
        # 계약 시작일, 계약 완료일
        mapped['계약 시작일'] = data.get('계약기간_시작일', '')
        mapped['계약 완료일'] = data.get('계약기간_종료일', '')
        
        # 상대 계약자
        mapped['상대 계약자'] = data.get('상대 계약자 정보', '')
        
        # 원본 보관 위치
        mapped['원본 보관 위치'] = data.get('원본 보관 위치', '')
        
        # 보안여부
        mapped['보안여부'] = data.get('보안여부', '')
        
        # 연관계약
        mapped['연관계약'] = data.get('연관 계약', '')
        
        # 계약 규모 (계약 규모 파싱)
        if '계약 규모' in data:
            # "10,000,000 / KRW 한국 / 부가세(10%) 별도" 형태 파싱
            value = data['계약 규모']
            parts = value.split(' / ')
            if len(parts) >= 2:
                mapped['계약 규모'] = parts[0].strip()
                mapped['통화'] = parts[1].strip()
                if len(parts) > 2:
                    mapped['계약규모 코멘트'] = parts[2].strip()
            else:
                mapped['계약 규모'] = value
        else:
            mapped['계약 규모'] = ''
            mapped['통화'] = ''
        
        # 주요 협의사항
        mapped['주요 협의사항'] = data.get('주요 협의사항', '')
        mapped['주요 협의사항_원본'] = data.get('주요 협의사항', '')
        
        # 계약의 배경 및 목적
        mapped['계약의 배경 및 목적'] = data.get('계약 배경/목적', '')
        mapped['계약의 배경 및 목적_원본'] = data.get('계약 배경/목적', '')
        
        # 계약 체결일
        mapped['계약 체결일'] = data.get('계약 체결일', '')
        
        # 자동 연장 여부
        mapped['자동 연장 여부'] = data.get('계약 자동 연장 여부', '')
        if '자동연장_여부' in data:
            mapped['자동 연장 여부'] = data.get('자동연장_여부', mapped['자동 연장 여부'])
        
        # 자동연장 코멘트
        mapped['통지(코멘트)'] = data.get('자동연장_코멘트', '')
        
        # 요청자 정보 (검토 요청자로 매핑)
        if '요청자_팀' in data or '요청자_이름' in data:
            mapped['검토 요청자 이름'] = data.get('요청자_이름', '')
        
        # 연관 문서 (첨부/별첨)
        mapped['관련문서'] = data.get('첨부/별첨', '')
        
        # 체결계약서 사본
        mapped['계약서 첨부 파일'] = data.get('체결 계약서 사본', '')
        
        return mapped
    
    def extract_contract_details(self, contract):
        """개별 계약서 상세 내용 추출 (재시도 로직 포함, 불필요한 텍스트 제거)"""
        if not contract.get('link'):
            return {}
        
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                print(f"  → 이동 URL: {contract['link']} (시도 {retry_count + 1}/{max_retries})")
                
                # 타임아웃 증가
                self.driver.set_page_load_timeout(180)
                
                self.driver.get(contract['link'])
                
                # 페이지 로딩 대기
                wait = WebDriverWait(self.driver, 30)
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                
                # 테이블이 로드될 때까지 대기 (최대 10초)
                for i in range(10):
                    time.sleep(1)
                    all_tables = self.driver.find_elements(By.TAG_NAME, "table")
                    if len(all_tables) > 0:
                        print(f"    → 테이블 로딩 완료: {len(all_tables)}개 (대기 {i+1}초)")
                        break
                    if i >= 4:  # 5초 후에도 없으면 조건부 대기
                        # 다른 데이터 구조 확인
                        main_exists = len(self.driver.find_elements(By.TAG_NAME, "main")) > 0
                        if main_exists:
                            break
                
                # 최종 테이블 개수 확인
                all_tables = self.driver.find_elements(By.TAG_NAME, "table")
                print(f"    → 페이지에서 {len(all_tables)}개의 테이블 발견")
                
                details = {}
                
                # 계약 정보 영역 추출 (테이블에서 직접 데이터 추출)
                try:
                    # 계약 정보 테이블 찾기 (첫 번째 테이블)
                    contract_table = None
                    
                    # 방법 1: 첫 번째 테이블
                    if len(all_tables) >= 1:
                        contract_table = all_tables[0]  # 인덱스 0 = 첫 번째 테이블
                        print(f"    ✓ 계약 정보 테이블 발견: 첫 번째 테이블 ({len(all_tables)}개 중)")
                    else:
                        # 방법 2: XPath로 시도
                        contract_xpaths = [
                            "//table[position()=1]",  # 첫 번째 테이블
                            "//table[contains(@class, 'border-spacing-0')]",  # 클래스 기반
                            "/html/body/div/div[1]/div[3]/div[2]/main/div/div[2]/div[2]/div[1]/div[1]/div[2]/table",  # 절대 경로
                            "//main//table[1]",  # 첫 번째 main의 첫 번째 테이블
                            "//div[contains(@class, '계약 정보')]//table",
                        ]
                        
                        for xpath in contract_xpaths:
                            try:
                                contract_table = self.driver.find_element(By.XPATH, xpath)
                                print(f"    ✓ 계약 정보 테이블 발견: {xpath[:80]}")
                                break
                            except:
                                continue
                    
                    if contract_table:
                        # 테이블의 각 행을 순회하면서 Key-Value 추출 (th→td 1:1)
                        kv = self._extract_table_key_values(contract_table)
                        print(f"    → {len(kv)}개 항목 추출")
                        details.update(kv)
                        # 특별 파싱 적용
                        details.update(self._parse_contract_info_special(details))
                    else:
                        print("    ⚠ 계약 정보 테이블을 찾을 수 없습니다.")
                        
                except Exception as e:
                    print(f"    ⚠ 계약 정보 영역 찾기 실패: {str(e)[:100]}")
                    import traceback
                    traceback.print_exc()
                
                # 상세 정보 영역 추출 (테이블에서 직접 데이터 추출)
                try:
                    # 상세 정보 테이블 찾기 (두 번째 테이블)
                    detail_table = None
                    
                    # 방법 1: 두 번째 테이블 (position 기반)
                    if len(all_tables) >= 2:
                        detail_table = all_tables[1]  # 인덱스 1 = 두 번째 테이블
                        print(f"    ✓ 상세 정보 테이블 발견: 두 번째 테이블 ({len(all_tables)}개 중)")
                    else:
                        # 방법 2: XPath로 시도
                        detail_xpaths = [
                            "//table[position()=2]",  # 두 번째 테이블
                            "//table[contains(@class, 'w-full')]",  # 클래스 기반
                            "/html/body/div/div[1]/div[3]/div[2]/main/div/div[2]/div[2]/div[1]/div[2]/div[2]/table",  # 절대 경로
                            "//main//table[2]",  # 첫 번째 main의 두 번째 테이블
                        ]
                        
                        for xpath in detail_xpaths:
                            try:
                                detail_table = self.driver.find_element(By.XPATH, xpath)
                                print(f"    ✓ 상세 정보 테이블 발견: {xpath[:80]}")
                                break
                            except:
                                continue
                    
                    if detail_table:
                        # 테이블의 각 행을 순회하면서 Key-Value 추출 (th→td 1:1)
                        kv = self._extract_table_key_values(detail_table)
                        print(f"    → {len(kv)}개 항목 추출")
                        details.update(kv)
                        # 특별 파싱 적용
                        details.update(self._parse_detail_info_special(details))
                    else:
                        print("    ⚠ 상세 정보 테이블을 찾을 수 없습니다.")
                        
                except Exception as e:
                    print(f"    ⚠ 상세 정보 영역 찾기 실패: {str(e)[:100]}")
                    import traceback
                    traceback.print_exc()
                
                # 페이지로 돌아가기
                self.driver.back()
                time.sleep(2)
                
                # 양식 파일 구조에 맞게 매핑 (계약명 안전 보정 포함)
                # 계약명 보정: '요청자' 등 잘못 들어가는 경우 방지
                if '계약명' in details and details.get('계약명'):
                    title_val = details['계약명']
                    # 비정상 패턴 필터링
                    suspicious_keywords = ['요청자', '검토 요청', '참조', '수신자']
                    if any(kw in title_val for kw in suspicious_keywords):
                        # 대안: 페이지 타이틀 혹은 링크 텍스트 재시도
                        try:
                            h_candidates = self.driver.find_elements(By.XPATH, "//main//h1 | //main//h2 | //h1 | //h2")
                            for h in h_candidates:
                                txt = h.text.strip()
                                if txt and not any(kw in txt for kw in suspicious_keywords):
                                    details['계약명'] = txt
                                    break
                        except Exception:
                            pass
                
                mapped_details = self._map_to_template_format(details)
                
                # 원본 데이터를 _original에 저장하고 매핑된 데이터를 추가
                if details:
                    original_data = {k: v for k, v in details.items()}
                    details.update(mapped_details)
                    details['_original_data'] = original_data
                
                return details
                
            except Exception as e:
                retry_count += 1
                error_msg = str(e)
                print(f"  ✗ 상세 추출 실패: {error_msg[:100]} (시도 {retry_count}/{max_retries})")
                
                if retry_count >= max_retries:
                    print(f"  ⚠ 최대 재시도 횟수 초과. 스킵합니다.")
                    try:
                        self.driver.back()
                        time.sleep(2)
                    except:
                        pass
                    return {'content': f'추출 실패 (재시도 {max_retries}회 초과): {error_msg[:100]}'}
                else:
                    time.sleep(2)
                    try:
                        self.driver.back()
                        time.sleep(1)
                    except:
                        pass
        
        return {'content': '추출 실패'}
    
    def save_data(self, timestamp=None, mode='w'):
        """추출된 데이터를 파일로 저장 (실시간 저장 지원)"""
        try:
            if not self.contract_data:
                print("⚠ 저장할 데이터가 없습니다.")
                return False
            
            if timestamp is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # CSV 저장
            csv_filename = f"contract_data_{timestamp}.csv"
            if self.contract_data:
                all_keys = set()
                for item in self.contract_data:
                    all_keys.update(item.keys())
                
                # 파일이 존재하면 헤더 생략
                file_exists = os.path.exists(csv_filename) and mode == 'a'
                
                with open(csv_filename, mode, newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=sorted(all_keys))
                    
                    # 헤더는 새 파일이거나 덮어쓰기 모드일 때만
                    if not file_exists:
                        writer.writeheader()
                    
                    for contract in self.contract_data:
                        row = {}
                        for key in sorted(all_keys):
                            value = contract.get(key, '')
                            row[key] = value
                        writer.writerow(row)
                
                if mode == 'w':
                    print(f"✓ CSV 파일 저장: {csv_filename}")
                else:
                    print(f"✓ CSV 파일 추가 저장: {csv_filename} ({len(self.contract_data)}개)")
            
            # Excel 저장 - 템플릿 컬럼 구조에 맞춰 적재
            template_path = "데이터 추출 양식.xlsx"
            template_excel_filename = f"데이터 추출 결과_{timestamp}.xlsx"
            try:
                # 템플릿 컬럼 로드
                template_df = pd.read_excel(template_path)
                template_cols = list(template_df.columns)
                # 템플릿 컬럼 기준으로 행 구성
                rows = []
                for contract in self.contract_data:
                    row = {col: contract.get(col, '') for col in template_cols}
                    rows.append(row)
                out_df = pd.DataFrame(rows, columns=template_cols)
                out_df.to_excel(template_excel_filename, index=False, engine='openpyxl')
                print(f"✓ 템플릿 기반 Excel 저장: {template_excel_filename} ({len(self.contract_data)}개)")
            except Exception as e:
                # 템플릿 저장 실패 시 일반 저장으로 폴백
                excel_filename = f"contract_data_{timestamp}.xlsx"
                df = pd.DataFrame(self.contract_data)
                df.to_excel(excel_filename, index=False, engine='openpyxl')
                print(f"⚠ 템플릿 저장 실패로 일반 Excel 저장: {excel_filename} - {e}")
            
            return True
            
        except Exception as e:
            print(f"✗ 데이터 저장 실패: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def run_full_process(self, username, password):
        """전체 프로세스 실행 - 페이지별로 계약서 상세 추출"""
        try:
            print("=== 계약서 데이터 추출 프로세스 시작 ===")
            
            # 1. 드라이버 설정
            if not self.setup_driver():
                return False
            
            # 2. 로그인
            if not self.login(username, password):
                return False
            
            # 3. 계약서 조회 페이지로 이동
            if not self.navigate_to_contracts():
                return False
            
            # 4. 페이지별로 계약서 링크 추출 및 상세 내용 추출 (실시간 저장)
            page_num = 0
            all_contracts = []
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            is_first_page = True
            
            while True:
                print(f"\n{'='*60}")
                print(f"--- page={page_num} 처리 중 ---")
                
                # 현재 페이지 URL로 이동
                current_url = f"{BASE_URL['PRODUCTION']}/clm/complete?page={page_num}"
                self.driver.get(current_url)
                time.sleep(3)
                
                print(f"URL: {current_url}")
                
                # "등록된 내용이 없습니다" 메시지 확인
                try:
                    page_text = self.driver.find_element(By.TAG_NAME, "body").text
                    if "등록된 내용이 없습니다" in page_text:
                        print(f"⚠ '등록된 내용이 없습니다' 메시지 발견. 추출 종료.")
                        break
                except:
                    pass
                
                # 현재 페이지의 계약서 링크 추출
                current_contracts = self.extract_current_page_contracts()
                
                if not current_contracts:
                    print(f"⚠ page={page_num}에 계약서가 없습니다.")
                    break
                
                print(f"✓ page={page_num}에서 {len(current_contracts)}개 계약서 발견")
                
                # 각 계약서 상세 내용 추출
                page_contracts = []
                success_count = 0
                fail_count = 0
                
                for i, contract in enumerate(current_contracts, 1):
                    print(f"\n  [{i}/{len(current_contracts)}] 계약서 상세 추출 중...")
                    
                    if contract.get('link'):
                        try:
                            details = self.extract_contract_details(contract)
                            contract.update(details)
                            
                            # 추출 성공 여부 확인
                            # 'content' 키가 없거나, 'content'에 '추출 실패'가 없으면 성공
                            if 'content' not in details or '추출 실패' not in details.get('content', ''):
                                # 데이터가 있는지 확인 (빈 딕셔너리가 아닌지)
                                if len(details) > 0:
                                    print(f"  ✓ 상세 정보 추출 완료")
                                    success_count += 1
                                else:
                                    print(f"  ⚠ 데이터가 비어있음 (계속 진행)")
                                    fail_count += 1
                            else:
                                print(f"  ⚠ 추출 실패: {details.get('content', '')[:50]} (계속 진행)")
                                fail_count += 1
                        except Exception as e:
                            print(f"  ✗ 예상치 못한 오류: {str(e)[:100]}")
                            contract['content'] = f"추출 실패: {str(e)[:100]}"
                            fail_count += 1
                    else:
                        print("  ℹ 링크가 없어 상세 정보를 추출할 수 없습니다.")
                        contract['content'] = "링크 없음"
                        fail_count += 1
                    
                    page_contracts.append(contract)
                    all_contracts.append(contract)
                
                print(f"\n  → page={page_num} 완료: 성공 {success_count}개, 실패 {fail_count}개")
                
                # 해당 페이지 데이터를 실시간으로 저장
                print(f"\n  📄 페이지 {page_num} 데이터 저장 중...")
                self.contract_data = all_contracts
                
                # 첫 번째 페이지는 새 파일, 이후는 추가 모드
                save_mode = 'a' if not is_first_page else 'w'
                if is_first_page:
                    is_first_page = False
                
                if self.save_data(timestamp=timestamp, mode='w'):  # 전체 데이터 덮어쓰기
                    print(f"  ✓ {len(all_contracts)}개 데이터 저장됨")
                
                # 다음 페이지로
                page_num += 1
                
                # 최대 100페이지 제한
                if page_num >= 100:
                    print("⚠ 최대 페이지 수 도달")
                    break
            
            print(f"\n{'='*60}")
            print(f"✓ 총 {len(all_contracts)}개 계약서 추출 완료")
            print(f"{'='*60}")
            
            self.contract_data = all_contracts
            
            print("=== 프로세스 완료 ===")
            return True
            
        except Exception as e:
            print(f"✗ 프로세스 실행 중 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            if self.driver:
                self.driver.quit()
                print("브라우저가 종료되었습니다.")

def main():
    """메인 함수"""
    # 설정 - .env 파일에서 환경변수 읽기
    username = env_vars.get('prod_ID', '').strip() or env_vars.get('dev_ID', '').strip()
    password = env_vars.get('prod_PW', '').strip() or env_vars.get('dev_PW', '').strip()
    
    print(f"\n환경변수 확인:")
    print(f"  - BASE_URL: {BASE_URL.get('PRODUCTION', '설정되지 않음')}")
    print(f"  - Username: {username}")
    print(f"  - Password: {'설정됨' if password else '설정되지 않음'}\n")
    
    # 추출기 생성 및 실행
    comparator = ContractComparator()
    success = comparator.run_full_process(username, password)
    
    if success:
        print("\n✓ 모든 작업이 성공적으로 완료되었습니다!")
    else:
        print("\n✗ 작업 중 오류가 발생했습니다.")

if __name__ == "__main__":
    main()

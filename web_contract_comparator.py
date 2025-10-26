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

from base_url import BASE_URL

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
            self.driver.get("https://harim.business.lawform.io")
            
            # 페이지 로딩 대기
            time.sleep(3)
            
            # 현재 페이지 정보 출력
            print(f"현재 URL: {self.driver.current_url}")
            print(f"페이지 제목: {self.driver.title}")
            
            # 로그인 페이지 로딩 대기
            wait = WebDriverWait(self.driver, 15)
            
            # 다양한 ID 입력 필드 셀렉터 시도
            id_selectors = [
                (By.NAME, "username"),
                (By.NAME, "email"),
                (By.NAME, "id"),
                (By.ID, "username"),
                (By.ID, "email"),
                (By.ID, "id"),
                (By.XPATH, "//input[@type='email']"),
                (By.XPATH, "//input[@type='text']"),
                (By.CSS_SELECTOR, "input[placeholder*='이메일']"),
                (By.CSS_SELECTOR, "input[placeholder*='아이디']"),
                (By.CSS_SELECTOR, "input[placeholder*='ID']")
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
            
            # 다양한 비밀번호 입력 필드 셀렉터 시도
            pw_selectors = [
                (By.NAME, "password"),
                (By.ID, "password"),
                (By.XPATH, "//input[@type='password']"),
                (By.CSS_SELECTOR, "input[placeholder*='비밀번호']"),
                (By.CSS_SELECTOR, "input[placeholder*='Password']")
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
            
            # 다양한 로그인 버튼 셀렉터 시도
            login_button_selectors = [
                (By.XPATH, "//button[@type='submit']"),
                (By.XPATH, "//input[@type='submit']"),
                (By.XPATH, "//button[contains(text(), '로그인')]"),
                (By.XPATH, "//button[contains(text(), 'Login')]"),
                (By.XPATH, "//button[contains(text(), 'Sign in')]"),
                (By.CSS_SELECTOR, "button[type='submit']"),
                (By.CSS_SELECTOR, "input[type='submit']"),
                (By.CLASS_NAME, "login-btn"),
                (By.CLASS_NAME, "submit-btn")
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
            
            # 로그인 성공 확인 (다양한 방법으로 시도)
            time.sleep(3)
            
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
            contract_url = BASE_URL.PRODUCTION + "/clm/complete?page=0"
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
                current_url = f"{BASE_URL.PRODUCTION}/clm/complete?page={page_num}"
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
                                # "[중분류명](대분류명 > [중분류명])" 형태 파싱
                                # 괄호 안의 경로를 추출
                                
                                # 괄호 안의 내용 추출: (대분류명 > [중분류명])
                                bracket_match = re.search(r'\((.*?)\)', value)
                                if bracket_match:
                                    bracket_content = bracket_match.group(1)  # "대분류명 > [중분류명]"
                                    
                                    # " > "로 분리
                                    if '>' in bracket_content:
                                        parts = bracket_content.split('>')
                                        data['계약분류_대분류'] = parts[0].strip()
                                        data['계약분류_중분류'] = parts[1].strip() if len(parts) > 1 else ''
                                    else:
                                        # 괄호는 있지만 > 없는 경우
                                        data['계약분류_대분류'] = bracket_content.strip()
                                        data['계약분류_중분류'] = ''
                                else:
                                    # 괄호가 없는 경우 원본 그대로 사용
                                    data['계약분류_대분류'] = value
                                    data['계약분류_중분류'] = ''
                            
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
                time.sleep(2)
                
                details = {}
                
                # 계약 정보 영역 추출 (정확한 XPath)
                try:
                    contract_xpath = "/html/body/div/div[1]/div[3]/div[2]/main/main/div/div[2]/div[2]/div[1]/div[1]/div[2]"
                    contract_element = self.driver.find_element(By.XPATH, contract_xpath)
                    contract_info_text = contract_element.text.strip()
                    
                    if contract_info_text:
                        print(f"    ✓ 계약 정보 영역 발견: {len(contract_info_text)}자")
                        details.update(self._parse_contract_info(contract_info_text))
                    else:
                        print(f"    ⚠ 계약 정보 영역이 비어있습니다.")
                        
                except Exception as e:
                    print(f"    ⚠ 계약 정보 영역 찾기 실패: {str(e)[:100]}")
                
                # 상세 정보 영역 추출 (정확한 XPath)
                try:
                    detail_xpath = "/html/body/div/div[1]/div[3]/div[2]/main/main/div/div[2]/div[2]/div[1]/div[2]"
                    detail_element = self.driver.find_element(By.XPATH, detail_xpath)
                    detail_info_text = detail_element.text.strip()
                    
                    if detail_info_text:
                        print(f"    ✓ 상세 정보 영역 발견: {len(detail_info_text)}자")
                        details.update(self._parse_detail_info(detail_info_text))
                    else:
                        print(f"    ⚠ 상세 정보 영역이 비어있습니다.")
                        
                except Exception as e:
                    print(f"    ⚠ 상세 정보 영역 찾기 실패: {str(e)[:100]}")
                
                # 페이지로 돌아가기
                self.driver.back()
                time.sleep(2)
                
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
            
            # Excel 저장 (항상 덮어쓰기)
            excel_filename = f"contract_data_{timestamp}.xlsx"
            if self.contract_data:
                df = pd.DataFrame(self.contract_data)
                df.to_excel(excel_filename, index=False, engine='openpyxl')
                
                if mode == 'w':
                    print(f"✓ Excel 파일 저장: {excel_filename}")
                else:
                    print(f"✓ Excel 파일 업데이트: {excel_filename} ({len(self.contract_data)}개)")
            
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
                current_url = f"{BASE_URL.PRODUCTION}/clm/complete?page={page_num}"
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
                            if details.get('content') and '추출 실패' not in details.get('content', ''):
                                print(f"  ✓ 상세 정보 추출 완료")
                                success_count += 1
                            else:
                                print(f"  ⚠ 추출 실패 (계속 진행)")
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
    # 설정
    username = "developer+id20251002103114449_m@amicuslex.net"
    password = "1q2w#E$R"
    
    # 추출기 생성 및 실행
    comparator = ContractComparator()
    success = comparator.run_full_process(username, password)
    
    if success:
        print("\n✓ 모든 작업이 성공적으로 완료되었습니다!")
    else:
        print("\n✗ 작업 중 오류가 발생했습니다.")

if __name__ == "__main__":
    main()

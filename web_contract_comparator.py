import pandas as pd
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from fuzzywuzzy import fuzz
import os
from datetime import datetime

class ContractComparator:
    def __init__(self):
        self.driver = None
        self.contract_data = []
        self.excel_data = None
        
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
            print("체결 계약서 조회 메뉴로 이동 중...")
            
            # 메뉴 찾기 및 클릭
            wait = WebDriverWait(self.driver, 10)
            
            # 체결 계약서 조회 메뉴 찾기 (다양한 방법으로 시도)
            menu_selectors = [
                "//a[contains(text(), '체결 계약서 조회')]",
                "//a[contains(text(), '계약서')]",
                "//a[contains(text(), '계약')]",
                "//li[contains(@class, 'contract')]//a",
                "//nav//a[contains(@href, 'contract')]"
            ]
            
            menu_found = False
            for selector in menu_selectors:
                try:
                    menu_element = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    menu_element.click()
                    menu_found = True
                    break
                except:
                    continue
            
            if not menu_found:
                print("✗ 체결 계약서 조회 메뉴를 찾을 수 없습니다.")
                return False
            
            # 페이지 로딩 대기
            time.sleep(3)
            print("✓ 체결 계약서 조회 페이지로 이동했습니다.")
            return True
            
        except Exception as e:
            print(f"✗ 메뉴 이동 실패: {str(e)}")
            return False
    
    def extract_contract_list(self):
        """계약서 목록 추출"""
        try:
            print("계약서 목록 추출 중...")
            
            wait = WebDriverWait(self.driver, 10)
            
            # 계약서 목록 테이블 찾기
            table_selectors = [
                "//table",
                "//div[contains(@class, 'table')]",
                "//div[contains(@class, 'list')]",
                "//div[contains(@class, 'contract')]"
            ]
            
            table_found = False
            for selector in table_selectors:
                try:
                    table = self.driver.find_element(By.XPATH, selector)
                    table_found = True
                    break
                except:
                    continue
            
            if not table_found:
                print("✗ 계약서 목록 테이블을 찾을 수 없습니다.")
                return False
            
            # 계약서 행들 찾기
            rows = table.find_elements(By.XPATH, ".//tr")
            
            if len(rows) <= 1:  # 헤더만 있는 경우
                print("✗ 계약서 데이터가 없습니다.")
                return False
            
            # 헤더 추출
            headers = []
            header_row = rows[0]
            header_cells = header_row.find_elements(By.XPATH, ".//th | .//td")
            for cell in header_cells:
                headers.append(cell.text.strip())
            
            print(f"발견된 컬럼: {headers}")
            
            # 데이터 행들 추출
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
                    
                    # 계약서 링크 찾기
                    try:
                        link_element = row.find_element(By.XPATH, ".//a")
                        row_data['link'] = link_element.get_attribute('href')
                        row_data['link_element'] = link_element
                    except:
                        row_data['link'] = None
                        row_data['link_element'] = None
                    
                    contract_list.append(row_data)
                    
                except Exception as e:
                    print(f"행 {i} 처리 중 오류: {str(e)}")
                    continue
            
            print(f"✓ {len(contract_list)}개의 계약서를 발견했습니다.")
            return contract_list
            
        except Exception as e:
            print(f"✗ 계약서 목록 추출 실패: {str(e)}")
            return False
    
    def extract_contract_details(self, contract_link_element):
        """개별 계약서 상세 내용 추출"""
        try:
            # 계약서 링크 클릭
            contract_link_element.click()
            time.sleep(2)
            
            # 상세 페이지 로딩 대기
            wait = WebDriverWait(self.driver, 10)
            
            # 계약서 상세 정보 추출
            details = {}
            
            # 다양한 방법으로 계약서 정보 추출 시도
            detail_selectors = [
                "//div[contains(@class, 'contract-detail')]",
                "//div[contains(@class, 'detail')]",
                "//div[contains(@class, 'content')]",
                "//div[contains(@class, 'info')]",
                "//table",
                "//div[contains(@class, 'form')]"
            ]
            
            for selector in detail_selectors:
                try:
                    detail_elements = self.driver.find_elements(By.XPATH, selector)
                    if detail_elements:
                        for element in detail_elements:
                            text = element.text.strip()
                            if text and len(text) > 10:  # 의미있는 텍스트만 추출
                                details['content'] = text
                                break
                        if 'content' in details:
                            break
                except:
                    continue
            
            # 페이지 전체 텍스트도 추출
            if 'content' not in details:
                details['content'] = self.driver.find_element(By.TAG_NAME, "body").text
            
            # 뒤로 가기
            self.driver.back()
            time.sleep(1)
            
            return details
            
        except Exception as e:
            print(f"✗ 계약서 상세 추출 실패: {str(e)}")
            # 뒤로 가기 시도
            try:
                self.driver.back()
            except:
                pass
            return {}
    
    def load_excel_data(self, excel_file_path):
        """엑셀 파일 로드"""
        try:
            print(f"엑셀 파일 로드 중: {excel_file_path}")
            
            # 엑셀 파일의 모든 시트 확인
            excel_file = pd.ExcelFile(excel_file_path)
            print(f"발견된 시트: {excel_file.sheet_names}")
            
            # 첫 번째 시트 로드 (또는 특정 시트 지정)
            self.excel_data = pd.read_excel(excel_file_path, sheet_name=0)
            print(f"✓ 엑셀 파일 로드 완료: {len(self.excel_data)}행")
            print(f"엑셀 컬럼: {list(self.excel_data.columns)}")
            
            return True
            
        except Exception as e:
            print(f"✗ 엑셀 파일 로드 실패: {str(e)}")
            return False
    
    def compare_data(self):
        """웹 데이터와 엑셀 데이터 비교"""
        try:
            print("데이터 비교 시작...")
            
            if not self.contract_data or self.excel_data is None:
                print("✗ 비교할 데이터가 없습니다.")
                return False
            
            comparison_results = []
            
            for web_contract in self.contract_data:
                best_match = None
                best_score = 0
                
                # 엑셀 데이터와 비교하여 가장 유사한 항목 찾기
                for idx, excel_row in self.excel_data.iterrows():
                    score = 0
                    match_count = 0
                    
                    # 각 필드별로 유사성 계산
                    for web_key, web_value in web_contract.items():
                        if web_key in ['link', 'link_element'] or not web_value:
                            continue
                            
                        for excel_col in self.excel_data.columns:
                            excel_value = str(excel_row[excel_col])
                            similarity = fuzz.ratio(str(web_value), excel_value)
                            
                            if similarity > 80:  # 80% 이상 유사하면 매치로 간주
                                score += similarity
                                match_count += 1
                    
                    if match_count > 0:
                        avg_score = score / match_count
                        if avg_score > best_score:
                            best_score = avg_score
                            best_match = {
                                'web_data': web_contract,
                                'excel_data': excel_row.to_dict(),
                                'similarity_score': avg_score,
                                'match_count': match_count
                            }
                
                comparison_results.append(best_match)
            
            # 결과를 DataFrame으로 변환
            results_df = pd.DataFrame(comparison_results)
            
            # 결과 저장
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"contract_comparison_results_{timestamp}.xlsx"
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                results_df.to_excel(writer, sheet_name='비교결과', index=False)
            
            print(f"✓ 비교 결과가 '{output_file}'에 저장되었습니다.")
            print(f"총 {len(comparison_results)}개 계약서 비교 완료")
            
            return True
            
        except Exception as e:
            print(f"✗ 데이터 비교 실패: {str(e)}")
            return False
    
    def run_full_process(self, username, password, excel_file_path):
        """전체 프로세스 실행"""
        try:
            print("=== 계약서 비교 프로세스 시작 ===")
            
            # 1. 드라이버 설정
            if not self.setup_driver():
                return False
            
            # 2. 로그인
            if not self.login(username, password):
                return False
            
            # 3. 계약서 조회 페이지로 이동
            if not self.navigate_to_contracts():
                return False
            
            # 4. 계약서 목록 추출
            contract_list = self.extract_contract_list()
            if not contract_list:
                return False
            
            # 5. 각 계약서 상세 내용 추출 (처음 5개만 테스트)
            print("계약서 상세 내용 추출 중...")
            for i, contract in enumerate(contract_list[:5]):  # 테스트를 위해 처음 5개만
                print(f"[{i+1}/5] 계약서 상세 추출 중...")
                if contract['link_element']:
                    details = self.extract_contract_details(contract['link_element'])
                    contract.update(details)
                else:
                    contract['content'] = "링크 없음"
            
            self.contract_data = contract_list[:5]  # 테스트용으로 5개만 저장
            
            # 6. 엑셀 파일 로드
            if not self.load_excel_data(excel_file_path):
                return False
            
            # 7. 데이터 비교
            if not self.compare_data():
                return False
            
            print("=== 프로세스 완료 ===")
            return True
            
        except Exception as e:
            print(f"✗ 프로세스 실행 중 오류: {str(e)}")
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
    excel_file_path = "/Users/ggpark/Desktop/python/로아이_통합파일.xlsx"
    
    # 비교기 생성 및 실행
    comparator = ContractComparator()
    success = comparator.run_full_process(username, password, excel_file_path)
    
    if success:
        print("✓ 모든 작업이 성공적으로 완료되었습니다!")
    else:
        print("✗ 작업 중 오류가 발생했습니다.")

if __name__ == "__main__":
    main()

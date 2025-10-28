import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def parse_custom_env():
    """.env 파일을 직접 파싱 (KEY : VALUE 형식 지원)"""
    env_vars = {}
    
    try:
        with open('.env', 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                if ':' in line:
                    key, value = line.split(':', 1)
                    key = key.strip()
                    value = value.strip()
                    value = value.strip('"').strip("'")
                    env_vars[key] = value
        return env_vars
    except Exception as e:
        print(f"⚠ .env 파일 읽기 실패: {e}")
        return {}

# .env 파일 파싱
env_vars = parse_custom_env()

# 환경변수 가져오기
BASE_URL = {
    'PRODUCTION': env_vars.get('prod_BASE_URL', '').strip() or env_vars.get('dev_BASE_URL', '').strip(),
}

def find_tables_in_page():
    """페이지의 모든 테이블 요소 찾기 및 분석"""
    driver = None
    
    try:
        # Chrome 드라이버 설정
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        
        driver = webdriver.Chrome(options=chrome_options)
        print("✓ Chrome 드라이버 설정 완료")
        
        # 1. 로그인
        print("\n=== 로그인 중 ===")
        username = env_vars.get('prod_ID', '').strip() or env_vars.get('dev_ID', '').strip()
        password = env_vars.get('prod_PW', '').strip() or env_vars.get('dev_PW', '').strip()
        
        driver.get(BASE_URL['PRODUCTION'])
        time.sleep(3)
        
        # ID 입력
        id_field = driver.find_element(By.CSS_SELECTOR, "input[type='email']")
        id_field.clear()
        id_field.send_keys(username)
        print("✓ ID 입력 완료")
        
        # 비밀번호 입력
        pw_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        pw_field.clear()
        pw_field.send_keys(password)
        print("✓ 비밀번호 입력 완료")
        
        # 로그인 버튼 클릭
        login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        login_button.click()
        print("✓ 로그인 버튼 클릭")
        
        time.sleep(3)
        print(f"현재 URL: {driver.current_url}\n")
        
        # 2. 계약서 목록 페이지로 이동
        print("=== 계약서 목록 페이지로 이동 ===")
        contract_url = BASE_URL['PRODUCTION'] + "/clm/complete?page=0"
        driver.get(contract_url)
        time.sleep(3)
        print(f"URL: {contract_url}\n")
        
        # 3. 첫 번째 계약서 링크 찾기
        print("=== 첫 번째 계약서 링크 찾기 ===")
        try:
            # 테이블 찾기
            table = driver.find_element(By.XPATH, "//table")
            rows = table.find_elements(By.XPATH, ".//tr")
            
            if len(rows) <= 1:
                print("✗ 계약서 목록이 비어있습니다.")
                return
            
            print(f"✓ 테이블 발견: {len(rows)}개 행")
            
            # 첫 번째 계약서 링크 추출
            first_contract_row = rows[1]  # rows[0]은 헤더
            cells = first_contract_row.find_elements(By.XPATH, ".//td")
            print(f"  → 첫 번째 행의 셀 개수: {len(cells)}")
            
            # 각 셀의 텍스트 출력
            for i, cell in enumerate(cells):
                text = cell.text.strip()
                print(f"  셀 {i+1}: {text[:50]}")
            
            # 링크 찾기
            link_element = first_contract_row.find_element(By.XPATH, ".//a")
            contract_link = link_element.get_attribute('href')
            print(f"\n✓ 첫 번째 계약서 링크: {contract_link}\n")
            
        except Exception as e:
            print(f"✗ 계약서 목록 추출 실패: {e}")
            return
        
        # 4. 계약서 상세 페이지로 이동
        print("=== 계약서 상세 페이지로 이동 ===")
        driver.get(contract_link)
        time.sleep(5)
        print(f"URL: {driver.current_url}\n")
        
        # 5. 페이지의 모든 테이블 찾기
        print("=== 페이지의 모든 테이블 요소 분석 ===\n")
        
        try:
            # 모든 table 요소 찾기
            all_tables = driver.find_elements(By.TAG_NAME, "table")
            print(f"✓ 총 {len(all_tables)}개의 <table> 태그 발견\n")
            
            for table_idx, table in enumerate(all_tables, 1):
                print(f"{'='*80}")
                print(f"테이블 #{table_idx}")
                print(f"{'='*80}")
                
                # 테이블의 경로 정보 추출
                table_id = table.get_attribute('id')
                table_class = table.get_attribute('class')
                print(f"ID: {table_id}")
                print(f"Class: {table_class}")
                
                # 테이블 내용 일부 추출
                rows = table.find_elements(By.XPATH, ".//tr")
                print(f"행 개수: {len(rows)}")
                
                # 처음 몇 개 행만 출력
                for row_idx, row in enumerate(rows[:5], 1):
                    cells = row.find_elements(By.XPATH, ".//td | .//th")
                    cell_texts = [cell.text.strip() for cell in cells]
                    print(f"  행 {row_idx}: {cell_texts[:5]}")  # 처음 5개 셀만
                
                if len(rows) > 5:
                    print(f"  ... (총 {len(rows)}개 행)")
                
                print()
        
        except Exception as e:
            print(f"✗ 테이블 분석 실패: {e}")
            import traceback
            traceback.print_exc()
        
        # 6. XPath 자동 생성 테스트
        print("\n=== XPath 자동 생성 테스트 ===\n")
        try:
            all_tables = driver.find_elements(By.TAG_NAME, "table")
            
            for table_idx, table in enumerate(all_tables, 1):
                print(f"테이블 #{table_idx}:")
                
                # 다양한 방식으로 XPath 찾기 시도
                xpaths = []
                
                # 방법 1: 테이블에서 직접 위치 찾기
                try:
                    # 절대 경로
                    xpath_absolute = driver.execute_script(
                        "return arguments[0].path;",
                        table
                    )
                    print(f"  - 절대 경로: {xpath_absolute}")
                except:
                    pass
                
                # 방법 2: 상대 경로
                table_id = table.get_attribute('id')
                if table_id:
                    xpaths.append(f"//table[@id='{table_id}']")
                
                table_class = table.get_attribute('class')
                if table_class:
                    class_list = table_class.split()
                    for cls in class_list:
                        xpaths.append(f"//table[@class='{cls}']")
                
                # 방법 3: 앞뒤 요소
                try:
                    preceding = table.find_elements(By.XPATH, "./preceding-sibling::*")
                    following = table.find_elements(By.XPATH, "./following-sibling::*")
                    print(f"  - 앞 요소 개수: {len(preceding)}")
                    print(f"  - 뒤 요소 개수: {len(following)}")
                except:
                    pass
                
                if xpaths:
                    print(f"  - 생성된 XPath: {xpaths[0]}")
                
                print()
        
        except Exception as e:
            print(f"✗ XPath 생성 실패: {e}")
            import traceback
            traceback.print_exc()
        
        # 7. 테이블 내용 상세 분석
        print("\n=== 테이블 내용 상세 분석 ===\n")
        try:
            all_tables = driver.find_elements(By.TAG_NAME, "table")
            
            for table_idx, table in enumerate(all_tables, 1):
                print(f"테이블 #{table_idx} 분석:")
                print("-" * 80)
                
                rows = table.find_elements(By.XPATH, ".//tr")
                
                # 각 행의 모든 셀 내용 출력
                for row_idx, row in enumerate(rows, 1):
                    cells = row.find_elements(By.XPATH, ".//td | .//th")
                    print(f"\n행 {row_idx}:")
                    
                    for cell_idx, cell in enumerate(cells):
                        cell_text = cell.text.strip()
                        cell_tag = cell.tag_name
                        print(f"  [{cell_idx}] ({cell_tag}): {cell_text}")
                
                print()
        
        except Exception as e:
            print(f"✗ 상세 분석 실패: {e}")
            import traceback
            traceback.print_exc()
        
        # 8. HTML 구조 출력 (첫 5000자)
        print("\n=== 페이지 HTML 구조 일부 ===\n")
        try:
            page_source = driver.page_source
            print(page_source[:5000])
        except:
            pass
        
        # 페이지 유지 (사용자가 확인할 수 있도록)
        print("\n" + "="*80)
        print("브라우저를 30초간 유지합니다. 확인 후 자동 종료됩니다.")
        print("="*80)
        time.sleep(30)
        
    except Exception as e:
        print(f"✗ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        if driver:
            driver.quit()
            print("\n✓ 브라우저 종료")

if __name__ == "__main__":
    find_tables_in_page()


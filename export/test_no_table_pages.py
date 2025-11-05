"""테이블 요소가 없는 계약서 상세 페이지를 탐지/분석하는 테스트 스크립트.

주요 기능
- 로그인 → 계약서 목록 순회 → 테이블 없는 상세 페이지 탐지
- dl, section, main 등 대체 구조의 존재 여부와 일부 내용을 출력
- HTML 전체를 파일로 저장해 사후 분석에 활용
"""

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

def parse_custom_env():
    """.env 파일을 직접 파싱"""
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

env_vars = parse_custom_env()
BASE_URL = {
    'PRODUCTION': env_vars.get('prod_BASE_URL', '').strip() or env_vars.get('dev_BASE_URL', '').strip(),
}

def analyze_no_table_page():
    """테이블이 없는 페이지 분석"""
    driver = None
    
    try:
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        
        driver = webdriver.Chrome(options=chrome_options)
        
        # 로그인
        username = env_vars.get('prod_ID', '').strip() or env_vars.get('dev_ID', '').strip()
        password = env_vars.get('prod_PW', '').strip() or env_vars.get('dev_PW', '').strip()
        
        driver.get(BASE_URL['PRODUCTION'])
        time.sleep(3)
        
        id_field = driver.find_element(By.CSS_SELECTOR, "input[type='email']")
        id_field.clear()
        id_field.send_keys(username)
        
        pw_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        pw_field.clear()
        pw_field.send_keys(password)
        
        login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        login_button.click()
        time.sleep(3)
        
        # 계약서 목록
        contract_url = BASE_URL['PRODUCTION'] + "/clm/complete?page=0"
        driver.get(contract_url)
        time.sleep(3)
        
        # 첫 번째 계약서 가져오기
        table = driver.find_element(By.XPATH, "//table")
        rows = table.find_elements(By.XPATH, ".//tr")
        
        # 테이블이 0개인 페이지 찾기
        no_table_contract = None
        
        for i, row in enumerate(rows[1:min(21, len(rows))], 1):
            try:
                cells = row.find_elements(By.XPATH, ".//td")
                if len(cells) == 0:
                    continue
                
                link_element = row.find_element(By.XPATH, ".//a")
                contract_link = link_element.get_attribute('href')
                contract_name = cells[1].text.strip() if len(cells) > 1 else "N/A"
                
                print(f"계약서 #{i}: {contract_name}")
                
                # 상세 페이지로 이동
                driver.get(contract_link)
                time.sleep(3)
                
                # 테이블 확인
                all_tables = driver.find_elements(By.TAG_NAME, "table")
                print(f"  → 테이블 {len(all_tables)}개")
                
                if len(all_tables) == 0:
                    print(f"  ✓ 테이블 없는 페이지 발견!")
                    
                    # 페이지 구조 분석
                    print(f"\n[구조 분석]")
                    
                    # dl 태그
                    dls = driver.find_elements(By.TAG_NAME, "dl")
                    print(f"  - dl 태그: {len(dls)}개")
                    
                    # main의 구조
                    main = driver.find_element(By.TAG_NAME, "main")
                    main_text = main.text[:500]  # 처음 500자
                    print(f"\n[main 내용 일부]\n{main_text}")
                    
                    # HTML 구조 일부
                    page_source = driver.page_source
                    
                    # "관리번호" 키워드 찾기
                    if "관리번호" in page_source:
                        idx = page_source.find("관리번호")
                        snippet = page_source[max(0, idx-200):min(len(page_source), idx+800)]
                        print(f"\n[HTML 구조: 관리번호 주변]\n{snippet[:1000]}")
                    
                    # 섹션 찾기
                    sections = driver.find_elements(By.XPATH, "//main//section | //main//div[contains(@class, 'section')]")
                    print(f"  - section 개수: {len(sections)}")
                    
                    # 전체 HTML 저장
                    with open(f"no_table_page_{i}.html", "w", encoding='utf-8') as f:
                        f.write(page_source)
                    print(f"\n✓ HTML 저장: no_table_page_{i}.html")
                    
                    break  # 하나만 분석
                
                driver.back()
                time.sleep(2)
                
            except Exception as e:
                print(f"  ✗ 오류: {e}")
                continue
        
        print("\n✓ 분석 완료")
        print("브라우저를 30초간 유지합니다...")
        time.sleep(30)
        
    except Exception as e:
        print(f"✗ 오류: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    analyze_no_table_page()


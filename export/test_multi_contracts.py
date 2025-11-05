"""여러 계약서 상세 페이지를 순회하며 구조를 점검/로그하는 테스트 스크립트.

주요 기능
- 로그인 후 계약서 목록에서 상위 N개 링크를 수집
- 각 상세 페이지의 테이블/구조/데이터 후보 요소를 스캔하여 콘솔에 요약 출력
- 초반 일부 페이지의 HTML 스니펫을 출력해 패턴 파악에 도움 제공
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

# 환경변수 로드
env_vars = parse_custom_env()
BASE_URL = {
    'PRODUCTION': env_vars.get('prod_BASE_URL', '').strip() or env_vars.get('dev_BASE_URL', '').strip(),
}

def analyze_multiple_contracts():
    """여러 계약서의 구조를 분석"""
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
        
        print("✓ 로그인 완료\n")
        
        # 계약서 목록 페이지로 이동
        contract_url = BASE_URL['PRODUCTION'] + "/clm/complete?page=0"
        driver.get(contract_url)
        time.sleep(3)
        
        print(f"=== 계약서 목록 페이지 분석 ===\n")
        
        # 먼저 모든 링크를 수집
        table = driver.find_element(By.XPATH, "//table")
        rows = table.find_elements(By.XPATH, ".//tr")
        
        print(f"총 {len(rows)-1}개 계약서 발견\n")
        
        contract_links = []
        for i, row in enumerate(rows[1:6], 1):  # 처음 5개만
            try:
                cells = row.find_elements(By.XPATH, ".//td")
                if len(cells) == 0:
                    continue
                
                link_element = row.find_element(By.XPATH, ".//a")
                contract_link = link_element.get_attribute('href')
                contract_name = cells[1].text.strip() if len(cells) > 1 else "N/A"
                
                contract_links.append({
                    'name': contract_name,
                    'link': contract_link
                })
            except Exception as e:
                print(f"  계약서 #{i} 링크 추출 실패: {e}")
                continue
        
        print(f"수집된 링크: {len(contract_links)}개\n")
        
        # 수집된 링크로 분석
        analyzed = 0
        for i, contract in enumerate(contract_links, 1):
            contract_name = contract['name']
            contract_link = contract['link']
            
            print(f"\n{'='*80}")
            print(f"계약서 #{i}: {contract_name}")
            print(f"URL: {contract_link}")
            print('='*80)
            
            # 계약서 상세 페이지로 이동
            driver.get(contract_link)
            time.sleep(3)
            
            # 페이지 구조 분석
            print(f"\n[구조 분석]")
            
            # 1. 테이블 개수 확인
            all_tables = driver.find_elements(By.TAG_NAME, "table")
            print(f"  - 테이블 개수: {len(all_tables)}")
            
            # 2. div 구조 확인
            main_divs = driver.find_elements(By.XPATH, "//main//div")
            print(f"  - main 안 div 개수: {len(main_divs)}")
            
            # 3. 폼 요소 확인
            forms = driver.find_elements(By.TAG_NAME, "form")
            print(f"  - form 개수: {len(forms)}")
            
            # 4. 리스트 요소 확인
            uls = driver.find_elements(By.TAG_NAME, "ul")
            print(f"  - ul 개수: {len(uls)}")
            
            # 5. 데이터를 담고 있을 수 있는 요소들 찾기
            print(f"\n[데이터 요소 탐지]")
            
            # dl 태그 (description list)
            dls = driver.find_elements(By.TAG_NAME, "dl")
            print(f"  - dl 태그 개수: {len(dls)}")
            
            # 테이블이 없는 경우를 위한 대안 찾기
            if len(all_tables) == 0:
                print(f"  ⚠ 테이블이 없습니다!")
                print(f"\n[대안 요소 찾기]")
                
                # dl 태그가 있는지 확인
                if len(dls) > 0:
                    print(f"    → dl 태그 발견! {len(dls)}개")
                    for dl_idx, dl in enumerate(dls, 1):
                        dt_elements = dl.find_elements(By.TAG_NAME, "dt")
                        dd_elements = dl.find_elements(By.TAG_NAME, "dd")
                        print(f"      dl #{dl_idx}: dt={len(dt_elements)}, dd={len(dd_elements)}")
                
                # div로 데이터가 표시되는지 확인
                data_divs = driver.find_elements(By.XPATH, "//div[contains(@class, 'grid') or contains(@class, 'flex')]")
                print(f"    - grid/flex div 개수: {len(data_divs)}")
                
                # li 태그
                lis = driver.find_elements(By.TAG_NAME, "li")
                print(f"    - li 태그 개수: {len(lis)}")
            
            # 6. 페이지 소스 일부 출력 (필요한 경우)
            if analyzed < 2:  # 처음 2개만
                print(f"\n[HTML 구조 일부]")
                page_source = driver.page_source
                # "관리번호", "계약명" 같은 키워드가 있는지 확인
                if "관리번호" in page_source:
                    idx = page_source.find("관리번호")
                    snippet = page_source[max(0, idx-100):min(len(page_source), idx+500)]
                    print(f"  관리번호 주변:\n{snippet[:600]}")
            
            # 목록으로 돌아가기
            driver.back()
            time.sleep(2)
            
            analyzed += 1
            
            # 5개까지만 분석
            if analyzed >= 5:
                break
        
        print(f"\n\n✓ 총 {analyzed}개 계약서 분석 완료")
        print("\n브라우저를 30초간 유지합니다...")
        time.sleep(30)
        
    except Exception as e:
        print(f"✗ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        if driver:
            driver.quit()
            print("✓ 브라우저 종료")

if __name__ == "__main__":
    analyze_multiple_contracts()


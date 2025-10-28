import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def test_extract_contract():
    """특정 계약서 페이지에서 데이터 추출 테스트"""
    
    # Chrome 드라이버 설정
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    
    driver = webdriver.Chrome(options=chrome_options)
    
    try:
        # 1. 로그인
        print("로그인 중...")
        driver.get("https://harim.business.lawform.io")
        time.sleep(3)
        
        # ID 입력
        id_field = driver.find_element(By.XPATH, "//input[@type='email']")
        id_field.send_keys("developer+id20251002103114449_m@amicuslex.net")
        time.sleep(1)
        
        # 비밀번호 입력
        pw_field = driver.find_element(By.XPATH, "//input[@type='password']")
        pw_field.send_keys("1q2w#E$R")
        time.sleep(1)
        
        # 로그인 버튼 클릭
        login_btn = driver.find_element(By.XPATH, "//button[@type='submit']")
        login_btn.click()
        time.sleep(5)
        
        print("OK 로그인 완료")
        
        # 2. 테스트할 계약서 페이지로 이동
        test_url = "https://harim.business.lawform.io/clm/10561"
        print(f"테스트 URL로 이동: {test_url}")
        driver.get(test_url)
        time.sleep(5)
        
        print(f"현재 URL: {driver.current_url}")
        
        # 3. 계약 정보 영역 추출
        try:
            contract_xpath = "/html/body/div/div[1]/div[3]/div[2]/main/main/div/div[2]/div[2]/div[1]/div[1]/div[2]"
            contract_element = driver.find_element(By.XPATH, contract_xpath)
            contract_info_text = contract_element.text.strip()
            
            print(f"\nOK 계약 정보 영역 발견: {len(contract_info_text)}자")
            print(f"\n계약 정보 내용:")
            print("="*60)
            print(contract_info_text)
            print("="*60)
            
        except Exception as e:
            print(f"\nFAIL 계약 정보 영역 찾기 실패: {str(e)}")
        
        # 4. 상세 정보 영역 추출
        try:
            detail_xpath = "/html/body/div/div[1]/div[3]/div[2]/main/main/div/div[2]/div[2]/div[1]/div[2]"
            detail_element = driver.find_element(By.XPATH, detail_xpath)
            detail_info_text = detail_element.text.strip()
            
            print(f"\nOK 상세 정보 영역 발견: {len(detail_info_text)}자")
            print(f"\n상세 정보 내용:")
            print("="*60)
            print(detail_info_text)
            print("="*60)
            
        except Exception as e:
            print(f"\nFAIL 상세 정보 영역 찾기 실패: {str(e)}")
        
        # 5. 스크린샷 저장
        driver.save_screenshot("test_page.png")
        print("\nOK 스크린샷 저장: test_page.png")
        
        # 페이지 소스 저장
        with open("test_page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("OK 페이지 소스 저장: test_page_source.html")
        
        print("\nOK 테스트 완료!")
        print("\n'q' 입력 후 Enter를 눌러 종료하세요.")
        input()
        
    except Exception as e:
        print(f"\nFAIL 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        driver.quit()

if __name__ == "__main__":
    test_extract_contract()


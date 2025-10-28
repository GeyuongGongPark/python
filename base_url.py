import os

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

# BASE_URL을 .env에서 가져오기
PRODUCTION_URL = env_vars.get('prod_BASE_URL', '').strip() or env_vars.get('dev_BASE_URL', '').strip()

class BASE_URL:
    PRODUCTION = PRODUCTION_URL


# 로그인 관련 URL
class LOGIN_URLS:
    HOME = PRODUCTION_URL
    LOGIN = f"{PRODUCTION_URL}/login"
    DASHBOARD = f"{PRODUCTION_URL}/dashboard"


# 계약서 생성 관련 URL
class DRIVE_URLS:
    DRIVE = f"{PRODUCTION_URL}/drive"  # My 계약서
    TEAM = f"{PRODUCTION_URL}/team_standard_contract"  # 기업 표준 계약서
    AUTO = f"{PRODUCTION_URL}/#documents_finder"  # 자동작성
    CHECKLIST = f"{PRODUCTION_URL}/ai/dchecklist"  # AI 필수조항검토
    GLD = "https://chatgld.io"


# 대량생성 관련 URL
class BULK_URLS:
    BULK = f"{PRODUCTION_URL}/bulk"  # 대량 생성


# CLM 관련 URL
class CLM_URLS:
    DRAFT = f"{PRODUCTION_URL}/clm/draft"  # 계약 검토 요청
    PROCESS = f"{PRODUCTION_URL}/clm/process"
    SEARCH = f"{PRODUCTION_URL}/clm/search"  # 통합검색
    REVIEW = f"{PRODUCTION_URL}/clm/review"  # 검토 요청 조회
    COMPLETE = f"{PRODUCTION_URL}/clm/complete"  # 체결 계약서 조회
    COMPARE = f"{PRODUCTION_URL}/document_compare"  # AI 계약 내용 비교
    PAUSE = f"{PRODUCTION_URL}/clm/complete?is_paused=2"  # 일시 중단 리스트


# SEAL 관련 URL
class SEAL_URLS:
    DRAFT = f"{PRODUCTION_URL}/seal/draft"  # 인감 사용 신청
    REVIEW = f"{PRODUCTION_URL}/seal"  # 인감 사용 신청 조회
    LEDGER = f"{PRODUCTION_URL}/seal/ledger"  # 인감 관리 대장


# Advice 관련 URL
class ADVICE_URLS:
    DRAFT = f"{PRODUCTION_URL}/advice/draft"  # 법률 지문 요청
    PROCESS = f"{PRODUCTION_URL}/advice/process"
    REVIEW = f"{PRODUCTION_URL}/advice"  # 법률 자문 조회


# Litigation 관련 URL
class LITIGATION_URLS:
    DRAFT = f"{PRODUCTION_URL}/litigation/draft"  # 송무 등록
    PROCESS = f"{PRODUCTION_URL}/litigation/process"
    REVIEW = f"{PRODUCTION_URL}/litigation"  # 송무 조회
    SCHEDULE = f"{PRODUCTION_URL}/litigation/schedule"  # 송무 전체 일정


# 법령 정보 관련 URL
class LAW_URLS:
    SCHEDULE = f"{PRODUCTION_URL}/law"  # 법령 캘린더


# 프로젝트 관련 URL
class PROJECT_URLS:
    PROJECT = f"{PRODUCTION_URL}/project"  # 프로젝트 조회


# 계약 정보 관리 관련 URL
class CONTRACT_URLS:
    CONTRACT = f"{PRODUCTION_URL}/contact"  # 계약처 관리
    STAMP = f"{PRODUCTION_URL}/template?type=stamp"  # 직인
    LOGO = f"{PRODUCTION_URL}/template?type=logo"  # 로고
    TEAM_STAMP = f"{PRODUCTION_URL}/template?type=team_stamp"  # 기업직인
    WATERMARK = f"{PRODUCTION_URL}/template?type=watermark"  # 워터마크


# 시스템 설정 관련 URL
class SETTING_URLS:
    TEAM = f"{PRODUCTION_URL}/teams"  # 구성원 관리
    ACCOUNT = f"{PRODUCTION_URL}/profile?type=account"  # 회원정보_계정 설정
    NOTIFICATION = f"{PRODUCTION_URL}/profile?type=notification"  # 회원정보_알림/이메일 수신 설정
    LOG = f"{PRODUCTION_URL}/profile?type=log"  # 회원정보_로그인 기록
    FAILEDLOG = f"{PRODUCTION_URL}/profile?type=failedLog"  # 회원정보_로그인 실패 기록
    FA = f"{PRODUCTION_URL}/profile?type=twoFA"  # 2단계 인증
    MANAGEMENT = f"{PRODUCTION_URL}/profile?type=deviceManagement"  # 회원정보_로그인 관리
    SETUP = f"{PRODUCTION_URL}/setup"  # 설정


# 모든 URL을 포함하는 통합 클래스
class URLS:
    BASE = BASE_URL
    LOGIN = LOGIN_URLS
    CLM = CLM_URLS
    SEAL = SEAL_URLS
    ADVICE = ADVICE_URLS
    LITIGATION = LITIGATION_URLS
    BULK = BULK_URLS
    LAW = LAW_URLS
    PROJECT = PROJECT_URLS
    CONTRACT = CONTRACT_URLS
    SETTING = SETTING_URLS


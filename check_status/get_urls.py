#!/usr/bin/env python3
"""base_url.py에서 모든 URL을 추출하여 JSON 형식으로 출력하는 스크립트"""

from __future__ import annotations  # Python 3.7+ 타입 힌트 호환성

import sys
import os
import json
import warnings

# 타입 체크 관련 경고 무시
warnings.filterwarnings('ignore', category=DeprecationWarning)

# 현재 스크립트의 디렉토리로 이동하여 utils 모듈을 import할 수 있도록 함
script_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(script_dir)
sys.path.insert(0, parent_dir)

try:
    # base_url.py import 시 발생할 수 있는 예외 처리
    # account_env.py의 타입 힌트 문제를 우회하기 위해 importlib 사용
    import traceback
    import importlib.util
    import types
    
    # ACCOUNT 환경변수가 없어도 import가 가능하도록 처리
    original_account = os.environ.get('ACCOUNT')
    if 'ACCOUNT' not in os.environ:
        os.environ['ACCOUNT'] = 'dummy'
    
    try:
        # 타입 체크를 우회하기 위해 직접 모듈 로드
        # 하지만 간단하게 import 시도
        import importlib
        import sys
        
        # 타입 체크 관련 설정
        if hasattr(sys, '_getframe'):
            # Python 3.10+ 타입 힌트 호환성 처리
            pass
        
        # 직접 import 시도
        from utils.base_url import (
            LOGIN_URLS, DRIVE_URLS, BULK_URLS, CLM_URLS, SEAL_URLS,
            ADVICE_URLS, LITIGATION_URLS, LAW_URLS, PROJECT_URLS,
            CONTRACT_URLS, SETTING_URLS, BASE_URL
        )
    except TypeError as type_error:
        # 타입 힌트 관련 에러인 경우 account_env.py를 우회하여 처리
        if "unsupported operand type(s) for |" in str(type_error):
            # account_env.py의 타입 힌트 문제를 우회하기 위해
            # 직접 base_url.py의 내용을 읽어서 처리하거나
            # 또는 account_env를 모킹
            print("타입 힌트 호환성 문제 감지. account_env.py를 우회합니다.", file=sys.stderr)
            
            # account_env 모듈을 모킹
            import types
            mock_account_env = types.ModuleType('account_env')
            def mock_load_account_env(*args, **kwargs):
                return {}
            mock_account_env.load_account_env = mock_load_account_env
            
            # sys.modules에 모킹된 모듈 등록
            sys.modules['utils.account_env'] = mock_account_env
            
            # 다시 import 시도
            from utils.base_url import (
                LOGIN_URLS, DRIVE_URLS, BULK_URLS, CLM_URLS, SEAL_URLS,
                ADVICE_URLS, LITIGATION_URLS, LAW_URLS, PROJECT_URLS,
                CONTRACT_URLS, SETTING_URLS, BASE_URL
            )
        else:
            raise
    except Exception as import_error:
        # import 실패 시 상세 에러 정보 출력
        error_msg = str(import_error)
        traceback_msg = traceback.format_exc()
        print(f"Import error: {error_msg}", file=sys.stderr)
        print(f"Traceback:\n{traceback_msg}", file=sys.stderr)
        raise
    finally:
        # 환경변수 복원
        if original_account is None:
            os.environ.pop('ACCOUNT', None)
        elif original_account != os.environ.get('ACCOUNT'):
            os.environ['ACCOUNT'] = original_account
    
    # 모든 URL을 카테고리별로 수집
    urls = {
        "BASE": {
            "PRODUCTION": BASE_URL.PRODUCTION
        },
        "LOGIN": {
            "HOME": LOGIN_URLS.HOME,
            "LOGIN": LOGIN_URLS.LOGIN,
            "DASHBOARD": LOGIN_URLS.DASHBOARD
        },
        "DRIVE": {
            "DRIVE": DRIVE_URLS.DRIVE,
            "TEAM": DRIVE_URLS.TEAM,
            "AUTO": DRIVE_URLS.AUTO,
            "CHECKLIST": DRIVE_URLS.CHECKLIST,
            "GLD": DRIVE_URLS.GLD
        },
        "BULK": {
            "BULK": BULK_URLS.BULK
        },
        "CLM": {
            "DRAFT": CLM_URLS.DRAFT,
            "PROCESS": CLM_URLS.PROCESS,
            "SEARCH": CLM_URLS.SEARCH,
            "REVIEW": CLM_URLS.REVIEW,
            "COMPLETE": CLM_URLS.COMPLETE,
            "COMPARE": CLM_URLS.COMPARE,
            "PAUSE": CLM_URLS.PAUSE
        },
        "SEAL": {
            "DRAFT": SEAL_URLS.DRAFT,
            "REVIEW": SEAL_URLS.REVIEW,
            "LEDGER": SEAL_URLS.LEDGER
        },
        "ADVICE": {
            "DRAFT": ADVICE_URLS.DRAFT,
            "PROCESS": ADVICE_URLS.PROCESS,
            "REVIEW": ADVICE_URLS.REVIEW
        },
        "LITIGATION": {
            "DRAFT": LITIGATION_URLS.DRAFT,
            "PROCESS": LITIGATION_URLS.PROCESS,
            "REVIEW": LITIGATION_URLS.REVIEW,
            "SCHEDULE": LITIGATION_URLS.SCHEDULE
        },
        "LAW": {
            "SCHEDULE": LAW_URLS.SCHEDULE
        },
        "PROJECT": {
            "PROJECT": PROJECT_URLS.PROJECT
        },
        "CONTRACT": {
            "CONTRACT": CONTRACT_URLS.CONTRACT,
            "STAMP": CONTRACT_URLS.STAMP,
            "LOGO": CONTRACT_URLS.LOGO,
            "TEAM_STAMP": CONTRACT_URLS.TEAM_STAMP,
            "WATERMARK": CONTRACT_URLS.WATERMARK
        },
        "SETTING": {
            "TEAM": SETTING_URLS.TEAM,
            "ACCOUNT": SETTING_URLS.ACCOUNT,
            "NOTIFICATION": SETTING_URLS.NOTIFICATION,
            "LOG": SETTING_URLS.LOG,
            "FAILEDLOG": SETTING_URLS.FAILEDLOG,
            "FA": SETTING_URLS.FA,
            "MANAGEMENT": SETTING_URLS.MANAGEMENT,
            "SETUP": SETTING_URLS.SETUP
        },
        "API": {
            "API_ROOT": "https://api.lawform.io/api/"
        }
    }
    
    # 평탄화된 URL 목록 생성 (카테고리.이름 형식)
    flat_urls = []
    for category, url_dict in urls.items():
        for name, url in url_dict.items():
            if url and url.strip():  # 빈 URL 제외
                flat_urls.append({
                    "category": category,
                    "name": name,
                    "url": url.strip(),
                    "key": f"{category}.{name}"
                })
    
    # JSON으로 출력
    print(json.dumps(flat_urls, indent=2, ensure_ascii=False))
    
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)
    sys.exit(1)


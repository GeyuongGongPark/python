import os
import json
from typing import Any, Dict

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None  # type: ignore


def _ensure_dotenv_loaded() -> None:
    """Load .env once if python-dotenv is available."""
    # load_dotenv가 없으면 조용히 패스 (컨테이너/배포 환경에서 이미 주입된 경우)
    if load_dotenv is not None:
        load_dotenv()


def _workspace_root() -> str:
    """프로젝트 루트 기준 경로 추정: 현재 파일 기준 상위로 계산."""
    # utils/account_env.py → 프로젝트 루트는 상위 상위 디렉토리로 가정
    return os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))


def _resolve_account_json_path(account_key: str) -> str:
    root = _workspace_root()
    candidate = os.path.join(root, "Account", f"{account_key}.json")
    return candidate


def load_account_env(account_key: str | None = None) -> Dict[str, Any]:
    """
    .env의 ACCOUNT 키(또는 인자로 받은 키)에 해당하는 Account/<key>.json을 로드해
    1차 필드들을 환경변수로 주입하고(dict도 반환)합니다.

    - account_key가 None이면 .env의 ACCOUNT를 사용
    - 문자열/숫자/불리언 값만 환경변수로 주입(문자열화). 중첩 구조는 반환 dict로 직접 사용하세요.
    """
    _ensure_dotenv_loaded()

    key = account_key or os.getenv("ACCOUNT")
    if not key:
        raise RuntimeError("ACCOUNT 환경변수가 필요합니다 (.env 또는 인자).")

    json_path = _resolve_account_json_path(key)
    if not os.path.exists(json_path):
        raise FileNotFoundError(f"계정 파일을 찾을 수 없습니다: {json_path}")

    with open(json_path, "r", encoding="utf-8") as f:
        data: Dict[str, Any] = json.load(f)

    for k, v in data.items():
        if isinstance(v, (str, int, float, bool)):
            os.environ.setdefault(str(k), str(v))

    return data


__all__ = ["load_account_env"]



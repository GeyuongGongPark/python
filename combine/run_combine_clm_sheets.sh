#!/bin/bash

# CLM 시트 통합 실행 스크립트
# 사용법:
#   ./run_combine_clm_sheets.sh [엑셀파일경로]

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Python 실행 경로 확인
if ! command -v python3 &> /dev/null; then
    if ! command -v python &> /dev/null; then
        echo "❌ 오류: Python이 설치되어 있지 않습니다."
        exit 1
    else
        PYTHON_CMD="python"
    fi
else
    PYTHON_CMD="python3"
fi

# 대상 경로: 인자가 없으면 현재 폴더, 있으면 전달값(파일 또는 디렉토리)
TARGET_PATH="${1:-$SCRIPT_DIR}"
if [[ ! "$TARGET_PATH" = /* ]]; then
    TARGET_PATH="$(realpath "$TARGET_PATH")"
fi

if [ ! -e "$TARGET_PATH" ]; then
    echo "❌ 오류: 경로를 찾을 수 없습니다: $TARGET_PATH"
    exit 1
fi

echo "🚀 통합 시작..."
echo "   입력 대상: $TARGET_PATH"
echo ""

"$PYTHON_CMD" "$SCRIPT_DIR/combine_clm_sheets.py" "$TARGET_PATH"

STATUS=$?
if [ $STATUS -eq 0 ]; then
    if [ -d "$TARGET_PATH" ]; then
        echo ""
        echo "✅ 통합 완료!"
        echo "   결과 폴더: $TARGET_PATH/done"
    else
        echo ""
        echo "✅ 통합 완료!"
        echo "   결과 파일은 입력 파일과 동일 폴더의 done 하위에 저장되었습니다."
    fi
else
    echo ""
    echo "❌ 처리 중 오류가 발생했습니다. (코드: $STATUS)"
    exit $STATUS
fi



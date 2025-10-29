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

# 파일 경로 확인
if [ $# -eq 0 ]; then
    # 인자가 없으면 현재 디렉토리에서 .xlsx 파일 자동 탐색 (이름에 CLM 포함 우선)
    EXCEL_FILE=$(ls -1 *.xlsx 2>/dev/null | grep -i "clm 이관_하림" | head -n 1)
    if [ -z "$EXCEL_FILE" ]; then
        EXCEL_FILE=$(ls -1 *.xlsx 2>/dev/null | head -n 1)
    fi

    if [ -z "$EXCEL_FILE" ]; then
        echo "📁 현재 디렉토리에서 엑셀 파일(.xlsx)을 찾을 수 없습니다."
        echo ""
        echo "사용법:"
        echo "  ./run_combine_clm_sheets.sh [엑셀파일경로]"
        echo ""
        echo "예시:"
        echo "  ./run_combine_clm_sheets.sh \"CLM 이관_하림_마이그레이션_애그리로보택.xlsx\""
        exit 1
    else
        echo "✅ 자동으로 파일을 찾았습니다: $EXCEL_FILE"
        EXCEL_FILE="$(realpath "$EXCEL_FILE")"
    fi
else
    EXCEL_FILE="$1"
    if [[ ! "$EXCEL_FILE" = /* ]]; then
        EXCEL_FILE="$(realpath "$EXCEL_FILE")"
    fi
    if [ ! -f "$EXCEL_FILE" ]; then
        echo "❌ 오류: 파일을 찾을 수 없습니다: $EXCEL_FILE"
        exit 1
    fi
fi

echo "🚀 통합 시작..."
echo "   입력 파일: $EXCEL_FILE"
echo ""

"$PYTHON_CMD" "$SCRIPT_DIR/combine_clm_sheets.py" "$EXCEL_FILE"

STATUS=$?
if [ $STATUS -eq 0 ]; then
    OUTPUT_FILE="$(dirname "$EXCEL_FILE")/CLM_통합.xlsx"
    echo ""
    echo "✅ 통합 완료!"
    echo "   결과 파일: $OUTPUT_FILE"
else
    echo ""
    echo "❌ 처리 중 오류가 발생했습니다. (코드: $STATUS)"
    exit $STATUS
fi



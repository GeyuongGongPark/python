#!/bin/bash

# CLM 처리 스크립트 실행 쉘 스크립트
# 사용법:
#   ./run_clm_process.sh [엑셀파일경로] [--explode]

# 스크립트가 있는 디렉토리로 이동
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
    # 인자가 없으면 현재 디렉토리에서 CLM 파일 찾기
    EXCEL_FILE=$(find . -maxdepth 1 -name "*CLM*.xlsx" -type f | head -n 1)
    
    if [ -z "$EXCEL_FILE" ]; then
        echo "📁 현재 디렉토리에서 CLM 엑셀 파일을 찾을 수 없습니다."
        echo ""
        echo "사용법:"
        echo "  ./run_clm_process.sh [엑셀파일경로] [--explode]"
        echo ""
        echo "예시:"
        echo "  ./run_clm_process.sh \"CLM 이관_하림_마이그레이션_애그리보텍_완.xlsx\""
        echo "  ./run_clm_process.sh \"CLM 이관_하림_마이그레이션_애그리보텍_완.xlsx\" --explode"
        exit 1
    else
        echo "✅ 자동으로 파일을 찾았습니다: $EXCEL_FILE"
        EXCEL_FILE="$(realpath "$EXCEL_FILE")"
    fi
else
    # 첫 번째 인자를 파일 경로로 사용
    EXCEL_FILE="$1"
    
    # 상대 경로인 경우 절대 경로로 변환
    if [[ ! "$EXCEL_FILE" = /* ]]; then
        EXCEL_FILE="$(realpath "$EXCEL_FILE")"
    fi
    
    # 파일 존재 확인
    if [ ! -f "$EXCEL_FILE" ]; then
        echo "❌ 오류: 파일을 찾을 수 없습니다: $EXCEL_FILE"
        exit 1
    fi
fi

# explode 옵션 확인
EXPLODE_OPT=""
if [ $# -ge 2 ] && [ "$2" == "--explode" ]; then
    EXPLODE_OPT="--explode"
    echo "📊 행 확장 모드로 실행합니다."
else
    echo "📊 병합 모드로 실행합니다. (기본값)"
fi

# Python 스크립트 실행
echo "🚀 CLM 처리 시작..."
echo "   입력 파일: $EXCEL_FILE"
echo ""

"$PYTHON_CMD" "$SCRIPT_DIR/clm_process.py" "$EXCEL_FILE" $EXPLODE_OPT

if [ $? -eq 0 ]; then
    OUTPUT_FILE="$(dirname "$EXCEL_FILE")/CLM_등록_처리결과.xlsx"
    echo ""
    echo "✅ 처리 완료!"
    echo "   결과 파일: $OUTPUT_FILE"
else
    echo ""
    echo "❌ 처리 중 오류가 발생했습니다."
    exit 1
fi


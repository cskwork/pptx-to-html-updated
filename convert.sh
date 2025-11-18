#!/bin/bash
################################################################################
# PPTX to HTML Converter - Shell Script (macOS/Linux)
# 사용법: ./convert.sh <input.pptx> [output_dir] [dpi]
################################################################################

# 색상 정의
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 스크립트 디렉토리 경로
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
CONVERTER="$SCRIPT_DIR/scripts/convert_pptx_to_html_v2.py"

# 도움말 표시
show_help() {
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${GREEN}  PPTX to HTML Converter${NC}"
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo ""
    echo "사용법:"
    echo "  ./convert.sh <input.pptx> [output_dir] [dpi]"
    echo ""
    echo "옵션:"
    echo "  input.pptx    변환할 PowerPoint 파일 (필수)"
    echo "  output_dir    출력 디렉토리 (기본값: output/)"
    echo "  dpi           이미지 품질 DPI (기본값: 150)"
    echo ""
    echo "DPI 설정 가이드:"
    echo "  72   - 빠른 변환, 작은 파일 크기"
    echo "  96   - 표준 웹 품질"
    echo "  150  - 권장 (기본값) - 고품질"
    echo "  300  - 최고 품질, 큰 파일 크기"
    echo ""
    echo "예시:"
    echo "  ./convert.sh presentation.pptx"
    echo "  ./convert.sh presentation.pptx my_output/"
    echo "  ./convert.sh presentation.pptx my_output/ 300"
    echo ""
}

# 헤더 출력
print_header() {
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${GREEN}  PPTX to HTML Converter v2.0${NC}"
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo ""
}

# 인자 확인 및 자동 감지
if [ -z "$1" ]; then
    # input/ 폴더에서 .pptx 파일 찾기
    shopt -s nullglob
    PPTX_FILES=(input/*.pptx)
    shopt -u nullglob
    
    if [ ${#PPTX_FILES[@]} -eq 1 ]; then
        INPUT_FILE="${PPTX_FILES[0]}"
        echo -e "${YELLOW}ℹ️  입력 파일이 지정되지 않았습니다. input/ 폴더의 파일을 자동으로 선택합니다.${NC}"
    elif [ ${#PPTX_FILES[@]} -gt 1 ]; then
        echo -e "${RED}❌ 오류: input/ 폴더에 여러 개의 PowerPoint 파일이 있습니다. 변환할 파일을 지정해주세요.${NC}"
        echo "발견된 파일:"
        for file in "${PPTX_FILES[@]}"; do
            echo "  - $file"
        done
        exit 1
    else
        echo -e "${RED}❌ 오류: 입력 파일을 지정해주세요${NC}"
        echo ""
        show_help
        exit 1
    fi
elif [ "$1" == "-h" ] || [ "$1" == "--help" ]; then
    show_help
    exit 0
else
    INPUT_FILE="$1"
fi

OUTPUT_DIR="${2:-output}"
DPI="${3:-150}"

# 헤더 출력
print_header

# 입력 파일 확인
if [ ! -f "$INPUT_FILE" ]; then
    echo -e "${RED}❌ 오류: 파일을 찾을 수 없습니다: $INPUT_FILE${NC}"
    exit 1
fi

# 파일 확장자 확인
if [[ ! "$INPUT_FILE" =~ \.pptx?$ ]]; then
    echo -e "${YELLOW}⚠️  경고: .pptx 파일이 아닙니다. 계속 진행합니다...${NC}"
fi

# Python 확인
if ! command -v python3 &> /dev/null; then
    if ! command -v python &> /dev/null; then
        echo -e "${RED}❌ 오류: Python이 설치되어 있지 않습니다${NC}"
        exit 1
    fi
    PYTHON_CMD="python"
else
    PYTHON_CMD="python3"
fi

# 변환기 스크립트 확인
if [ ! -f "$CONVERTER" ]; then
    echo -e "${RED}❌ 오류: 변환기를 찾을 수 없습니다: $CONVERTER${NC}"
    exit 1
fi

# 출력 디렉토리 생성
mkdir -p "$OUTPUT_DIR"

# 설정 정보 출력
echo -e "${BLUE}📄 입력 파일:${NC} $INPUT_FILE"
echo -e "${BLUE}📁 출력 디렉토리:${NC} $OUTPUT_DIR"
echo -e "${BLUE}🖼️  이미지 DPI:${NC} $DPI"
echo -e "${BLUE}🐍 Python:${NC} $PYTHON_CMD"
echo ""
echo -e "${YELLOW}🔄 변환 시작...${NC}"
echo ""

# 변환 실행
START_TIME=$(date +%s)

$PYTHON_CMD "$CONVERTER" "$INPUT_FILE" "$OUTPUT_DIR" $DPI

EXIT_CODE=$?
END_TIME=$(date +%s)
DURATION=$((END_TIME - START_TIME))

echo ""
echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"

# 결과 확인
    if [ $EXIT_CODE -eq 0 ]; then
    echo -e "${GREEN}✅ 변환 완료!${NC} (소요 시간: ${DURATION}초)"
    echo ""
    echo -e "${BLUE}📂 출력 파일:${NC}"

    # 입력 파일명에서 확장자 제거 및 기본 이름 추출
    BASENAME=$(basename "$INPUT_FILE")
    FILENAME="${BASENAME%.*}"
    HTML_FILE="$OUTPUT_DIR/$FILENAME.html"
    REPORT_FILE="$OUTPUT_DIR/${FILENAME}_report.md"

    # 생성된 파일 목록
    if [ -f "$HTML_FILE" ]; then
        echo -e "  ${GREEN}✓${NC} $HTML_FILE"
    elif [ -f "$OUTPUT_DIR/presentation.html" ]; then
        # 하위 호환성 또는 기본값
        HTML_FILE="$OUTPUT_DIR/presentation.html"
        echo -e "  ${GREEN}✓${NC} $HTML_FILE"
    fi
    
    if [ -f "$REPORT_FILE" ]; then
        echo -e "  ${GREEN}✓${NC} $REPORT_FILE"
    elif [ -f "$OUTPUT_DIR/presentation_report.md" ]; then
        echo -e "  ${GREEN}✓${NC} $OUTPUT_DIR/presentation_report.md"
    fi

    if [ -f "$OUTPUT_DIR/conversion.log" ]; then
        echo -e "  ${GREEN}✓${NC} $OUTPUT_DIR/conversion.log"
    fi
    if [ -d "$OUTPUT_DIR/assets" ]; then
        ASSET_COUNT=$(find "$OUTPUT_DIR/assets" -type f | wc -l)
        echo -e "  ${GREEN}✓${NC} $OUTPUT_DIR/assets/ (${ASSET_COUNT}개 파일)"
    fi

    echo ""
    echo -e "${BLUE}🌐 HTML 파일 열기:${NC}"
    echo -e "  open \"$HTML_FILE\""
    echo ""

    # macOS에서 자동으로 열기 제안
    read -p "지금 브라우저에서 열까요? (y/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        if command -v open &> /dev/null; then
            open "$HTML_FILE"
        elif command -v xdg-open &> /dev/null; then
            xdg-open "$HTML_FILE"
        else
            echo -e "${YELLOW}⚠️  브라우저를 자동으로 열 수 없습니다. 수동으로 파일을 열어주세요.${NC}"
        fi
    fi
else
    echo -e "${RED}❌ 변환 실패 (종료 코드: $EXIT_CODE)${NC}"
    echo ""
    echo -e "${YELLOW}📋 로그 파일을 확인하세요:${NC}"
    if [ -f "$OUTPUT_DIR/conversion.log" ]; then
        echo -e "  $OUTPUT_DIR/conversion.log"
    fi
    exit $EXIT_CODE
fi

echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"

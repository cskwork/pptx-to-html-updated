#!/bin/bash
################################################################################
# PPTX to HTML Converter - Shell Script (macOS/Linux)
# 사용법: ./convert.sh [input.pptx] [output_dir] [dpi]
# 인자가 없으면 input/ 폴더의 모든 .pptx 파일을 변환합니다.
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
INPUT_DIR="$SCRIPT_DIR/input"

# 도움말 표시
show_help() {
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${GREEN}  PPTX to HTML Converter${NC}"
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo ""
    echo "사용법:"
    echo "  ./convert.sh [input.pptx] [output_dir] [dpi]"
    echo ""
    echo "옵션:"
    echo "  input.pptx    변환할 PowerPoint 파일 (선택, 없으면 input/ 폴더의 모든 파일 변환)"
    echo "  output_dir    출력 기본 디렉토리 (기본값: output/)"
    echo "  dpi           이미지 품질 DPI (기본값: 150)"
    echo ""
    echo "DPI 설정 가이드:"
    echo "  72   - 빠른 변환, 작은 파일 크기"
    echo "  96   - 표준 웹 품질"
    echo "  150  - 권장 (기본값) - 고품질"
    echo "  300  - 최고 품질, 큰 파일 크기"
    echo ""
    echo "예시:"
    echo "  ./convert.sh                                    # input/ 폴더의 모든 파일 변환"
    echo "  ./convert.sh presentation.pptx                  # 단일 파일 변환"
    echo "  ./convert.sh presentation.pptx my_output/ 300   # 단일 파일, 커스텀 출력 및 DPI"
    echo ""
}

# 헤더 출력
print_header() {
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${GREEN}  PPTX to HTML Converter v2.0${NC}"
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo ""
}

# 단일 파일 변환 함수
convert_file() {
    local INPUT_FILE="$1"
    local OUTPUT_DIR="$2"
    local DPI="$3"
    
    # 입력 파일 확인
    if [ ! -f "$INPUT_FILE" ]; then
        echo -e "${RED}❌ 오류: 파일을 찾을 수 없습니다: $INPUT_FILE${NC}"
        return 1
    fi

    # 파일 확장자 확인
    if [[ ! "$INPUT_FILE" =~ \.pptx?$ ]]; then
        echo -e "${YELLOW}⚠️  경고: .pptx 파일이 아닙니다. 계속 진행합니다...${NC}"
    fi

    # 출력 디렉토리 생성
    mkdir -p "$OUTPUT_DIR"

    # 변환 실행
    local START_TIME=$(date +%s)
    $PYTHON_CMD "$CONVERTER" "$INPUT_FILE" "$OUTPUT_DIR" $DPI
    local EXIT_CODE=$?
    local END_TIME=$(date +%s)
    local DURATION=$((END_TIME - START_TIME))

    if [ $EXIT_CODE -eq 0 ]; then
        echo -e "${GREEN}✅ 변환 완료!${NC} (소요 시간: ${DURATION}초)"
        
        # 입력 파일명에서 확장자 제거 및 기본 이름 추출
        local BASENAME=$(basename "$INPUT_FILE")
        local FILENAME="${BASENAME%.*}"
        local HTML_FILE="$OUTPUT_DIR/$FILENAME.html"
        
        if [ -f "$HTML_FILE" ]; then
            echo -e "  ${GREEN}✓${NC} $HTML_FILE"
        fi
        return 0
    else
        echo -e "${RED}❌ 변환 실패 (종료 코드: $EXIT_CODE)${NC}"
        return $EXIT_CODE
    fi
}

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

# 헤더 출력
print_header

# 인자 확인 및 파일 목록 결정
if [ "$1" == "-h" ] || [ "$1" == "--help" ]; then
    show_help
    exit 0
fi

# 파일 목록 결정 및 출력 디렉토리 설정
# $1이 .pptx 파일이거나 존재하는 파일인지 확인
if [ -z "$1" ] || ([ ! -f "$1" ] && [ ! -f "$INPUT_DIR/$1" ] && [[ ! "$1" =~ \.pptx?$ ]]); then
    # 배치 모드: $1이 없거나 파일이 아닌 경우 (output_dir로 간주)
    OUTPUT_BASE_DIR="${1:-output}"
    DPI="${2:-150}"
    # input/ 폴더에서 .pptx 파일 찾기
    shopt -s nullglob
    PPTX_FILES=("$INPUT_DIR"/*.pptx)
    shopt -u nullglob
    
    if [ ${#PPTX_FILES[@]} -eq 0 ]; then
        echo -e "${RED}❌ 오류: input/ 폴더에 .pptx 파일이 없습니다${NC}"
        echo ""
        show_help
        exit 1
    fi
    
    echo -e "${YELLOW}ℹ️  입력 파일이 지정되지 않았습니다. input/ 폴더의 모든 파일을 변환합니다.${NC}"
    echo -e "${BLUE}📁 발견된 파일: ${#PPTX_FILES[@]}개${NC}"
    echo ""
    
    # 배치 모드: 모든 파일 변환
    TOTAL_FILES=${#PPTX_FILES[@]}
    SUCCESS_COUNT=0
    FAIL_COUNT=0
    
    for i in "${!PPTX_FILES[@]}"; do
        INPUT_FILE="${PPTX_FILES[$i]}"
        BASENAME=$(basename "$INPUT_FILE")
        FILENAME="${BASENAME%.*}"
        
        # 각 파일마다 별도의 출력 폴더 생성
        FILE_OUTPUT_DIR="$OUTPUT_BASE_DIR/$FILENAME"
        
        echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
        echo -e "${BLUE}[$((i+1))/$TOTAL_FILES] 변환 중: ${NC}$BASENAME"
        echo -e "${BLUE}📁 출력 폴더:${NC} $FILE_OUTPUT_DIR"
        echo ""
        
        if convert_file "$INPUT_FILE" "$FILE_OUTPUT_DIR" "$DPI"; then
            SUCCESS_COUNT=$((SUCCESS_COUNT + 1))
        else
            FAIL_COUNT=$((FAIL_COUNT + 1))
        fi
        echo ""
    done
    
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${GREEN}✅ 전체 변환 완료!${NC}"
    echo -e "${GREEN}   성공: $SUCCESS_COUNT개${NC}"
    if [ $FAIL_COUNT -gt 0 ]; then
        echo -e "${RED}   실패: $FAIL_COUNT개${NC}"
    fi
    echo ""
    
    if [ $FAIL_COUNT -eq 0 ]; then
        exit 0
    else
        exit 1
    fi
else
    # 단일 파일 모드: $1은 파일명, $2는 output_dir, $3는 dpi
    INPUT_FILE="$1"
    OUTPUT_BASE_DIR="${2:-output}"
    DPI="${3:-150}"
    
    # 상대 경로인 경우 input/ 폴더에서 찾기
    if [ ! -f "$INPUT_FILE" ] && [ -f "$INPUT_DIR/$INPUT_FILE" ]; then
        INPUT_FILE="$INPUT_DIR/$INPUT_FILE"
    fi
    
    BASENAME=$(basename "$INPUT_FILE")
    FILENAME="${BASENAME%.*}"
    
    # 단일 파일도 별도 폴더에 저장
    FILE_OUTPUT_DIR="$OUTPUT_BASE_DIR/$FILENAME"
    
    echo -e "${BLUE}📄 입력 파일:${NC} $INPUT_FILE"
    echo -e "${BLUE}📁 출력 디렉토리:${NC} $FILE_OUTPUT_DIR"
    echo -e "${BLUE}🖼️  이미지 DPI:${NC} $DPI"
    echo -e "${BLUE}🐍 Python:${NC} $PYTHON_CMD"
    echo ""
    echo -e "${YELLOW}🔄 변환 시작...${NC}"
    echo ""
    
    if convert_file "$INPUT_FILE" "$FILE_OUTPUT_DIR" "$DPI"; then
        HTML_FILE="$FILE_OUTPUT_DIR/$FILENAME.html"
        echo ""
        echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
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
        echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
        exit 0
    else
        exit 1
    fi
fi

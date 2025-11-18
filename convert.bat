@echo off
REM ============================================================================
REM PPTX to HTML Converter - Batch File (Windows)
REM 사용법: convert.bat [input.pptx] [output_dir] [dpi]
REM 인자가 없으면 input/ 폴더의 모든 .pptx 파일을 변환합니다.
REM ============================================================================

setlocal enabledelayedexpansion

REM 스크립트 디렉토리 경로
set "SCRIPT_DIR=%~dp0"
set "CONVERTER=%SCRIPT_DIR%scripts\convert_pptx_to_html_v2.py"
set "INPUT_DIR=%SCRIPT_DIR%input"

REM 도움말 표시
if "%1"=="-h" goto :show_help
if "%1"=="--help" goto :show_help
if "%1"=="/?" goto :show_help

REM Python 확인
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [91m❌ 오류: Python이 설치되어 있지 않습니다[0m
    echo.
    echo Python 3.7 이상을 설치해주세요:
    echo https://www.python.org/downloads/
    exit /b 1
)

REM 변환기 스크립트 확인
if not exist "%CONVERTER%" (
    echo [91m❌ 오류: 변환기를 찾을 수 없습니다: %CONVERTER%[0m
    exit /b 1
)

REM 헤더 출력
call :print_header

REM 파일 목록 결정
if "%1"=="" (
    REM 배치 모드: input/ 폴더의 모든 .pptx 파일 변환
    set "OUTPUT_BASE_DIR=%~2"
    if "!OUTPUT_BASE_DIR!"=="" set "OUTPUT_BASE_DIR=output"
    set "DPI=%~3"
    if "!DPI!"=="" set "DPI=150"
    
    REM input/ 폴더에서 .pptx 파일 찾기
    set "FILE_COUNT=0"
    for %%F in ("%INPUT_DIR%\*.pptx") do (
        set /a FILE_COUNT+=1
        set "FILE_!FILE_COUNT!=%%F"
    )
    
    if !FILE_COUNT! equ 0 (
        echo [91m❌ 오류: input/ 폴더에 .pptx 파일이 없습니다[0m
        echo.
        goto :show_help
        exit /b 1
    )
    
    echo [93mℹ️  입력 파일이 지정되지 않았습니다. input/ 폴더의 모든 파일을 변환합니다.[0m
    echo [94m📁 발견된 파일: !FILE_COUNT!개[0m
    echo.
    
    set "SUCCESS_COUNT=0"
    set "FAIL_COUNT=0"
    
    for /L %%I in (1,1,!FILE_COUNT!) do (
        set "INPUT_FILE=!FILE_%%I!"
        for %%F in ("!INPUT_FILE!") do (
            set "BASENAME=%%~nxF"
            set "FILENAME=%%~nF"
        )
        
        set "FILE_OUTPUT_DIR=!OUTPUT_BASE_DIR!\!FILENAME!"
        
        echo ================================================================
        echo [%%I/!FILE_COUNT!] 변환 중: !BASENAME!
        echo [94m📁 출력 폴더:[0m !FILE_OUTPUT_DIR!
        echo.
        
        call :convert_file "!INPUT_FILE!" "!FILE_OUTPUT_DIR!" !DPI!
        if !ERRORLEVEL! equ 0 (
            set /a SUCCESS_COUNT+=1
        ) else (
            set /a FAIL_COUNT+=1
        )
        echo.
    )
    
    echo ================================================================
    echo [92m✅ 전체 변환 완료![0m
    echo [92m   성공: !SUCCESS_COUNT!개[0m
    if !FAIL_COUNT! gtr 0 (
        echo [91m   실패: !FAIL_COUNT!개[0m
    )
    echo.
    
    if !FAIL_COUNT! equ 0 (
        exit /b 0
    ) else (
        exit /b 1
    )
) else (
    REM 단일 파일 모드
    set "INPUT_FILE=%~1"
    set "OUTPUT_BASE_DIR=%~2"
    if "!OUTPUT_BASE_DIR!"=="" set "OUTPUT_BASE_DIR=output"
    set "DPI=%~3"
    if "!DPI!"=="" set "DPI=150"
    
    REM 상대 경로인 경우 input/ 폴더에서 찾기
    if not exist "!INPUT_FILE!" (
        if exist "%INPUT_DIR%\!INPUT_FILE!" (
            set "INPUT_FILE=%INPUT_DIR%\!INPUT_FILE!"
        )
    )
    
    REM 입력 파일 확인
    if not exist "!INPUT_FILE!" (
        echo [91m❌ 오류: 파일을 찾을 수 없습니다: %~1[0m
        exit /b 1
    )
    
    for %%F in ("!INPUT_FILE!") do (
        set "BASENAME=%%~nxF"
        set "FILENAME=%%~nF"
    )
    
    set "FILE_OUTPUT_DIR=!OUTPUT_BASE_DIR!\!FILENAME!"
    
    echo [94m📄 입력 파일:[0m !INPUT_FILE!
    echo [94m📁 출력 디렉토리:[0m !FILE_OUTPUT_DIR!
    echo [94m🖼️  이미지 DPI:[0m !DPI!
    echo.
    echo [93m🔄 변환 시작...[0m
    echo.
    
    call :convert_file "!INPUT_FILE!" "!FILE_OUTPUT_DIR!" !DPI!
    set "EXIT_CODE=!ERRORLEVEL!"
    
    if !EXIT_CODE! equ 0 (
        set "HTML_FILE=!FILE_OUTPUT_DIR!\!FILENAME!.html"
        echo.
        echo ================================================================
        echo [94m🌐 HTML 파일 열기:[0m
        echo   start "" "!HTML_FILE!"
        echo.
        
        REM 자동으로 열기 제안
        set /p "OPEN_NOW=지금 브라우저에서 열까요? (y/n): "
        if /i "!OPEN_NOW!"=="y" (
            start "" "!HTML_FILE!"
        )
        echo ================================================================
        exit /b 0
    ) else (
        exit /b !EXIT_CODE!
    )
)

goto :eof

REM 단일 파일 변환 함수
:convert_file
set "INPUT_FILE=%~1"
set "OUTPUT_DIR=%~2"
set "DPI=%~3"

REM 출력 디렉토리 생성
if not exist "!OUTPUT_DIR!" mkdir "!OUTPUT_DIR!"

REM 변환 실행
python "%CONVERTER%" "!INPUT_FILE!" "!OUTPUT_DIR!" !DPI!
set "EXIT_CODE=!ERRORLEVEL!"

if !EXIT_CODE! equ 0 (
    echo [92m✅ 변환 완료![0m
    
    for %%F in ("!INPUT_FILE!") do set "FILENAME=%%~nF"
    set "HTML_FILE=!OUTPUT_DIR!\!FILENAME!.html"
    
    if exist "!HTML_FILE!" (
        echo   [92m✓[0m !HTML_FILE!
    )
    exit /b 0
) else (
    echo [91m❌ 변환 실패 (종료 코드: !EXIT_CODE!)[0m
    exit /b !EXIT_CODE!
)

:show_help
echo ================================================================
echo   PPTX to HTML Converter
echo ================================================================
echo.
echo 사용법:
echo   convert.bat [input.pptx] [output_dir] [dpi]
echo.
echo 옵션:
echo   input.pptx    변환할 PowerPoint 파일 (선택, 없으면 input/ 폴더의 모든 파일 변환)
echo   output_dir    출력 기본 디렉토리 (기본값: output)
echo   dpi           이미지 품질 DPI (기본값: 150)
echo.
echo DPI 설정 가이드:
echo   72   - 빠른 변환, 작은 파일 크기
echo   96   - 표준 웹 품질
echo   150  - 권장 (기본값) - 고품질
echo   300  - 최고 품질, 큰 파일 크기
echo.
echo 예시:
echo   convert.bat                                    # input/ 폴더의 모든 파일 변환
echo   convert.bat presentation.pptx                  # 단일 파일 변환
echo   convert.bat presentation.pptx my_output 300   # 단일 파일, 커스텀 출력 및 DPI
echo.
exit /b 1

:print_header
echo ================================================================
echo   PPTX to HTML Converter v2.0
echo ================================================================
echo.
goto :eof

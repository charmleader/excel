@echo off
chcp 65001 >nul
title 엑셀 통합기 - 안전 실행

echo.
echo ========================================
echo    🚀 스마트 엑셀 파일 통합기
echo    안전 실행 모드
echo ========================================
echo.

:: Python 실행 권한 확인
echo [1/4] Python 환경 확인 중...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되어 있지 않습니다.
    echo    Python을 설치한 후 다시 실행해주세요.
    pause
    exit /b 1
)
echo ✅ Python 환경 확인 완료
echo.

:: 현재 폴더 권한 설정 시도
echo [2/4] 폴더 권한 설정 시도 중...
icacls . /grant %USERNAME%:F /T >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ 폴더 권한 설정 완료
) else (
    echo ⚠️  폴더 권한 설정 실패 (계속 진행)
)
echo.

:: 필요한 패키지 설치 확인
echo [3/4] 필요한 패키지 확인 중...
python -c "import pandas, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo 📦 필요한 패키지를 설치합니다...
    pip install pandas openpyxl xlrd >nul 2>&1
    if %errorlevel% equ 0 (
        echo ✅ 패키지 설치 완료
    ) else (
        echo ⚠️  패키지 설치 실패 (계속 진행)
    )
) else (
    echo ✅ 필요한 패키지 확인 완료
)
echo.

:: 엑셀 통합기 실행
echo [4/4] 엑셀 통합기 실행 중...
echo.
python "엑셀 통합기.py"

echo.
echo ========================================
echo           프로그램이 종료되었습니다
echo ========================================
pause

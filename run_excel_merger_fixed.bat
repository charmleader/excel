@echo off
chcp 65001 >nul
title 엑셀 통합기 - 오류 수정 버전

echo.
echo ========================================
echo    🚀 스마트 엑셀 파일 통합기
echo    오류 수정 및 안전 실행
echo ========================================
echo.

:: Python 환경 확인
echo [1/5] Python 환경 확인 중...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되어 있지 않습니다.
    pause
    exit /b 1
)
echo ✅ Python 환경 확인 완료
echo.

:: 필요한 패키지 설치
echo [2/5] 필요한 패키지 설치 중...
pip install pandas openpyxl xlrd >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ 패키지 설치 완료
) else (
    echo ⚠️  패키지 설치 실패 (계속 진행)
)
echo.

:: Excel 파일 오류 수정
echo [3/5] Excel 파일 오류 수정 중...
python fix_excel_errors.py
echo.

:: 폴더 권한 설정
echo [4/5] 폴더 권한 설정 중...
icacls . /grant %USERNAME%:F /T >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ 폴더 권한 설정 완료
) else (
    echo ⚠️  폴더 권한 설정 실패 (계속 진행)
)
echo.

:: 엑셀 통합기 실행
echo [5/5] 엑셀 통합기 실행 중...
echo.
python "엑셀 통합기.py"

echo.
echo ========================================
echo           프로그램이 종료되었습니다
echo ========================================
pause

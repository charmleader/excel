@echo off
chcp 65001 >nul
title 엑셀 통합기 - 관리자 권한 실행

echo.
echo ========================================
echo    🚀 스마트 엑셀 파일 통합기
echo    관리자 권한으로 실행
echo ========================================
echo.

:: 관리자 권한 확인
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ⚠️  관리자 권한이 필요합니다.
    echo    이 스크립트를 관리자로 실행해주세요.
    echo.
    echo 🔧 해결 방법:
    echo    1. 이 파일을 마우스 우클릭
    echo    2. "관리자 권한으로 실행" 선택
    echo.
    pause
    exit /b 1
)

echo ✅ 관리자 권한 확인 완료
echo.

:: 현재 폴더의 파일 권한 설정
echo [1/3] 파일 권한 설정 중...
icacls . /grant Everyone:F /T >nul 2>&1
echo ✅ 파일 권한 설정 완료
echo.

:: Python 실행 권한 확인
echo [2/3] Python 환경 확인 중...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되어 있지 않습니다.
    echo    Python을 설치한 후 다시 실행해주세요.
    pause
    exit /b 1
)
echo ✅ Python 환경 확인 완료
echo.

:: 엑셀 통합기 실행
echo [3/3] 엑셀 통합기 실행 중...
echo.
python "엑셀 통합기.py"

echo.
echo ========================================
echo           프로그램이 종료되었습니다
echo ========================================
pause

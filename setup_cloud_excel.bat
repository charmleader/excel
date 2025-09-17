@echo off
chcp 65001 >nul
title 클라우드 Excel 처리 시스템 설정

echo.
echo ========================================
echo    🌐 클라우드 Excel 처리 시스템
echo    Supabase 기반 설정 및 실행
echo ========================================
echo.

:: Python 환경 확인
echo [1/6] Python 환경 확인 중...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되어 있지 않습니다.
    echo 📥 Python을 설치한 후 다시 실행해주세요.
    pause
    exit /b 1
)
echo ✅ Python 환경 확인 완료
echo.

:: 필요한 패키지 설치
echo [2/6] 필요한 패키지 설치 중...
pip install streamlit pandas openpyxl xlrd requests >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ 패키지 설치 완료
) else (
    echo ⚠️  패키지 설치 실패 (계속 진행)
)
echo.

:: Supabase 설정 파일 확인
echo [3/6] Supabase 설정 확인 중...
if not exist "supabase_config.json" (
    echo ⚠️  Supabase 설정 파일이 없습니다.
    echo 📝 설정 파일을 생성합니다...
    
    echo {> supabase_config.json
    echo   "supabase_url": "https://your-project.supabase.co",>> supabase_config.json
    echo   "supabase_key": "your-anon-key",>> supabase_config.json
    echo   "project_name": "excel_processor">> supabase_config.json
    echo }>> supabase_config.json
    
    echo ✅ 설정 파일 생성 완료
    echo.
    echo 📋 Supabase 설정 방법:
    echo 1. https://supabase.com 에서 계정 생성
    echo 2. 새 프로젝트 생성
    echo 3. Settings ^> API에서 URL과 Key 복사
    echo 4. supabase_config.json 파일을 수정
    echo.
    echo ⏸️  설정을 완료한 후 Enter 키를 누르세요...
    pause
) else (
    echo ✅ Supabase 설정 파일 확인 완료
)
echo.

:: 데이터베이스 스키마 설정 안내
echo [4/6] 데이터베우스 스키마 설정 안내
echo 📋 다음 단계를 따라 데이터베이스를 설정하세요:
echo.
echo 1. Supabase 대시보드에서 SQL Editor 열기
echo 2. supabase_setup.sql 파일의 내용을 복사
echo 3. SQL Editor에 붙여넣기 후 실행
echo 4. 테이블 생성 완료 확인
echo.
echo ⏸️  데이터베이스 설정을 완료한 후 Enter 키를 누르세요...
pause
echo.

:: 웹 인터페이스 생성
echo [5/6] 웹 인터페이스 생성 중...
python cloud_excel_launcher.py --setup-only >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ 웹 인터페이스 생성 완료
) else (
    echo ⚠️  웹 인터페이스 생성 실패 (수동 생성)
)
echo.

:: Supabase 연결 테스트
echo [6/7] Supabase 연결 테스트 중...
python test_supabase_connection.py
if %errorlevel% equ 0 (
    echo ✅ Supabase 연결 테스트 성공
) else (
    echo ⚠️  Supabase 연결 테스트 실패
    echo 📋 설정을 확인한 후 다시 시도해주세요
)
echo.

:: 실행 옵션 선택
echo [7/7] 실행 옵션 선택
echo.
echo 1. 웹 인터페이스 실행 (권장)
echo 2. 명령줄 인터페이스 실행
echo 3. 연결 테스트만 실행
echo 4. 설정만 완료하고 종료
echo.
set /p choice="선택하세요 (1-4): "

if "%choice%"=="1" (
    echo.
    echo 🚀 웹 인터페이스를 시작합니다...
    echo 📱 브라우저에서 http://localhost:8501 로 접속하세요
    echo ⏹️  종료하려면 Ctrl+C를 누르세요
    echo.
    python cloud_excel_launcher.py
) else if "%choice%"=="2" (
    echo.
    echo 🖥️  명령줄 인터페이스를 시작합니다...
    echo.
    python excel_cloud_processor.py
) else if "%choice%"=="3" (
    echo.
    echo 🔍 Supabase 연결 테스트를 실행합니다...
    echo.
    python test_supabase_connection.py
) else if "%choice%"=="4" (
    echo.
    echo ✅ 설정이 완료되었습니다!
    echo 📋 다음 단계:
    echo 1. supabase_config.json 파일 수정
    echo 2. 데이터베이스 스키마 설정
    echo 3. test_supabase_connection.py 실행
    echo 4. cloud_excel_launcher.py 실행
    echo.
) else (
    echo ❌ 잘못된 선택입니다.
)

echo.
echo ========================================
echo           설정이 완료되었습니다
echo ========================================
pause

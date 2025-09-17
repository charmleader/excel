@echo off
chcp 65001 >nul
title 엑셀 파일 통합기 - 실행기

echo.
echo ========================================
echo    🚀 스마트 엑셀 파일 통합기
echo ========================================
echo.

echo [1/3] 기존 프로세스 종료 중...
taskkill /f /im python.exe >nul 2>&1
echo ✅ 기존 프로세스 정리 완료
echo.

echo [2/3] Streamlit 앱 시작 중...
echo 🌐 앱이 시작되면 브라우저가 자동으로 열립니다...
echo.

start "" "excel_merger_launcher.html"

echo [3/3] 서버 실행 중...
streamlit run excel_merger_web.py --server.port 8501 --server.headless true --browser.gatherUsageStats false

echo.
echo ========================================
echo           앱이 종료되었습니다
echo ========================================
pause

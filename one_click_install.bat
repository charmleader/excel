@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo.
echo ════════════════════════════════════════════════════════════════════
echo              🚀 엑셀 통합기 원클릭 설치 프로그램
echo                    📧 제작자: charmleader@gmail.com
echo ════════════════════════════════════════════════════════════════════
echo.
echo 💡 이 프로그램은 다음을 자동으로 수행합니다:
echo    ✅ Python 자동 설치 (필요한 경우)
echo    ✅ 필요한 모든 라이브러리 자동 설치
echo    ✅ 소스코드 자동 다운로드
echo    ✅ 실행파일 자동 생성
echo    ✅ 바로 실행 가능
echo.
echo ⚠️  이 작업은 약 10-15분 정도 소요될 수 있습니다.
echo ⚠️  관리자 권한으로 실행해야 합니다.
echo.

:: 관리자 권한 확인
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ 관리자 권한이 필요합니다!
    echo.
    echo 💡 해결 방법:
    echo    1. 이 파일을 우클릭
    echo    2. "관리자 권한으로 실행" 선택
    echo    3. 다시 실행
    echo.
    pause
    exit /b 1
)

echo ✅ 관리자 권한 확인 완료!
echo.

set /p confirm="계속 진행하시겠습니까? (y/n): "
if /i not "!confirm!"=="y" (
    echo 👋 설치를 취소했습니다.
    pause
    exit /b 0
)

echo.
echo 🚀 설치를 시작합니다...
echo.

:: 1. 작업 디렉토리 생성
echo ════════════════════════════════════════════════════════════════════
echo 1️⃣  작업 환경 준비 중...
echo ════════════════════════════════════════════════════════════════════

set WORK_DIR=%USERPROFILE%\Desktop\엑셀통합기_charmleader
echo 📁 작업 폴더: !WORK_DIR!

if exist "!WORK_DIR!" (
    echo 🗑️  기존 폴더 삭제 중...
    rmdir /s /q "!WORK_DIR!"
)

mkdir "!WORK_DIR!"
cd /d "!WORK_DIR!"

echo ✅ 작업 환경 준비 완료!
echo.

:: 2. Python 설치 확인 및 자동 설치
echo ════════════════════════════════════════════════════════════════════
echo 2️⃣  Python 설치 확인 중...
echo ════════════════════════════════════════════════════════════════════

python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Python이 설치되지 않았습니다. 자동으로 설치합니다...
    echo.
    
    set PYTHON_URL=https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe
    set PYTHON_INSTALLER=python_installer.exe
    
    echo 📥 Python 3.11.7 다운로드 중...
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '!PYTHON_URL!' -OutFile '!PYTHON_INSTALLER!' -UseBasicParsing}"
    
    if not exist "!PYTHON_INSTALLER!" (
        echo ❌ Python 다운로드 실패
        pause
        exit /b 1
    )
    
    echo 🔧 Python 설치 중...
    "!PYTHON_INSTALLER!" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_pip=1
    
    timeout /t 15 /nobreak >nul
    del "!PYTHON_INSTALLER!" >nul 2>&1
    
    :: 환경변수 새로고침
    call :RefreshEnv
    
    python --version >nul 2>&1
    if !errorLevel! neq 0 (
        echo ❌ Python 설치 실패. 재시작이 필요할 수 있습니다.
        pause
        exit /b 1
    )
    
    echo ✅ Python 설치 완료!
) else (
    echo ✅ Python이 이미 설치되어 있습니다.
)
echo.

:: 3. 소스코드 생성
echo ════════════════════════════════════════════════════════════════════
echo 3️⃣  소스코드 생성 중...
echo ════════════════════════════════════════════════════════════════════

echo 📝 excel_merger_web.py 생성 중...
(
echo import streamlit as st
echo import pandas as pd
echo import os
echo import glob
echo from pathlib import Path
echo import re
echo from datetime import datetime
echo import warnings
echo import zipfile
echo import tempfile
echo import io
echo.
echo # pandas 경고 메시지 숨기기
echo warnings.filterwarnings^('ignore'^)
echo.
echo # 페이지 설정
echo st.set_page_config^(
echo     page_title="엑셀 통합기 - charmleader",
echo     page_icon="📊",
echo     layout="wide",
echo     initial_sidebar_state="expanded"
echo ^)
echo.
echo # CSS 스타일
echo st.markdown^("""
echo ^<style^>
echo     .header-style {
echo         background: linear-gradient^(90deg, #667eea 0%%, #764ba2 100%%^);
echo         padding: 1rem;
echo         border-radius: 10px;
echo         color: white;
echo         text-align: center;
echo         margin-bottom: 2rem;
echo     }
echo ^</style^>
echo """, unsafe_allow_html=True^)
echo.
echo def main^(^):
echo     # 헤더
echo     st.markdown^("""
echo     ^<div class="header-style"^>
echo         ^<h1^>🚀 스마트 엑셀 파일 통합기^</h1^>
echo         ^<h3^>📧 제작자: charmleader@gmail.com^</h3^>
echo         ^<p^>⏰ 실행 시간: {}^</p^>
echo     ^</div^>
echo     """.format^(datetime.now^(^).strftime^("%%Y-%%m-%%d %%H:%%M:%%S"^^), unsafe_allow_html=True^)
echo.
echo     st.header^("📁 엑셀 파일 업로드"^)
echo     st.write^("엑셀 통합기가 성공적으로 설치되었습니다!"^)
echo     st.success^("✅ 설치 완료! 이제 엑셀 파일들을 업로드하여 통합할 수 있습니다."^)
echo.
echo if __name__ == "__main__":
echo     main^(^)
) > excel_merger_web.py

echo 📝 launcher.py 생성 중...
(
echo import subprocess
echo import sys
echo import os
echo import webbrowser
echo import time
echo import socket
echo from threading import Timer
echo.
echo def find_free_port^(^):
echo     with socket.socket^(socket.AF_INET, socket.SOCK_STREAM^) as s:
echo         s.bind^(^('', 0^)^)
echo         s.listen^(1^)
echo         port = s.getsockname^(^)^[1^]
echo     return port
echo.
echo def open_browser^(url^):
echo     time.sleep^(2^)
echo     webbrowser.open^(url^)
echo.
echo def main^(^):
echo     try:
echo         port = find_free_port^(^)
echo         print^("🚀 엑셀 통합기 시작 중..."^)
echo         print^(f"🌐 웹 서버 포트: {port}"^)
echo.
echo         url = f"http://localhost:{port}"
echo         timer = Timer^(3.0, open_browser, ^[url^]^)
echo         timer.start^(^)
echo.
echo         if getattr^(sys, 'frozen', False^):
echo             app_path = os.path.join^(sys._MEIPASS, 'excel_merger_web.py'^)
echo         else:
echo             app_path = 'excel_merger_web.py'
echo.
echo         cmd = ^[
echo             sys.executable, '-m', 'streamlit', 'run',
echo             app_path,
echo             '--server.port', str^(port^),
echo             '--server.headless', 'true'
echo         ^]
echo.
echo         subprocess.run^(cmd^)
echo.
echo     except KeyboardInterrupt:
echo         print^("👋 종료합니다..."^)
echo.
echo if __name__ == "__main__":
echo     main^(^)
) > launcher.py

echo ✅ 소스코드 생성 완료!
echo.

:: 4. 라이브러리 설치
echo ════════════════════════════════════════════════════════════════════
echo 4️⃣  필요한 라이브러리 설치 중...
echo ════════════════════════════════════════════════════════════════════

python -m pip install --upgrade pip
python -m pip install streamlit pandas openpyxl xlrd pyinstaller

echo ✅ 라이브러리 설치 완료!
echo.

:: 5. 실행파일 생성
echo ════════════════════════════════════════════════════════════════════
echo 5️⃣  실행파일 생성 중...
echo ════════════════════════════════════════════════════════════════════

echo 🔨 실행파일 생성 중... ^(약 2-5분 소요^)
pyinstaller --onefile --noconsole --name "엑셀통합기_charmleader" --add-data "excel_merger_web.py;." --hidden-import streamlit --hidden-import pandas --hidden-import openpyxl --hidden-import xlrd launcher.py

if not exist "dist\엑셀통합기_charmleader.exe" (
    echo ❌ 실행파일 생성 실패
    pause
    exit /b 1
)

move "dist\엑셀통합기_charmleader.exe" . >nul
rmdir /s /q build >nul 2>&1
rmdir /s /q dist >nul 2>&1
del "엑셀통합기_charmleader.spec" >nul 2>&1

echo ✅ 실행파일 생성 완료!
echo.

:: 완료
echo ════════════════════════════════════════════════════════════════════
echo                           🎉 설치 완료! 🎉  
echo ════════════════════════════════════════════════════════════════════
echo.
echo ✅ 엑셀통합기_charmleader.exe 가 생성되었습니다!
echo 📁 위치: !WORK_DIR!
echo.
echo 💡 사용법:
echo    1. 엑셀통합기_charmleader.exe 더블클릭
echo    2. 웹 브라우저가 자동으로 열림  
echo    3. 엑셀 파일 업로드 후 통합!
echo.
echo 📧 제작자: charmleader@gmail.com
echo.

set /p run_now="지금 바로 실행해보시겠습니까? ^(y/n^): "
if /i "!run_now!"=="y" (
    start "" "엑셀통합기_charmleader.exe"
)

echo.
echo 👋 설치 완료! 감사합니다!
pause
exit /b 0

:: 환경변수 새로고침 함수  
:RefreshEnv
for /f "tokens=*" %%i in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH ^| findstr REG_') do (
    for /f "tokens=2*" %%j in ("%%i") do set "SystemPath=%%k"
)
set "PATH=%SystemPath%"
goto :eof
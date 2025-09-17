@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo.
echo ════════════════════════════════════════════════════════════════════
echo               🚀 엑셀 통합기 자동 설치 및 빌드 스크립트
echo                    📧 제작자: charmleader@gmail.com
echo ════════════════════════════════════════════════════════════════════
echo.

:: 관리자 권한 확인
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ⚠️  관리자 권한이 필요합니다!
    echo    이 스크립트를 우클릭 → "관리자 권한으로 실행"해주세요.
    echo.
    pause
    exit /b 1
)

echo ✅ 관리자 권한 확인 완료
echo.

:: 1. Python 설치 확인 및 자동 설치
echo ════════════════════════════════════════════════════════════════════
echo 1️⃣  Python 설치 상태 확인 중...
echo ════════════════════════════════════════════════════════════════════

python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Python이 설치되지 않았습니다. 자동으로 설치합니다...
    echo.
    
    :: Python 다운로드 URL (최신 3.11 버전)
    set PYTHON_URL=https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe
    set PYTHON_INSTALLER=python_installer.exe
    
    echo 📥 Python 3.11.7 다운로드 중... (약 2-5분 소요)
    echo    URL: !PYTHON_URL!
    echo.
    
    :: PowerShell을 사용하여 Python 다운로드
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '!PYTHON_URL!' -OutFile '!PYTHON_INSTALLER!' -UseBasicParsing}"
    
    if not exist "!PYTHON_INSTALLER!" (
        echo ❌ Python 다운로드에 실패했습니다.
        echo 💡 수동으로 https://python.org 에서 Python을 설치해주세요.
        pause
        exit /b 1
    )
    
    echo ✅ Python 다운로드 완료!
    echo.
    echo 🔧 Python 설치 중... (약 3-7분 소요)
    echo    ⚠️  설치 중에는 창을 닫지 마세요!
    echo.
    
    :: Python 자동 설치 (PATH 추가, pip 설치, 모든 사용자용)
    "!PYTHON_INSTALLER!" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_pip=1
    
    :: 설치 완료 대기
    timeout /t 10 /nobreak >nul
    
    echo 🗑️  임시 파일 정리...
    del "!PYTHON_INSTALLER!" >nul 2>&1
    
    echo.
    echo 🔄 시스템 환경변수 새로고침...
    echo    잠시 후 시스템을 다시 시작해야 할 수도 있습니다.
    
    :: 환경변수 새로고침
    call :RefreshEnv
    
    :: Python 설치 확인
    python --version >nul 2>&1
    if !errorLevel! neq 0 (
        echo ❌ Python 설치가 완료되지 않았습니다.
        echo 💡 컴퓨터를 재시작한 후 다시 실행해주세요.
        echo.
        set /p restart="지금 재시작하시겠습니까? (y/n): "
        if /i "!restart!"=="y" (
            shutdown /r /t 10 /c "Python 설치 완료 후 재시작"
            echo 🔄 10초 후 재시작됩니다...
        )
        pause
        exit /b 1
    )
    
    echo ✅ Python 설치 완료!
) else (
    echo ✅ Python이 이미 설치되어 있습니다.
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo    버전: %%i
)
echo.

:: 2. pip 업그레이드
echo ════════════════════════════════════════════════════════════════════
echo 2️⃣  pip 업그레이드 중...
echo ════════════════════════════════════════════════════════════════════
python -m pip install --upgrade pip
echo ✅ pip 업그레이드 완료!
echo.

:: 3. 가상환경 생성
echo ════════════════════════════════════════════════════════════════════
echo 3️⃣  가상환경 생성 중...
echo ════════════════════════════════════════════════════════════════════

if exist "excel_merger_env" (
    echo 🗑️  기존 가상환경 삭제 중...
    rmdir /s /q excel_merger_env
)

echo 📦 새로운 가상환경 생성 중...
python -m venv excel_merger_env

if not exist "excel_merger_env" (
    echo ❌ 가상환경 생성에 실패했습니다.
    pause
    exit /b 1
)

echo ✅ 가상환경 생성 완료!
echo.

:: 4. 가상환경 활성화 및 라이브러리 설치
echo ════════════════════════════════════════════════════════════════════
echo 4️⃣  필수 라이브러리 설치 중...
echo ════════════════════════════════════════════════════════════════════

call excel_merger_env\Scripts\activate

echo 📥 라이브러리 설치 중... (약 3-5분 소요)
echo    - streamlit (웹 인터페이스)
echo    - pandas (데이터 처리)  
echo    - openpyxl (엑셀 파일 처리)
echo    - xlrd (구형 엑셀 파일 지원)
echo    - pyinstaller (실행파일 생성)
echo.

pip install streamlit==1.28.1
pip install pandas==2.1.3  
pip install openpyxl==3.1.2
pip install xlrd==2.0.1
pip install pyinstaller==6.2.0

echo ✅ 라이브러리 설치 완료!
echo.

:: 5. 필수 파일 확인
echo ════════════════════════════════════════════════════════════════════
echo 5️⃣  필수 파일 확인 중...
echo ════════════════════════════════════════════════════════════════════

set MISSING_FILES=0

if not exist "excel_merger_web.py" (
    echo ❌ excel_merger_web.py 파일이 없습니다.
    set MISSING_FILES=1
)

if not exist "launcher.py" (
    echo ❌ launcher.py 파일이 없습니다.
    set MISSING_FILES=1
)

if !MISSING_FILES! equ 1 (
    echo.
    echo ⚠️  필수 파일이 누락되었습니다!
    echo 💡 다음 파일들을 같은 폴더에 넣어주세요:
    echo    - excel_merger_web.py (메인 웹 애플리케이션)
    echo    - launcher.py (런처 스크립트)
    echo.
    pause
    exit /b 1
)

echo ✅ 모든 필수 파일이 준비되었습니다!
echo.

:: 6. 실행파일 생성
echo ════════════════════════════════════════════════════════════════════
echo 6️⃣  실행파일 생성 중... (약 2-5분 소요)
echo ════════════════════════════════════════════════════════════════════

echo 🔨 PyInstaller로 실행파일 생성 중...
echo    이 과정에서 시간이 많이 걸릴 수 있습니다. 기다려주세요...
echo.

pyinstaller --onefile --noconsole --name "엑셀통합기_charmleader" --add-data "excel_merger_web.py;." --hidden-import streamlit --hidden-import pandas --hidden-import openpyxl --hidden-import xlrd launcher.py

if not exist "dist\엑셀통합기_charmleader.exe" (
    echo ❌ 실행파일 생성에 실패했습니다.
    echo 💡 오류 메시지를 확인하고 charmleader@gmail.com으로 문의해주세요.
    pause
    exit /b 1
)

echo ✅ 실행파일 생성 완료!
echo.

:: 7. 정리 작업
echo ════════════════════════════════════════════════════════════════════
echo 7️⃣  정리 작업 중...
echo ════════════════════════════════════════════════════════════════════

echo 📁 실행파일을 현재 폴더로 이동 중...
move "dist\엑셀통합기_charmleader.exe" . >nul 2>&1

echo 🗑️  임시 폴더 정리 중...
rmdir /s /q build >nul 2>&1
rmdir /s /q dist >nul 2>&1  
del "엑셀통합기_charmleader.spec" >nul 2>&1

echo ✅ 정리 작업 완료!
echo.

:: 완료 메시지
echo ════════════════════════════════════════════════════════════════════
echo                           🎉 설치 완료! 🎉
echo ════════════════════════════════════════════════════════════════════
echo.
echo ✅ 모든 작업이 성공적으로 완료되었습니다!
echo.
echo 📄 생성된 파일: 엑셀통합기_charmleader.exe
echo 💡 사용법:
echo    1. 위 실행파일을 더블클릭
echo    2. 웹 브라우저가 자동으로 열림
echo    3. 엑셀 파일들을 업로드하고 통합!
echo.
echo 🎁 특징:
echo    - 완전 자동화된 버전 관리
echo    - 웹 기반 사용자 친화적 인터페이스  
echo    - 안전한 파일 처리
echo    - 하나의 실행파일로 어디서든 사용 가능
echo.
echo 📧 제작자: charmleader@gmail.com
echo 💌 문의 및 피드백 언제든 환영합니다!
echo.
echo ════════════════════════════════════════════════════════════════════

set /p run_now="지금 바로 실행해보시겠습니까? (y/n): "
if /i "!run_now!"=="y" (
    echo.
    echo 🚀 엑셀 통합기를 실행합니다...
    start "" "엑셀통합기_charmleader.exe"
)

echo.
echo 👋 설치가 완료되었습니다. 감사합니다!
pause
exit /b 0

:: 환경변수 새로고침 함수
:RefreshEnv
for /f "tokens=*" %%i in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH ^| findstr REG_') do (
    for /f "tokens=2*" %%j in ("%%i") do set "SystemPath=%%k"
)
for /f "tokens=*" %%i in ('reg query "HKCU\Environment" /v PATH 2^>nul ^| findstr REG_') do (
    for /f "tokens=2*" %%j in ("%%i") do set "UserPath=%%k"
)
if defined UserPath (
    set "PATH=%SystemPath%;%UserPath%"
) else (
    set "PATH=%SystemPath%"
)
goto :eof
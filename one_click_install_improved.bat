@echo off
chcp 65001 >nul
title 스마트 엑셀 파일 통합기 - 원클릭 설치

echo.
echo ========================================
echo    🚀 스마트 엑셀 파일 통합기
echo    원클릭 설치 및 실행 스크립트
echo ========================================
echo.

:: 관리자 권한 확인
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ⚠️  관리자 권한이 필요합니다.
    echo    이 스크립트를 관리자로 실행해주세요.
    echo.
    pause
    exit /b 1
)

echo [1/8] 시스템 환경 확인 중...
echo.

:: Python 설치 확인
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되어 있지 않습니다.
    echo.
    echo [1-1] Python 자동 설치를 시작합니다...
    echo.
    
    :: Python 다운로드 및 설치
    echo Python 3.11.7 다운로드 중...
    powershell -Command "& {Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe' -OutFile 'python_installer.exe'}"
    
    if exist python_installer.exe (
        echo Python 설치 중... (자동 설치)
        python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
        del python_installer.exe
        
        :: PATH 새로고침
        call refreshenv
        
        echo ✅ Python 설치 완료
    ) else (
        echo ❌ Python 다운로드에 실패했습니다.
        echo    수동으로 Python을 설치해주세요: https://www.python.org/downloads/
        pause
        exit /b 1
    )
) else (
    echo ✅ Python이 이미 설치되어 있습니다.
)

echo.
echo [2/8] 필요한 패키지 설치 중...
echo.

:: pip 업그레이드
python -m pip install --upgrade pip

:: 필수 패키지 설치
echo Streamlit 설치 중...
pip install streamlit

echo Pandas 설치 중...
pip install pandas

echo OpenPyXL 설치 중...
pip install openpyxl

echo XLRD 설치 중...
pip install xlrd

echo PyInstaller 설치 중...
pip install pyinstaller

echo ✅ 모든 패키지 설치 완료
echo.

echo [3/8] 프로젝트 파일 확인 중...
if not exist "excel_merger_web.py" (
    echo ❌ excel_merger_web.py 파일을 찾을 수 없습니다.
    echo    프로젝트 폴더에서 실행해주세요.
    pause
    exit /b 1
)
echo ✅ 프로젝트 파일 확인 완료
echo.

echo [4/8] 실행 파일 생성 중...
echo PyInstaller로 실행 파일 생성 중...
pyinstaller --onefile --windowed --name "엑셀통합기" --icon=NONE excel_merger_web.py

if %errorlevel% neq 0 (
    echo ❌ 실행 파일 생성에 실패했습니다.
    echo    수동으로 실행해주세요: streamlit run excel_merger_web.py
    pause
    exit /b 1
)
echo ✅ 실행 파일 생성 완료
echo.

echo [5/8] 바탕화면 바로가기 생성 중...
set "desktop=%USERPROFILE%\Desktop"
set "shortcut=%desktop%\엑셀통합기.lnk"

powershell -Command "& {$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%shortcut%'); $Shortcut.TargetPath = '%CD%\dist\엑셀통합기.exe'; $Shortcut.WorkingDirectory = '%CD%'; $Shortcut.Description = '스마트 엑셀 파일 통합기'; $Shortcut.Save()}"

echo ✅ 바탕화면 바로가기 생성 완료
echo.

echo [6/8] 시작 메뉴 등록 중...
set "startmenu=%APPDATA%\Microsoft\Windows\Start Menu\Programs"
copy "dist\엑셀통합기.exe" "%startmenu%\엑셀통합기.exe" >nul 2>&1
echo ✅ 시작 메뉴 등록 완료
echo.

echo [7/8] 방화벽 규칙 추가 중...
netsh advfirewall firewall add rule name="엑셀통합기" dir=in action=allow protocol=TCP localport=8501 >nul 2>&1
echo ✅ 방화벽 규칙 추가 완료
echo.

echo [8/8] 설치 완료!
echo.
echo ========================================
echo           🎉 설치가 완료되었습니다!
echo ========================================
echo.
echo 📁 설치된 파일들:
echo    - dist\엑셀통합기.exe (실행 파일)
echo    - 바탕화면 바로가기
echo    - 시작 메뉴 등록
echo.
echo 🚀 실행 방법:
echo    1. 바탕화면의 '엑셀통합기' 바로가기 더블클릭
echo    2. 또는 dist 폴더의 엑셀통합기.exe 실행
echo    3. 또는 명령어: streamlit run excel_merger_web.py
echo.
echo 🌐 웹 브라우저에서 http://localhost:8501 접속
echo.

:: 자동 실행 여부 확인
set /p auto_run="지금 바로 실행하시겠습니까? (Y/N): "
if /i "%auto_run%"=="Y" (
    echo.
    echo 🚀 엑셀통합기를 실행합니다...
    start "" "dist\엑셀통합기.exe"
) else (
    echo.
    echo 💡 언제든지 바탕화면의 바로가기를 통해 실행할 수 있습니다.
)

echo.
echo ========================================
echo           설치 스크립트 완료
echo ========================================
pause

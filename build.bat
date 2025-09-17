@echo off
echo ========================================
echo    스마트 엑셀 파일 통합기 빌드 스크립트
echo ========================================
echo.

echo [1/5] Python 환경 확인 중...
python --version
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되어 있지 않습니다.
    echo    Python 3.7 이상을 설치해주세요.
    pause
    exit /b 1
)
echo ✅ Python 환경 확인 완료
echo.

echo [2/5] 필요한 패키지 설치 중...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ❌ 패키지 설치에 실패했습니다.
    pause
    exit /b 1
)
echo ✅ 패키지 설치 완료
echo.

echo [3/5] 파일 권한 설정 중...
icacls . /grant Everyone:F /T
echo ✅ 파일 권한 설정 완료
echo.

echo [4/5] 실행 파일 생성 중...
pyinstaller --onefile --windowed --name "엑셀통합기" excel_merger_web.py
if %errorlevel% neq 0 (
    echo ❌ 실행 파일 생성에 실패했습니다.
    pause
    exit /b 1
)
echo ✅ 실행 파일 생성 완료
echo.

echo [5/5] 빌드 완료!
echo.
echo 📁 생성된 파일들:
echo    - dist/엑셀통합기.exe
echo    - build/ (임시 파일들)
echo.
echo 🚀 실행 방법:
echo    1. dist 폴더로 이동
echo    2. 엑셀통합기.exe 실행
echo.
echo ========================================
echo           빌드가 완료되었습니다!
echo ========================================
pause
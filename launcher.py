import subprocess
import sys
import os
import webbrowser
import time
import socket
from threading import Timer

# Windows에서 유니코드 출력을 위한 인코딩 설정
if sys.platform.startswith('win'):
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())

def find_free_port():
    """사용 가능한 포트 찾기"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port

def open_browser(url):
    """브라우저 열기 (약간의 지연 후)"""
    time.sleep(2)
    webbrowser.open(url)

def main():
    try:
        # 사용 가능한 포트 찾기
        port = find_free_port()
        
        print("=" * 60)
        print("[시작] 엑셀 통합기 시작 중...")
        print("[정보] 제작자: charmleader.com")
        print("=" * 60)
        print(f"[포트] 웹 서버 포트: {port}")
        print("[대기] 잠시 후 브라우저가 자동으로 열립니다...")
        print("[종료] 종료하려면 Ctrl+C를 누르세요")
        print("=" * 60)
        
        # 브라우저 자동 열기 타이머 설정
        url = f"http://localhost:{port}"
        timer = Timer(3.0, open_browser, [url])
        timer.start()
        
        # Streamlit 앱 실행
        # 실행파일 경로에서 스크립트 파일 찾기
        if getattr(sys, 'frozen', False):
            # 실행파일로 실행된 경우
            app_path = os.path.join(sys._MEIPASS, 'excel_merger_web.py')
        else:
            # 스크립트로 실행된 경우
            app_path = 'excel_merger_web.py'
        
        # Streamlit 실행
        cmd = [
            sys.executable, '-m', 'streamlit', 'run',
            app_path,
            '--server.port', str(port),
            '--server.headless', 'true',
            '--browser.gatherUsageStats', 'false',
            '--server.fileWatcherType', 'none'
        ]
        
        subprocess.run(cmd)
        
    except KeyboardInterrupt:
        print("\n" + "=" * 60)
        print("[종료] 엑셀 통합기를 종료합니다...")
        print("[감사] 감사합니다! - charmleader.com")
        print("=" * 60)
    except Exception as e:
        print(f"\n[오류] 오류 발생: {e}")
        print("[신고] 이 오류를 charmleader.com으로 신고해주세요.")
        input("Enter 키를 눌러 종료하세요...")

if __name__ == "__main__":
    main()
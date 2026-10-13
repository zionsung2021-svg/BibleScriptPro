import sys
import os
import io
import time
import socket
import subprocess
import webbrowser
import threading
import multiprocessing
import uvicorn

# --windowed 모드에서 표준 입출력이 None이 되어 터지는 현상 방지
class DummyWriter:
    def write(self, data):
        pass
    def flush(self):
        pass
    def isatty(self):
        return False

if sys.stdout is None or not hasattr(sys.stdout, "isatty"):
    sys.stdout = DummyWriter()
if sys.stderr is None or not hasattr(sys.stderr, "isatty"):
    sys.stderr = DummyWriter()
if sys.stdin is None:
    sys.stdin = io.StringIO()

def is_port_in_use(port: int = 8000) -> bool:
    """8000번 포트가 이미 실행 중인지 확인"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def launch_app_window(url: str = "http://127.0.0.1:8000"):
    """웹 브라우저 대신 독립된 데스크톱 윈도우 창(App 모드)으로 실행"""
    # 윈도우 기본 엣지 브라우저 및 크롬 경로 탐색
    edge_paths = [
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe"),
    ]
    chrome_paths = [
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
    ]

    browser_exe = None
    for path in edge_paths + chrome_paths:
        if os.path.exists(path):
            browser_exe = path
            break

    if browser_exe:
        # 독립 윈도우 창 모드 실행 (--app 옵션)
        # 1280x820 기본 창 크기 지정
        subprocess.Popen([browser_exe, f"--app={url}", "--window-size=1280,820"])
    else:
        # 전용 브라우저 실행 실패 시 기본 브라우저 열기 fallback
        webbrowser.open(url)

def open_window_delayed():
    """서버 준비 시간 확보 후 윈도우 창 띄우기"""
    time.sleep(1.2)
    launch_app_window("http://127.0.0.1:8000")

if __name__ == "__main__":
    multiprocessing.freeze_support()

    # 이미 백그라운드에 서버가 돌아가고 있다면 창만 띄우고 즉시 종료
    if is_port_in_use(8000):
        launch_app_window("http://127.0.0.1:8000")
        sys.exit(0)

    # 윈도우 창 자동 실행 스레드
    threading.Thread(target=open_window_delayed, daemon=True).start()

    # 실행 경로 기준 패키지 등록
    if getattr(sys, 'frozen', False):
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)

    from backend.main import app

    no_log_config = {
        "version": 1,
        "disable_existing_loggers": False,
        "formatters": {},
        "handlers": {
            "default": {
                "class": "logging.NullHandler",
            },
        },
        "loggers": {
            "uvicorn": {"handlers": ["default"], "level": "CRITICAL"},
            "uvicorn.error": {"handlers": ["default"], "level": "CRITICAL"},
            "uvicorn.access": {"handlers": ["default"], "level": "CRITICAL"},
        },
    }

    uvicorn.run(
        app,
        host="127.0.0.1",
        port=8000,
        log_config=no_log_config,
        access_log=False
    )

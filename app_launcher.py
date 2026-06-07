"""
Windows 실행파일 진입점.
서버를 백그라운드 스레드로 시작하고 브라우저를 자동으로 엽니다.
"""
from __future__ import annotations

import sys
import threading
import time
import webbrowser
from pathlib import Path


def _ensure_data_dir() -> None:
    """
    PyInstaller --onefile 모드에서 _MEIPASS 안의 초기 데이터를
    exe 옆의 data/ 폴더로 복사합니다 (처음 실행 시 한 번만).
    """
    if not getattr(sys, "frozen", False):
        return  # 개발 환경에서는 불필요

    meipass = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    exe_dir = Path(sys.executable).parent
    src_data = meipass / "data"
    dst_data = exe_dir / "data"

    if src_data.exists() and not dst_data.exists():
        import shutil
        shutil.copytree(src_data, dst_data)


def main() -> None:
    _ensure_data_dir()

    import server  # noqa: PLC0415

    port = 5174

    def run_server() -> None:
        import os
        os.environ.setdefault("PORT", str(port))
        server.main()

    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()

    # 서버가 뜰 때까지 잠깐 대기 후 브라우저 오픈
    time.sleep(1.5)
    webbrowser.open(f"http://127.0.0.1:{port}")

    # 메인 스레드가 종료되지 않도록 유지
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()

import sys
import threading
import time
import webbrowser
import uvicorn
from main import app

HOST = "127.0.0.1"
PORT = 8765


def _open_browser():
    # 給 server 一點時間啟動再開瀏覽器
    time.sleep(1.2)
    webbrowser.open(f"http://{HOST}:{PORT}")


if __name__ == "__main__":
    threading.Thread(target=_open_browser, daemon=True).start()
    uvicorn.run(app, host=HOST, port=PORT)

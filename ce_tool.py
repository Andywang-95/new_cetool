import logging
import multiprocessing
import os
import threading
from pathlib import Path

# Force pywebview to use CEF backend instead of WinForms (to avoid CLR/pythonnet issues)
os.environ["PYWEBVIEW_WEBVIEW_BACKEND"] = "cef"

import webview
from screeninfo import get_monitors

from app import create_app
from app.desktop_api import Api, JsApi

# Setup logging to file for debugging packaged exe
log_dir = Path.home() / "CE_Tool_Logs"
log_dir.mkdir(exist_ok=True)
log_file = log_dir / "ce_tool.log"
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(log_file), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)
logger.info(f"CE Tool started. Log file: {log_file}")

multiprocessing.freeze_support()
monitor = get_monitors()[0]
screen_width = monitor.width
screen_height = monitor.height

app = create_app()
logger.info("Flask app created successfully")


def start_flask():
    logger.info("Starting Flask server on port 5001")
    app.run(port=5001, use_reloader=False)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        logger.info("Creating Flask thread")
        flask_thread = threading.Thread(target=start_flask)
        flask_thread.daemon = True
        flask_thread.start()

        logger.info("Creating webview API instance")
        api = Api(app)
        js_api = JsApi(api)

        logger.info("Creating webview window")
        window = webview.create_window(
            "CE BOM Tool",
            "http://127.0.0.1:5001",
            js_api=js_api,
            width=int(screen_width * 0.4),
            height=int(screen_height * 0.7),
        )
        api.window = window
        logger.info("Starting webview")
        webview.start()
    except Exception as e:
        import sys
        import traceback

        logger.error("Exception occurred", exc_info=True)
        title = "Webview 啟動失敗："
        try:
            print(title)
        except UnicodeEncodeError:
            try:
                sys.stderr.write(title + "\n")
            except Exception:
                pass

        tb = traceback.format_exc()
        try:
            print(tb)
        except UnicodeEncodeError:
            try:
                sys.stderr.buffer.write(tb.encode("utf-8", errors="replace"))
            except Exception:
                # last resort: write ascii-safe fallback
                try:
                    sys.stderr.write(
                        tb.encode("ascii", errors="replace").decode("ascii")
                    )
                except Exception:
                    pass

        try:
            input("按下 Enter 關閉...")
        except Exception:
            pass

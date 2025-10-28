import multiprocessing
import threading

import webview
from screeninfo import get_monitors

from app import create_app
from app.desktop_api import Api, JsApi

multiprocessing.freeze_support()
monitor = get_monitors()[0]
screen_width = monitor.width
screen_height = monitor.height

app = create_app()


def start_flask():
    app.run(port=5001, use_reloader=False)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        flask_thread = threading.Thread(target=start_flask)
        flask_thread.daemon = True
        flask_thread.start()

        api = Api(app)
        js_api = JsApi(api)
        window = webview.create_window(
            "CE BOM Tool",
            "http://127.0.0.1:5001",
            js_api=js_api,
            width=int(screen_width * 0.4),
            height=int(screen_height * 0.7),
        )
        api.window = window
        webview.start()
    except Exception as e:
        import traceback

        print("❌ Webview 啟動失敗：")
        print(traceback.format_exc())
        input("按下 Enter 關閉...")

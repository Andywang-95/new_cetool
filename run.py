import threading

import webview

from app import create_app
from app.desktop_api import Api, JsApi

app = create_app()


def start_flask():
    app.run(port=5001, debug=True, use_reloader=False)


if __name__ == "__main__":
    flask_thread = threading.Thread(target=start_flask)
    flask_thread.daemon = True
    flask_thread.start()

    api = Api(app)
    js_api = JsApi(api)
    window = webview.create_window(
        "CE BOM Tool", "http://127.0.0.1:5001", js_api=js_api
    )
    api.window = window
    webview.start(debug=True)

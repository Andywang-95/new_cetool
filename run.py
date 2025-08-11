import threading

import webview

from app import create_app
from app.desktop_api import Api

app = create_app()


def start_flask():
    app.run(port=5000, debug=True, use_reloader=False)


if __name__ == "__main__":
    flask_thread = threading.Thread(target=start_flask)
    flask_thread.daemon = True
    flask_thread.start()

    api = Api(app)
    window = webview.create_window("CE BOM Tool", "http://127.0.0.1:5001", js_api=api)
    api.window = window
    webview.start(debug=True)

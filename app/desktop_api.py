from typing import Optional

import webview
from flask import Flask

from .services import db_settings


class Api:
    def __init__(self, app: Flask):
        self.window: Optional[webview.Window] = None  # 這樣 Pylance 就知道了
        self.app = app

    def select_bom_path(self):
        if self.window:
            result = self.window.create_file_dialog(
                webview.OPEN_DIALOG, file_types=["Excel files (*.xlsx;*.xls)"]
            )
            return result[0] if result else None
        else:
            return None

    def save_settings(self, settings: dict):
        db_settings.save_settings(settings)
        with self.app.app_context():
            self.app.config.update(settings)

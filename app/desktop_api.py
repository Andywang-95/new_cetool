import json
import traceback
from typing import Optional

import webview
from flask import Flask

from app.services.review import ReviewService

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

    def logs(self, type: str, msg: str):
        if self.window:
            self.window.evaluate_js(
                f"""
                Alpine.store('logStore').addLog({json.dumps(type)}, {json.dumps(msg)});
                document.querySelectorAll('.log-textarea').forEach(el => {{
                    el.scrollTop = el.scrollHeight;
                }});
                """
            )
        else:
            print(msg)

    def run_review(self, method, bom_path, col, row):
        try:
            review = ReviewService(self.app.config, bom_path, self.logs)
            if method == "BOM_TipTop_PTC":
                review.run("C", 7, "main")
            elif method == "Result":
                review.run("B", 5, "result")
            elif method == "系統BOM":
                review.run("C", 2, "system")
            elif method == "自定義":
                review.run(col, int(row), "custom")
        except Exception as e:
            self.logs("review", f"Review failed: \n{traceback.format_exc()}")


class JsApi:
    """
    暴露接口給前端使用
    """

    def __init__(self, api: Api):
        self.select_bom_path = api.select_bom_path
        self.save_settings = api.save_settings
        self.run_review = api.run_review

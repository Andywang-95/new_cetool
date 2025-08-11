from flask import Flask

from app.services import db_settings


def create_app():
    app = Flask(__name__)

    # 預先載入設定檔（會自動建立）
    app.config.update(db_settings.load_settings())

    from . import routes

    app.register_blueprint(routes.bp)

    return app

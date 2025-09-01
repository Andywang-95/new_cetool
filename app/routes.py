from flask import Blueprint, current_app, jsonify, render_template

bp = Blueprint("main", __name__)


@bp.route("/")
def index():
    return render_template("index.html", **current_app.config)


@bp.route("/api/settings")
def get_settings():
    settings = {
        "database_path": current_app.config.get("database_path", ""),
        # "pn_location": current_app.config.get("pn_location", ""),
    }
    return jsonify(settings)

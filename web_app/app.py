"""web_app/app.py - Flaskアプリケーションファクトリ"""
from __future__ import annotations

import os
import secrets

from flask import Flask, redirect, session, url_for, Response

from .config import Config, APP_VERSION, APP_RELEASE_DATE
from .database import close_db, init_db


def create_app() -> Flask:
    """Flaskアプリケーションを生成・設定して返す。"""
    app = Flask(__name__)
    app.config.from_object(Config)

    with app.app_context():
        init_db(app)
    app.teardown_appcontext(close_db)

    from .routes.auth import auth_bp
    from .routes.schedule import schedule_bp
    from .routes.tasks import tasks_bp
    from .routes.admin import admin_bp
    from .routes.export import export_bp
    from .routes.daily import daily_bp
    from .routes.help import help_bp
    from .routes.mail_report import mail_report_bp
    from .routes.project_tasks import project_tasks_bp
    from .routes.api import api_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(schedule_bp)
    app.register_blueprint(tasks_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(export_bp)
    app.register_blueprint(daily_bp)
    app.register_blueprint(help_bp)
    app.register_blueprint(mail_report_bp)
    app.register_blueprint(project_tasks_bp)
    app.register_blueprint(api_bp)

    @app.after_request
    def _no_cache(response: Response) -> Response:
        """HTMLレスポンスのブラウザキャッシュを無効化する。"""
        if "text/html" in response.content_type:
            response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
            response.headers["Pragma"] = "no-cache"
        return response

    @app.before_request
    def _ensure_csrf_token() -> None:
        """全リクエスト前にCSRFトークンをセッションに生成する（なければ）。"""
        if "csrf_token" not in session:
            session["csrf_token"] = secrets.token_hex(32)

    @app.context_processor
    def _inject_globals() -> dict:
        """全テンプレートに csrf_token・アプリバージョンを注入する。"""
        return {
            "csrf_token": session.get("csrf_token", ""),
            "app_version": APP_VERSION,
            "app_release_date": APP_RELEASE_DATE,
            "env_label": os.environ.get("FLASK_ENV_LABEL", ""),
        }

    @app.route("/")
    def index():
        """ルートURL: ログイン済みなら週間予定へ、未ログインならログインへ"""
        if "user_id" in session:
            return redirect(url_for("schedule.weekly"))
        return redirect(url_for("auth.login"))

    return app

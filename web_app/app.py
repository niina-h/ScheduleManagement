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
    from .routes.planner import planner_bp
    from .routes.survey import survey_bp

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
    app.register_blueprint(planner_bp)
    app.register_blueprint(survey_bp)

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

    @app.before_request
    def _require_survey() -> Response | None:
        """未回答のログイン中ユーザーをアンケート画面へ強制的に誘導する。

        アンケート画面自体・ログアウト・静的ファイルは無限リダイレクトを避けるため対象外。
        「次回回答する」で見送った場合は、session["survey_skipped"] が立ち、
        今回のログインセッション中は再表示しない（ログアウト・再ログインで再度表示される）。
        """
        from flask import request as _request

        if _request.endpoint in (None, "survey_bp.survey", "survey_bp.survey_skip", "auth.logout", "static"):
            return None
        user_id = session.get("user_id")
        if not user_id:
            return None
        if session.get("survey_skipped"):
            return None
        from .models import has_answered_survey
        if not has_answered_survey(int(user_id)):
            return redirect(url_for("survey_bp.survey"))
        return None

    @app.context_processor
    def _inject_globals() -> dict:
        """全テンプレートに csrf_token・アプリバージョン・役職判定を注入する。

        役職はセッションに旧値（マスタ/ユーザー）が残っていても正しく判定できるよう、
        normalize_role で正規化した値と、判定済みフラグを渡す。テンプレート側は
        文字列比較ではなく role_is_* フラグを使うことで新旧の揺れに影響されない。
        """
        from .auth_helpers import (
            normalize_role, is_system_admin, is_dept_head, is_privileged,
        )
        raw_role = session.get("user_role", "")
        # ユーザー名メニューの所属切替用に、切替可能な所属一覧を注入（未ログイン時は空）。
        switchable_depts: list[str] = []
        if session.get("user_id"):
            try:
                from .routes.auth import get_switchable_depts
                switchable_depts = get_switchable_depts()
            except Exception:
                switchable_depts = []
        return {
            "csrf_token": session.get("csrf_token", ""),
            "app_version": APP_VERSION,
            "app_release_date": APP_RELEASE_DATE,
            "env_label": os.environ.get("FLASK_ENV_LABEL", ""),
            # 正規化済み役職と判定フラグ（テンプレートはこれらを使う）
            "current_role": normalize_role(raw_role),
            "role_is_system_admin": is_system_admin(raw_role),
            "role_is_dept_head": is_dept_head(raw_role),
            "role_is_privileged": is_privileged(raw_role),
            "switchable_depts": switchable_depts,
        }

    @app.route("/")
    def index():
        """ルートURL: ログイン済みなら週間予定へ、未ログインならログインへ"""
        if "user_id" in session:
            return redirect(url_for("schedule.weekly"))
        return redirect(url_for("auth.login"))

    return app

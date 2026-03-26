"""web_app/routes/help.py - ヘルプ・手順書ページ"""
from __future__ import annotations

from flask import Blueprint, abort, render_template

help_bp = Blueprint("help_bp", __name__)

# 許可されたページ名のセット（パストラバーサル対策）
_VALID_PAGES: frozenset[str] = frozenset({
    "index", "login", "schedule", "daily", "admin", "export", "version",
    "tasks", "mail_report", "project_tasks",
})


@help_bp.route("/help")
@help_bp.route("/help/<page_name>")
def help_page(page_name: str = "index") -> str:
    """ヘルプページを表示する。

    Args:
        page_name: 表示するヘルプページ名。許可リスト外は404を返す。

    Returns:
        str: レンダリングされたHTMLレスポンス。
    """
    if page_name not in _VALID_PAGES:
        abort(404)
    return render_template("help.html", page_name=page_name)

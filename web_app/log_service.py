"""web_app/log_service.py - 操作ログ記録サービス"""
from __future__ import annotations

import logging
from datetime import datetime

from flask import request, session

from .database import get_db

logger = logging.getLogger(__name__)

# アクション種別定数
ACTION_LOGIN          = "LOGIN"
ACTION_LOGOUT         = "LOGOUT"
ACTION_SCHEDULE_SAVE  = "SCHEDULE_SAVE"
ACTION_DAILY_SAVE     = "DAILY_SAVE"
ACTION_USER_ADD       = "USER_ADD"
ACTION_USER_UPDATE    = "USER_UPDATE"
ACTION_USER_DELETE    = "USER_DELETE"
ACTION_EXPORT         = "EXPORT"


def record_operation(action_type: str, detail: str = "") -> None:
    """操作ログをDBとロガーに記録する。

    Args:
        action_type: アクション種別（ACTION_* 定数を使用）。
        detail: 補足情報（ユーザー名・対象日付など）。
    """
    user_id: int | None = session.get("user_id")
    user_name: str = session.get("user_name", "")
    ip_address: str = request.remote_addr or ""
    created_at: str = datetime.now().isoformat(timespec="seconds")

    try:
        db = get_db()
        db.execute(
            "INSERT INTO operation_log (user_id, user_name, action_type, detail, ip_address, created_at)"
            " VALUES (?, ?, ?, ?, ?, ?)",
            (user_id, user_name, action_type, detail, ip_address, created_at),
        )
        db.commit()
    except Exception:
        logger.exception("操作ログのDB記録に失敗しました。")

    logger.info("[OP] %s user_id=%s ip=%s", action_type, user_id, ip_address)

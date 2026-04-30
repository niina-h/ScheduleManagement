"""ブラウザ JavaScript からの非同期取得用 API ルート群。

ログインユーザー本人のデータを JSON で返却する。
社内 LAN 内のみで完結し、外部通信は発生しない。
"""
from __future__ import annotations

from datetime import date
from typing import Any

from flask import Blueprint, jsonify, session

from ..models import get_events_for_user_date

api_bp = Blueprint("api_bp", __name__, url_prefix="/api")


@api_bp.route("/today-events")
def today_events() -> Any:
    """ログインユーザーが担当する本日のイベント一覧を JSON で返す。

    ブラウザ通知（開始10分前のリマインダー）の元データとして使用する。
    未ログイン時は 401 を返す。

    Returns:
        Response: ``{"events": [{...}, ...]}`` 形式の JSON レスポンス。
                  各イベントは task_name, event_start_time, event_end_time,
                  start_date, id, status を含む。
    """
    if not session.get("user_id"):
        return jsonify({"events": []}), 401

    user_id: int = int(session["user_id"])
    today_str: str = date.today().isoformat()
    events = get_events_for_user_date(user_id, today_str)

    # 通知に必要な最小フィールドのみ返す
    payload = [
        {
            "id": ev["id"],
            "task_name": ev.get("task_name", ""),
            "event_start_time": ev.get("event_start_time", ""),
            "event_end_time": ev.get("event_end_time", ""),
            "start_date": ev.get("start_date", ""),
        }
        for ev in events
        if ev.get("event_start_time")  # 開始時刻が登録されているもののみ
    ]
    return jsonify({"events": payload})

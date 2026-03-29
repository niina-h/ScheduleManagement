"""web_app/routes/daily.py - 日次実績管理ルート。

/daily/today          : 当日の日次実績へリダイレクト（GET）
/daily/<date_str>     : 日次実績の表示（GET）
/daily/save           : 日次実績の保存（POST）
"""
from __future__ import annotations

import json
from datetime import date, datetime, timedelta
from typing import Any

from flask import Blueprint, abort, flash, redirect, render_template, request, session, url_for

from ..models import (
    add_carryover,
    defer_task_to_weekly_schedule,
    get_accessible_users,
    get_all_users,
    get_daily_comment,
    get_daily_result,
    get_daily_result_meta,
    get_events_for_user_date,
    get_pending_carryovers,
    get_task_master,
    get_user_by_id,
    get_weekly_leave,
    get_weekly_schedule,
    remove_rescheduled_daily_result,
    remove_rescheduled_task,
    resolve_carryovers_by_task,
    save_admin_comment,
    save_daily_comment,
    save_daily_result,
    sync_daily_progress_to_task,
)
from ..auth_helpers import is_privileged, is_master, can_access_user

daily_bp = Blueprint("daily_bp", __name__)


# ---------------------------------------------------------------------------
# ヘルパー関数
# ---------------------------------------------------------------------------


def _get_monday(d: date) -> date:
    """指定日の週の月曜日を返す。

    Args:
        d: 基準となる日付。

    Returns:
        その週の月曜日の日付。
    """
    return d - timedelta(days=d.weekday())


def _get_week_start_for_date(d: date) -> str:
    """指定日が属する週の月曜日を ISO 形式文字列で返す。

    Args:
        d: 基準となる日付。

    Returns:
        週開始日（月曜日）の 'YYYY-MM-DD' 文字列。
    """
    return _get_monday(d).isoformat()


def _prev_weekday(d: date) -> date:
    """指定日の前の平日（土日をスキップ）を返す。

    Args:
        d: 基準となる日付。

    Returns:
        前の平日の日付。
    """
    prev = d - timedelta(days=1)
    while prev.weekday() >= 5:
        prev -= timedelta(days=1)
    return prev


def _next_weekday(d: date) -> date:
    """指定日の次の平日（土日をスキップ）を返す。

    Args:
        d: 基準となる日付。

    Returns:
        次の平日の日付。
    """
    nxt = d + timedelta(days=1)
    while nxt.weekday() >= 5:
        nxt += timedelta(days=1)
    return nxt


def _ensure_weekday(d: date) -> date:
    """土日の場合は直前の平日に丸める。

    Args:
        d: 対象日付。

    Returns:
        平日の日付。
    """
    if d.weekday() >= 5:
        return _prev_weekday(d)
    return d


# ---------------------------------------------------------------------------
# ルート
# ---------------------------------------------------------------------------


@daily_bp.route("/daily/today")
def daily_today() -> Any:
    """今日の日付で日次実績ページへリダイレクトする。

    Returns:
        本日の日次実績ページへのリダイレクト、または未ログイン時はログインページへのリダイレクト。
    """
    if "user_id" not in session:
        return redirect(url_for("auth.login"))

    today = _ensure_weekday(date.today())
    # ユーザー切り替え中はuser_idを引き継ぐ
    req_user_id = request.args.get("user_id", "").strip()
    if req_user_id:
        return redirect(url_for("daily_bp.daily_view", date_str=today.isoformat()) + f"?user_id={req_user_id}")
    return redirect(url_for("daily_bp.daily_view", date_str=today.isoformat()))


@daily_bp.route("/daily/<date_str>")
def daily_view(date_str: str) -> Any:
    """指定日の日次実績を表示する（GET）。

    管理者は ?user_id=<id> クエリパラメータで他ユーザーの実績を閲覧・編集できる。

    Args:
        date_str: 表示対象の日付文字列（'YYYY-MM-DD' 形式）。

    Returns:
        daily.html のレンダリング結果、またはログインページへのリダイレクト。
    """
    if "user_id" not in session:
        return redirect(url_for("auth.login"))

    # 日付パース（不正値は今日にフォールバック）
    try:
        target_date = date.fromisoformat(date_str)
    except ValueError:
        target_date = date.today()

    # 土日は直前の平日に丸める
    target_date = _ensure_weekday(target_date)
    date_str = target_date.isoformat()

    # 対象ユーザーIDの決定（管理職・マスタは他ユーザーを閲覧可能・スコープ制限あり）
    req_user_id = request.args.get("user_id", "").strip()
    if req_user_id and is_privileged(session.get("user_role", "")):
        try:
            target_user_id: int = int(req_user_id)
        except ValueError:
            target_user_id = int(session["user_id"])
        # スコープ制限チェック
        _target_check = get_user_by_id(target_user_id)
        if _target_check is not None:
            _login_user_dict = {"id": int(session["user_id"]), "role": session.get("user_role", ""), "dept": session.get("user_dept", "")}
            if not can_access_user(_login_user_dict, _target_check):
                abort(403)
    else:
        target_user_id = int(session["user_id"])

    is_admin_view: bool = (target_user_id != int(session["user_id"]))

    # ユーザー情報取得
    login_user = get_user_by_id(int(session["user_id"]))
    target_user = get_user_by_id(target_user_id)
    if target_user is None:
        target_user = login_user
        target_user_id = int(session["user_id"])
        is_admin_view = False

    # 週開始日・曜日インデックス
    week_start: str = _get_week_start_for_date(target_date)
    day_of_week: int = target_date.weekday()  # 0=月 〜 4=金

    # 休暇区分取得
    leave_data: dict = get_weekly_leave(target_user_id, week_start)
    leave_type: str = leave_data.get(day_of_week, "")

    # 週間予定から当日分を取得
    schedule = get_weekly_schedule(target_user_id, week_start)
    day_schedule = schedule.get(day_of_week, {"am": [], "pm": []})
    schedule_am: list = day_schedule.get("am", [])
    schedule_pm: list = day_schedule.get("pm", [])

    # 日次実績・コメント・メタ情報取得
    result: dict = get_daily_result(target_user_id, date_str)
    comment: dict = get_daily_comment(target_user_id, date_str)
    result_meta: dict | None = get_daily_result_meta(target_user_id, date_str)

    # タスクマスター取得・JSON化（JS用）
    task_master: list = get_task_master(target_user_id)
    task_map_json: str = json.dumps(
        {t["task_name"]: t["default_hours"] for t in task_master},
        ensure_ascii=False,
    )

    # 曜日ラベル
    day_label: str = ["月", "火", "水", "木", "金"][day_of_week]

    # 管理職・マスタ用ユーザー一覧（切り替えドロップダウン用・スコープ制限あり）
    login_role_nav: str = session.get("user_role", "")
    login_dept_nav: str = session.get("user_dept", "")
    all_users_list: list[dict] = []
    if is_privileged(login_role_nav):
        all_users_list = get_accessible_users(int(session["user_id"]), login_role_nav, login_dept_nav)

    # 管理職・マスタが自分自身のページを見ている場合、スコープ内メンバーの振り返りコメントを取得
    # 表示対象ロール:
    #   管理職 → ユーザー のみ（他の管理職・マスタは除外）
    #   マスタ  → 管理職 + ユーザー（他のマスタは除外）
    login_role_sub: str = session.get("user_role", "")
    if is_master(login_role_sub):
        _allowed_roles: frozenset[str] = frozenset({"管理職", "ユーザー"})
    else:
        _allowed_roles = frozenset({"ユーザー"})

    subordinate_comments: list[dict] = []
    if is_privileged(login_role_sub) and not is_admin_view:
        all_users = all_users_list  # スコープ制限済みリストを再利用
        for sub in all_users:
            if sub["id"] == target_user_id:
                continue  # 自分自身はスキップ
            if sub.get("role", "") not in _allowed_roles:
                continue  # 役職フィルター
            sub_comment = get_daily_comment(sub["id"], date_str)
            subordinate_comments.append({
                "user": sub,
                "comment": sub_comment,
            })

    # 当日のイベント一覧を取得
    day_events: list[dict] = get_events_for_user_date(target_user_id, date_str)

    # 繰越タスク（保留中）を取得
    pending_carryovers: list[dict] = get_pending_carryovers(target_user_id)

    # リスケカレンダー用: 今日から13週分の休暇日を {date_str: leave_type} で収集
    leave_dates: dict[str, str] = {}
    today_d = date.today()
    for w in range(14):
        ws = _get_monday(today_d + timedelta(weeks=w)).isoformat()
        wl = get_weekly_leave(target_user_id, ws)
        for dow, ltype in wl.items():
            if ltype:
                d = date.fromisoformat(ws) + timedelta(days=int(dow))
                leave_dates[d.isoformat()] = ltype
    leave_dates_json: str = json.dumps(leave_dates, ensure_ascii=False)

    return render_template(
        "daily.html",
        user=login_user,
        target_user=target_user,
        date_str=date_str,
        day_of_week=day_of_week,
        day_label=day_label,
        leave_type=leave_type,
        schedule_am=schedule_am,
        schedule_pm=schedule_pm,
        result=result,
        comment=comment,
        result_meta=result_meta,
        task_master=task_master,
        task_map_json=task_map_json,
        prev_date=_prev_weekday(target_date).isoformat(),
        next_date=_next_weekday(target_date).isoformat(),
        today=date.today().isoformat(),
        is_admin_view=is_admin_view,
        subordinate_comments=subordinate_comments,
        all_users=all_users_list,
        pending_carryovers=pending_carryovers,
        leave_dates_json=leave_dates_json,
        day_events=day_events,
    )


@daily_bp.route("/daily/save", methods=["POST"])
def daily_save() -> Any:
    """日次実績とコメントを保存する（POST）。

    フォームから実績データとコメントを受け取り、DBへ保存したうえで
    日次実績ページへリダイレクトする。

    Returns:
        日次実績ページへのリダイレクト、またはログインページへのリダイレクト。
    """
    if "user_id" not in session:
        return redirect(url_for("auth.login"))

    raw_date: str = request.form.get("date_str", "").strip()
    # date_str を strptime で厳格にバリデートする（CLAUDE.md セクション4-2）
    try:
        datetime.strptime(raw_date, "%Y-%m-%d")
        date_str: str = raw_date
    except ValueError:
        flash("不正な日付が指定されました", "warning")
        form_uid = request.form.get("target_user_id", "").strip()
        if form_uid:
            return redirect(url_for("daily_bp.daily_today") + f"?user_id={form_uid}")
        return redirect(url_for("daily_bp.daily_today"))

    # 対象ユーザーIDの決定
    form_user_id: str = request.form.get("target_user_id", "").strip()
    if form_user_id and is_privileged(session.get("user_role", "")):
        try:
            target_user_id: int = int(form_user_id)
        except ValueError:
            target_user_id = int(session["user_id"])
    else:
        target_user_id = int(session["user_id"])

    is_admin_view: bool = (target_user_id != int(session["user_id"]))

    # フォームから実績データを解析
    data: dict[str, list[dict[str, Any]]] = {"am": [], "pm": []}
    for slot in ("am", "pm"):
        for i in range(5):
            task: str = request.form.get(f"result_task_{slot}_{i}", "").strip()
            subcategory_name: str = request.form.get(f"subcategory_{slot}_{i}", "").strip()
            # 作業名がない場合は defer_date・carryover もクリア
            raw_defer: str = request.form.get(f"defer_date_{slot}_{i}", "").strip()
            defer_date: str = raw_defer if task else ""
            is_carryover: int = 1 if (task and request.form.get(f"carryover_{slot}_{i}", "") == "1") else 0
            # project_task_id の取得（タスク管理と紐づいている場合）
            pt_id_raw: str = request.form.get(f"project_task_id_{slot}_{i}", "").strip()
            try:
                project_task_id: int | None = int(pt_id_raw) if pt_id_raw else None
            except ValueError:
                project_task_id = None
            try:
                hours: float = float(request.form.get(f"result_hours_{slot}_{i}", 0) or 0)
                if defer_date:
                    hours = 0.0  # リスケ: 時間は0固定
                else:
                    hours = max(0.0, min(hours, 24.0))
            except ValueError:
                hours = 0.0
            data[slot].append({
                "task_name": task,
                "hours": hours,
                "defer_date": defer_date,
                "is_carryover": is_carryover,
                "subcategory_name": subcategory_name,
                "project_task_id": project_task_id,
            })

    updated_by: str = session.get("user_name", "")

    # リスケ解除検出：保存前の defer_date と今回の defer_date を比較
    old_result: dict = get_daily_result(target_user_id, date_str)
    orig_date_obj = date.fromisoformat(date_str)
    reschedule_prefix: str = f"【{orig_date_obj.month:02d}/{orig_date_obj.day:02d} リスケ】"
    for slot in ("am", "pm"):
        for i in range(5):
            old_entry = old_result[slot][i]
            new_entry = data[slot][i]
            old_defer: str = old_entry.get("defer_date") or ""
            new_defer: str = new_entry.get("defer_date") or ""
            old_task: str = old_entry.get("task_name") or ""
            if old_defer and not new_defer and old_task:
                # defer_date がクリアされた → リスケ解除
                defer_task_name: str = f"{reschedule_prefix}{old_task}"
                remove_rescheduled_task(target_user_id, old_defer, defer_task_name)
                remove_rescheduled_daily_result(target_user_id, old_defer, defer_task_name)

    # 実績保存
    save_daily_result(target_user_id, date_str, data, updated_by)

    # 後日対応：defer_date が指定されたタスクを週間予定に追加
    for slot in ("am", "pm"):
        for i in range(5):
            defer_date_raw: str = request.form.get(f"defer_date_{slot}_{i}", "").strip()
            if not defer_date_raw:
                continue
            try:
                datetime.strptime(defer_date_raw, "%Y-%m-%d")
            except ValueError:
                continue
            task_name: str = request.form.get(f"result_task_{slot}_{i}", "").strip()
            if not task_name:
                continue
            # 元日付を mm/dd 形式で付加（例: 【03/20 リスケ】作業名）
            orig_date = date.fromisoformat(date_str)
            reschedule_prefix = f"【{orig_date.month:02d}/{orig_date.day:02d} リスケ】"
            defer_task_name: str = f"{reschedule_prefix}{task_name}"
            try:
                defer_hours: float = float(
                    request.form.get(f"result_hours_{slot}_{i}", 0) or 0
                )
                defer_hours = max(0.0, min(defer_hours, 24.0))
            except ValueError:
                defer_hours = 0.0
            defer_task_to_weekly_schedule(
                target_user_id, defer_date_raw, defer_task_name, defer_hours, updated_by
            )

    # is_carryover フラグに基づいて carryover テーブルを更新
    # 全スロットを走査して「どのスロットでも⏩ONのタスク」を収集
    carryover_on_tasks: set[str] = set()
    all_result_tasks: set[str] = set()
    for slot in ("am", "pm"):
        for entry in data[slot]:
            task_nm: str = entry.get("task_name") or ""
            if not task_nm:
                continue
            all_result_tasks.add(task_nm)
            if entry.get("is_carryover"):
                co_hours: float = float(entry.get("hours") or 0)
                add_carryover(target_user_id, date_str, task_nm, co_hours)
                carryover_on_tasks.add(task_nm)

    # 全スロットで⏩OFFかつリスケでもないタスクの繰越を解決
    for task_nm in all_result_tasks - carryover_on_tasks:
        resolve_carryovers_by_task(target_user_id, task_nm, date_str)

    # タスク管理の進捗連動（project_task_id が紐づいている実績の工数を反映）
    sync_daily_progress_to_task(target_user_id, date_str)

    # コメント保存
    reflection: str = request.form.get("reflection", "").strip()
    action: str = request.form.get("action", "").strip()
    save_daily_comment(target_user_id, date_str, reflection, action, updated_by)

    flash("日次実績を保存しました", "success")

    # リダイレクト先の決定
    if is_admin_view:
        return redirect(
            url_for("daily_bp.daily_view", date_str=date_str) + f"?user_id={target_user_id}"
        )
    return redirect(url_for("daily_bp.daily_view", date_str=date_str))


@daily_bp.route("/daily/resolve_carryover/<int:carryover_id>", methods=["POST"])
def resolve_carryover(carryover_id: int) -> Any:
    """繰越タスクを手動で解決済みにする（POST）。

    Args:
        carryover_id: 解決する繰越レコードのID。

    Returns:
        日次実績ページへのリダイレクト。
    """
    if "user_id" not in session:
        return redirect(url_for("auth.login"))

    date_str: str = request.form.get("date_str", date.today().isoformat())
    form_user_id: str = request.form.get("target_user_id", "").strip()
    target_uid: int = int(form_user_id) if form_user_id else int(session["user_id"])
    from ..models import resolve_carryover_by_id
    resolve_carryover_by_id(target_uid, carryover_id)
    flash("繰り越しを解決済みにしました", "success")
    if target_uid != int(session["user_id"]):
        return redirect(url_for("daily_bp.daily_view", date_str=date_str) + f"?user_id={target_uid}")
    return redirect(url_for("daily_bp.daily_view", date_str=date_str))


@daily_bp.route("/daily/save_admin_comment", methods=["POST"])
def daily_save_admin_comment() -> Any:
    """管理者（上長）コメントを保存する（POST）。

    管理者のみ実行可能。対象ユーザーの実績画面に表示されるコメントを保存する。

    Returns:
        日次実績ページへのリダイレクト。
    """
    if "user_id" not in session:
        return redirect(url_for("auth.login"))
    if not is_privileged(session.get("user_role", "")):
        abort(403)

    raw_date: str = request.form.get("date_str", "").strip()
    try:
        datetime.strptime(raw_date, "%Y-%m-%d")
        date_str: str = raw_date
    except ValueError:
        flash("不正な日付が指定されました", "warning")
        return redirect(url_for("daily_bp.daily_today"))

    form_user_id: str = request.form.get("target_user_id", "").strip()
    try:
        target_user_id: int = int(form_user_id)
    except ValueError:
        target_user_id = int(session["user_id"])

    # 自身へのコメントは禁止
    if target_user_id == int(session["user_id"]):
        if request.form.get("ajax") == "1":
            return {"ok": False, "error": "自己コメントは入力できません"}, 400
        flash("自身への上長コメントは入力できません", "warning")
        return redirect(url_for("daily_bp.daily_view", date_str=date_str))

    # 管理職は同一部署のみ（マスタは全員OK）
    if not is_master(session.get("user_role", "")):
        target_check = get_user_by_id(target_user_id)
        if target_check is None:
            abort(404)
        login_user_dict = {"id": int(session["user_id"]), "role": session.get("user_role", ""), "dept": session.get("user_dept", "")}
        if not can_access_user(login_user_dict, dict(target_check)):
            abort(403)

    admin_comment: str = request.form.get("admin_comment", "").strip()
    updated_by: str = session.get("user_name", "")

    save_admin_comment(target_user_id, date_str, admin_comment, updated_by)

    # AJAXリクエストの場合はJSONで返す（フルページリロード不要）
    if request.form.get("ajax") == "1":
        from datetime import datetime as _dt
        updated_at_display = _dt.now().strftime("%Y-%m-%d %H:%M")
        return {"ok": True, "updated_at": updated_at_display, "updated_by": updated_by}, 200

    flash("管理者コメントを保存しました", "success")

    # 管理者自身のページから保存した場合は自分のページへ戻る
    from_own_page: str = request.form.get("from_own_page", "")
    if from_own_page == "1":
        return redirect(url_for("daily_bp.daily_view", date_str=date_str))
    return redirect(
        url_for("daily_bp.daily_view", date_str=date_str) + f"?user_id={target_user_id}"
    )

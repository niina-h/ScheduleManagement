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
    get_active_tasks_for_user,
    get_all_project_tasks,
    get_all_users,
    get_comments_in_range,
    get_daily_comment,
    get_daily_plan_vs_actual_for_users,
    get_daily_result,
    get_daily_result_meta,
    get_events_for_user_date,
    get_pending_carryovers,
    get_task_master,
    get_task_plan_vs_actual,
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
from ..auth_helpers import is_privileged, is_master, is_manager, can_access_user, normalize_role

daily_bp = Blueprint("daily_bp", __name__)


# ---------------------------------------------------------------------------
# ヘルパー関数
# ---------------------------------------------------------------------------


def _build_project_tasks_json(user_id: int) -> str:
    """該当ユーザーに割り当てられたプロジェクトタスクをJSONに変換する。

    実績画面でのタスク自動マッチング用。タスク名 → ID のマッピング。

    Args:
        user_id: 対象ユーザーID

    Returns:
        str: [{id, name, status, progress}] のJSON文字列
    """
    import json
    tasks = get_all_project_tasks(assigned_to=user_id)
    result = [
        {
            "id": t["id"],
            "name": t["task_name"],
            "status": t["status"],
            "progress": t.get("progress", 0),
        }
        for t in tasks
        if t["status"] not in ("完了", "中断")
    ]
    return json.dumps(result, ensure_ascii=False)


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
    if is_privileged(session.get("user_role", "")):
        if req_user_id:
            try:
                target_user_id: int = int(req_user_id)
            except ValueError:
                target_user_id = int(session["user_id"])
        elif session.get("selected_user_id"):
            target_user_id = int(session["selected_user_id"])
        else:
            target_user_id = int(session["user_id"])
        # スコープ制限チェック
        _target_check = get_user_by_id(target_user_id)
        if _target_check is not None:
            _login_user_dict = {"id": int(session["user_id"]), "role": session.get("user_role", ""), "dept": session.get("user_dept", "")}
            if not can_access_user(_login_user_dict, _target_check):
                target_user_id = int(session["user_id"])
                session.pop("selected_user_id", None)
            else:
                session["selected_user_id"] = target_user_id
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
    # ドロップダウンの選択肢はログイン中の本人の権限範囲（誰の画面に切り替えられるか）。
    # システム管理者が所属切替していない場合、get_accessible_users は全所属を返すが、
    # 部署未設定のテスト用アカウント等が混入するのを避けるため、自分自身の所属で絞り込む。
    login_role_nav: str = session.get("user_role", "")
    login_dept_nav: str = session.get("user_dept", "")
    all_users_list: list[dict] = []
    if is_privileged(login_role_nav):
        all_users_list = get_accessible_users(int(session["user_id"]), login_role_nav, login_dept_nav)
        if login_dept_nav:
            all_users_list = [u for u in all_users_list if (u.get("dept") or "") == login_dept_nav]

    # 部下の振り返りコメント一覧: 「切り替え先（target_user）」が本人ログインした場合と
    # 全く同じ画面になるよう、target_user 自身の権限・所属を基準に計算する
    # （ログイン中の本人ではなく、成り代わっている相手の視点で部下一覧を出す）。
    # 表示対象ロール:
    #   管理職        → 一般 のみ（他の管理職・所属長・システム管理者は除外）
    #   所属長/システム管理者 → 管理職 + 一般（他の所属長・システム管理者は除外）
    target_role: str = target_user.get("role", "") if target_user else ""
    target_dept: str = target_user.get("dept", "") if target_user else ""
    if is_manager(target_role):
        _allowed_roles: frozenset[str] = frozenset({"一般"})
    else:
        _allowed_roles = frozenset({"管理職", "一般"})

    subordinate_comments: list[dict] = []
    if is_privileged(target_role):
        target_accessible = get_accessible_users(target_user_id, target_role, target_dept)
        if target_dept:
            target_accessible = [u for u in target_accessible if (u.get("dept") or "") == target_dept]
        for sub in target_accessible:
            if sub["id"] == target_user_id:
                continue  # 本人はスキップ
            if normalize_role(sub.get("role", "")) not in _allowed_roles:
                continue  # 役職フィルター
            sub_comment = get_daily_comment(sub["id"], date_str)
            subordinate_comments.append({
                "user": sub,
                "comment": sub_comment,
            })

    # ログイン中の本人が「表示中の人物（target_user）」に上長コメントできるか。
    # 自分自身へのコメントは不可。切り替え中でも常にこの判定（本人の権限）で行う。
    can_comment_target: bool = False
    if target_user_id != int(session["user_id"]) and is_privileged(session.get("user_role", "")):
        if is_master(session.get("user_role", "")):
            can_comment_target = True
        else:
            _login_user_dict2 = {
                "id": int(session["user_id"]), "role": session.get("user_role", ""),
                "dept": session.get("user_dept", ""),
            }
            can_comment_target = can_access_user(_login_user_dict2, dict(target_user)) if target_user else False

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
        can_comment_target=can_comment_target,
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
        project_tasks_json=_build_project_tasks_json(target_user_id),
        active_project_tasks=get_active_tasks_for_user(target_user_id),
        # 進捗（予定 vs 実績）は「本日時点」の指標のため、当日を表示している時のみ算出する。
        task_progress_json=json.dumps(
            get_task_plan_vs_actual(target_user_id) if target_date == date.today() else {},
            ensure_ascii=False,
        ),
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

    # project_task_id の妥当性検証用（タスク名が変わったのに古い紐付けが送信された
    # 場合に備え、id→task_name のマップで一致を確認する。フロントの change イベント
    # に依存しない、サーバー側での最終防御）。
    pt_name_by_id: dict[int, str] = {
        t["id"]: t["task_name"] for t in get_all_project_tasks()
    }

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
            # project_task_id の取得（タスク管理と紐づいている場合）。
            # 紐付き先のタスク名が現在の入力値と一致しない場合は、古い紐付けと
            # 判断して無視する（フロントのJS修正が効かなかった場合の保険）。
            pt_id_raw: str = request.form.get(f"project_task_id_{slot}_{i}", "").strip()
            try:
                project_task_id: int | None = int(pt_id_raw) if pt_id_raw else None
            except ValueError:
                project_task_id = None
            if project_task_id is not None and pt_name_by_id.get(project_task_id) != task:
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


@daily_bp.route("/daily/comments-review")
def comments_review() -> Any:
    """配下メンバーの振り返りコメントを期間横断で一覧する（管理職・所属長・システム管理者）。

    今日1日分しか見られなかった上長コメント確認を、週／任意期間で俯瞰できるようにする。
    上長コメント未入力の行を強調し、フィードバック漏れを防ぐ。

    クエリ: ?start=YYYY-MM-DD&end=YYYY-MM-DD（省略時は今週の月〜日）

    Returns:
        Any: コメントレビュー画面のHTML、権限がなければ403。
    """
    login_role = session.get("user_role", "")
    if not is_privileged(login_role):
        abort(403)

    login_id = int(session["user_id"])
    login_dept = session.get("user_dept", "")

    # 期間（既定＝今週 月〜日）
    today = date.today()
    default_start = today - timedelta(days=today.weekday())
    default_end = default_start + timedelta(days=6)

    def _parse(qs: str, fallback: date) -> date:
        try:
            return date.fromisoformat(qs) if qs else fallback
        except ValueError:
            return fallback

    start_d = _parse(request.args.get("start", ""), default_start)
    end_d = _parse(request.args.get("end", ""), default_end)
    if end_d < start_d:
        start_d, end_d = end_d, start_d

    # 閲覧可能な配下ユーザー（システム管理者の所属切替・所属長の担当所属に連動）
    members = get_accessible_users(login_id, login_role, login_dept)
    member_map = {u["id"]: u for u in members}
    member_ids = list(member_map.keys())

    # 期間内のコメントを取得し、ユーザー名を添える
    raw = get_comments_in_range(member_ids, start_d.isoformat(), end_d.isoformat())
    comments: list[dict] = []
    for c in raw:
        u = member_map.get(c["user_id"], {})
        comments.append({
            **c,
            "name": u.get("name", "?"),
            "dept": u.get("dept", ""),
            "needs_admin_comment": not (c.get("admin_comment") or "").strip(),
        })

    pending_count = sum(1 for c in comments if c["needs_admin_comment"])

    return render_template(
        "comments_review.html",
        comments=comments,
        start_date=start_d.isoformat(),
        end_date=end_d.isoformat(),
        prev_week=(start_d - timedelta(days=7)).isoformat(),
        next_week=(start_d + timedelta(days=7)).isoformat(),
        this_week=default_start.isoformat(),
        pending_count=pending_count,
        total_count=len(comments),
    )


@daily_bp.route("/daily/team-progress")
def team_progress() -> Any:
    """配下メンバー全員の当日の予定 vs 実績の差分を一覧する（管理職・所属長・システム管理者）。

    振り返り機能：本人の日次画面（daily.html）で1人ずつ見ていた予定/実績差分を、
    所属長・管理職が部下全員分まとめて俯瞰できるようにする。

    クエリ: ?date=YYYY-MM-DD（省略時は当日）

    Returns:
        Any: 部下状況一覧画面のHTML、権限がなければ403。
    """
    login_role = session.get("user_role", "")
    if not is_privileged(login_role):
        abort(403)

    login_id = int(session["user_id"])
    login_dept = session.get("user_dept", "")

    raw_date = request.args.get("date", "").strip()
    try:
        target_date = date.fromisoformat(raw_date) if raw_date else date.today()
    except ValueError:
        target_date = date.today()
    date_str = target_date.isoformat()

    # 閲覧可能な配下ユーザー（システム管理者の所属切替・所属長の担当所属に連動）
    members = get_accessible_users(login_id, login_role, login_dept)
    member_map = {u["id"]: u for u in members}
    member_ids = list(member_map.keys())

    diffs = get_daily_plan_vs_actual_for_users(member_ids, date_str)
    members_view: list[dict] = []
    for d in diffs:
        u = member_map.get(d["user_id"], {})
        members_view.append({
            **d,
            "name": u.get("name", "?"),
            "dept": u.get("dept", ""),
        })
    # 遅れが大きい順に並べて、要注意のメンバーを目立たせる
    members_view.sort(key=lambda m: m["diff_hours"])

    behind_count = sum(1 for m in members_view if m["state"] == "behind")

    return render_template(
        "team_progress.html",
        members=members_view,
        date_str=date_str,
        prev_date=(target_date - timedelta(days=1)).isoformat(),
        next_date=(target_date + timedelta(days=1)).isoformat(),
        today=date.today().isoformat(),
        behind_count=behind_count,
        total_count=len(members_view),
    )

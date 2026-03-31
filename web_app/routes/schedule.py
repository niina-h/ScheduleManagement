"""
週間予定管理ルート。

/schedule       : 週間予定の表示（GET）
/schedule/save  : 週間予定の保存（POST）
/schedule/copy_last_week : 先週の予定をコピー（POST）
"""

from __future__ import annotations

import json
from datetime import date, timedelta
from typing import Any

from flask import (
    Blueprint,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

from ..models import (
    copy_last_week_schedule,
    get_accessible_users,
    get_active_tasks_for_user,
    get_all_categories,
    get_all_subcategories,
    get_all_users,
    get_daily_comment,
    get_events_for_user_date,
    get_task_master,
    get_user_by_id,
    get_week_daily_results,
    get_weekly_leave,
    get_weekly_schedule,
    get_weekly_schedule_meta,
    apply_routine_to_week,
    get_routine_schedules,
    import_events_to_weekly_schedule,
    import_tasks_to_weekly_schedule,
    save_weekly_leave,
    save_weekly_schedule,
)
from ..log_service import record_operation, ACTION_SCHEDULE_SAVE
from ..auth_helpers import is_privileged, is_master, can_access_user

schedule_bp = Blueprint("schedule", __name__)


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


def _get_default_week_start() -> str:
    """デフォルトの週開始日（常に今週月曜日）を返す。

    Returns:
        今週月曜日の 'YYYY-MM-DD' 文字列。
    """
    return _get_monday(date.today()).isoformat()


def _week_dates(week_start: str) -> list[str]:
    """week_start から月〜金の日付文字列リストを返す。

    Args:
        week_start: 週開始日（月曜日）の 'YYYY-MM-DD' 文字列。

    Returns:
        月曜〜金曜の 'YYYY-MM-DD' 文字列を格納した長さ5のリスト。
    """
    start = date.fromisoformat(week_start)
    return [(start + timedelta(days=i)).isoformat() for i in range(5)]


def _parse_schedule_form(form: Any) -> dict[int, dict[str, list[dict[str, Any]]]]:
    """フォームデータを save_weekly_schedule 用の dict 形式に変換する。

    フォームキーの命名規則:
        task_{day}_{slot}_{i}         : タスク名（day=0-4, slot='am'/'pm', i=0-4）
        hours_{day}_{slot}_{i}        : 作業時間（同上）
        subcategory_{day}_{slot}_{i}  : 中区分名（同上）

    Args:
        form: Flask の request.form オブジェクト（ImmutableMultiDict 相当）。

    Returns:
        {day: {'am': [{task_name, hours, subcategory_name} × 5], 'pm': [同]}} 形式の辞書。
    """
    data: dict[int, dict[str, list[dict[str, Any]]]] = {}
    for day in range(5):
        data[day] = {"am": [], "pm": []}
        for slot in ["am", "pm"]:
            for i in range(5):
                task = form.get(f"task_{day}_{slot}_{i}", "").strip()
                subcategory_name = form.get(f"subcategory_{day}_{slot}_{i}", "").strip()
                try:
                    hours = float(form.get(f"hours_{day}_{slot}_{i}", 0) or 0)
                except ValueError:
                    hours = 0.0
                data[day][slot].append({
                    "task_name": task,
                    "hours": hours,
                    "subcategory_name": subcategory_name,
                })
    return data


# ---------------------------------------------------------------------------
# ルート
# ---------------------------------------------------------------------------


@schedule_bp.route("/schedule")
def weekly() -> Any:
    """週間予定ページを表示する（GET）。

    クエリパラメータ `week` で表示週を指定できる。未指定時はデフォルト週
    （木曜以降なら翌週月曜、それ以前なら今週月曜）を使用する。
    管理者は `user_id` クエリパラメータで他ユーザーの週間予定を閲覧・編集できる。

    Returns:
        schedule.html のレンダリング結果、またはログインページへのリダイレクト。
    """
    login_user_id: int | None = session.get("user_id")
    if not login_user_id:
        return redirect(url_for("auth.login"))

    # 管理職・マスタは他ユーザーの予定を表示できる（スコープ制限あり）
    req_user_id: str = request.args.get("user_id", "").strip()
    if is_privileged(session.get("user_role", "")):
        if req_user_id:
            # URLに user_id が指定された場合はセッションに保存
            try:
                target_user_id: int = int(req_user_id)
            except ValueError:
                target_user_id = int(login_user_id)
        elif session.get("selected_user_id"):
            # URLに指定がなければセッションの選択ユーザーを維持
            target_user_id = int(session["selected_user_id"])
        else:
            target_user_id = int(login_user_id)
        # スコープ制限チェック
        target = get_user_by_id(target_user_id)
        if target is None:
            from flask import abort
            abort(404)
        login_user_dict = {"id": int(login_user_id), "role": session.get("user_role", ""), "dept": session.get("user_dept", "")}
        if not can_access_user(login_user_dict, target):
            target_user_id = int(login_user_id)
            session.pop("selected_user_id", None)
        else:
            session["selected_user_id"] = target_user_id
    else:
        target_user_id = int(login_user_id)

    is_admin_view: bool = (target_user_id != int(login_user_id))

    week_start: str = request.args.get("week", "") or _get_default_week_start()

    # week_start が月曜日か検証（不正な入力は今週月曜にフォールバック）
    try:
        ws_date = date.fromisoformat(week_start)
        if ws_date.weekday() != 0:
            ws_date = _get_monday(ws_date)
            week_start = ws_date.isoformat()
    except ValueError:
        week_start = _get_default_week_start()
        ws_date = date.fromisoformat(week_start)

    schedule: dict = get_weekly_schedule(target_user_id, week_start)
    task_master: list[dict] = get_task_master(target_user_id)
    user: dict | None = get_user_by_id(target_user_id)
    login_user: dict | None = get_user_by_id(int(login_user_id))

    # 休暇設定・更新者情報取得
    leave_data: dict = get_weekly_leave(target_user_id, week_start)
    schedule_meta: dict | None = get_weekly_schedule_meta(target_user_id, week_start)

    # 管理職・マスタ用: ユーザーリスト（ユーザー切り替えセレクト用・スコープ制限あり）
    login_role: str = session.get("user_role", "")
    login_dept: str = session.get("user_dept", "")
    if is_privileged(login_role):
        all_users: list[dict] = get_accessible_users(int(login_user_id), login_role, login_dept)
    else:
        all_users = []

    prev_week: str = (ws_date - timedelta(weeks=1)).isoformat()
    next_week: str = (ws_date + timedelta(weeks=1)).isoformat()
    dates: list[str] = _week_dates(week_start)

    # タスク名 → default_hours のマップをJSON文字列に変換する（JS用）
    task_map_json: str = json.dumps(
        {t["task_name"]: t["default_hours"] for t in task_master},
        ensure_ascii=False,
    )

    # 当週にデータが保存されているか確認する
    has_data: bool = any(
        entry["task_name"] or entry["hours"] > 0
        for day_slots in schedule.values()
        for slot_entries in day_slots.values()
        for entry in slot_entries
    )

    # 過去日の実績データを取得（日付ヘッダーに「実績」バッジ、セルに実績値を表示するため）
    week_daily_results: dict = get_week_daily_results(target_user_id, dates)

    # 上長コメントデータを取得（週間予定ヘッダーの「上長」バッジ表示用）
    week_admin_comments: dict[int, dict] = {}
    for i, d in enumerate(dates):
        c = get_daily_comment(target_user_id, d)
        week_admin_comments[i] = {
            "has_admin_comment": bool(c.get("admin_comment")),
            "reflection": c.get("reflection", ""),
            "action": c.get("action", ""),
            "admin_comment": c.get("admin_comment", ""),
            "updated_by": c.get("updated_by", ""),
        }

    categories = get_all_categories()
    all_subcategories = get_all_subcategories()

    # タスク管理からインポート可能なタスク一覧
    active_project_tasks: list[dict] = get_active_tasks_for_user(target_user_id)

    # 各曜日のイベント一覧
    week_events: dict[int, list[dict]] = {}
    for i, d in enumerate(dates):
        week_events[i] = get_events_for_user_date(target_user_id, d)

    return render_template(
        "schedule.html",
        user=user,
        login_user=login_user,
        schedule=schedule,
        task_master=task_master,
        week_start=week_start,
        week_dates=dates,
        prev_week=prev_week,
        next_week=next_week,
        task_map_json=task_map_json,
        has_data=has_data,
        today=date.today().isoformat(),
        leave_data=leave_data,
        schedule_meta=schedule_meta,
        is_admin_view=is_admin_view,
        target_user_id=target_user_id,
        all_users=all_users,
        week_daily_results=week_daily_results,
        week_admin_comments=week_admin_comments,
        categories=categories,
        all_subcategories=all_subcategories,
        active_project_tasks=active_project_tasks,
        week_events=week_events,
    )


@schedule_bp.route("/schedule/save", methods=["POST"])
def save() -> Any:
    """週間予定を保存する（POST）。

    フォームから各日・各スロット・各枠のタスク名と時間を受け取り、
    DBへ保存したうえで週間予定ページへリダイレクトする。

    Returns:
        週間予定ページへのリダイレクト、またはログインページへのリダイレクト。
    """
    login_user_id: int | None = session.get("user_id")
    if not login_user_id:
        return redirect(url_for("auth.login"))

    week_start: str = request.form.get("week_start", "").strip()
    if not week_start:
        week_start = _get_default_week_start()

    # 管理職・マスタによる他ユーザー編集対応
    form_user_id: str = request.form.get("target_user_id", "").strip()
    if form_user_id and is_privileged(session.get("user_role", "")):
        try:
            target_user_id: int = int(form_user_id)
        except ValueError:
            target_user_id = int(login_user_id)
    else:
        target_user_id = int(login_user_id)

    is_admin_view: bool = (target_user_id != int(login_user_id))
    updated_by: str = session.get("user_name", "")

    # 過去週への書き込みをサーバー側で禁止（管理職・マスタを除く）
    try:
        week_start_date = date.fromisoformat(week_start)
    except ValueError:
        flash("不正な週指定です", "warning")
        return redirect(url_for("schedule.weekly"))
    current_monday = date.today() - timedelta(days=date.today().weekday())
    if week_start_date < current_monday and not is_privileged(session.get("user_role", "")):
        flash("過去の週は編集できません", "warning")
        return redirect(url_for("schedule.weekly") + f"?week={week_start}")

    # 週間予定保存
    data = _parse_schedule_form(request.form)
    save_weekly_schedule(target_user_id, week_start, data, updated_by)

    # 休暇種別保存
    leave_data: dict[int, str] = {}
    for day in range(5):
        leave_type = request.form.get(f"leave_{day}", "").strip()
        leave_data[day] = leave_type
    save_weekly_leave(target_user_id, week_start, leave_data)

    record_operation(ACTION_SCHEDULE_SAVE, f"week={week_start} user_id={target_user_id}")
    flash("週間予定を保存しました", "success")
    redirect_url = url_for("schedule.weekly") + f"?week={week_start}"
    if is_admin_view:
        redirect_url += f"&user_id={target_user_id}"
    return redirect(redirect_url)


@schedule_bp.route("/schedule/copy_last_week", methods=["POST"])
def copy_last_week() -> Any:
    """先週の週間予定をコピーする（POST）。

    フォームから week_start を受け取り、1週前のデータをコピーする。
    コピー成功時は success フラッシュ、対象データなし時は warning フラッシュを表示する。

    Returns:
        週間予定ページへのリダイレクト、またはログインページへのリダイレクト。
    """
    login_user_id: int | None = session.get("user_id")
    if not login_user_id:
        return redirect(url_for("auth.login"))

    week_start: str = request.form.get("week_start", "").strip()
    if not week_start:
        week_start = _get_default_week_start()

    # 管理職・マスタによる他ユーザー編集対応
    form_user_id: str = request.form.get("target_user_id", "").strip()
    if form_user_id and is_privileged(session.get("user_role", "")):
        try:
            target_user_id: int = int(form_user_id)
        except ValueError:
            target_user_id = int(login_user_id)
    else:
        target_user_id = int(login_user_id)

    is_admin_view: bool = (target_user_id != int(login_user_id))

    success: bool = copy_last_week_schedule(target_user_id, week_start)
    if success:
        flash("先週の予定をコピーしました", "success")
    else:
        flash("コピー元の予定が見つかりませんでした", "warning")

    redirect_url = url_for("schedule.weekly") + f"?week={week_start}"
    if is_admin_view:
        redirect_url += f"&user_id={target_user_id}"
    return redirect(redirect_url)


@schedule_bp.route("/schedule/clear", methods=["POST"])
def clear_schedule() -> Any:
    """指定週の週間予定を全クリアする（POST）。

    フォームから week_start を受け取り、全スロットを空で上書き保存する。

    Returns:
        週間予定ページへのリダイレクト、またはログインページへのリダイレクト。
    """
    login_user_id: int | None = session.get("user_id")
    if not login_user_id:
        return redirect(url_for("auth.login"))

    week_start: str = request.form.get("week_start", "").strip()
    if not week_start:
        week_start = _get_default_week_start()

    # 管理職・マスタによる他ユーザー編集対応
    form_user_id: str = request.form.get("target_user_id", "").strip()
    if form_user_id and is_privileged(session.get("user_role", "")):
        try:
            target_user_id: int = int(form_user_id)
        except ValueError:
            target_user_id = int(login_user_id)
    else:
        target_user_id = int(login_user_id)

    is_admin_view: bool = (target_user_id != int(login_user_id))
    updated_by: str = session.get("user_name", "")

    empty_data: dict = {
        day: {
            slot: [{"task_name": "", "hours": 0.0} for _ in range(5)]
            for slot in ["am", "pm"]
        }
        for day in range(5)
    }
    save_weekly_schedule(target_user_id, week_start, empty_data, updated_by)
    flash("週間予定をクリアしました", "success")
    redirect_url = url_for("schedule.weekly") + f"?week={week_start}"
    if is_admin_view:
        redirect_url += f"&user_id={target_user_id}"
    return redirect(redirect_url)


@schedule_bp.route("/schedule/import_tasks", methods=["POST"])
def import_tasks() -> Any:
    """タスク管理のタスクを週間予定にインポートする（POST）。

    フォームから選択されたタスクIDを受け取り、AMスロットの空き枠に配置する。

    Returns:
        週間予定ページへのリダイレクト。
    """
    login_user_id: int | None = session.get("user_id")
    if not login_user_id:
        return redirect(url_for("auth.login"))

    week_start: str = request.form.get("week_start", "").strip()
    if not week_start:
        week_start = _get_default_week_start()

    # 対象ユーザー判定
    form_user_id: str = request.form.get("target_user_id", "").strip()
    if form_user_id and is_privileged(session.get("user_role", "")):
        try:
            target_user_id: int = int(form_user_id)
        except ValueError:
            target_user_id = int(login_user_id)
    else:
        target_user_id = int(login_user_id)

    is_admin_view: bool = (target_user_id != int(login_user_id))
    updated_by: str = session.get("user_name", "")

    # 選択されたタスクID
    task_ids_raw: list[str] = request.form.getlist("import_task_ids")
    task_ids: list[int] = []
    for raw in task_ids_raw:
        try:
            task_ids.append(int(raw))
        except ValueError:
            continue

    if task_ids:
        count = import_tasks_to_weekly_schedule(
            target_user_id, week_start, task_ids, updated_by,
        )
        flash(f"{count}件のタスクをインポートしました", "success")
    else:
        flash("インポートするタスクが選択されていません", "warning")

    redirect_url = url_for("schedule.weekly") + f"?week={week_start}"
    if is_admin_view:
        redirect_url += f"&user_id={target_user_id}"
    return redirect(redirect_url)


@schedule_bp.route("/schedule/import_events", methods=["POST"])
def import_events() -> Any:
    """イベントを週間予定に自動配置する（POST）。

    タスク管理のイベント（is_event=1）を、開始時刻のAM/PM判定に基づき配置する。

    Returns:
        週間予定ページへのリダイレクト。
    """
    login_user_id: int | None = session.get("user_id")
    if not login_user_id:
        return redirect(url_for("auth.login"))

    week_start: str = request.form.get("week_start", "").strip()
    if not week_start:
        week_start = _get_default_week_start()

    # 対象ユーザー判定
    form_user_id: str = request.form.get("target_user_id", "").strip()
    if form_user_id and is_privileged(session.get("user_role", "")):
        try:
            target_user_id: int = int(form_user_id)
        except ValueError:
            target_user_id = int(login_user_id)
    else:
        target_user_id = int(login_user_id)

    is_admin_view: bool = (target_user_id != int(login_user_id))
    updated_by: str = session.get("user_name", "")

    count = import_events_to_weekly_schedule(
        target_user_id, week_start, updated_by,
    )
    if count > 0:
        flash(f"{count}件のイベントを配置しました", "success")
    else:
        flash("配置するイベントがありません", "info")

    redirect_url = url_for("schedule.weekly") + f"?week={week_start}"
    if is_admin_view:
        redirect_url += f"&user_id={target_user_id}"
    return redirect(redirect_url)


@schedule_bp.route("/schedule/import_tasks_and_events", methods=["POST"])
def import_tasks_and_events() -> Any:
    """タスクとイベントを同時に週間予定へ取り込む（タスク管理反映）。

    選択されたタスクをAMスロットに配置し、さらにイベントも自動配置する。

    Returns:
        週間予定ページへのリダイレクト。
    """
    login_user_id: int | None = session.get("user_id")
    if not login_user_id:
        return redirect(url_for("auth.login"))

    week_start: str = request.form.get("week_start", "").strip()
    if not week_start:
        week_start = _get_default_week_start()

    # 対象ユーザー判定
    form_user_id: str = request.form.get("target_user_id", "").strip()
    if form_user_id and is_privileged(session.get("user_role", "")):
        try:
            target_user_id: int = int(form_user_id)
        except ValueError:
            target_user_id = int(login_user_id)
    else:
        target_user_id = int(login_user_id)

    is_admin_view: bool = (target_user_id != int(login_user_id))
    updated_by: str = session.get("user_name", "")

    # タスク取込（担当者が一致し週と期間が重なるタスクを自動検索）
    task_count = import_tasks_to_weekly_schedule(
        target_user_id, week_start, updated_by,
    )

    # イベント取込
    event_count = import_events_to_weekly_schedule(
        target_user_id, week_start, updated_by,
    )

    # 定例スケジュールを空き行に適用（反映後に定例行を補完）
    routine_count = len(get_routine_schedules(target_user_id))
    apply_routine_to_week(target_user_id, week_start, updated_by)

    msgs = []
    if task_count > 0:
        msgs.append(f"タスク{task_count}件")
    if event_count > 0:
        msgs.append(f"イベント{event_count}件")
    if routine_count > 0:
        msgs.append(f"定例{routine_count}件")

    if msgs:
        flash("・".join(msgs) + "を取り込みました", "success")
    else:
        flash("取り込む対象がありませんでした", "info")

    redirect_url = url_for("schedule.weekly") + f"?week={week_start}"
    if is_admin_view:
        redirect_url += f"&user_id={target_user_id}"
    return redirect(redirect_url)

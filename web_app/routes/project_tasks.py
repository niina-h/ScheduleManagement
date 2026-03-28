"""プロジェクトタスク管理ルート。

管理職・マスタは全タスクを閲覧・編集できる。
一般ユーザーは自分に割り当てられたタスクのみ参照可能（編集不可）。
"""
from __future__ import annotations

import json

from flask import (
    Blueprint, abort, flash, jsonify, redirect, render_template,
    request, session, url_for,
)

from ..auth_helpers import is_privileged
from ..models import (
    PROJECT_TASK_STATUSES,
    add_project_task,
    delete_project_task,
    get_accessible_users_for_dashboard,
    get_all_categories,
    get_all_project_tasks,
    get_all_subcategories,
    get_all_users,
    get_project_task_by_id,
    get_task_progress_summary,
    update_project_task,
)

project_tasks_bp = Blueprint(
    "project_tasks_bp", __name__, url_prefix="/project-tasks",
)


@project_tasks_bp.before_request
def _check_login() -> object | None:
    """未ログインならログイン画面へリダイレクトする。"""
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    return None


@project_tasks_bp.route("/")
def task_list() -> str:
    """プロジェクトタスク一覧画面を表示する。

    管理職・マスタ: 全タスク表示（編集可）
    一般ユーザー: 自分のタスクのみ表示（参照のみ）

    Returns:
        str: レンダリング済みHTML
    """
    login_role = session.get("user_role", "")
    login_id = int(session["user_id"])
    privileged = is_privileged(login_role)

    if privileged:
        tasks = get_all_project_tasks()
    else:
        tasks = get_all_project_tasks(assigned_to=login_id)

    categories = get_all_categories()
    subcategories = get_all_subcategories()

    # 担当者選択用のユーザーリスト（管理職・マスタのみ使用）
    users = get_all_users(dept_filter=session.get("user_dept")) if privileged else []

    return render_template(
        "project_tasks.html",
        tasks=tasks,
        categories=categories,
        subcategories=subcategories,
        statuses=PROJECT_TASK_STATUSES,
        privileged=privileged,
        users=users,
    )


@project_tasks_bp.route("/add", methods=["POST"])
def add_task() -> object:
    """プロジェクトタスクを追加する（管理職・マスタのみ）。

    Returns:
        object: 一覧画面へのリダイレクト
    """
    if not is_privileged(session.get("user_role", "")):
        abort(403)
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    cat_id = request.form.get("category_id", "")
    subcat_id = request.form.get("subcategory_id", "")
    task_name = request.form.get("task_name", "").strip()
    description = request.form.get("description", "").strip()
    assigned_to_str = request.form.get("assigned_to", "").strip()
    start_date = request.form.get("start_date", "").strip()
    end_date = request.form.get("end_date", "").strip()
    status = request.form.get("status", "未着手")
    progress_str = request.form.get("progress", "0").strip()
    delay_str = request.form.get("delay_days", "0").strip()

    if not task_name or not start_date or not end_date:
        flash("タスク名・開始日・終了日は必須です。", "warning")
        return redirect(url_for("project_tasks_bp.task_list"))

    if status not in PROJECT_TASK_STATUSES:
        status = "未着手"

    try:
        progress = max(0, min(int(progress_str), 100))
    except ValueError:
        progress = 0
    try:
        delay_days = max(0, int(delay_str))
    except ValueError:
        delay_days = 0
    try:
        assigned_to = int(assigned_to_str) if assigned_to_str else None
    except ValueError:
        assigned_to = None
    assigned_to_2_str = request.form.get("assigned_to_2", "").strip()
    try:
        assigned_to_2 = int(assigned_to_2_str) if assigned_to_2_str else None
    except ValueError:
        assigned_to_2 = None

    is_milestone = 1 if request.form.get("is_milestone") else 0

    add_project_task(
        category_id=int(cat_id) if cat_id else None,
        subcategory_id=int(subcat_id) if subcat_id else None,
        task_name=task_name,
        description=description,
        start_date=start_date,
        end_date=end_date,
        status=status,
        progress=progress,
        delay_days=delay_days,
        created_by=int(session["user_id"]),
        updated_by=session.get("user_name", ""),
        assigned_to=assigned_to,
        assigned_to_2=assigned_to_2,
        is_milestone=is_milestone,
    )
    flash("タスクを追加しました。", "success")
    return redirect(url_for("project_tasks_bp.task_list"))


@project_tasks_bp.route("/update/<int:task_id>", methods=["POST"])
def update_task(task_id: int) -> object:
    """プロジェクトタスクを更新する（管理職・マスタのみ）。

    Args:
        task_id: タスクID

    Returns:
        object: 一覧画面へのリダイレクト
    """
    if not is_privileged(session.get("user_role", "")):
        abort(403)
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    existing = get_project_task_by_id(task_id)
    if not existing:
        abort(404)

    cat_id = request.form.get("category_id", "")
    subcat_id = request.form.get("subcategory_id", "")
    task_name = request.form.get("task_name", "").strip()
    description = request.form.get("description", "").strip()
    assigned_to_str = request.form.get("assigned_to", "").strip()
    start_date = request.form.get("start_date", "").strip()
    end_date = request.form.get("end_date", "").strip()
    status = request.form.get("status", "未着手")
    progress_str = request.form.get("progress", "0").strip()
    delay_str = request.form.get("delay_days", "0").strip()

    if not task_name or not start_date or not end_date:
        flash("タスク名・開始日・終了日は必須です。", "warning")
        return redirect(url_for("project_tasks_bp.task_list"))

    if status not in PROJECT_TASK_STATUSES:
        status = "未着手"

    try:
        progress = max(0, min(int(progress_str), 100))
    except ValueError:
        progress = 0
    try:
        delay_days = max(0, int(delay_str))
    except ValueError:
        delay_days = 0
    try:
        assigned_to = int(assigned_to_str) if assigned_to_str else None
    except ValueError:
        assigned_to = None
    assigned_to_2_str = request.form.get("assigned_to_2", "").strip()
    try:
        assigned_to_2 = int(assigned_to_2_str) if assigned_to_2_str else None
    except ValueError:
        assigned_to_2 = None

    is_milestone = 1 if request.form.get("is_milestone") else 0

    update_project_task(
        task_id=task_id,
        category_id=int(cat_id) if cat_id else None,
        subcategory_id=int(subcat_id) if subcat_id else None,
        task_name=task_name,
        description=description,
        start_date=start_date,
        end_date=end_date,
        status=status,
        progress=progress,
        delay_days=delay_days,
        updated_by=session.get("user_name", ""),
        assigned_to=assigned_to,
        assigned_to_2=assigned_to_2,
        is_milestone=is_milestone,
    )
    flash("タスクを更新しました。", "success")
    return redirect(url_for("project_tasks_bp.task_list"))


@project_tasks_bp.route("/bulk-update", methods=["POST"])
def bulk_update_tasks() -> object:
    """プロジェクトタスクを一括更新する（管理職・マスタのみ）。

    Returns:
        object: 一覧画面へのリダイレクト
    """
    if not is_privileged(session.get("user_role", "")):
        abort(403)
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    task_ids_raw = request.form.getlist("task_id")
    updated_count = 0

    for raw_id in task_ids_raw:
        try:
            task_id = int(raw_id)
        except ValueError:
            continue

        existing = get_project_task_by_id(task_id)
        if not existing:
            continue

        sfx = f"_{task_id}"
        task_name = request.form.get(f"task_name{sfx}", "").strip()
        description = request.form.get(f"description{sfx}", "").strip()
        start_date = request.form.get(f"start_date{sfx}", "").strip()
        end_date = request.form.get(f"end_date{sfx}", "").strip()
        status = request.form.get(f"status{sfx}", "未着手")
        progress_str = request.form.get(f"progress{sfx}", "0").strip()
        delay_str = request.form.get(f"delay_days{sfx}", "0").strip()
        assigned_to_str = request.form.get(f"assigned_to{sfx}", "").strip()
        assigned_to_2_str = request.form.get(f"assigned_to_2{sfx}", "").strip()
        cat_id = request.form.get(f"category_id{sfx}", "")
        subcat_id = request.form.get(f"subcategory_id{sfx}", "")

        if not task_name or not start_date or not end_date:
            continue

        if status not in PROJECT_TASK_STATUSES:
            status = "未着手"
        try:
            progress = max(0, min(int(progress_str), 100))
        except ValueError:
            progress = 0
        try:
            delay_days = max(0, int(delay_str))
        except ValueError:
            delay_days = 0
        try:
            assigned_to = int(assigned_to_str) if assigned_to_str else None
        except ValueError:
            assigned_to = None
        try:
            assigned_to_2 = int(assigned_to_2_str) if assigned_to_2_str else None
        except ValueError:
            assigned_to_2 = None

        is_milestone = 1 if request.form.get(f"is_milestone{sfx}") else 0

        update_project_task(
            task_id=task_id,
            category_id=int(cat_id) if cat_id else None,
            subcategory_id=int(subcat_id) if subcat_id else None,
            task_name=task_name,
            description=description,
            start_date=start_date,
            end_date=end_date,
            status=status,
            progress=progress,
            delay_days=delay_days,
            updated_by=session.get("user_name", ""),
            assigned_to=assigned_to,
            assigned_to_2=assigned_to_2,
            is_milestone=is_milestone,
        )
        updated_count += 1

    flash(f"{updated_count}件のタスクを更新しました。", "success")
    return redirect(url_for("project_tasks_bp.task_list"))


@project_tasks_bp.route("/delete/<int:task_id>", methods=["POST"])
def delete_task(task_id: int) -> object:
    """プロジェクトタスクを削除する（管理職・マスタのみ）。

    Args:
        task_id: タスクID

    Returns:
        object: 一覧画面へのリダイレクト
    """
    if not is_privileged(session.get("user_role", "")):
        abort(403)
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    delete_project_task(task_id)
    flash("タスクを削除しました。", "success")
    return redirect(url_for("project_tasks_bp.task_list"))


# -- ステータス→色マッピング（ダッシュボード用） --
_STATUS_COLOR_MAP: dict[str, str] = {
    "未着手": "#9ca3af",
    "着手": "#60a5fa",
    "順調": "#34d399",
    "遅れ": "#f87171",
    "完了": "#10b981",
    "停止": "#d1d5db",
}


def _build_chart_json(summary: dict) -> dict:
    """サマリー情報からグラフ描画用のJSON構造を構築する。

    Args:
        summary: get_task_progress_summary() の戻り値。

    Returns:
        dict: ステータス別集計とタスク別進捗を含むグラフ用辞書。
    """
    status_breakdown: dict[str, int] = summary["status_breakdown"]
    status_labels: list[str] = list(status_breakdown.keys())
    status_counts: list[int] = list(status_breakdown.values())
    status_colors: list[str] = [
        _STATUS_COLOR_MAP.get(s, "#9ca3af") for s in status_labels
    ]

    task_names: list[str] = []
    task_progresses: list[float] = []
    task_colors: list[str] = []
    task_statuses: list[str] = []

    for task in summary["tasks"]:
        task_names.append(task.get("task_name", ""))
        task_progresses.append(float(task.get("progress", 0) or 0))
        status: str = task.get("status", "未着手")
        task_statuses.append(status)
        task_colors.append(_STATUS_COLOR_MAP.get(status, "#9ca3af"))

    return {
        "status_labels": status_labels,
        "status_counts": status_counts,
        "status_colors": status_colors,
        "task_names": task_names,
        "task_progresses": task_progresses,
        "task_colors": task_colors,
        "task_statuses": task_statuses,
    }


def _resolve_dashboard_target(
    login_id: int, login_role: str, login_dept: str,
) -> tuple[int, str, list[dict]]:
    """ダッシュボード表示対象のユーザーIDと名前、選択可能ユーザーリストを返す。

    Args:
        login_id: ログインユーザーID。
        login_role: ログインユーザーの役職。
        login_dept: ログインユーザーの部署。

    Returns:
        tuple: (target_user_id, target_user_name, selectable_users)

    Raises:
        Werkzeug 403: 権限外のユーザーを指定した場合。
    """
    selectable_users: list[dict] = get_accessible_users_for_dashboard(
        login_id, login_role, login_dept,
    )
    privileged: bool = is_privileged(login_role)

    # クエリパラメータからユーザーIDを取得
    raw_user_id: str | None = request.args.get("user_id")
    if raw_user_id is not None:
        try:
            target_user_id: int = int(raw_user_id)
        except (ValueError, TypeError):
            abort(400)
    else:
        target_user_id = login_id

    # 権限チェック（自分自身は常に許可）
    if target_user_id != login_id:
        if not privileged:
            abort(403)
        else:
            accessible_ids: set[int] = {u["id"] for u in selectable_users}
            if target_user_id not in accessible_ids:
                abort(403)

    # 対象ユーザー名を解決
    target_user_name: str = session.get("user_name", "")
    for u in selectable_users:
        if u["id"] == target_user_id:
            target_user_name = u["name"]
            break

    return target_user_id, target_user_name, selectable_users


@project_tasks_bp.route("/dashboard")
def progress_dashboard() -> str:
    """進捗ダッシュボード画面を表示する。

    セッションの権限に応じて、自分または他ユーザーのタスク進捗を
    グラフ付きで閲覧できる。

    Returns:
        str: レンダリングされたHTMLテンプレート。
    """
    login_role: str = session.get("user_role", "")
    login_id: int = int(session["user_id"])
    login_dept: str = session.get("user_dept", "")
    privileged: bool = is_privileged(login_role)

    target_user_id, target_user_name, selectable_users = _resolve_dashboard_target(
        login_id, login_role, login_dept,
    )

    summary: dict = get_task_progress_summary(target_user_id)
    chart_json: dict = _build_chart_json(summary)

    return render_template(
        "project_tasks_dashboard.html",
        summary=summary,
        chart_json=chart_json,
        privileged=privileged,
        selectable_users=selectable_users,
        selected_user_id=target_user_id,
        selected_user_name=target_user_name,
    )


@project_tasks_bp.route("/dashboard/api")
def progress_dashboard_api() -> tuple:
    """ユーザー切替時にダッシュボードデータをJSONで返すAPI。

    クエリパラメータ user_id で対象ユーザーを指定する。
    権限チェックは progress_dashboard() と同一ロジック。

    Returns:
        tuple: (Response, status_code) JSON形式のレスポンス。
    """
    login_role: str = session.get("user_role", "")
    login_id: int = int(session["user_id"])
    login_dept: str = session.get("user_dept", "")
    privileged: bool = is_privileged(login_role)

    target_user_id, target_user_name, selectable_users = _resolve_dashboard_target(
        login_id, login_role, login_dept,
    )

    summary: dict = get_task_progress_summary(target_user_id)
    chart_json: dict = _build_chart_json(summary)

    return jsonify({
        "summary": summary,
        "chart_json": chart_json,
        "privileged": privileged,
        "selectable_users": selectable_users,
        "selected_user_id": target_user_id,
        "selected_user_name": target_user_name,
    }), 200


@project_tasks_bp.route("/gantt/update-dates/<int:task_id>", methods=["POST"])
def gantt_update_dates(task_id: int) -> tuple:
    """ガントチャートのドラッグ操作で開始日・終了日を更新するAPI。

    管理職・マスタのみ使用可能。JSON形式で start_date / end_date を受け取る。

    Args:
        task_id: 更新対象のタスクID。

    Returns:
        tuple: (Response, status_code) JSON形式のレスポンス。
    """
    if not is_privileged(session.get("user_role", "")):
        abort(403)

    # JSON APIのCSRFチェック（X-CSRF-Token ヘッダー）
    csrf = request.headers.get("X-CSRF-Token", "")
    if csrf != session.get("csrf_token"):
        abort(400)

    existing = get_project_task_by_id(task_id)
    if not existing:
        abort(404)

    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "JSONデータが必要です"}), 400

    start_date: str = data.get("start_date", "").strip()
    end_date: str = data.get("end_date", "").strip()

    if not start_date or not end_date:
        return jsonify({"error": "開始日・終了日は必須です"}), 400

    # 日付形式の検証
    from datetime import datetime
    try:
        datetime.strptime(start_date, "%Y-%m-%d")
        datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError:
        return jsonify({"error": "日付形式が不正です（YYYY-MM-DD）"}), 400

    update_project_task(
        task_id=task_id,
        category_id=existing.get("category_id"),
        subcategory_id=existing.get("subcategory_id"),
        task_name=existing["task_name"],
        description=existing.get("description", ""),
        start_date=start_date,
        end_date=end_date,
        status=existing["status"],
        progress=existing.get("progress", 0),
        delay_days=existing.get("delay_days", 0),
        updated_by=session.get("user_name", ""),
        assigned_to=existing.get("assigned_to"),
        assigned_to_2=existing.get("assigned_to_2"),
        is_milestone=existing.get("is_milestone", 0),
    )

    return jsonify({"ok": True, "start_date": start_date, "end_date": end_date}), 200


@project_tasks_bp.route("/gantt")
def gantt() -> str:
    """ガントチャート画面を表示する。

    全権限で閲覧可能。一般ユーザーは自分のタスクのみ表示。

    Returns:
        str: レンダリング済みHTML
    """
    login_role = session.get("user_role", "")
    login_id = int(session["user_id"])
    privileged = is_privileged(login_role)

    if privileged:
        tasks = get_all_project_tasks()
    else:
        tasks = get_all_project_tasks(assigned_to=login_id)

    # テンプレートに渡すJSON用データ
    gantt_data = []
    for t in tasks:
        # 担当者表示（姓のみ、2名対応）
        names = []
        ln1 = t.get("assigned_last_name") or t.get("assigned_name") or ""
        ln2 = t.get("assigned_last_name_2") or t.get("assigned_name_2") or ""
        if ln1:
            names.append(ln1)
        if ln2:
            names.append(ln2)
        gantt_data.append({
            "id": t["id"],
            "name": t["task_name"],
            "assigned": "・".join(names),
            "category": t.get("category_name") or "",
            "subcategory": t.get("subcategory_name") or "",
            "start": t["start_date"],
            "end": t["end_date"],
            "progress": t.get("progress", 0),
            "status": t["status"],
            "is_milestone": t.get("is_milestone", 0),
        })

    return render_template(
        "project_tasks_gantt.html",
        gantt_json=json.dumps(gantt_data, ensure_ascii=False),
        privileged=privileged,
        csrf_token=session.get("csrf_token", ""),
    )

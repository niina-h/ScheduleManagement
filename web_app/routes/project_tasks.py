"""プロジェクトタスク管理ルート（管理職・マスタ専用）。"""
from __future__ import annotations

from flask import (
    Blueprint, abort, flash, redirect, render_template,
    request, session, url_for,
)

from ..auth_helpers import is_privileged
from ..models import (
    PROJECT_TASK_STATUSES,
    add_project_task,
    delete_project_task,
    get_all_categories,
    get_all_project_tasks,
    get_all_subcategories,
    get_project_task_by_id,
    update_project_task,
)

project_tasks_bp = Blueprint(
    "project_tasks_bp", __name__, url_prefix="/project-tasks",
)


@project_tasks_bp.before_request
def _check_privileged() -> object | None:
    """管理職またはマスタでなければリダイレクトする。"""
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    if not is_privileged(session.get("user_role", "")):
        abort(403)
    return None


@project_tasks_bp.route("/")
def task_list() -> str:
    """プロジェクトタスク一覧画面を表示する。

    Returns:
        str: レンダリング済みHTML
    """
    tasks = get_all_project_tasks()
    categories = get_all_categories()
    subcategories = get_all_subcategories()
    return render_template(
        "project_tasks.html",
        tasks=tasks,
        categories=categories,
        subcategories=subcategories,
        statuses=PROJECT_TASK_STATUSES,
    )


@project_tasks_bp.route("/add", methods=["POST"])
def add_task() -> object:
    """プロジェクトタスクを追加する。

    Returns:
        object: 一覧画面へのリダイレクト
    """
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    cat_id = request.form.get("category_id", "")
    subcat_id = request.form.get("subcategory_id", "")
    task_name = request.form.get("task_name", "").strip()
    description = request.form.get("description", "").strip()
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
    )
    flash("タスクを追加しました。", "success")
    return redirect(url_for("project_tasks_bp.task_list"))


@project_tasks_bp.route("/update/<int:task_id>", methods=["POST"])
def update_task(task_id: int) -> object:
    """プロジェクトタスクを更新する。

    Args:
        task_id: タスクID

    Returns:
        object: 一覧画面へのリダイレクト
    """
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    existing = get_project_task_by_id(task_id)
    if not existing:
        abort(404)

    cat_id = request.form.get("category_id", "")
    subcat_id = request.form.get("subcategory_id", "")
    task_name = request.form.get("task_name", "").strip()
    description = request.form.get("description", "").strip()
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
    )
    flash("タスクを更新しました。", "success")
    return redirect(url_for("project_tasks_bp.task_list"))


@project_tasks_bp.route("/delete/<int:task_id>", methods=["POST"])
def delete_task(task_id: int) -> object:
    """プロジェクトタスクを削除する。

    Args:
        task_id: タスクID

    Returns:
        object: 一覧画面へのリダイレクト
    """
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    delete_project_task(task_id)
    flash("タスクを削除しました。", "success")
    return redirect(url_for("project_tasks_bp.task_list"))

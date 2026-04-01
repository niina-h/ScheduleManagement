"""プロジェクトタスク管理ルート。

管理職・マスタは全タスクを閲覧・編集できる。
一般ユーザーは自分に割り当てられたタスクのみ参照可能（編集不可）。
"""
from __future__ import annotations

import io
import json
import logging
from datetime import date, timedelta

import openpyxl
from flask import (
    Blueprint, abort, flash, jsonify, redirect, render_template,
    request, send_file, session, url_for,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..auth_helpers import is_master, is_privileged
from ..models import (
    PROJECT_TASK_STATUSES,
    add_project_task,
    delete_project_task,
    delete_routine_task,
    get_accessible_users,
    get_accessible_users_for_dashboard,
    get_all_categories,
    get_all_project_tasks,
    get_all_subcategories,
    get_all_users,
    get_project_task_by_id,
    get_routine_schedules,
    get_task_master,
    get_task_overview_summary,
    get_task_progress_summary,
    import_brabio_excel,
    save_routine_task,
    update_project_task,
)

logger = logging.getLogger(__name__)

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
    login_dept = session.get("user_dept", "")
    privileged = is_privileged(login_role)

    # 権限別タスク取得
    if login_role == "マスタ":
        # マスタ: 同一部署のメンバーのタスクのみ
        accessible = get_accessible_users(login_id, login_role, login_dept)
        accessible_ids = [u["id"] for u in accessible]
        tasks = get_all_project_tasks(user_ids=accessible_ids)
    elif login_role == "管理職":
        # 管理職: 担当メンバー全員のタスク
        accessible = get_accessible_users(login_id, login_role, login_dept)
        accessible_ids = [u["id"] for u in accessible]
        tasks = get_all_project_tasks(user_ids=accessible_ids)
    elif privileged:
        tasks = get_all_project_tasks()
    else:
        tasks = get_all_project_tasks(assigned_to=login_id)

    categories = get_all_categories()
    subcategories = get_all_subcategories()

    # 担当者選択用のユーザーリスト（権限別に絞り込み）
    if login_role in ("マスタ", "管理職"):
        users = get_accessible_users(login_id, login_role, login_dept)
    elif privileged:
        users = get_all_users()
    else:
        users = []

    # 定例スケジュール
    routine_schedules = get_routine_schedules(login_id)
    # 作業登録から「定例作業」カテゴリの作業を取得
    all_task_master = get_task_master(login_id)
    routine_task_options = [
        t for t in all_task_master
        if t.get("category_name") in ("定例", "定例作業")
        or t.get("subcategory_name") in ("定例", "定例作業")
    ]
    used_rows = {r["row_number"] for r in routine_schedules}

    return render_template(
        "project_tasks.html",
        tasks=tasks,
        categories=categories,
        subcategories=subcategories,
        statuses=PROJECT_TASK_STATUSES,
        privileged=privileged,
        users=users,
        routine_schedules=routine_schedules,
        routine_task_options=routine_task_options,
        used_rows=used_rows,
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
    is_event_add = 1 if request.form.get("is_event") else 0
    # イベントは開催日のみのため end_date が空の場合は start_date を使用
    if is_event_add and not end_date:
        end_date = start_date
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
    is_event = 1 if request.form.get("is_event") else 0
    event_start_time = request.form.get("event_start_time", "").strip()
    event_end_time = request.form.get("event_end_time", "").strip()
    planned_hours_str = request.form.get("planned_hours", "0").strip()
    try:
        planned_hours = max(0.0, float(planned_hours_str))
    except ValueError:
        planned_hours = 0.0

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
        is_event=is_event,
        event_start_time=event_start_time,
        event_end_time=event_end_time,
        planned_hours=planned_hours,
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
    is_event = 1 if request.form.get("is_event") else 0
    event_start_time = request.form.get("event_start_time", "").strip()
    event_end_time = request.form.get("event_end_time", "").strip()
    planned_hours_str = request.form.get("planned_hours", "0").strip()
    try:
        planned_hours = max(0.0, float(planned_hours_str))
    except ValueError:
        planned_hours = 0.0

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
        is_event=is_event,
        event_start_time=event_start_time,
        event_end_time=event_end_time,
        planned_hours=planned_hours,
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
    deleted_count = 0

    for raw_id in task_ids_raw:
        try:
            task_id = int(raw_id)
        except ValueError:
            continue

        existing = get_project_task_by_id(task_id)
        if not existing:
            continue

        # 削除チェックボックスが ON の場合は削除して次へ
        if request.form.get(f"delete_{task_id}"):
            delete_project_task(task_id)
            deleted_count += 1
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
        is_event = 1 if request.form.get(f"is_event{sfx}") else 0
        event_start_time = request.form.get(f"event_start_time{sfx}", "").strip()
        event_end_time = request.form.get(f"event_end_time{sfx}", "").strip()
        planned_hours_str = request.form.get(f"planned_hours{sfx}", "0").strip()
        try:
            planned_hours_val = max(0.0, float(planned_hours_str))
        except ValueError:
            planned_hours_val = 0.0

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
            is_event=is_event,
            event_start_time=event_start_time,
            event_end_time=event_end_time,
            planned_hours=planned_hours_val,
        )
        updated_count += 1

    msgs = []
    if updated_count:
        msgs.append(f"{updated_count}件更新")
    if deleted_count:
        msgs.append(f"{deleted_count}件削除")
    flash("、".join(msgs) + "しました。" if msgs else "変更はありませんでした。", "success")
    return redirect(url_for("project_tasks_bp.task_list"))


@project_tasks_bp.route("/import-brabio", methods=["POST"])
def import_brabio() -> object:
    """ブラビオExcelからタスクをインポートする（管理職・マスタのみ）。

    data/ ディレクトリ内のExcelファイルを読み込み、
    タスク管理にインポートする。

    Returns:
        object: 一覧画面へのリダイレクト
    """
    import pathlib

    if not is_privileged(session.get("user_role", "")):
        abort(403)
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    # アップロードファイル or デフォルトファイル
    uploaded = request.files.get("brabio_file")
    if uploaded and uploaded.filename:
        # アップロードされたファイルを一時保存
        import tempfile
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        uploaded.save(tmp.name)
        file_path = tmp.name
    else:
        # data/ または reports/ ディレクトリのデフォルトファイルを検索
        project_root = pathlib.Path(__file__).resolve().parent.parent.parent
        candidates = [
            project_root / "reports" / "商品開発業務.xlsx",
            project_root / "data" / "ブラビオ商品開発業務.xlsx",
        ]
        # reports/ 内の .xlsx を追加検索
        reports_dir = project_root / "reports"
        if reports_dir.is_dir():
            for f in reports_dir.glob("*.xlsx"):
                if f not in candidates:
                    candidates.append(f)
        file_path = None
        for cand in candidates:
            if cand.exists():
                file_path = str(cand)
                break
        if file_path is None:
            flash("インポートファイルが見つかりません", "warning")
            return redirect(url_for("project_tasks_bp.task_list"))

    result = import_brabio_excel(
        file_path=file_path,
        created_by=int(session["user_id"]),
        updated_by=session.get("user_name", ""),
    )

    parts = []
    if result["imported"]:
        parts.append(f"{result['imported']}件取込")
    if result.get("updated"):
        parts.append(f"{result['updated']}件更新")
    msg = f"インポート完了: {' / '.join(parts) if parts else '変更なし'}"
    if result["errors"]:
        msg += f" / {len(result['errors'])}件エラー"
    flash(msg, "success" if (result["imported"] > 0 or result.get("updated", 0) > 0) else "info")
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


@project_tasks_bp.route("/routine/save", methods=["POST"])
def save_routine() -> object:
    """定例スケジュールを登録する（ログインユーザー自身）。

    Returns:
        object: タスク管理画面へのリダイレクト
    """
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)
    user_id = int(session["user_id"])
    task_name = request.form.get("task_name", "").strip()
    subcategory_name = request.form.get("subcategory_name", "").strip()
    row_number_str = request.form.get("row_number", "0").strip()
    default_hours_str = request.form.get("default_hours", "0").strip()

    if not task_name:
        flash("作業名を選択してください。", "warning")
        return redirect(url_for("project_tasks_bp.task_list"))
    try:
        row_number = int(row_number_str)
        if not 1 <= row_number <= 10:
            raise ValueError
    except ValueError:
        flash("行番号は1〜10で指定してください。", "warning")
        return redirect(url_for("project_tasks_bp.task_list"))
    try:
        default_hours = max(0.0, float(default_hours_str))
    except ValueError:
        default_hours = 0.0

    ok = save_routine_task(user_id, task_name, subcategory_name, default_hours, row_number)
    flash("定例スケジュールを登録しました。" if ok else "登録に失敗しました（行番号重複の可能性）。",
          "success" if ok else "warning")
    return redirect(url_for("project_tasks_bp.task_list"))


@project_tasks_bp.route("/routine/delete/<int:routine_id>", methods=["POST"])
def delete_routine(routine_id: int) -> object:
    """定例スケジュールを削除する（ログインユーザー自身）。

    Args:
        routine_id: 定例スケジュールID

    Returns:
        object: タスク管理画面へのリダイレクト
    """
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)
    user_id = int(session["user_id"])
    delete_routine_task(routine_id, user_id)
    flash("定例スケジュールを削除しました。", "success")
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

    # クエリパラメータからユーザーIDを取得（なければセッションの選択ユーザーを維持）
    raw_user_id: str | None = request.args.get("user_id")
    if raw_user_id is not None:
        try:
            target_user_id: int = int(raw_user_id)
        except (ValueError, TypeError):
            abort(400)
        session["selected_user_id"] = target_user_id
    elif privileged and session.get("selected_user_id"):
        target_user_id = int(session["selected_user_id"])
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


@project_tasks_bp.route("/overview")
def task_overview() -> str:
    """管理者・マスタ向けタスク全体俯瞰ダッシュボードを表示する。

    全タスクのステータス別円グラフ、カテゴリ別進捗バーチャート、
    担当者別集計を表示する。管理職・マスタのみアクセス可。

    Returns:
        str: レンダリング済みHTML
    """
    login_role: str = session.get("user_role", "")
    if not is_privileged(login_role):
        abort(403)

    summary: dict = get_task_overview_summary()
    chart_json: dict = _build_overview_chart_json(summary)

    return render_template(
        "project_tasks_overview.html",
        summary=summary,
        chart_json=chart_json,
    )


def _build_overview_chart_json(summary: dict) -> dict:
    """全体俯瞰ダッシュボード用のグラフJSON構造を構築する。

    Args:
        summary: get_task_overview_summary() の戻り値。

    Returns:
        dict: ステータス別・カテゴリ別・担当者別のグラフ描画データ。
    """
    status_breakdown: dict[str, int] = summary["status_breakdown"]
    status_labels: list[str] = list(status_breakdown.keys())
    status_counts: list[int] = list(status_breakdown.values())
    status_colors: list[str] = [
        _STATUS_COLOR_MAP.get(s, "#9ca3af") for s in status_labels
    ]

    # カテゴリ別
    cat_names: list[str] = [c["name"] for c in summary["category_summary"]]
    cat_avg_progress: list[float] = [c["avg_progress"] for c in summary["category_summary"]]
    cat_counts: list[int] = [c["count"] for c in summary["category_summary"]]
    cat_completed: list[int] = [c["completed"] for c in summary["category_summary"]]
    cat_delayed: list[int] = [c["delayed"] for c in summary["category_summary"]]

    # 担当者別
    user_names: list[str] = [u["name"] for u in summary["user_summary"]]
    user_avg_progress: list[float] = [u["avg_progress"] for u in summary["user_summary"]]
    user_counts: list[int] = [u["count"] for u in summary["user_summary"]]
    user_completed: list[int] = [u["completed"] for u in summary["user_summary"]]
    user_delayed: list[int] = [u["delayed"] for u in summary["user_summary"]]

    return {
        "status_labels": status_labels,
        "status_counts": status_counts,
        "status_colors": status_colors,
        "cat_names": cat_names,
        "cat_avg_progress": cat_avg_progress,
        "cat_counts": cat_counts,
        "cat_completed": cat_completed,
        "cat_delayed": cat_delayed,
        "user_names": user_names,
        "user_avg_progress": user_avg_progress,
        "user_counts": user_counts,
        "user_completed": user_completed,
        "user_delayed": user_delayed,
    }


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

    # 管理職・マスタは全体ステータスも同時表示
    overview_summary: dict | None = None
    overview_chart_json: dict | None = None
    if privileged:
        overview_summary = get_task_overview_summary()
        overview_chart_json = _build_overview_chart_json(overview_summary)

    return render_template(
        "project_tasks_dashboard.html",
        summary=summary,
        chart_json=chart_json,
        privileged=privileged,
        selectable_users=selectable_users,
        selected_user_id=target_user_id,
        selected_user_name=target_user_name,
        overview_summary=overview_summary,
        overview_chart_json=overview_chart_json,
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


@project_tasks_bp.route("/gantt/update-fields/<int:task_id>", methods=["POST"])
def gantt_update_fields(task_id: int) -> tuple:
    """ガントチャートから状態・進捗・遅延日を更新するAPI。

    管理職・マスタのみ使用可能。JSON形式で status / progress / delay_days を受け取る。

    Args:
        task_id: 更新対象のタスクID。

    Returns:
        tuple: (Response, status_code) JSON形式のレスポンス。
    """
    if not is_privileged(session.get("user_role", "")):
        abort(403)

    csrf = request.headers.get("X-CSRF-Token", "")
    if csrf != session.get("csrf_token"):
        abort(400)

    existing = get_project_task_by_id(task_id)
    if not existing:
        abort(404)

    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "JSONデータが必要です"}), 400

    status: str = data.get("status", existing["status"])
    if status not in PROJECT_TASK_STATUSES:
        status = existing["status"]

    try:
        progress: int = max(0, min(int(data.get("progress", existing.get("progress", 0))), 100))
    except (ValueError, TypeError):
        progress = existing.get("progress", 0) or 0

    try:
        delay_days: int = max(0, int(data.get("delay_days", existing.get("delay_days", 0))))
    except (ValueError, TypeError):
        delay_days = existing.get("delay_days", 0) or 0

    update_project_task(
        task_id=task_id,
        category_id=existing.get("category_id"),
        subcategory_id=existing.get("subcategory_id"),
        task_name=existing["task_name"],
        description=existing.get("description", ""),
        start_date=existing["start_date"],
        end_date=existing["end_date"],
        status=status,
        progress=progress,
        delay_days=delay_days,
        updated_by=session.get("user_name", ""),
        assigned_to=existing.get("assigned_to"),
        assigned_to_2=existing.get("assigned_to_2"),
        is_milestone=existing.get("is_milestone", 0),
    )

    return jsonify({
        "ok": True, "status": status,
        "progress": progress, "delay_days": delay_days,
    }), 200


@project_tasks_bp.route("/gantt/reorder", methods=["POST"])
def gantt_reorder() -> tuple:
    """ガントチャート上でタスクの表示順を入れ替えるAPI。

    JSON: { "order": [id1, id2, ...] } — カテゴリ内の表示順を更新する。

    Returns:
        tuple: (Response, status_code)
    """
    if not is_privileged(session.get("user_role", "")):
        abort(403)

    csrf = request.headers.get("X-CSRF-Token", "")
    if csrf != session.get("csrf_token"):
        abort(400)

    data = request.get_json(silent=True)
    if not data or "order" not in data:
        return jsonify({"error": "orderが必要です"}), 400

    db = get_db()
    order_list = data["order"]
    for idx, task_id in enumerate(order_list):
        db.execute(
            "UPDATE project_task SET display_order = ? WHERE id = ?",
            (idx, int(task_id)),
        )
    db.commit()
    return jsonify({"ok": True}), 200


@project_tasks_bp.route("/gantt")
def gantt() -> str:
    """ガントチャート画面を表示する。

    全権限で閲覧可能。一般ユーザーは自分のタスクのみ表示。

    Returns:
        str: レンダリング済みHTML
    """
    login_role = session.get("user_role", "")
    login_id = int(session["user_id"])
    login_dept = session.get("user_dept", "")
    privileged = is_privileged(login_role)

    # タスク一覧と同じスコープで取得
    if login_role == "マスタ":
        accessible = get_accessible_users(login_id, login_role, login_dept)
        accessible_ids = [u["id"] for u in accessible]
        tasks = get_all_project_tasks(user_ids=accessible_ids)
    elif login_role == "管理職":
        accessible = get_accessible_users(login_id, login_role, login_dept)
        accessible_ids = [u["id"] for u in accessible]
        tasks = get_all_project_tasks(user_ids=accessible_ids)
    elif privileged:
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
        # イベントはマイルストーンとして表示
        is_event = t.get("is_event", 0)
        is_ms = t.get("is_milestone", 0) or is_event
        gantt_data.append({
            "id": t["id"],
            "name": t["task_name"],
            "assigned": "・".join(names),
            "category": (t.get("category_name") or "（イベント）") if is_event else (t.get("category_name") or ""),
            "subcategory": t.get("subcategory_name") or "",
            "start": t["start_date"],
            "end": t["end_date"],
            "progress": t.get("progress", 0),
            "status": t["status"],
            "delay_days": t.get("delay_days", 0) or 0,
            "is_milestone": is_ms,
        })

    return render_template(
        "project_tasks_gantt.html",
        gantt_json=json.dumps(gantt_data, ensure_ascii=False),
        privileged=privileged,
        csrf_token=session.get("csrf_token", ""),
    )


# ---------------------------------------------------------------------------
# ガントチャート Excel エクスポート
# ---------------------------------------------------------------------------

# ステータスに対応するExcelセル色（RRGGBB）
_STATUS_FILL: dict[str, str] = {
    "未着手": "D1D5DB",
    "着手":   "93C5FD",
    "順調":   "6EE7B7",
    "遅れ":   "FCA5A5",
    "完了":   "10B981",
    "停止":   "E5E7EB",
}

# 進捗バーの色
_PROGRESS_FILL: dict[str, str] = {
    "遅れ": "EF4444",
    "_default": "059669",
}


def _get_monday(d: date) -> date:
    """指定日が属する週の月曜日を返す。"""
    return d - timedelta(days=d.weekday())


def _build_gantt_excel(
    tasks: list[dict],
    start_date: date,
    display_days: int,
) -> openpyxl.Workbook:
    """ガントチャート付きExcelワークブックを生成する。

    左側に大項目・タスク名・担当・状態・進捗を配置し、
    右側に日付ごとのセル塗りつぶしでガントバーを描画する。

    Args:
        tasks: タスク一覧（get_all_project_tasks の戻り値）
        start_date: 表示開始日
        display_days: 表示日数

    Returns:
        openpyxl.Workbook: 生成済みワークブック
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ガントチャート"

    # -- スタイル定義 --
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    header_fill = PatternFill("solid", fgColor="334155")
    header_font = Font(bold=True, color="FFFFFF", size=9, name="游ゴシック")
    cat_fill = PatternFill("solid", fgColor="E2E8F0")
    cat_font = Font(bold=True, size=10, name="游ゴシック")
    body_font = Font(size=9, name="游ゴシック")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    sun_fill = PatternFill("solid", fgColor="FEF2F2")
    sat_fill = PatternFill("solid", fgColor="EFF6FF")
    today_fill = PatternFill("solid", fgColor="FEE2E2")
    milestone_fill = PatternFill("solid", fgColor="C4B5FD")

    # 固定列: A=大項目, B=タスク名, C=担当, D=状態, E=進捗%
    fixed_cols = 5
    col_widths = {"A": 14, "B": 24, "C": 10, "D": 8, "E": 7}

    # -- ヘッダー行1: 固定列名 + 日付 --
    headers = ["大項目", "タスク名", "担当", "状態", "進捗%"]
    weekday_ja = ["月", "火", "水", "木", "金", "土", "日"]

    # 行1: 日付ヘッダー (月/日)
    # 行2: 曜日ヘッダー
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(1, ci, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        # 行2もヘッダー背景
        cell2 = ws.cell(2, ci, "")
        cell2.fill = header_fill
        cell2.border = thin_border

    today_d = date.today()
    for di in range(display_days):
        dt = start_date + timedelta(days=di)
        col = fixed_cols + di + 1
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = 3.5

        # 行1: 月/日
        cell1 = ws.cell(1, col, f"{dt.month}/{dt.day}")
        cell1.font = Font(size=7, name="游ゴシック", bold=True,
                          color="EF4444" if dt.weekday() == 6 else
                          "3B82F6" if dt.weekday() == 5 else "FFFFFF")
        cell1.fill = header_fill
        cell1.alignment = center_align
        cell1.border = thin_border

        # 行2: 曜日
        cell2 = ws.cell(2, col, weekday_ja[dt.weekday()])
        cell2.font = Font(size=7, name="游ゴシック",
                          color="EF4444" if dt.weekday() == 6 else
                          "3B82F6" if dt.weekday() == 5 else "333333")
        cell2.fill = PatternFill("solid", fgColor="F1F5F9")
        cell2.alignment = center_align
        cell2.border = thin_border

    # 列幅
    for letter, w in col_widths.items():
        ws.column_dimensions[letter].width = w

    # -- カテゴリ別にグループ化 --
    cat_order: list[str] = []
    cat_map: dict[str, list[dict]] = {}
    for t in tasks:
        cat = t.get("category_name") or "（未分類）"
        if cat not in cat_map:
            cat_map[cat] = []
            cat_order.append(cat)
        cat_map[cat].append(t)

    row = 3  # データ開始行
    for cat in cat_order:
        task_list = cat_map[cat]

        # カテゴリ行
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=fixed_cols + display_days)
        cell = ws.cell(row, 1, f"■ {cat}（{len(task_list)}件）")
        cell.fill = cat_fill
        cell.font = cat_font
        cell.alignment = left_align
        cell.border = thin_border
        row += 1

        # タスク行
        for t in task_list:
            # 担当者名
            names = []
            ln1 = t.get("assigned_last_name") or t.get("assigned_name") or ""
            ln2 = t.get("assigned_last_name_2") or t.get("assigned_name_2") or ""
            if ln1:
                names.append(ln1)
            if ln2:
                names.append(ln2)
            assigned = "・".join(names)

            progress = t.get("progress", 0) or 0
            status = t.get("status", "")

            # 固定列
            ws.cell(row, 1, cat).font = body_font
            ws.cell(row, 1).alignment = left_align
            ws.cell(row, 1).border = thin_border

            name_prefix = "◆ " if t.get("is_milestone") else ""
            ws.cell(row, 2, name_prefix + t["task_name"]).font = body_font
            ws.cell(row, 2).alignment = left_align
            ws.cell(row, 2).border = thin_border

            ws.cell(row, 3, assigned).font = body_font
            ws.cell(row, 3).alignment = center_align
            ws.cell(row, 3).border = thin_border

            status_cell = ws.cell(row, 4, status)
            status_cell.font = body_font
            status_cell.alignment = center_align
            status_cell.border = thin_border

            prog_cell = ws.cell(row, 5, f"{progress}%")
            prog_cell.font = body_font
            prog_cell.alignment = center_align
            prog_cell.border = thin_border

            # ガントバー描画
            try:
                t_start = date.fromisoformat(t["start_date"])
                t_end = date.fromisoformat(t["end_date"])
            except (ValueError, TypeError, KeyError):
                row += 1
                continue

            for di in range(display_days):
                dt = start_date + timedelta(days=di)
                col = fixed_cols + di + 1
                cell = ws.cell(row, col)
                cell.border = thin_border

                # 土日背景
                if dt.weekday() == 5:
                    cell.fill = sat_fill
                elif dt.weekday() == 6:
                    cell.fill = sun_fill

                # 今日ハイライト
                if dt == today_d:
                    cell.fill = today_fill

                # タスク期間内
                if t_start <= dt <= t_end:
                    if t.get("is_milestone"):
                        cell.fill = PatternFill("solid", fgColor="C4B5FD")
                        if dt == t_start:
                            cell.value = "◆"
                            cell.font = Font(size=8, color="6D28D9", bold=True)
                            cell.alignment = center_align
                    elif status == "完了":
                        cell.fill = PatternFill("solid", fgColor="10B981")
                    elif status == "停止":
                        cell.fill = PatternFill("solid", fgColor="E5E7EB")
                    else:
                        # 進捗バー計算
                        total_days = (t_end - t_start).days + 1
                        day_idx = (dt - t_start).days
                        prog_days = round(total_days * progress / 100)

                        if day_idx < prog_days:
                            # 進捗済み部分
                            fill_color = _PROGRESS_FILL.get(
                                status, _PROGRESS_FILL["_default"]
                            )
                            cell.fill = PatternFill("solid", fgColor=fill_color)
                        else:
                            # 予定部分
                            cell.fill = PatternFill("solid", fgColor="93C5FD")

            row += 1

    # 行1-2を固定
    ws.freeze_panes = "F3"

    return wb


@project_tasks_bp.route("/gantt/export")
def export_gantt() -> object:
    """ガントチャートをExcelファイルとしてエクスポートする。

    権限別の出力範囲:
      - ユーザー: 自分が担当のタスクのみ
      - 管理職: 自部署メンバーが担当のタスク
      - マスタ: 全タスク

    クエリパラメータ:
      - start: 表示開始日（YYYY-MM-DD）。省略時は前週月曜日。
      - days: 表示日数。省略時は28。

    Returns:
        object: Excelファイルのダウンロードレスポンス
    """
    login_role: str = session.get("user_role", "")
    login_id: int = int(session["user_id"])
    login_dept: str = session.get("user_dept", "")
    privileged: bool = is_privileged(login_role)

    # 表示期間
    start_param: str = request.args.get("start", "")
    days_param: str = request.args.get("days", "28")

    try:
        display_days = max(7, min(90, int(days_param)))
    except ValueError:
        display_days = 28

    if start_param:
        try:
            start_d = date.fromisoformat(start_param)
        except ValueError:
            start_d = _get_monday(date.today()) - timedelta(days=7)
    else:
        start_d = _get_monday(date.today()) - timedelta(days=7)

    # 権限別タスク取得
    if is_master(login_role):
        # マスタ: 全タスク
        tasks = get_all_project_tasks()
    elif privileged:
        # 管理職: 自部署メンバーのタスク
        dept_users = get_all_users(dept_filter=login_dept)
        dept_user_ids: set[int] = {u["id"] for u in dept_users}
        all_tasks = get_all_project_tasks()
        tasks = [
            t for t in all_tasks
            if (t.get("assigned_to") in dept_user_ids
                or t.get("assigned_to_2") in dept_user_ids)
        ]
    else:
        # 一般ユーザー: 自分のタスクのみ
        tasks = get_all_project_tasks(assigned_to=login_id)

    wb = _build_gantt_excel(tasks, start_d, display_days)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"ガントチャート_{start_d.isoformat()}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

"""管理者向けダッシュボード・ユーザー管理ルートを提供するBlueprintモジュール。"""
from __future__ import annotations

import logging
import re
import secrets
from datetime import date, timedelta

logger = logging.getLogger(__name__)

from flask import (
    Blueprint,
    abort,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

from ..models import (
    add_dept,
    add_user,
    clear_user_password,
    count_operation_logs,
    delete_dept,
    delete_user,
    get_all_depts,
    get_all_users,
    get_all_users_daily_status,
    get_all_users_schedule_status,
    get_operation_logs,
    get_user_by_id,
    save_user_manager,
    set_user_password,
    update_dept,
    update_user,
    update_user_std_hours,
    update_users_order,
    user_has_password,
)
from ..log_service import record_operation, ACTION_USER_ADD, ACTION_USER_DELETE, ACTION_USER_UPDATE
from ..auth_helpers import is_privileged, is_master, can_access_user, can_set_password_for

admin_bp = Blueprint("admin_bp", __name__, url_prefix="/admin")


@admin_bp.before_request
def _require_admin_hook() -> None:
    """全ルートの前に管理職・マスタチェックとCSRFトークン検証を行う。

    - 未ログイン → ログイン画面へリダイレクト
    - 管理職・マスタ以外 → 403
    - POST/PUT/DELETE リクエスト → CSRFトークン検証
    """
    if "user_id" not in session:
        return redirect(url_for("auth.login"))
    if not is_privileged(session.get("user_role", "")):
        abort(403)
    if request.method in ("POST", "PUT", "DELETE"):
        token = request.form.get("csrf_token", "")
        if not secrets.compare_digest(token, session.get("csrf_token", "")):
            abort(400)


def _redirect_dashboard():
    """管理者ダッシュボードへのリダイレクトレスポンスを返す。"""
    return redirect(url_for("admin_bp.dashboard"))


def _get_monday(d: date) -> date:
    """指定日が属する週の月曜日を返す。

    Args:
        d (date): 基準日。

    Returns:
        date: その週の月曜日。
    """
    return d - timedelta(days=d.weekday())


def _get_current_week_start() -> str:
    """今週の月曜日をISO形式文字列で返す。

    Returns:
        str: 今週月曜日の 'YYYY-MM-DD' 形式文字列。
    """
    return _get_monday(date.today()).isoformat()


@admin_bp.route("/", endpoint="dashboard")
def dashboard() -> str:
    """管理者ダッシュボードを表示する。

    クエリパラメータ `week` で週を指定できる（未指定時は今週月曜）。
    全ユーザーの週間予定登録状況とユーザー一覧を表示する。

    Returns:
        str: 管理者ダッシュボードのHTMLレスポンス。
    """
    week_param: str = request.args.get("week", "")
    if week_param:
        try:
            week_date = date.fromisoformat(week_param)
            week_start = _get_monday(week_date).isoformat()
        except ValueError:
            week_start = _get_current_week_start()
    else:
        week_start = _get_current_week_start()

    ws_date = date.fromisoformat(week_start)
    prev_week: str = (ws_date - timedelta(weeks=1)).isoformat()
    next_week: str = (ws_date + timedelta(weeks=1)).isoformat()

    login_role: str = session.get("user_role", "")
    login_dept: str = session.get("user_dept", "")

    # スコープ制限: マスタ・管理職ともに自部署のみ
    dept_filter: str | None = login_dept if login_dept else None
    status_list = get_all_users_schedule_status(week_start, dept_filter=dept_filter)
    all_users = get_all_users(dept_filter=dept_filter)

    # 当日の実績入力状況
    today_str: str = date.today().isoformat()
    daily_status = get_all_users_daily_status(today_str, dept_filter=dept_filter)

    # パスワード設定済みユーザーIDのセット（全ユーザー対象）
    users_with_password: set[int] = {
        u["id"] for u in all_users
        if user_has_password(u["id"])
    }

    all_depts = get_all_depts()

    # 担当メンバー設定用: 同一部署の管理職・マスタ一覧（上長候補）/ 対象ユーザー
    login_id: int = int(session.get("user_id", 0))
    manager_candidates: list[dict] = [
        u for u in all_users if is_privileged(u.get("role", ""))
    ]
    assignable_users: list[dict] = [
        u for u in all_users
        if u.get("role") == "ユーザー" and u.get("dept") == login_dept
    ]

    return render_template(
        "admin.html",
        status_list=status_list,
        week_start=week_start,
        prev_week=prev_week,
        next_week=next_week,
        all_users=all_users,
        today_str=today_str,
        daily_status=daily_status,
        users_with_password=users_with_password,
        all_depts=all_depts,
        login_role=login_role,
        login_id=login_id,
        manager_candidates=manager_candidates,
        assignable_users=assignable_users,
    )


@admin_bp.route("/users/add", methods=["POST"])
def add_user_route() -> str:
    """新しいユーザーを追加する。

    フォームから名前・役職・部署・基本勤務時間を受け取り登録する。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    name: str = request.form.get("name", "").strip()
    role: str = request.form.get("role", "").strip()
    dept: str = request.form.get("dept", "").strip()

    try:
        std_hours = float(request.form.get("std_hours", 8.0))
    except ValueError:
        flash("基本勤務時間は数値で入力してください", "warning")
        return _redirect_dashboard()

    if not name:
        flash("名前を入力してください", "warning")
        return _redirect_dashboard()

    success: bool = add_user(name, role, dept, std_hours)
    if success:
        record_operation(ACTION_USER_ADD, f"role={role}")
        flash(f"ユーザー「{name}」を追加しました", "success")
    else:
        flash("ユーザーの追加に失敗しました（名前が重複している可能性があります）", "warning")

    return _redirect_dashboard()


@admin_bp.route("/users/delete/<int:user_id>", methods=["POST"])
def delete_user_route(user_id: int) -> str:
    """指定したユーザーを削除する。

    Args:
        user_id (int): 削除対象のユーザーID。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    delete_user(user_id)
    record_operation(ACTION_USER_DELETE, f"user_id={user_id}")
    flash("ユーザーを削除しました", "success")
    return _redirect_dashboard()


@admin_bp.route("/users/update_hours/<int:user_id>", methods=["POST"])
def update_hours(user_id: int) -> str:
    """指定ユーザーの基本勤務時間を更新する。

    Args:
        user_id (int): 更新対象のユーザーID。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    try:
        std_hours = float(request.form.get("std_hours", 8.0))
    except ValueError:
        flash("基本勤務時間は数値で入力してください", "warning")
        return _redirect_dashboard()

    update_user_std_hours(user_id, std_hours)
    flash("基本勤務時間を更新しました", "success")
    return _redirect_dashboard()


@admin_bp.route("/users/set_password/<int:user_id>", methods=["POST"])
def set_password(user_id: int) -> str:
    """指定した管理者ユーザーのパスワードを設定・変更する。

    Args:
        user_id (int): パスワードを設定する管理者ユーザーのID。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    target = get_user_by_id(user_id)
    if target is None:
        abort(404)
    operator_role: str = session.get("user_role", "")
    if not can_set_password_for(operator_role, target["role"]):
        abort(403)

    # パスワードは strip() しない（前後の空白も意図した文字として扱う）
    password: str = request.form.get("password", "")
    password_confirm: str = request.form.get("password_confirm", "")

    if not password:
        flash("パスワードを入力してください", "warning")
        return _redirect_dashboard()

    if password != password_confirm:
        flash("パスワードと確認用パスワードが一致しません", "warning")
        return _redirect_dashboard()

    if not re.fullmatch(r'\d{4}', password):
        flash("パスワードは4桁の数字で入力してください", "warning")
        return _redirect_dashboard()

    success = set_user_password(user_id, password)
    if success:
        flash("パスワードを設定しました", "success")
    else:
        flash("パスワードの設定に失敗しました", "warning")

    return _redirect_dashboard()


@admin_bp.route("/users/clear_password/<int:user_id>", methods=["POST"])
def clear_password(user_id: int) -> str:
    """指定した管理者ユーザーのパスワードを削除する（パスワードなしに戻す）。

    Args:
        user_id (int): パスワードを削除する管理者ユーザーのID。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    success = clear_user_password(user_id)
    if success:
        flash("パスワードを削除しました", "success")
    else:
        flash("パスワードの削除に失敗しました", "warning")

    return _redirect_dashboard()


@admin_bp.route("/users/bulk_update", methods=["POST"])
def bulk_update_users() -> str:
    """全ユーザーの情報を一括更新する。

    フォームから user_count 件数分のユーザー情報を受け取り、
    各ユーザーの名前・役職・部署・基本勤務時間を一括更新する。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    try:
        user_count = int(request.form.get("user_count", 0))
    except ValueError:
        flash("更新データが不正です", "warning")
        return _redirect_dashboard()

    login_role: str = session.get("user_role", "")
    login_dept: str = session.get("user_dept", "")

    error_ids: list[str] = []
    skipped_count: int = 0
    for i in range(user_count):
        raw_id = request.form.get(f"uid_{i}", "")
        name: str = request.form.get(f"name_{i}", "").strip()
        role: str = request.form.get(f"role_{i}", "ユーザー").strip()
        dept: str = request.form.get(f"dept_{i}", "").strip()
        try:
            user_id = int(raw_id)
            std_hours = float(request.form.get(f"std_hours_{i}", 8.0))
        except ValueError:
            continue

        if not name:
            continue

        # 管理職は自部署以外のユーザーを更新不可
        if not is_master(login_role):
            target_user = get_user_by_id(user_id)
            if target_user is None or target_user.get("dept") != login_dept:
                skipped_count += 1
                continue

        success = update_user(user_id, name, role, dept, std_hours)
        if not success:
            error_ids.append(raw_id)

    if error_ids:
        flash(f"一部のユーザー更新に失敗しました（ID: {', '.join(error_ids)}）", "warning")
    elif skipped_count > 0:
        flash(f"権限範囲外のユーザー {skipped_count} 件はスキップされました", "warning")
    else:
        record_operation(ACTION_USER_UPDATE, f"user_count={user_count}")
        flash("ユーザー情報を一括更新しました", "success")

    return _redirect_dashboard()


@admin_bp.route("/depts/add", methods=["POST"])
def add_dept_route() -> str:
    """部署を追加する（マスタのみ）。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    if not is_master(session.get("user_role", "")):
        abort(403)
    dept_name: str = request.form.get("dept_name", "").strip()
    try:
        display_order = int(request.form.get("display_order", 0))
    except ValueError:
        display_order = 0
    if not dept_name:
        flash("部署名を入力してください", "warning")
        return _redirect_dashboard()
    try:
        add_dept(dept_name, display_order)
        flash(f"部署「{dept_name}」を追加しました", "success")
    except Exception:
        logger.exception("部署追加に失敗しました (dept_name=%s)", dept_name)
        flash("部署の追加に失敗しました（名前が重複している可能性があります）", "warning")
    return _redirect_dashboard()


@admin_bp.route("/depts/delete/<int:dept_id>", methods=["POST"])
def delete_dept_route(dept_id: int) -> str:
    """部署を削除する（マスタのみ・使用中は削除不可）。

    Args:
        dept_id (int): 削除対象の部署ID。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    if not is_master(session.get("user_role", "")):
        abort(403)
    success = delete_dept(dept_id)
    if success:
        flash("部署を削除しました", "success")
    else:
        flash("この部署には所属ユーザーがいるため削除できません", "warning")
    return _redirect_dashboard()


@admin_bp.route("/depts/update/<int:dept_id>", methods=["POST"])
def update_dept_route(dept_id: int) -> str:
    """部署名・表示順を更新する（マスタのみ）。

    Args:
        dept_id (int): 更新対象の部署ID。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    if not is_master(session.get("user_role", "")):
        abort(403)
    dept_name: str = request.form.get("dept_name", "").strip()
    try:
        display_order = int(request.form.get("display_order", 0))
    except ValueError:
        display_order = 0
    if not dept_name:
        flash("部署名を入力してください", "warning")
        return _redirect_dashboard()
    update_dept(dept_id, dept_name, display_order)
    flash("部署情報を更新しました", "success")
    return _redirect_dashboard()


@admin_bp.route("/assignments/save", methods=["POST"])
def save_assignments() -> str:
    """担当メンバーの上長（manager_id）を一括保存する。

    フォームから manager_{user_id} の形式で上長IDを受け取り、
    各ユーザーの manager_id を更新する。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    login_role: str = session.get("user_role", "")
    if not is_privileged(login_role):
        abort(403)

    login_id: int = int(session.get("user_id", 0))
    login_dept: str = session.get("user_dept", "")

    # 更新可能なユーザーIDセット（スコープ制限: マスタ・管理職ともに自部署のみ）
    all_u = get_all_users(dept_filter=login_dept if login_dept else None)
    allowed_ids: set[int] = {
        u["id"] for u in all_u
        if u.get("role") == "ユーザー" and u.get("dept") == login_dept
    }

    updated = 0
    for key, val in request.form.items():
        if not key.startswith("manager_"):
            continue
        try:
            target_uid = int(key[len("manager_"):])
        except ValueError:
            continue
        if target_uid not in allowed_ids:
            continue
        manager_id: int | None = int(val) if val else None
        save_user_manager(target_uid, manager_id)
        updated += 1

    flash(f"担当メンバー設定を保存しました（{updated}件）", "success")
    return _redirect_dashboard()


@admin_bp.route("/users/reorder", methods=["POST"])
def reorder_users() -> str:
    """ユーザーの表示順を更新する（マスタのみ）。

    フォームから order_0, order_1, ... の形式でユーザーIDを受け取り、
    その順番で display_order を更新する。

    Returns:
        str: 管理者ダッシュボードへのリダイレクトレスポンス。
    """
    if not is_master(session.get("user_role", "")):
        abort(403)

    user_ids: list[int] = []
    i = 0
    while True:
        raw = request.form.get(f"order_{i}")
        if raw is None:
            break
        try:
            user_ids.append(int(raw))
        except ValueError:
            pass
        i += 1

    if not user_ids:
        flash("並び順データが取得できませんでした", "warning")
        return _redirect_dashboard()

    update_users_order(user_ids)
    flash("ユーザーの並び順を保存しました", "success")
    return _redirect_dashboard()


@admin_bp.route("/logs")
def operation_logs() -> str:
    """操作ログ一覧画面（管理者専用）。

    Returns:
        str: 操作ログ一覧のHTMLレスポンス。
    """
    page = max(1, int(request.args.get("page", 1)))
    action_type = request.args.get("action_type", "")
    per_page = 50
    offset = (page - 1) * per_page
    logs = get_operation_logs(limit=per_page, offset=offset, action_type=action_type)
    total = count_operation_logs(action_type=action_type)
    total_pages = max(1, (total + per_page - 1) // per_page)
    action_types = [
        "LOGIN", "LOGOUT", "SCHEDULE_SAVE", "DAILY_SAVE",
        "USER_ADD", "USER_UPDATE", "USER_DELETE", "EXPORT",
    ]
    return render_template(
        "admin_logs.html",
        logs=logs,
        page=page,
        total_pages=total_pages,
        total=total,
        action_type=action_type,
        action_types=action_types,
    )

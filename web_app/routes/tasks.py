"""作業マスタ管理（一覧・追加・削除・区分管理）のルートを提供するBlueprintモジュール。"""
from __future__ import annotations

from flask import (
    Blueprint,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

from ..models import (
    add_task,
    delete_task,
    get_task_master,
    update_task_order,
    get_all_categories,
    get_all_subcategories,
    add_category,
    add_subcategory,
    delete_category,
    delete_subcategory,
)

tasks_bp = Blueprint("tasks", __name__, url_prefix="/tasks")


def _require_login() -> None:
    """ログイン済みかどうかを確認する。

    未ログインの場合はログインページへリダイレクトする。

    Returns:
        None
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    return None


@tasks_bp.route("/", endpoint="task_list")
def task_list() -> str:
    """作業マスタ一覧ページを表示する。

    ログインユーザーの作業マスタを取得してテンプレートへ渡す。

    Returns:
        str: 作業マスタ一覧のHTMLレスポンス、または未ログイン時はリダイレクト。
    """
    redir = _require_login()
    if redir is not None:
        return redir

    user_id: int = session["user_id"]
    user_name: str = session["user_name"]
    login_role: str = session.get("user_role", "")
    task_master = get_task_master(user_id)
    # 全ユーザーに区分データを渡す（登録フォームで使用）
    categories = get_all_categories()
    all_subcategories = get_all_subcategories()
    return render_template(
        "tasks.html",
        task_master=task_master,
        user_name=user_name,
        login_role=login_role,
        categories=categories,
        all_subcategories=all_subcategories,
    )


@tasks_bp.route("/add", methods=["POST"])
def add() -> str:
    """作業マスタへ新しい作業名を追加する。

    フォームから受け取った `task_name` を検証し、重複チェック後に登録する。

    Returns:
        str: 作業マスタ一覧ページへのリダイレクトレスポンス。
    """
    redir = _require_login()
    if redir is not None:
        return redir

    task_name: str = request.form.get("task_name", "").strip()
    if not task_name:
        flash("作業名を入力してください", "warning")
        return redirect(url_for("tasks.task_list"))

    # デフォルト作業時間を取得・バリデートする
    try:
        default_hours = float(request.form.get("default_hours", 0.0) or 0.0)
    except ValueError:
        default_hours = 0.0

    # 大区分・中区分IDを取得・バリデートする
    try:
        category_id: int | None = int(request.form.get("category_id") or 0) or None
    except ValueError:
        category_id = None
    try:
        subcategory_id: int | None = int(request.form.get("subcategory_id") or 0) or None
    except ValueError:
        subcategory_id = None

    user_id: int = session["user_id"]
    success: bool = add_task(user_id, task_name, default_hours, category_id, subcategory_id)

    if success:
        flash(f"「{task_name}」を追加しました", "success")
    else:
        flash("その作業名はすでに登録されています", "warning")

    return redirect(url_for("tasks.task_list"))


@tasks_bp.route("/delete/<int:task_id>", methods=["POST"])
def delete(task_id: int) -> str:
    """指定した作業IDの作業を削除する。

    Args:
        task_id (int): 削除対象の作業マスタID。

    Returns:
        str: 作業マスタ一覧ページへのリダイレクトレスポンス。
    """
    redir = _require_login()
    if redir is not None:
        return redir

    user_id: int = session["user_id"]
    delete_task(task_id, user_id)
    return redirect(url_for("tasks.task_list"))


@tasks_bp.route("/move/<int:task_id>/<direction>", methods=["POST"])
def move(task_id: int, direction: str) -> str:
    """作業の表示順を1つ上または下に移動する。

    Args:
        task_id (int): 移動対象の作業マスタID。
        direction (str): 移動方向（'up' または 'down'）。

    Returns:
        str: 作業マスタ一覧ページへのリダイレクトレスポンス。
    """
    redir = _require_login()
    if redir is not None:
        return redir

    user_id: int = session["user_id"]
    tasks = get_task_master(user_id)

    # 対象タスクのインデックスを取得
    idx = next((i for i, t in enumerate(tasks) if t["id"] == task_id), None)
    if idx is None:
        return redirect(url_for("tasks.task_list"))

    if direction == "up" and idx > 0:
        swap_idx = idx - 1
    elif direction == "down" and idx < len(tasks) - 1:
        swap_idx = idx + 1
    else:
        return redirect(url_for("tasks.task_list"))

    # 表示順を入れ替える（現在のdisplay_orderを基準に連番で振り直し）
    task_ids = [t["id"] for t in tasks]
    task_ids[idx], task_ids[swap_idx] = task_ids[swap_idx], task_ids[idx]
    for order, tid in enumerate(task_ids):
        update_task_order(tid, user_id, order)

    return redirect(url_for("tasks.task_list"))


@tasks_bp.route("/swap-order", methods=["POST"])
def swap_order():
    """2つの作業のdisplay_orderを入れ替える（AJAX対応・JSON応答）。

    Returns:
        Response: 成功時は {"ok": true}、失敗時は {"ok": false} のJSON。
    """
    redir = _require_login()
    if redir is not None:
        return jsonify(ok=False), 401

    user_id: int = session["user_id"]
    data = request.get_json(silent=True) or {}
    task_id_a = data.get("task_id_a")
    task_id_b = data.get("task_id_b")
    if not task_id_a or not task_id_b:
        return jsonify(ok=False), 400

    tasks = get_task_master(user_id)
    order_map = {t["id"]: i for i, t in enumerate(tasks)}

    if task_id_a not in order_map or task_id_b not in order_map:
        return jsonify(ok=False), 404

    # 2つのタスクのdisplay_orderを入れ替え
    update_task_order(task_id_a, user_id, order_map[task_id_b])
    update_task_order(task_id_b, user_id, order_map[task_id_a])

    return jsonify(ok=True)


# ---------------------------------------------------------------------------
# 作業区分管理（管理職・マスタ専用）
# ---------------------------------------------------------------------------


def _require_privileged() -> None:
    """管理職またはマスタ権限かどうかを確認する。

    未ログインまたは権限不足の場合は 403 を返す。

    Returns:
        None
    """
    role: str = session.get("user_role", "")
    if role not in ("管理職", "マスタ"):
        abort(403)


@tasks_bp.route("/categories", endpoint="category_list")
def category_list() -> str:
    """大区分・中区分管理ページを表示する。

    管理職・マスタのみアクセス可。

    Returns:
        str: 区分管理ページのHTMLレスポンス。
    """
    redir = _require_login()
    if redir is not None:
        return redir
    _require_privileged()

    categories = get_all_categories()
    all_subcategories = get_all_subcategories()
    return render_template(
        "tasks/categories.html",
        categories=categories,
        all_subcategories=all_subcategories,
    )


@tasks_bp.route("/categories/add", methods=["POST"], endpoint="category_add")
def category_add() -> str:
    """大区分を追加する。

    フォームから受け取った `name` を検証し、重複チェック後に登録する。

    Returns:
        str: 区分管理ページへのリダイレクトレスポンス。
    """
    redir = _require_login()
    if redir is not None:
        return redir
    _require_privileged()

    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    name: str = request.form.get("name", "").strip()
    if not name:
        flash("大区分名を入力してください", "warning")
        return redirect(url_for("tasks.task_list"))

    success: bool = add_category(name)
    if success:
        flash(f"大区分「{name}」を追加しました", "success")
    else:
        flash("その大区分名はすでに登録されています", "warning")

    return redirect(url_for("tasks.task_list"))


@tasks_bp.route(
    "/categories/delete/<int:cat_id>",
    methods=["POST"],
    endpoint="category_delete",
)
def category_delete(cat_id: int) -> str:
    """指定した大区分を削除する（中区分もCASCADE削除）。

    Args:
        cat_id: 削除対象の大区分ID。

    Returns:
        str: 区分管理ページへのリダイレクトレスポンス。
    """
    redir = _require_login()
    if redir is not None:
        return redir
    _require_privileged()

    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    delete_category(cat_id)
    flash("大区分を削除しました", "success")
    return redirect(url_for("tasks.task_list"))


@tasks_bp.route("/subcategories/add", methods=["POST"], endpoint="subcategory_add")
def subcategory_add() -> str:
    """中区分を追加する。

    フォームから受け取った `category_id` と `name` を検証し、重複チェック後に登録する。

    Returns:
        str: 区分管理ページへのリダイレクトレスポンス。
    """
    redir = _require_login()
    if redir is not None:
        return redir
    _require_privileged()

    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    try:
        category_id: int = int(request.form.get("category_id", 0))
    except ValueError:
        flash("大区分の選択が不正です", "warning")
        return redirect(url_for("tasks.task_list"))

    name: str = request.form.get("name", "").strip()
    if not name or category_id <= 0:
        flash("大区分と中区分名を入力してください", "warning")
        return redirect(url_for("tasks.task_list"))

    success: bool = add_subcategory(category_id, name)
    if success:
        flash(f"中区分「{name}」を追加しました", "success")
    else:
        flash("その中区分名はすでにこの大区分に登録されています", "warning")

    return redirect(url_for("tasks.task_list"))


@tasks_bp.route(
    "/subcategories/delete/<int:sub_id>",
    methods=["POST"],
    endpoint="subcategory_delete",
)
def subcategory_delete(sub_id: int) -> str:
    """指定した中区分を削除する。

    Args:
        sub_id: 削除対象の中区分ID。

    Returns:
        str: 区分管理ページへのリダイレクトレスポンス。
    """
    redir = _require_login()
    if redir is not None:
        return redir
    _require_privileged()

    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    delete_subcategory(sub_id)
    flash("中区分を削除しました", "success")
    return redirect(url_for("tasks.task_list"))

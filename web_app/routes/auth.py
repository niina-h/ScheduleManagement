"""認証関連のルート（ログイン・ログアウト）を提供するBlueprintモジュール。"""
from __future__ import annotations

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
    check_user_password,
    clear_remember_token,
    get_all_users,
    get_user_by_id,
    get_user_by_remember_token,
    set_remember_token,
    set_user_password,
    user_has_password,
)
from ..log_service import record_operation, ACTION_LOGIN, ACTION_LOGOUT
from ..auth_helpers import is_privileged

auth_bp = Blueprint("auth", __name__)


@auth_bp.route("/login", methods=["GET", "POST"])
def login() -> str:
    """ログインページの表示とログイン処理を行う。

    GET  : ユーザー一覧付きのログインフォームを表示する。
    POST : フォームで選択されたユーザーIDを検証し、セッションを設定してスケジュール画面へリダイレクトする。

    Returns:
        str: GETの場合はHTMLレスポンス、POSTの場合はリダイレクトレスポンス。
    """
    remember_token = request.cookies.get("remember_token", "")
    remembered_user = get_user_by_remember_token(remember_token)

    # 記憶トークンによる自動ログイン（GET時・ログアウト直後はスキップ）
    if request.method == "GET" and not request.args.get("logged_out"):
        if remembered_user is not None:
            session["user_id"] = remembered_user["id"]
            session["user_name"] = remembered_user["name"]
            session["user_role"] = remembered_user["role"]
            session.permanent = True
            return redirect(url_for("schedule.weekly"))

    if request.method == "POST":
        raw_id = request.form.get("user_id", "").strip()
        if not raw_id:
            flash("ユーザーを選択してください", "warning")
            return redirect(url_for("auth.login"))

        try:
            user_id = int(raw_id)
        except ValueError:
            flash("無効なユーザーIDです", "warning")
            return redirect(url_for("auth.login"))

        user = get_user_by_id(user_id)
        if user is None:
            flash("ユーザーが見つかりません", "warning")
            return redirect(url_for("auth.login"))

        # 管理職・マスタでパスワードが設定されている場合は検証する
        # ただし有効なremember_tokenがある場合はスキップ
        if is_privileged(user["role"]) and user_has_password(user_id):
            use_token = request.form.get("use_remember_token", "") == "1"
            token_user = get_user_by_remember_token(request.cookies.get("remember_token", ""))
            if use_token and token_user and token_user["id"] == user_id:
                pass  # トークン認証済み・パスワードスキップ
            else:
                password = request.form.get("password", "")
                if not check_user_password(user_id, password):
                    flash("パスワードが正しくありません", "danger")
                    return redirect(url_for("auth.login"))

        session["user_id"] = user["id"]
        session["user_name"] = user["name"]
        session["user_role"] = user["role"]
        session["user_dept"] = user.get("dept", "")
        session.permanent = True

        # ログイン成功時: 管理職・マスタはトークンを保存、一般ユーザーはトークンなし
        resp = redirect(url_for("schedule.weekly"))
        if is_privileged(user["role"]):
            resp.set_cookie("last_user_id", str(user_id), max_age=365 * 24 * 3600)
            token = set_remember_token(user_id)
            resp.set_cookie("remember_token", token, max_age=365 * 24 * 3600, httponly=True)
        else:
            clear_remember_token(user_id)
            resp.delete_cookie("remember_token")
        record_operation(ACTION_LOGIN, f"user_id={user_id}")
        return resp

    users = get_all_users()
    # ログイン画面用: 部署名順 → ロール順（マスタ→管理職→ユーザー）
    _role_order = {"マスタ": 0, "管理職": 1, "ユーザー": 2}
    users.sort(key=lambda u: (u.get("dept", ""), _role_order.get(u.get("role", ""), 9)))

    # パスワードが設定されている管理職・マスタIDのセット（JS用）
    password_required_ids = [
        u["id"] for u in users
        if is_privileged(u["role"]) and user_has_password(u["id"])
    ]

    # クッキーから前回ログインユーザーIDを取得する
    last_user_id = request.cookies.get("last_user_id", "")
    remembered_user_id: int | None = remembered_user["id"] if remembered_user else None
    return render_template(
        "login.html",
        users=users,
        last_user_id=last_user_id,
        password_required_ids=password_required_ids,
        remembered_user_id=remembered_user_id,
    )


@auth_bp.route("/reset_password_and_login", methods=["POST"])
def reset_password_and_login() -> str:
    """ログイン画面からパスワードをリセットして即座にログインする。

    パスワード未設定ユーザー、またはパスワードを忘れたユーザーが
    新しいパスワードを設定して同時にログインできる。

    Returns:
        str: 週間予定へのリダイレクト、またはエラー時はログイン画面へのリダイレクト。
    """
    import re as _re
    raw_id = request.form.get("user_id", "").strip()
    new_pw = request.form.get("new_password", "")
    confirm_pw = request.form.get("new_password_confirm", "")

    if not raw_id:
        flash("ユーザーを選択してください", "warning")
        return redirect(url_for("auth.login"))

    try:
        user_id = int(raw_id)
    except ValueError:
        flash("無効なユーザーIDです", "warning")
        return redirect(url_for("auth.login"))

    user = get_user_by_id(user_id)
    if user is None:
        flash("ユーザーが見つかりません", "warning")
        return redirect(url_for("auth.login"))

    if not is_privileged(user["role"]):
        flash("パスワードの設定は管理職・マスタのみ対象です", "warning")
        return redirect(url_for("auth.login"))

    if not _re.fullmatch(r'\d{4}', new_pw):
        flash("パスワードは4桁の数字で入力してください", "warning")
        return redirect(url_for("auth.login"))

    if new_pw != confirm_pw:
        flash("パスワードと確認用パスワードが一致しません", "warning")
        return redirect(url_for("auth.login"))

    set_user_password(user_id, new_pw)

    session["user_id"] = user["id"]
    session["user_name"] = user["name"]
    session["user_role"] = user["role"]
    session["user_dept"] = user.get("dept", "")
    session.permanent = True

    resp = redirect(url_for("schedule.weekly"))
    resp.set_cookie("last_user_id", str(user_id), max_age=365 * 24 * 3600)
    token = set_remember_token(user_id)
    resp.set_cookie("remember_token", token, max_age=365 * 24 * 3600, httponly=True)
    record_operation(ACTION_LOGIN, f"user_id={user_id} (password_reset)")
    flash("パスワードを設定してログインしました", "success")
    return resp


@auth_bp.route("/logout")
def logout() -> str:
    """ログアウト処理を行い、ログインページへリダイレクトする。

    セッションを全消去してからログインページへ誘導する。

    Returns:
        str: ログインページへのリダイレクトレスポンス。
    """
    record_operation(ACTION_LOGOUT, "")
    session.clear()
    return redirect(url_for("auth.login", logged_out=1))

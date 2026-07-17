"""認証関連のルート（ログイン・ログアウト）を提供するBlueprintモジュール。"""
from __future__ import annotations

import secrets

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
    get_user_by_login_id,
    get_user_by_remember_token,
    set_remember_token,
    set_user_password,
    user_has_password,
)
from ..log_service import record_operation, ACTION_LOGIN, ACTION_LOGOUT
from ..auth_helpers import normalize_role, is_system_admin, is_dept_head

auth_bp = Blueprint("auth", __name__)

# 記憶トークンクッキーの有効期間（1年）
_TOKEN_MAX_AGE = 365 * 24 * 3600

# ── ログイン試行のレート制限（4桁PINの総当たり対策） ──
# login_id ごとに連続失敗回数と最終失敗時刻をメモリに保持する。
# 社内LAN・少人数運用のため、専用テーブルは設けずプロセス内メモリで十分。
_MAX_FAILURES = 5           # 連続失敗の上限
_LOCKOUT_SECONDS = 300      # ロックアウト時間（5分）
_login_failures: dict[str, tuple[int, float]] = {}  # login_id(小文字) → (失敗回数, 最終失敗epoch)


def _is_locked_out(login_id: str) -> int:
    """login_id がロックアウト中なら残り秒数を返す（0=ロックなし）。"""
    import time
    rec = _login_failures.get(login_id.lower())
    if not rec:
        return 0
    count, last = rec
    if count < _MAX_FAILURES:
        return 0
    remaining = int(_LOCKOUT_SECONDS - (time.time() - last))
    if remaining <= 0:
        _login_failures.pop(login_id.lower(), None)  # 期限切れは解除
        return 0
    return remaining


def _record_failure(login_id: str) -> None:
    """ログイン失敗を記録する（連続失敗回数を加算）。"""
    import time
    key = login_id.lower()
    count = _login_failures.get(key, (0, 0.0))[0] + 1
    _login_failures[key] = (count, time.time())


def _clear_failures(login_id: str) -> None:
    """ログイン成功時に失敗カウントをクリアする。"""
    _login_failures.pop(login_id.lower(), None)


def _csrf_ok() -> bool:
    """POSTフォームのCSRFトークンがセッションのものと一致するか検証する。

    Returns:
        bool: 一致すれば True。
    """
    token = request.form.get("csrf_token", "")
    return bool(token) and secrets.compare_digest(token, session.get("csrf_token", ""))


def _establish_session(user: dict) -> None:
    """ログイン成功時にセッションへユーザー情報を設定する。

    セッション固定化攻撃を防ぐため、ログイン前のセッションを破棄してから
    新しい CSRF トークンとユーザー情報を設定する。

    Args:
        user: ユーザー情報の dict（id, name, role, dept を含む）。
    """
    session.clear()  # セッション固定化対策：ログイン前の状態を破棄
    session["csrf_token"] = secrets.token_hex(32)  # トークン再生成
    session["user_id"] = user["id"]
    session["user_name"] = user["name"]
    # 役職は正規化した新値（システム管理者・一般 等）で格納し、以後の判定・表示を統一する。
    session["user_role"] = normalize_role(user["role"])
    session["user_dept"] = user.get("dept", "")
    session.permanent = True


@auth_bp.route("/login", methods=["GET", "POST"])
def login() -> str:
    """ログインページの表示とログイン処理を行う（ログインID＋パスワード方式）。

    GET  : ログインID・パスワードの入力フォームを表示する。
           有効な記憶トークンがあれば自動ログインする（ログアウト直後を除く）。
    POST : ログインIDでユーザーを特定し、パスワードを検証してセッションを張る。

    Returns:
        str: GETはHTMLレスポンス、POSTはリダイレクトレスポンス。
    """
    remember_token = request.cookies.get("remember_token", "")
    remembered_user = get_user_by_remember_token(remember_token)

    # 記憶トークンによる自動ログイン（GET時・ログアウト直後はスキップ）
    if request.method == "GET" and not request.args.get("logged_out"):
        if remembered_user is not None:
            _establish_session(remembered_user)
            return redirect(url_for("schedule.weekly"))

    if request.method == "POST":
        if not _csrf_ok():
            flash("セッションが無効です。もう一度お試しください", "danger")
            return redirect(url_for("auth.login"))

        login_id = request.form.get("login_id", "").strip()
        password = request.form.get("password", "")
        remember = request.form.get("remember_me", "") == "1"

        if not login_id or not password:
            flash("ログインIDとパスワードを入力してください", "warning")
            return redirect(url_for("auth.login"))

        # レート制限：連続失敗が上限に達していればロックアウト
        locked = _is_locked_out(login_id)
        if locked > 0:
            flash(f"ログインの失敗が続いたため、約{(locked + 59) // 60}分間ロックされています", "danger")
            return redirect(url_for("auth.login"))

        user = get_user_by_login_id(login_id)
        # ユーザー不明・パスワード不一致は同一メッセージ（ID存在の推測を防ぐ）
        if user is None or not check_user_password(user["id"], password):
            _record_failure(login_id)
            flash("ログインIDまたはパスワードが正しくありません", "danger")
            return redirect(url_for("auth.login"))

        user_id = user["id"]
        _clear_failures(login_id)
        _establish_session(user)

        resp = redirect(url_for("schedule.weekly"))
        resp.set_cookie("last_login_id", login_id, max_age=_TOKEN_MAX_AGE)
        # 「次回から自動ログイン」チェック時のみトークンを発行（全ロール対象）。
        if remember:
            token = set_remember_token(user_id)
            resp.set_cookie(
                "remember_token", token, max_age=_TOKEN_MAX_AGE,
                httponly=True, samesite="Lax",
            )
        else:
            clear_remember_token(user_id)
            resp.delete_cookie("remember_token")
        record_operation(ACTION_LOGIN, f"user_id={user_id}")
        return resp

    # 前回ログインIDをクッキーから復元（入力欄の初期値に使う）
    last_login_id = request.cookies.get("last_login_id", "")
    return render_template("login.html", last_login_id=last_login_id)


@auth_bp.route("/reset_password_and_login", methods=["POST"])
def reset_password_and_login() -> str:
    """初回パスワード設定を行い、そのままログインする（全ロール対象）。

    パスワード未設定のユーザーが、ログインIDと現行パスワードの確認を経ずに
    初回パスワードを設定してログインする。既にパスワードが設定済みの場合は拒否する
    （パスワード変更は管理画面またはログイン後の機能で行う）。

    Returns:
        str: 週間予定へのリダイレクト、またはエラー時はログイン画面へのリダイレクト。
    """
    import re as _re
    if not _csrf_ok():
        flash("セッションが無効です。もう一度お試しください", "danger")
        return redirect(url_for("auth.login"))

    login_id = request.form.get("login_id", "").strip()
    new_pw = request.form.get("new_password", "")
    confirm_pw = request.form.get("new_password_confirm", "")

    if not login_id:
        flash("ログインIDを入力してください", "warning")
        return redirect(url_for("auth.login"))

    user = get_user_by_login_id(login_id)
    if user is None:
        flash("ログインIDが正しくありません", "warning")
        return redirect(url_for("auth.login"))

    # 既にパスワード設定済みなら初回設定は不可（乗っ取り防止）。
    if user_has_password(user["id"]):
        flash("このユーザーは既にパスワードが設定されています。ログインしてください", "warning")
        return redirect(url_for("auth.login"))

    if not _re.fullmatch(r"\d{4}", new_pw):
        flash("パスワードは4桁の数字で入力してください", "warning")
        return redirect(url_for("auth.login"))

    if new_pw != confirm_pw:
        flash("パスワードと確認用パスワードが一致しません", "warning")
        return redirect(url_for("auth.login"))

    user_id = user["id"]
    set_user_password(user_id, new_pw)
    _establish_session(user)

    resp = redirect(url_for("schedule.weekly"))
    resp.set_cookie("last_login_id", login_id, max_age=_TOKEN_MAX_AGE)
    token = set_remember_token(user_id)
    resp.set_cookie(
        "remember_token", token, max_age=_TOKEN_MAX_AGE,
        httponly=True, samesite="Lax",
    )
    record_operation(ACTION_LOGIN, f"user_id={user_id} (initial_password_set)")
    flash("パスワードを設定してログインしました", "success")
    return resp


@auth_bp.route("/logout")
def logout() -> str:
    """ログアウト処理を行い、ログインページへリダイレクトする。

    セッションを全消去し、記憶トークンをDB・クッキー両方から失効させてから
    ログインページへ誘導する。

    Returns:
        str: ログインページへのリダイレクトレスポンス。
    """
    record_operation(ACTION_LOGOUT, "")
    # 記憶トークンをDBから失効（自動ログインを確実に無効化）。
    user_id = session.get("user_id")
    if user_id:
        clear_remember_token(int(user_id))
    session.clear()
    resp = redirect(url_for("auth.login", logged_out=1))
    resp.delete_cookie("remember_token")
    return resp


def get_switchable_depts() -> list[str]:
    """ログイン中ユーザーが切り替え可能な所属名の一覧を返す。

    - システム管理者: 全部署（dept_master）。
    - 所属長: 自分の担当所属（user_affiliation）。
    - その他: 空。

    Returns:
        list[str]: 切替可能な所属名のリスト。
    """
    from ..models import get_all_depts, get_user_affiliations
    role = session.get("user_role", "")
    if is_system_admin(role):
        return [d["dept_name"] for d in get_all_depts()]
    if is_dept_head(role):
        uid = session.get("user_id")
        return get_user_affiliations(int(uid)) if uid else []
    return []


@auth_bp.route("/switch_dept")
def switch_dept() -> str:
    """システム管理者・所属長が操作対象の所属を切り替える。

    指定所属を session["active_dept"] に保存する。切替可能な所属以外は無視する。
    'all'（空）を指定するとシステム管理者は全所属表示へ戻る。

    Returns:
        str: 直前のページ、なければ週間予定へのリダイレクト。
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    dept = request.args.get("dept", "").strip()
    allowed = get_switchable_depts()
    if dept == "" or dept.lower() == "all":
        session.pop("active_dept", None)  # 全所属表示へ
    elif dept in allowed:
        session["active_dept"] = dept
    # それ以外（許可外）は無視して現状維持
    return redirect(request.referrer or url_for("schedule.weekly"))

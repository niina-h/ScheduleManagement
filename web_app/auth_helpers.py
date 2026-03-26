"""web_app/auth_helpers.py - 役職・権限チェックヘルパー

役職体系:
  マスタ   : 最高権限。全ユーザー参照可。他管理職のパスワード設定・部署マスタ管理可。
  管理職   : 自部署のメンバーのみ参照・編集可。
  ユーザー : 自分自身のデータのみ参照・編集可。
"""
from __future__ import annotations

# 権限あり役職のセット
_PRIVILEGED_ROLES: frozenset[str] = frozenset({"管理職", "マスタ"})


def is_privileged(role: str) -> bool:
    """管理職またはマスタであれば True を返す。

    Args:
        role: ユーザーの役職文字列（session.get("user_role") の値）。

    Returns:
        bool: 管理権限を持つ場合 True。
    """
    return role in _PRIVILEGED_ROLES


def is_master(role: str) -> bool:
    """マスタ役職であれば True を返す。

    Args:
        role: ユーザーの役職文字列。

    Returns:
        bool: マスタの場合 True。
    """
    return role == "マスタ"


def can_access_user(login_user: dict, target_user: dict) -> bool:
    """ログインユーザーが対象ユーザーを参照・編集できるか返す。

    マスタは全ユーザー（マスタ以外）にアクセス可。
    管理職は、対象ユーザーの manager_id が自分のIDと一致する場合にアクセス可。
    manager_id が未設定の場合は同一部署であればアクセス可（後方互換）。
    一般ユーザーはこの関数を呼び出す前に弾くこと。

    Args:
        login_user: ログインユーザーの dict（id, role, dept を含む）。
        target_user: 参照対象ユーザーの dict（manager_id, dept を含む）。

    Returns:
        bool: アクセス可能な場合 True。
    """
    role: str = login_user.get("role", "")
    if not is_privileged(role):
        return False
    # マスタは自分自身＋全部署のマスタ以外全員にアクセス可
    if role == "マスタ":
        if target_user.get("id") == login_user.get("id"):
            return True
        if target_user.get("role", "") == "マスタ":
            return False
        return True
    # 管理職: manager_id が設定されていれば自分のIDと一致するか確認
    target_manager_id = target_user.get("manager_id")
    if target_manager_id is not None:
        return target_manager_id == login_user.get("id")
    # manager_id 未設定の場合のフォールバック: 同一部署チェック
    if target_user.get("role", "") == "マスタ":
        return False
    return (login_user.get("dept") or "") == (target_user.get("dept") or "")


def can_set_password_for(operator_role: str, target_role: str) -> bool:
    """パスワード設定の可否を返す。

    マスタは全役職に設定可。
    管理職はユーザー・管理職（自分含む）に設定可。マスタには設定不可。

    Args:
        operator_role: 操作者の役職。
        target_role: 設定対象ユーザーの役職。

    Returns:
        bool: 設定可能な場合 True。
    """
    if is_master(operator_role):
        return True
    if operator_role == "管理職":
        return target_role in ("ユーザー", "管理職")
    return False

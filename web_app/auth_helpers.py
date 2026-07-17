"""web_app/auth_helpers.py - 役職・権限チェックヘルパー

役職体系（4段階）:
  システム管理者 : 最高権限。全所属を参照可。所属を切り替えて所属長と同じ操作ができる。
                   部署マスタ管理・ユーザー並び替え等のシステム設定も可能。
  所属長         : 自分に設定された所属（複数可）のみ参照。その所属の全ユーザーを
                   切替・代理修正できる。他所属には一切関与しない。
  管理職         : 自部署（直属部下）のメンバーのみ参照・編集可。
  一般           : 自分自身のデータのみ参照・編集可。

旧役職（マスタ／ユーザー）は normalize_role() で新役職へ正規化するため、
DB・セッションに旧値が残っていても全判定関数が正しく動作する（後方互換）。
"""
from __future__ import annotations

# ── 新しい正規の役職値（DB・セッションに格納する値） ──
ROLE_SYSTEM_ADMIN = "システム管理者"   # 旧「マスタ」
ROLE_DEPT_HEAD = "所属長"              # 新設
ROLE_MANAGER = "管理職"                # 現状維持
ROLE_GENERAL = "一般"                  # 旧「ユーザー」

# 旧役職文字列 → 新役職文字列の対応（後方互換・DB変換の両方で使用）
_LEGACY_ROLE_MAP: dict[str, str] = {
    "マスタ": ROLE_SYSTEM_ADMIN,
    "ユーザー": ROLE_GENERAL,
    # 「管理職」は不変
}

# 管理権限を持つ役職（一般のみ False）
_PRIVILEGED_ROLES: frozenset[str] = frozenset(
    {ROLE_SYSTEM_ADMIN, ROLE_DEPT_HEAD, ROLE_MANAGER}
)


def normalize_role(role: str) -> str:
    """旧役職文字列を新役職文字列へ正規化する。

    未知の値・新役職はそのまま返す。DB／セッションに旧値（マスタ・ユーザー）が
    残っていても、判定関数の入口でこれを通すことで一貫した判定ができる。

    Args:
        role: 役職文字列（旧・新どちらでも可）。

    Returns:
        str: 正規化後の役職文字列。
    """
    return _LEGACY_ROLE_MAP.get(role or "", role or "")


def is_system_admin(role: str) -> bool:
    """システム管理者（旧「マスタ」）であれば True を返す。

    Args:
        role: ユーザーの役職文字列。

    Returns:
        bool: システム管理者の場合 True。
    """
    return normalize_role(role) == ROLE_SYSTEM_ADMIN


def is_dept_head(role: str) -> bool:
    """所属長であれば True を返す。

    Args:
        role: ユーザーの役職文字列。

    Returns:
        bool: 所属長の場合 True。
    """
    return normalize_role(role) == ROLE_DEPT_HEAD


def is_manager(role: str) -> bool:
    """管理職であれば True を返す。

    Args:
        role: ユーザーの役職文字列。

    Returns:
        bool: 管理職の場合 True。
    """
    return normalize_role(role) == ROLE_MANAGER


def is_privileged(role: str) -> bool:
    """管理権限（システム管理者・所属長・管理職）を持てば True を返す。

    一般のみ False。既存の呼び出し互換を維持する（意味は「管理職以上」から
    「一般以外」へ拡張され、所属長も権限ありとして扱う）。

    Args:
        role: ユーザーの役職文字列（session.get("user_role") の値）。

    Returns:
        bool: 管理権限を持つ場合 True。
    """
    return normalize_role(role) in _PRIVILEGED_ROLES


def is_master(role: str) -> bool:
    """[非推奨] is_system_admin へのエイリアス（旧コード互換のため残置）。

    Args:
        role: ユーザーの役職文字列。

    Returns:
        bool: システム管理者の場合 True。
    """
    return is_system_admin(role)


def can_access_user(login_user: dict, target_user: dict) -> bool:
    """ログインユーザーが対象ユーザーを参照・編集できるか返す。

    システム管理者は全ユーザーにアクセス可。
    所属長は対象ユーザーの所属が自分の担当所属（user_affiliation）に含まれればアクセス可。
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
    # システム管理者は全ユーザーにアクセス可
    if is_system_admin(role):
        return True
    # 所属長: 対象の所属が自分の担当所属に含まれるか（他所属には一切関与しない）
    if is_dept_head(role):
        from .models import get_user_affiliations
        affils = set(get_user_affiliations(login_user.get("id")))
        return (target_user.get("dept") or "") in affils
    # 管理職: manager_id が設定されていれば自分のIDと一致するか確認
    target_manager_id = target_user.get("manager_id")
    if target_manager_id is not None:
        return target_manager_id == login_user.get("id")
    # manager_id 未設定の場合のフォールバック: 同一部署チェック
    if is_system_admin(target_user.get("role", "")):
        return False
    return (login_user.get("dept") or "") == (target_user.get("dept") or "")


def get_effective_depts(
    login_user: dict,
    affiliations: list[str] | None = None,
    active_dept: str | None = None,
) -> list[str] | None:
    """ログインユーザーが操作対象にできる所属名の集合を返す。

    - システム管理者: active_dept 指定時はその1所属。未指定時は None（＝全所属）。
    - 所属長: affiliations（担当所属）すべて。active_dept 指定時はその1所属に絞る。
    - 管理職: [login_user['dept']]（従来の同一部署）。
    - 一般: []（自分のみ）。

    Args:
        login_user: ログインユーザー dict（role, dept を含む）。
        affiliations: 所属長の担当所属リスト（models.get_user_affiliations の結果）。
        active_dept: セッションで選択中の所属（切替状態）。

    Returns:
        list[str] | None: 対象所属名のリスト。None は「全所属（無制限）」の番兵。
    """
    role = login_user.get("role", "")
    if is_system_admin(role):
        return [active_dept] if active_dept else None
    if is_dept_head(role):
        depts = list(affiliations or [])
        if active_dept and active_dept in depts:
            return [active_dept]
        return depts
    if is_manager(role):
        return [login_user.get("dept") or ""]
    return []


def get_active_scope_dept(login_role: str, login_dept: str, active_dept: str | None) -> str | None:
    """画面のデータ絞り込みに使う「実効所属」を返す。

    - システム管理者: active_dept があればその所属、無ければ None（全所属）。
    - 所属長・管理職・一般: login_dept（自分の所属）。※所属長の複数所属は
      get_accessible_users 側で担当所属全体を扱うため、単一所属フィルタが要る画面では
      主所属を返す。

    Args:
        login_role: ログインユーザーの役職。
        login_dept: ログインユーザーの主所属。
        active_dept: セッションで選択中の所属（システム管理者の切替状態）。

    Returns:
        str | None: 絞り込みに使う所属名。None は全所属を意味する。
    """
    if is_system_admin(login_role):
        return active_dept or None
    return login_dept or None


def can_set_password_for(operator_role: str, target_role: str) -> bool:
    """パスワード設定の可否を返す。

    システム管理者は全役職に設定可。
    管理職は一般・管理職（自分含む）に設定可。システム管理者には設定不可。

    Args:
        operator_role: 操作者の役職。
        target_role: 設定対象ユーザーの役職。

    Returns:
        bool: 設定可能な場合 True。
    """
    if is_system_admin(operator_role):
        return True
    if is_manager(operator_role):
        # 一般・管理職に設定可（新旧どちらの文字列でも normalize で吸収）
        return normalize_role(target_role) in (ROLE_GENERAL, ROLE_MANAGER)
    return False

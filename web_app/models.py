"""web_app/models.py - DBアクセス関数群（ドメインモデル層）"""
from __future__ import annotations

import logging
import secrets
import sqlite3
from datetime import date, datetime, timedelta

from werkzeug.security import check_password_hash, generate_password_hash

from .database import get_db

logger = logging.getLogger(__name__)

# 1週間あたりの曜日数（月〜金）
_DAYS = range(5)
# タイムスロット
_SLOTS = ("am", "pm")
# 1スロットあたりの枠数
_SLOT_SIZE = 5


# ---------------------------------------------------------------------------
# ユーザー関連
# ---------------------------------------------------------------------------


def get_all_users(dept_filter: str | None = None) -> list[dict]:
    """全ユーザー一覧を取得する。

    Args:
        dept_filter: 指定時はその部署のユーザーのみ返す。None なら全件。

    Returns:
        list[dict]: ユーザー情報のリスト（id, name, role, dept, std_hours_am, std_hours_pm, std_hours）
    """
    db = get_db()
    if dept_filter is not None:
        rows = db.execute(
            "SELECT id, name, role, dept, std_hours_am, std_hours_pm, std_hours, manager_id "
            "FROM users WHERE dept = ? ORDER BY display_order ASC, id ASC",
            (dept_filter,),
        ).fetchall()
    else:
        rows = db.execute(
            "SELECT id, name, role, dept, std_hours_am, std_hours_pm, std_hours, manager_id "
            "FROM users ORDER BY display_order ASC, id ASC"
        ).fetchall()
    return [dict(row) for row in rows]


def update_users_order(user_ids: list[int]) -> None:
    """ユーザーの表示順を更新する。

    Args:
        user_ids: 表示したい順番に並べたユーザーIDのリスト。
                  リストのインデックスが display_order の値となる。
    """
    db = get_db()
    for order, uid in enumerate(user_ids):
        db.execute(
            "UPDATE users SET display_order = ? WHERE id = ?",
            (order, uid),
        )
    db.commit()


def get_user_by_id(user_id: int) -> dict | None:
    """
    IDでユーザーを取得する。

    Args:
        user_id: ユーザーID

    Returns:
        dict | None: ユーザー情報。存在しない場合はNone
    """
    db = get_db()
    row = db.execute(
        "SELECT id, name, role, dept, std_hours_am, std_hours_pm, std_hours, manager_id "
        "FROM users WHERE id = ?",
        (user_id,),
    ).fetchone()
    return dict(row) if row else None


def get_user_by_name(name: str) -> dict | None:
    """
    名前でユーザーを取得する。

    Args:
        name: ユーザー名

    Returns:
        dict | None: ユーザー情報。存在しない場合はNone
    """
    db = get_db()
    row = db.execute(
        "SELECT id, name, role, dept, std_hours_am, std_hours_pm, std_hours, manager_id "
        "FROM users WHERE name = ?",
        (name,),
    ).fetchone()
    return dict(row) if row else None


def get_direct_reports(manager_id: int) -> list[dict]:
    """直属の部下ユーザー一覧を取得する（manager_id が一致するユーザー）。

    Args:
        manager_id: 上長のユーザーID。

    Returns:
        list[dict]: 直属部下のユーザー情報リスト。
    """
    db = get_db()
    rows = db.execute(
        "SELECT id, name, role, dept, std_hours_am, std_hours_pm, std_hours, manager_id "
        "FROM users WHERE manager_id = ? ORDER BY display_order ASC, id ASC",
        (manager_id,),
    ).fetchall()
    return [dict(row) for row in rows]


def get_accessible_users(login_id: int, login_role: str, login_dept: str) -> list[dict]:
    """ログインユーザーがアクセスできるユーザーリストを返す。

    管理職は直属部下（manager_id=login_id）が設定されていればそれのみ、
    未設定の場合は同一部署（マスタ除く）を返す。
    マスタはマスタ以外の全ユーザーを返す。

    Args:
        login_id: ログインユーザーのID。
        login_role: ログインユーザーの役職。
        login_dept: ログインユーザーの部署。

    Returns:
        list[dict]: アクセス可能なユーザーリスト。
    """
    if login_role == "マスタ":
        # マスタは自分自身＋同一部署全員（他マスタ除く）を返す
        all_dept = get_all_users(dept_filter=login_dept if login_dept else None)
        others = [u for u in all_dept if u.get("role") != "マスタ"]
        self_user = next((u for u in all_dept if u.get("id") == login_id), None)
        if self_user and not any(u.get("id") == login_id for u in others):
            return [self_user] + others
        return others
    if login_role == "管理職":
        all_dept = get_all_users(dept_filter=login_dept if login_dept else None)
        self_user = next((u for u in all_dept if u.get("id") == login_id), None)
        direct = get_direct_reports(login_id)
        if direct:
            members = direct
        else:
            # 直属部下未設定の場合: 同一部署（マスタ除く）
            members = [u for u in all_dept if u.get("role") != "マスタ"]
        # 自分自身をリスト先頭に追加（未含の場合）
        if self_user and not any(u.get("id") == login_id for u in members):
            return [self_user] + members
        return members
    return []


def save_user_manager(user_id: int, manager_id: int | None) -> None:
    """ユーザーの直属の上長（manager_id）を設定する。

    Args:
        user_id: 対象ユーザーのID。
        manager_id: 上長のユーザーID。None の場合は割り当てを解除する。
    """
    db = get_db()
    db.execute(
        "UPDATE users SET manager_id = ? WHERE id = ?",
        (manager_id, user_id),
    )
    db.commit()


def set_user_password(user_id: int, password: str) -> bool:
    """ユーザーのパスワードハッシュを設定する。

    Args:
        user_id: 対象ユーザーID
        password: 平文パスワード

    Returns:
        bool: 更新成功時 True
    """
    db = get_db()
    try:
        hashed = generate_password_hash(password)
        db.execute(
            "UPDATE users SET password_hash = ? WHERE id = ?",
            (hashed, user_id),
        )
        db.commit()
        return True
    except Exception:
        db.rollback()
        logger.exception("パスワード設定中にエラーが発生しました (user_id=%s)", user_id)
        return False


def clear_user_password(user_id: int) -> bool:
    """ユーザーのパスワードを削除する（パスワード不要に戻す）。

    Args:
        user_id: 対象ユーザーID

    Returns:
        bool: 更新成功時 True
    """
    db = get_db()
    try:
        db.execute(
            "UPDATE users SET password_hash = '' WHERE id = ?",
            (user_id,),
        )
        db.commit()
        return True
    except Exception:
        db.rollback()
        logger.exception("パスワード削除中にエラーが発生しました (user_id=%s)", user_id)
        return False


def check_user_password(user_id: int, password: str) -> bool:
    """ユーザーIDとパスワードを検証する。

    Args:
        user_id: ユーザーID
        password: 平文パスワード

    Returns:
        bool: パスワードが正しい場合 True
    """
    db = get_db()
    row = db.execute(
        "SELECT password_hash FROM users WHERE id = ?",
        (user_id,),
    ).fetchone()
    if row is None:
        return False
    hashed: str = row["password_hash"] or ""
    if not hashed:
        return False
    return check_password_hash(hashed, password)


def user_has_password(user_id: int) -> bool:
    """ユーザーにパスワードが設定されているか確認する。

    Args:
        user_id: ユーザーID

    Returns:
        bool: パスワードが設定されている場合 True
    """
    db = get_db()
    row = db.execute(
        "SELECT password_hash FROM users WHERE id = ?",
        (user_id,),
    ).fetchone()
    if row is None:
        return False
    return bool(row["password_hash"])


def add_user(
    name: str,
    role: str,
    dept: str,
    std_hours: float,
) -> bool:
    """
    ユーザーを追加する。

    std_hours_am / std_hours_pm は互換性のため std_hours / 2 を設定する。

    Args:
        name: ユーザー名
        role: ロール
        dept: 部署
        std_hours: 1日あたりの標準勤務時間

    Returns:
        bool: 追加成功時True、重複時False
    """
    db = get_db()
    try:
        # 新規ユーザーは全表示で最後尾に配置（display_order の最大値 + 1）
        max_order = db.execute(
            "SELECT COALESCE(MAX(display_order), 0) FROM users"
        ).fetchone()[0]
        db.execute(
            "INSERT INTO users (name, role, dept, std_hours_am, std_hours_pm, std_hours, display_order) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (name, role, dept, std_hours / 2, std_hours / 2, std_hours, max_order + 1),
        )
        db.commit()
        return True
    except Exception:
        db.rollback()
        return False


def update_user(
    user_id: int,
    name: str,
    role: str,
    dept: str,
    std_hours: float,
) -> bool:
    """ユーザーの全情報（名前・役職・部署・基本勤務時間）を更新する。

    名前が他ユーザーと重複する場合は False を返す。

    Args:
        user_id: 更新対象のユーザーID
        name: 新しいユーザー名
        role: 新しい役職
        dept: 新しい部署
        std_hours: 新しい1日あたりの基本勤務時間

    Returns:
        bool: 更新成功時 True、名前重複などエラー時 False
    """
    db = get_db()
    try:
        db.execute(
            "UPDATE users SET name = ?, role = ?, dept = ?, "
            "std_hours = ?, std_hours_am = ?, std_hours_pm = ? WHERE id = ?",
            (name, role, dept, std_hours, std_hours / 2, std_hours / 2, user_id),
        )
        db.commit()
        return True
    except Exception:
        db.rollback()
        return False


def delete_user(user_id: int) -> None:
    """
    ユーザーを削除する（関連データはCASCADE削除）。

    Args:
        user_id: 削除対象のユーザーID
    """
    db = get_db()
    db.execute("DELETE FROM users WHERE id = ?", (user_id,))
    db.commit()


def update_user_std_hours(
    user_id: int,
    std_hours: float,
) -> None:
    """
    ユーザーの標準勤務時間を更新する。

    std_hours_am / std_hours_pm は互換性のため std_hours / 2 で同時更新する。

    Args:
        user_id: ユーザーID
        std_hours: 1日あたりの標準勤務時間
    """
    db = get_db()
    db.execute(
        "UPDATE users SET std_hours = ?, std_hours_am = ?, std_hours_pm = ? WHERE id = ?",
        (std_hours, std_hours / 2, std_hours / 2, user_id),
    )
    db.commit()


# ---------------------------------------------------------------------------
# タスクマスター関連
# ---------------------------------------------------------------------------


def get_task_master(user_id: int) -> list[dict]:
    """ユーザーのタスクマスターを取得する。

    大区分・中区分名も JOIN して返す。
    display_order昇順、task_name昇順でソートする。

    Args:
        user_id: ユーザーID

    Returns:
        list[dict]: タスク情報のリスト（category_name / subcategory_name を含む）
    """
    db = get_db()
    rows = db.execute(
        "SELECT tm.id, tm.user_id, tm.task_name, tm.display_order, tm.default_hours,"
        " tm.category_id, tc.name AS category_name,"
        " tm.subcategory_id, ts.name AS subcategory_name"
        " FROM task_master tm"
        " LEFT JOIN task_category tc ON tm.category_id = tc.id"
        " LEFT JOIN task_subcategory ts ON tm.subcategory_id = ts.id"
        " WHERE tm.user_id = ?"
        " ORDER BY tm.display_order ASC, tm.task_name ASC",
        (user_id,),
    ).fetchall()
    return [dict(row) for row in rows]


def get_global_task_category_map() -> dict[str, dict[str, str]]:
    """全ユーザー横断でタスク名→区分情報のマップを取得する。

    同名タスクが複数ユーザーに登録されている場合、最初に見つかった
    区分情報を採用する（区分が設定されているもの優先）。

    Returns:
        dict[str, dict[str, str]]: {task_name: {"category_name": ..., "subcategory_name": ...}}
    """
    db = get_db()
    rows = db.execute(
        "SELECT DISTINCT tm.task_name,"
        " tc.name AS category_name,"
        " ts.name AS subcategory_name"
        " FROM task_master tm"
        " LEFT JOIN task_category tc ON tm.category_id = tc.id"
        " LEFT JOIN task_subcategory ts ON tm.subcategory_id = ts.id"
        " WHERE tc.name IS NOT NULL"
        " ORDER BY tm.task_name",
    ).fetchall()
    result: dict[str, dict[str, str]] = {}
    for row in rows:
        name = row["task_name"]
        if name not in result:
            result[name] = {
                "category_name": row["category_name"] or "",
                "subcategory_name": row["subcategory_name"] or "",
            }
    return result


def add_task(
    user_id: int,
    task_name: str,
    default_hours: float = 0.0,
    category_id: int | None = None,
    subcategory_id: int | None = None,
) -> bool:
    """タスクマスターにタスクを追加する。

    Args:
        user_id: ユーザーID
        task_name: タスク名
        default_hours: デフォルト作業時間（省略時は0.0）
        category_id: 大区分ID（省略時はNone）
        subcategory_id: 中区分ID（省略時はNone）

    Returns:
        bool: 追加成功時True、重複時False
    """
    db = get_db()
    try:
        db.execute(
            "INSERT INTO task_master (user_id, task_name, default_hours, category_id, subcategory_id)"
            " VALUES (?, ?, ?, ?, ?)",
            (user_id, task_name, default_hours, category_id or None, subcategory_id or None),
        )
        db.commit()
        return True
    except Exception:
        db.rollback()
        return False


def update_task_order(task_id: int, user_id: int, display_order: int) -> None:
    """タスクの表示順を更新する。

    Args:
        task_id: 更新対象のタスクID
        user_id: 所有者ユーザーID（所有確認用）
        display_order: 新しい表示順
    """
    db = get_db()
    db.execute(
        "UPDATE task_master SET display_order = ? WHERE id = ? AND user_id = ?",
        (display_order, task_id, user_id),
    )
    db.commit()


def delete_task(task_id: int, user_id: int) -> bool:
    """
    タスクマスターからタスクを削除する。

    user_idが一致する場合のみ削除する。

    Args:
        task_id: タスクID
        user_id: ユーザーID（所有確認用）

    Returns:
        bool: 削除成功時True、対象なし時False
    """
    db = get_db()
    cursor = db.execute(
        "DELETE FROM task_master WHERE id = ? AND user_id = ?",
        (task_id, user_id),
    )
    db.commit()
    return cursor.rowcount > 0


# ---------------------------------------------------------------------------
# 週間予定関連
# ---------------------------------------------------------------------------


def _empty_schedule() -> dict:
    """
    空の週間予定構造体を生成する。

    Returns:
        dict: {day(0-4): {'am': [{task_name:'', hours:0.0}×5], 'pm': [同]}}
    """
    return {
        day: {
            slot: [{"task_name": "", "hours": 0.0} for _ in range(_SLOT_SIZE)]
            for slot in _SLOTS
        }
        for day in _DAYS
    }


def get_weekly_schedule(user_id: int, week_start: str) -> dict:
    """
    指定ユーザー・週の週間予定を取得する。

    データが存在しない枠は task_name=''、hours=0.0 で埋める。

    Args:
        user_id: ユーザーID
        week_start: 週開始日（'YYYY-MM-DD'形式）

    Returns:
        dict: {day(0-4): {'am': [{task_name:str, hours:float}×5], 'pm': [同]}}
    """
    db = get_db()
    rows = db.execute(
        "SELECT day_of_week, time_slot, slot_index, task_name, hours "
        "FROM weekly_schedule "
        "WHERE user_id = ? AND week_start = ?",
        (user_id, week_start),
    ).fetchall()

    schedule = _empty_schedule()
    for row in rows:
        day = row["day_of_week"]
        slot = row["time_slot"]
        idx = row["slot_index"]
        if day in _DAYS and slot in _SLOTS and 0 <= idx < _SLOT_SIZE:
            schedule[day][slot][idx] = {
                "task_name": row["task_name"] or "",
                "hours": row["hours"] or 0.0,
            }
    return schedule


def save_weekly_schedule(
    user_id: int,
    week_start: str,
    data: dict,
    updated_by: str = "",
) -> None:
    """
    週間予定を保存する（UPSERT）。

    created_at は既存レコードがあれば維持し、新規なら現在時刻を設定する。
    updated_at は常に現在時刻で上書きする。

    Args:
        user_id: ユーザーID
        week_start: 週開始日（'YYYY-MM-DD'形式）
        data: {day: {'am': [{task_name:str, hours:float}×5], 'pm': [同]}}
        updated_by: 更新者名（省略可）
    """
    db = get_db()
    now = datetime.now().isoformat()

    # 既存レコードのcreated_atをキャッシュ
    existing_rows = db.execute(
        "SELECT day_of_week, time_slot, slot_index, created_at "
        "FROM weekly_schedule "
        "WHERE user_id = ? AND week_start = ?",
        (user_id, week_start),
    ).fetchall()
    created_at_map: dict[tuple[int, str, int], str] = {
        (r["day_of_week"], r["time_slot"], r["slot_index"]): r["created_at"]
        for r in existing_rows
    }

    for day, slots in data.items():
        day_int = int(day)
        for slot, entries in slots.items():
            for idx, entry in enumerate(entries):
                key = (day_int, slot, idx)
                created_at = created_at_map.get(key) or now
                db.execute(
                    "INSERT OR REPLACE INTO weekly_schedule "
                    "(user_id, week_start, day_of_week, time_slot, slot_index, "
                    " task_name, hours, subcategory_name, created_at, updated_at, updated_by) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (
                        user_id,
                        week_start,
                        day_int,
                        slot,
                        idx,
                        entry.get("task_name", ""),
                        entry.get("hours", 0.0),
                        entry.get("subcategory_name", ""),
                        created_at,
                        now,
                        updated_by,
                    ),
                )
    db.commit()


def get_weekly_schedule_meta(user_id: int, week_start: str) -> dict | None:
    """週間予定の最終更新日時・更新者を取得する。

    複数レコードがある場合は最も新しい updated_at を返す。

    Args:
        user_id: ユーザーID
        week_start: 週開始日（'YYYY-MM-DD'形式）

    Returns:
        dict | None: {'updated_at': str, 'updated_by': str} または None
    """
    db = get_db()
    row = db.execute(
        "SELECT updated_at, updated_by FROM weekly_schedule "
        "WHERE user_id = ? AND week_start = ? "
        "ORDER BY updated_at DESC LIMIT 1",
        (user_id, week_start),
    ).fetchone()
    return dict(row) if row else None


def copy_last_week_schedule(user_id: int, week_start: str) -> bool:
    """
    1週間前の週間予定をコピーする。

    Args:
        user_id: ユーザーID
        week_start: コピー先の週開始日（'YYYY-MM-DD'形式）

    Returns:
        bool: コピー成功時True、コピー元データなし時False
    """
    from datetime import timedelta

    target_date = date.fromisoformat(week_start)
    last_week_start = (target_date - timedelta(days=7)).isoformat()

    last_week_data = get_weekly_schedule(user_id, last_week_start)

    # 全枠が空かどうか確認
    has_data = any(
        entry["task_name"] or entry["hours"] > 0.0
        for slots in last_week_data.values()
        for entries in slots.values()
        for entry in entries
    )
    if not has_data:
        return False

    save_weekly_schedule(user_id, week_start, last_week_data)
    return True


def get_all_users_schedule_status(
    week_start: str, dept_filter: str | None = None
) -> list[dict]:
    """全ユーザーの週間予定の入力状況を取得する。

    Args:
        week_start: 週開始日（'YYYY-MM-DD'形式）
        dept_filter: 部署名でフィルタリングする場合に指定。None の場合は全部署。

    Returns:
        list[dict]: [{id, name, dept, role, std_hours_am, std_hours_pm, std_hours,
                      has_schedule:bool, filled_slots:int}]
                    filled_slots は hours > 0 の枠数
    """
    db = get_db()
    if dept_filter:
        users = db.execute(
            "SELECT id, name, dept, role, std_hours_am, std_hours_pm, std_hours "
            "FROM users WHERE dept = ? ORDER BY display_order ASC, id ASC",
            (dept_filter,),
        ).fetchall()
    else:
        users = db.execute(
            "SELECT id, name, dept, role, std_hours_am, std_hours_pm, std_hours "
            "FROM users ORDER BY display_order ASC, id ASC"
        ).fetchall()

    result: list[dict] = []
    for user in users:
        user_id = user["id"]
        schedule_rows = db.execute(
            "SELECT COUNT(*) AS cnt, "
            "SUM(CASE WHEN hours > 0.0 THEN 1 ELSE 0 END) AS filled "
            "FROM weekly_schedule "
            "WHERE user_id = ? AND week_start = ?",
            (user_id, week_start),
        ).fetchone()

        total = schedule_rows["cnt"] or 0
        filled = schedule_rows["filled"] or 0

        result.append(
            {
                "id": user["id"],
                "name": user["name"],
                "dept": user["dept"],
                "role": user["role"],
                "std_hours_am": user["std_hours_am"],
                "std_hours_pm": user["std_hours_pm"],
                "std_hours": user["std_hours"] or 8.0,
                "has_schedule": total > 0,
                "filled_slots": int(filled),
            }
        )
    return result


# ---------------------------------------------------------------------------
# 日次実績関連
# ---------------------------------------------------------------------------


def _empty_daily_result() -> dict:
    """空の日次実績構造体を生成する。

    Returns:
        dict: {'am': [{task_name:'', hours:0.0}×5], 'pm': [同]}
    """
    return {
        slot: [{"task_name": "", "hours": 0.0, "defer_date": "", "is_carryover": 0} for _ in range(_SLOT_SIZE)]
        for slot in _SLOTS
    }


def get_daily_result(user_id: int, date_str: str) -> dict:
    """指定ユーザー・日付の日次実績を取得する。

    データが存在しない枠は task_name=''、hours=0.0 で埋める。

    Args:
        user_id: ユーザーID
        date_str: 日付（'YYYY-MM-DD'形式）

    Returns:
        dict: {'am': [{task_name:str, hours:float}×5], 'pm': [同]}
    """
    db = get_db()
    rows = db.execute(
        "SELECT time_slot, slot_index, task_name, hours, defer_date, is_carryover "
        "FROM daily_result "
        "WHERE user_id = ? AND date = ?",
        (user_id, date_str),
    ).fetchall()

    result = _empty_daily_result()
    for row in rows:
        slot = row["time_slot"]
        idx = row["slot_index"]
        if slot in _SLOTS and 0 <= idx < _SLOT_SIZE:
            result[slot][idx] = {
                "task_name": row["task_name"] or "",
                "hours": row["hours"] or 0.0,
                "defer_date": row["defer_date"] or "",
                "is_carryover": int(row["is_carryover"] or 0),
            }
    return result


def get_daily_result_meta(user_id: int, date_str: str) -> dict | None:
    """日次実績の最終更新日時・更新者を取得する。

    Args:
        user_id: ユーザーID
        date_str: 日付（'YYYY-MM-DD'形式）

    Returns:
        dict | None: {'updated_at': str, 'updated_by': str} または None
    """
    db = get_db()
    row = db.execute(
        "SELECT updated_at, updated_by FROM daily_result "
        "WHERE user_id = ? AND date = ? "
        "ORDER BY updated_at DESC LIMIT 1",
        (user_id, date_str),
    ).fetchone()
    return dict(row) if row else None


def save_daily_result(
    user_id: int,
    date_str: str,
    data: dict,
    updated_by: str = "",
) -> None:
    """日次実績を保存する（UPSERT）。

    Args:
        user_id: ユーザーID
        date_str: 日付（'YYYY-MM-DD'形式）
        data: {'am': [{task_name:str, hours:float}×5], 'pm': [同]}
        updated_by: 更新者名
    """
    import sqlite3 as _sqlite3

    db = get_db()
    now = datetime.now().isoformat()
    try:
        for slot, entries in data.items():
            for idx, entry in enumerate(entries):
                db.execute(
                    "INSERT OR REPLACE INTO daily_result "
                    "(user_id, date, time_slot, slot_index, task_name, hours, "
                    " subcategory_name, defer_date, is_carryover, updated_at, updated_by) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (
                        user_id,
                        date_str,
                        slot,
                        idx,
                        entry.get("task_name", ""),
                        entry.get("hours", 0.0),
                        entry.get("subcategory_name", ""),
                        entry.get("defer_date", ""),
                        int(entry.get("is_carryover", 0)),
                        now,
                        updated_by,
                    ),
                )
        db.commit()
    except _sqlite3.DatabaseError as exc:
        db.rollback()
        logger.error("日次実績の保存に失敗しました: %s", exc)
        raise


def get_daily_comment(user_id: int, date_str: str) -> dict:
    """日次コメント（振り返り・対策）を取得する。

    Args:
        user_id: ユーザーID
        date_str: 日付（'YYYY-MM-DD'形式）

    Returns:
        dict: {'reflection': str, 'action': str, 'updated_at': str, 'updated_by': str}
              レコードが存在しない場合は空文字列を返す。
    """
    db = get_db()
    row = db.execute(
        "SELECT reflection, action, admin_comment, updated_at, updated_by "
        "FROM daily_comment "
        "WHERE user_id = ? AND date = ?",
        (user_id, date_str),
    ).fetchone()
    if row:
        return {
            "reflection": row["reflection"] or "",
            "action": row["action"] or "",
            "admin_comment": row["admin_comment"] or "",
            "updated_at": row["updated_at"] or "",
            "updated_by": row["updated_by"] or "",
        }
    return {"reflection": "", "action": "", "admin_comment": "", "updated_at": "", "updated_by": ""}


def save_daily_comment(
    user_id: int,
    date_str: str,
    reflection: str,
    action: str,
    updated_by: str = "",
) -> None:
    """日次コメント（振り返り・対策）を保存する（UPSERT）。

    admin_comment は上書きしない（別関数で管理）。

    Args:
        user_id: ユーザーID
        date_str: 日付（'YYYY-MM-DD'形式）
        reflection: 本日の振り返り
        action: 今後の対策・懸念事項
        updated_by: 更新者名
    """
    import sqlite3 as _sqlite3

    db = get_db()
    now = datetime.now().isoformat()
    try:
        db.execute(
            "INSERT INTO daily_comment "
            "(user_id, date, reflection, action, updated_at, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(user_id, date) DO UPDATE SET "
            "reflection=excluded.reflection, action=excluded.action, "
            "updated_at=excluded.updated_at, updated_by=excluded.updated_by",
            (user_id, date_str, reflection, action, now, updated_by),
        )
        db.commit()
    except _sqlite3.DatabaseError as exc:
        db.rollback()
        logger.error("日次コメントの保存に失敗しました: %s", exc)
        raise


def save_admin_comment(
    user_id: int,
    date_str: str,
    admin_comment: str,
    updated_by: str = "",
) -> None:
    """管理者（上長）コメントを保存する（UPSERT）。

    ユーザー本人のコメント（reflection/action）は上書きしない。

    Args:
        user_id: コメント対象ユーザーID
        date_str: 日付（'YYYY-MM-DD'形式）
        admin_comment: 管理者コメント
        updated_by: 更新者名（管理者名）
    """
    import sqlite3 as _sqlite3

    db = get_db()
    now = datetime.now().isoformat()
    try:
        db.execute(
            "INSERT INTO daily_comment "
            "(user_id, date, reflection, action, admin_comment, updated_at, updated_by) "
            "VALUES (?, ?, '', '', ?, ?, ?) "
            "ON CONFLICT(user_id, date) DO UPDATE SET "
            "admin_comment=excluded.admin_comment, "
            "updated_at=excluded.updated_at, updated_by=excluded.updated_by",
            (user_id, date_str, admin_comment, now, updated_by),
        )
        db.commit()
    except _sqlite3.DatabaseError as exc:
        db.rollback()
        logger.error("管理者コメントの保存に失敗しました: %s", exc)
        raise


# ---------------------------------------------------------------------------
# 週間休暇設定関連
# ---------------------------------------------------------------------------


def get_weekly_leave(user_id: int, week_start: str) -> dict:
    """週間休暇設定を取得する。

    Args:
        user_id: ユーザーID
        week_start: 週開始日（'YYYY-MM-DD'形式）

    Returns:
        dict[int, str]: {day_of_week(0-4): leave_type}
                        未設定の曜日は '' を返す。
    """
    db = get_db()
    rows = db.execute(
        "SELECT day_of_week, leave_type FROM weekly_leave "
        "WHERE user_id = ? AND week_start = ?",
        (user_id, week_start),
    ).fetchall()
    result: dict[int, str] = {i: "" for i in range(5)}
    for row in rows:
        day = row["day_of_week"]
        if 0 <= day <= 4:
            result[day] = row["leave_type"] or ""
    return result


def save_weekly_leave(user_id: int, week_start: str, leave_data: dict) -> None:
    """週間休暇設定を保存する（UPSERT）。

    leave_type が空文字の場合はレコードを削除する。

    Args:
        user_id: ユーザーID
        week_start: 週開始日（'YYYY-MM-DD'形式）
        leave_data: {day_of_week(int): leave_type(str)}
    """
    import sqlite3 as _sqlite3

    db = get_db()
    try:
        for day, leave_type in leave_data.items():
            day_int = int(day)
            if leave_type:
                db.execute(
                    "INSERT OR REPLACE INTO weekly_leave "
                    "(user_id, week_start, day_of_week, leave_type) "
                    "VALUES (?, ?, ?, ?)",
                    (user_id, week_start, day_int, leave_type),
                )
            else:
                db.execute(
                    "DELETE FROM weekly_leave "
                    "WHERE user_id = ? AND week_start = ? AND day_of_week = ?",
                    (user_id, week_start, day_int),
                )
        db.commit()
    except _sqlite3.DatabaseError as exc:
        db.rollback()
        logger.error("週間休暇設定の保存に失敗しました: %s", exc)
        raise


def remove_rescheduled_task(user_id: int, target_date_str: str, task_name: str) -> None:
    """リスケ解除時に、振り替え先の週間予定からタスクをクリアする。

    Args:
        user_id: ユーザーID
        target_date_str: リスケ先の日付（'YYYY-MM-DD'形式）
        task_name: 削除するタスク名（例: 「【03/20 リスケ】作業名」）
    """
    target_date = date.fromisoformat(target_date_str)
    day_of_week = target_date.weekday()
    week_start = (target_date - timedelta(days=day_of_week)).isoformat()
    db = get_db()
    db.execute(
        "UPDATE weekly_schedule SET task_name = '', hours = 0 "
        "WHERE user_id = ? AND week_start = ? AND day_of_week = ? AND task_name = ?",
        (user_id, week_start, day_of_week, task_name),
    )
    db.commit()
    logger.info("リスケ解除: weekly_schedule から削除 user=%d task=%s date=%s", user_id, task_name, target_date_str)


def remove_rescheduled_daily_result(user_id: int, target_date_str: str, task_name: str) -> None:
    """リスケ解除時に、振り替え先の日次実績からタスクをクリアする。

    Args:
        user_id: ユーザーID
        target_date_str: リスケ先の日付（'YYYY-MM-DD'形式）
        task_name: クリアするタスク名（例: 「【03/20 リスケ】作業名」）
    """
    db = get_db()
    db.execute(
        "UPDATE daily_result SET task_name = '', hours = 0, defer_date = '', is_carryover = 0 "
        "WHERE user_id = ? AND date = ? AND task_name = ?",
        (user_id, target_date_str, task_name),
    )
    db.commit()
    logger.info("リスケ解除: daily_result からクリア user=%d task=%s date=%s", user_id, task_name, target_date_str)


def defer_task_to_weekly_schedule(
    user_id: int,
    target_date_str: str,
    task_name: str,
    hours: float,
    updated_by: str = "",
) -> None:
    """指定日の週間予定に作業を追加する（後日対応用）。

    空きスロット（task_nameが空のスロット）を午前→午後の順で探して挿入する。
    全スロットが埋まっている場合はスキップする。

    Args:
        user_id: ユーザーID
        target_date_str: 後日対応日（'YYYY-MM-DD'形式）
        task_name: 作業名
        hours: 予定時間
        updated_by: 更新者名
    """
    target_date = date.fromisoformat(target_date_str)
    day_of_week = target_date.weekday()  # 0=月〜4=金
    week_start = (target_date - timedelta(days=day_of_week)).isoformat()

    db = get_db()
    now = datetime.now().isoformat()

    # 同名タスクが対象日の週間予定に既に登録済みであればスキップ（重複防止）
    already = db.execute(
        "SELECT id FROM weekly_schedule "
        "WHERE user_id=? AND week_start=? AND day_of_week=? AND task_name=?",
        (user_id, week_start, day_of_week, task_name),
    ).fetchone()
    if already:
        logger.info(
            "後日対応: 既に登録済みのためスキップ user=%d task=%s date=%s",
            user_id, task_name, target_date_str,
        )
        return

    for slot in ("am", "pm"):
        for idx in range(5):
            existing = db.execute(
                "SELECT task_name FROM weekly_schedule "
                "WHERE user_id=? AND week_start=? AND day_of_week=? AND time_slot=? AND slot_index=?",
                (user_id, week_start, day_of_week, slot, idx),
            ).fetchone()
            if existing is None or not existing["task_name"]:
                db.execute(
                    "INSERT OR REPLACE INTO weekly_schedule "
                    "(user_id, week_start, day_of_week, time_slot, slot_index, "
                    " task_name, hours, created_at, updated_at, updated_by) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (user_id, week_start, day_of_week, slot, idx,
                     task_name, hours, now, now, updated_by),
                )
                db.commit()
                logger.info(
                    "後日対応: user=%d task=%s → %s %s[%d]",
                    user_id, task_name, target_date_str, slot, idx,
                )
                return
    logger.warning(
        "後日対応: 空きスロットなし user=%d date=%s task=%s",
        user_id, target_date_str, task_name,
    )


def get_week_daily_results(user_id: int, week_dates: list[str]) -> dict[int, dict]:
    """週の各日の日次実績を一括取得する。

    Args:
        user_id: ユーザーID
        week_dates: 月〜金の日付文字列リスト（長さ5）

    Returns:
        dict: {day_index(0-4): {'am': [...], 'pm': [...], 'has_result': bool}}
    """
    result: dict[int, dict] = {}
    for i, date_str in enumerate(week_dates):
        data = get_daily_result(user_id, date_str)
        has_result = any(
            entry["task_name"] or entry["hours"] > 0
            for slot in data.values()
            for entry in slot
        )
        result[i] = {**data, "has_result": has_result}
    return result


def get_all_users_daily_status(
    date_str: str, dept_filter: str | None = None
) -> list[dict]:
    """全ユーザーの日次実績入力状況を取得する。

    Args:
        date_str: 日付（'YYYY-MM-DD'形式）
        dept_filter: 部署名でフィルタリングする場合に指定。None の場合は全部署。

    Returns:
        list[dict]: [{id, name, dept, role, has_result:bool, filled_slots:int}]
                    filled_slots は hours > 0 の枠数
    """
    db = get_db()
    if dept_filter:
        users = db.execute(
            "SELECT id, name, dept, role FROM users WHERE dept = ? ORDER BY name ASC",
            (dept_filter,),
        ).fetchall()
    else:
        users = db.execute(
            "SELECT id, name, dept, role FROM users ORDER BY name ASC"
        ).fetchall()

    result: list[dict] = []
    for user in users:
        row = db.execute(
            "SELECT COUNT(*) AS cnt, "
            "SUM(CASE WHEN hours > 0.0 THEN 1 ELSE 0 END) AS filled "
            "FROM daily_result "
            "WHERE user_id = ? AND date = ?",
            (user["id"], date_str),
        ).fetchone()
        total = row["cnt"] or 0
        filled = row["filled"] or 0
        result.append(
            {
                "id": user["id"],
                "name": user["name"],
                "dept": user["dept"],
                "role": user["role"],
                "has_result": total > 0,
                "filled_slots": int(filled),
            }
        )
    return result


def get_pending_carryovers(user_id: int) -> list[dict]:
    """保留中の繰越タスク一覧を取得する。

    Args:
        user_id: ユーザーID

    Returns:
        list[dict]: [{id, task_name, original_date, planned_hours}] 日付昇順
    """
    db = get_db()
    rows = db.execute(
        "SELECT id, task_name, original_date, planned_hours "
        "FROM carryover "
        "WHERE user_id = ? AND resolved = 0 "
        "ORDER BY original_date ASC, task_name ASC",
        (user_id,),
    ).fetchall()
    return [dict(row) for row in rows]


def add_carryover(
    user_id: int, original_date: str, task_name: str, planned_hours: float
) -> None:
    """繰越タスクを登録する。既存レコードが解決済みの場合は未解決に戻す。

    Args:
        user_id: ユーザーID
        original_date: 予定日（'YYYY-MM-DD'形式）
        task_name: タスク名
        planned_hours: 予定時間
    """
    now = datetime.now().isoformat()
    db = get_db()
    db.execute(
        "INSERT OR IGNORE INTO carryover "
        "(user_id, task_name, original_date, planned_hours, created_at) "
        "VALUES (?, ?, ?, ?, ?)",
        (user_id, task_name, original_date, planned_hours, now),
    )
    # 既存レコードが解決済みだった場合も未解決に戻す
    db.execute(
        "UPDATE carryover SET resolved = 0, planned_hours = ? "
        "WHERE user_id = ? AND task_name = ? AND original_date = ?",
        (planned_hours, user_id, task_name, original_date),
    )
    db.commit()


def resolve_carryover_by_slot(
    user_id: int, original_date: str, time_slot: str, slot_index: int
) -> None:
    """指定スロットの繰越タスクを解決済みにする（daily_result の is_carryover と連動）。

    Args:
        user_id: ユーザーID
        original_date: 対象日付（'YYYY-MM-DD'形式）
        time_slot: スロット ('am' または 'pm')
        slot_index: スロット内インデックス（0〜4）
    """
    db = get_db()
    # 対象スロットのタスク名を取得
    row = db.execute(
        "SELECT task_name FROM daily_result "
        "WHERE user_id=? AND date=? AND time_slot=? AND slot_index=?",
        (user_id, original_date, time_slot, slot_index),
    ).fetchone()
    if not row or not row["task_name"]:
        return
    db.execute(
        "UPDATE carryover SET resolved = 1 "
        "WHERE user_id=? AND task_name=? AND original_date=? AND resolved=0",
        (user_id, row["task_name"], original_date),
    )
    db.commit()


# ---------------------------------------------------------------------------
# 部署マスタ関連
# ---------------------------------------------------------------------------

def get_all_depts() -> list[dict]:
    """部署マスタの全件を display_order 昇順で取得する。

    Returns:
        list[dict]: [{id, dept_name, display_order}]
    """
    db = get_db()
    rows = db.execute(
        "SELECT id, dept_name, display_order FROM dept_master ORDER BY display_order ASC, id ASC"
    ).fetchall()
    return [dict(row) for row in rows]


def add_dept(dept_name: str, display_order: int = 0) -> None:
    """部署マスタに新しい部署を登録する。

    Args:
        dept_name: 部署名
        display_order: 表示順
    """
    db = get_db()
    db.execute(
        "INSERT INTO dept_master (dept_name, display_order) VALUES (?, ?)",
        (dept_name.strip(), display_order),
    )
    db.commit()


def update_dept(dept_id: int, dept_name: str, display_order: int = 0) -> None:
    """部署マスタの部署名・表示順を更新する。

    Args:
        dept_id: 更新対象の部署ID
        dept_name: 新しい部署名
        display_order: 表示順
    """
    db = get_db()
    db.execute(
        "UPDATE dept_master SET dept_name = ?, display_order = ? WHERE id = ?",
        (dept_name.strip(), display_order, dept_id),
    )
    db.commit()


def delete_dept(dept_id: int) -> bool:
    """部署マスタから部署を削除する。所属ユーザーがいれば削除不可。

    Args:
        dept_id: 削除対象の部署ID

    Returns:
        bool: 削除成功なら True、所属ユーザーがいれば False。
    """
    db = get_db()
    row = db.execute("SELECT dept_name FROM dept_master WHERE id = ?", (dept_id,)).fetchone()
    if not row:
        return False
    dept_name = row["dept_name"]
    count = db.execute(
        "SELECT COUNT(*) FROM users WHERE dept = ?", (dept_name,)
    ).fetchone()[0]
    if count > 0:
        return False
    db.execute("DELETE FROM dept_master WHERE id = ?", (dept_id,))
    db.commit()
    return True


def dept_has_users(dept_id: int) -> bool:
    """指定部署にユーザーが所属しているか確認する。

    Args:
        dept_id: 部署ID

    Returns:
        bool: 所属ユーザーがいれば True。
    """
    db = get_db()
    row = db.execute("SELECT dept_name FROM dept_master WHERE id = ?", (dept_id,)).fetchone()
    if not row:
        return False
    count = db.execute(
        "SELECT COUNT(*) FROM users WHERE dept = ?", (row["dept_name"],)
    ).fetchone()[0]
    return count > 0


def resolve_carryover_by_id(user_id: int, carryover_id: int) -> None:
    """指定IDの繰越タスクを解決済みにし、daily_result の is_carryover もリセットする。

    Args:
        user_id: ユーザーID（権限確認用）
        carryover_id: 繰越レコードのID
    """
    db = get_db()
    row = db.execute(
        "SELECT original_date, task_name FROM carryover WHERE id = ? AND user_id = ?",
        (carryover_id, user_id),
    ).fetchone()
    if row:
        # daily_result の対象スロットの is_carryover を 0 にリセット
        db.execute(
            "UPDATE daily_result SET is_carryover = 0 "
            "WHERE user_id = ? AND date = ? AND task_name = ?",
            (user_id, row["original_date"], row["task_name"]),
        )
    db.execute(
        "UPDATE carryover SET resolved = 1 WHERE id = ? AND user_id = ?",
        (carryover_id, user_id),
    )
    db.commit()


def resolve_carryovers_by_task(user_id: int, task_name: str, original_date: str = "") -> None:
    """タスク名（と日付）が一致する保留中の繰越を解決済みにする。

    Args:
        user_id: ユーザーID
        task_name: 解決するタスク名
        original_date: 対象日付（指定時はその日付のものだけ解決、省略時は全日付）
    """
    db = get_db()
    if original_date:
        db.execute(
            "UPDATE carryover SET resolved = 1 "
            "WHERE user_id = ? AND task_name = ? AND original_date = ? AND resolved = 0",
            (user_id, task_name, original_date),
        )
    else:
        db.execute(
            "UPDATE carryover SET resolved = 1 "
            "WHERE user_id = ? AND task_name = ? AND resolved = 0",
            (user_id, task_name),
        )
    db.commit()


def set_remember_token(user_id: int) -> str:
    """ログイン記憶用のセキュアトークンを生成してDBに保存する。

    Args:
        user_id: ユーザーID

    Returns:
        生成したトークン文字列（64文字の16進数）
    """
    token = secrets.token_hex(32)
    expiry = (datetime.now() + timedelta(days=365)).isoformat()
    db = get_db()
    db.execute(
        "UPDATE users SET remember_token = ?, remember_token_expiry = ? WHERE id = ?",
        (token, expiry, user_id),
    )
    db.commit()
    return token


def get_user_by_remember_token(token: str) -> dict | None:
    """記憶トークンからユーザーを取得する。期限切れは無効とする。

    Args:
        token: クッキーから取得したトークン文字列

    Returns:
        有効なトークンに対応するユーザー辞書、無効・期限切れは None
    """
    if not token:
        return None
    db = get_db()
    row = db.execute(
        "SELECT * FROM users WHERE remember_token = ?", (token,)
    ).fetchone()
    if row is None:
        return None
    expiry: str = row["remember_token_expiry"] or ""
    if expiry and datetime.fromisoformat(expiry) < datetime.now():
        return None
    return dict(row)


def clear_remember_token(user_id: int) -> None:
    """ユーザーの記憶トークンを削除する（ログアウト時に呼び出す）。

    Args:
        user_id: ユーザーID
    """
    db = get_db()
    db.execute(
        "UPDATE users SET remember_token = '', remember_token_expiry = '' WHERE id = ?",
        (user_id,),
    )
    db.commit()


def get_operation_logs(
    limit: int = 50,
    offset: int = 0,
    action_type: str = "",
) -> list[dict]:
    """操作ログを取得する（管理者画面用）。

    Args:
        limit: 取得件数。
        offset: 取得開始位置。
        action_type: アクション種別フィルタ（空文字なら全件）。

    Returns:
        list[dict]: ログレコードのリスト。
    """
    db = get_db()
    if action_type:
        rows = db.execute(
            "SELECT id, user_id, user_name, action_type, detail, ip_address, created_at"
            " FROM operation_log WHERE action_type = ?"
            " ORDER BY id DESC LIMIT ? OFFSET ?",
            (action_type, limit, offset),
        ).fetchall()
    else:
        rows = db.execute(
            "SELECT id, user_id, user_name, action_type, detail, ip_address, created_at"
            " FROM operation_log ORDER BY id DESC LIMIT ? OFFSET ?",
            (limit, offset),
        ).fetchall()
    return [dict(r) for r in rows]


def count_operation_logs(action_type: str = "") -> int:
    """操作ログの総件数を返す。

    Args:
        action_type: アクション種別フィルタ（空文字なら全件）。

    Returns:
        int: 件数。
    """
    db = get_db()
    if action_type:
        row = db.execute(
            "SELECT COUNT(*) FROM operation_log WHERE action_type = ?",
            (action_type,),
        ).fetchone()
    else:
        row = db.execute("SELECT COUNT(*) FROM operation_log").fetchone()
    return row[0] if row else 0


# ---------------------------------------------------------------------------
# 作業区分（大区分・中区分）関連
# ---------------------------------------------------------------------------


def get_all_categories() -> list[dict]:
    """全大区分を取得する（display_order昇順）。

    Returns:
        list[dict]: 大区分情報のリスト（id, name, display_order）
    """
    db = get_db()
    rows = db.execute(
        "SELECT id, name, display_order FROM task_category ORDER BY display_order ASC, id ASC"
    ).fetchall()
    return [dict(row) for row in rows]


def get_subcategories(category_id: int) -> list[dict]:
    """指定大区分の中区分一覧を取得する（display_order昇順）。

    Args:
        category_id: 大区分ID

    Returns:
        list[dict]: 中区分情報のリスト（id, category_id, name, display_order）
    """
    db = get_db()
    rows = db.execute(
        "SELECT id, category_id, name, display_order FROM task_subcategory"
        " WHERE category_id = ? ORDER BY display_order ASC, id ASC",
        (category_id,),
    ).fetchall()
    return [dict(row) for row in rows]


def get_all_subcategories() -> list[dict]:
    """全中区分を取得する（category_id, display_order昇順）。カテゴリ名も含む。

    task_category と JOIN し、各行に category_id, category_name, id, name,
    display_order を含む。

    Returns:
        list[dict]: 中区分情報のリスト
    """
    db = get_db()
    rows = db.execute(
        "SELECT s.id, s.category_id, c.name AS category_name,"
        " s.name, s.display_order"
        " FROM task_subcategory s"
        " JOIN task_category c ON s.category_id = c.id"
        " ORDER BY s.category_id ASC, s.display_order ASC, s.id ASC"
    ).fetchall()
    return [dict(row) for row in rows]


def add_category(name: str) -> bool:
    """大区分を追加する。重複時はFalseを返す。

    Args:
        name: 大区分名

    Returns:
        bool: 追加成功時True、重複時False
    """
    db = get_db()
    try:
        db.execute(
            "INSERT INTO task_category (name) VALUES (?)",
            (name,),
        )
        db.commit()
        return True
    except sqlite3.IntegrityError:
        db.rollback()
        return False


def add_subcategory(category_id: int, name: str) -> bool:
    """中区分を追加する。重複時はFalseを返す。

    Args:
        category_id: 所属する大区分ID
        name: 中区分名

    Returns:
        bool: 追加成功時True、重複時False
    """
    db = get_db()
    try:
        db.execute(
            "INSERT INTO task_subcategory (category_id, name) VALUES (?, ?)",
            (category_id, name),
        )
        db.commit()
        return True
    except sqlite3.IntegrityError:
        db.rollback()
        return False


def delete_category(category_id: int) -> None:
    """大区分を削除する（中区分もCASCADE削除）。

    Args:
        category_id: 削除対象の大区分ID
    """
    db = get_db()
    db.execute("DELETE FROM task_category WHERE id = ?", (category_id,))
    db.commit()


def delete_subcategory(subcategory_id: int) -> None:
    """中区分を削除する。

    Args:
        subcategory_id: 削除対象の中区分ID
    """
    db = get_db()
    db.execute("DELETE FROM task_subcategory WHERE id = ?", (subcategory_id,))
    db.commit()


def update_category_name(category_id: int, name: str) -> bool:
    """大区分名を更新する。重複時はFalseを返す。

    Args:
        category_id: 更新対象の大区分ID
        name: 新しい大区分名

    Returns:
        bool: 更新成功時True、重複時False
    """
    db = get_db()
    try:
        db.execute(
            "UPDATE task_category SET name = ? WHERE id = ?",
            (name, category_id),
        )
        db.commit()
        return True
    except sqlite3.IntegrityError:
        db.rollback()
        return False


def update_subcategory_name(subcategory_id: int, name: str) -> bool:
    """中区分名を更新する。重複時はFalseを返す。

    Args:
        subcategory_id: 更新対象の中区分ID
        name: 新しい中区分名

    Returns:
        bool: 更新成功時True、重複時False
    """
    db = get_db()
    try:
        db.execute(
            "UPDATE task_subcategory SET name = ? WHERE id = ?",
            (name, subcategory_id),
        )
        db.commit()
        return True
    except sqlite3.IntegrityError:
        db.rollback()
        return False


# ---------------------------------------------------------------------------
# メール設定関連
# ---------------------------------------------------------------------------


def get_mail_setting(role: str) -> dict:
    """指定役職のメール設定を取得する。

    Args:
        role: 役職名（'管理職' または 'マスタ'）

    Returns:
        dict: {role, to_address, cc_address, subject_template, body_template}
              レコードが存在しない場合は空文字列を返す。
    """
    db = get_db()
    row = db.execute(
        "SELECT role, to_address, cc_address, subject_template, body_template"
        " FROM mail_settings WHERE role = ?",
        (role,),
    ).fetchone()
    if row:
        return dict(row)
    return {
        "role": role,
        "to_address": "",
        "cc_address": "",
        "subject_template": "",
        "body_template": "",
    }


def save_mail_setting(
    role: str,
    to_address: str,
    cc_address: str,
    subject_template: str,
    body_template: str,
) -> None:
    """指定役職のメール設定を保存する（INSERT OR REPLACE）。

    Args:
        role: 役職名（'管理職' または 'マスタ'）
        to_address: TO宛先（複数はセミコロン区切り）
        cc_address: CC宛先（複数はセミコロン区切り）
        subject_template: 件名テンプレート
        body_template: 本文テンプレート
    """
    db = get_db()
    db.execute(
        "INSERT OR REPLACE INTO mail_settings"
        " (role, to_address, cc_address, subject_template, body_template)"
        " VALUES (?, ?, ?, ?, ?)",
        (role, to_address, cc_address, subject_template, body_template),
    )
    db.commit()

"""web_app/database.py - SQLiteデータベース接続・初期化ユーティリティ"""
from __future__ import annotations

import json
import logging
import pathlib
import sqlite3

from flask import Flask, g

logger = logging.getLogger(__name__)

# DBスキーマ定義
_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    role TEXT NOT NULL DEFAULT 'ユーザー',
    dept TEXT DEFAULT '',
    std_hours_am REAL NOT NULL DEFAULT 4.0,
    std_hours_pm REAL NOT NULL DEFAULT 4.0,
    std_hours REAL NOT NULL DEFAULT 8.0
);

CREATE TABLE IF NOT EXISTS task_master (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    task_name TEXT NOT NULL,
    display_order INTEGER DEFAULT 0,
    default_hours REAL DEFAULT 0.0,
    UNIQUE(user_id, task_name)
);

CREATE TABLE IF NOT EXISTS weekly_schedule (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    week_start TEXT NOT NULL,
    day_of_week INTEGER NOT NULL,
    time_slot TEXT NOT NULL,
    slot_index INTEGER NOT NULL,
    task_name TEXT DEFAULT '',
    hours REAL DEFAULT 0.0,
    created_at TEXT,
    updated_at TEXT,
    updated_by TEXT DEFAULT '',
    UNIQUE(user_id, week_start, day_of_week, time_slot, slot_index)
);

CREATE TABLE IF NOT EXISTS daily_result (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    time_slot TEXT NOT NULL,
    slot_index INTEGER NOT NULL,
    task_name TEXT DEFAULT '',
    hours REAL DEFAULT 0.0,
    updated_at TEXT,
    updated_by TEXT DEFAULT '',
    UNIQUE(user_id, date, time_slot, slot_index)
);

CREATE TABLE IF NOT EXISTS daily_comment (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    reflection TEXT DEFAULT '',
    action TEXT DEFAULT '',
    updated_at TEXT,
    updated_by TEXT DEFAULT '',
    UNIQUE(user_id, date)
);

CREATE TABLE IF NOT EXISTS weekly_leave (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    week_start TEXT NOT NULL,
    day_of_week INTEGER NOT NULL,
    leave_type TEXT DEFAULT '',
    UNIQUE(user_id, week_start, day_of_week)
);

CREATE TABLE IF NOT EXISTS carryover (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    task_name TEXT NOT NULL,
    original_date TEXT NOT NULL,
    planned_hours REAL DEFAULT 0,
    resolved INTEGER DEFAULT 0,
    created_at TEXT NOT NULL,
    UNIQUE(user_id, task_name, original_date)
);

CREATE TABLE IF NOT EXISTS dept_master (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dept_name TEXT NOT NULL UNIQUE,
    display_order INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS operation_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    user_name TEXT NOT NULL DEFAULT '',
    action_type TEXT NOT NULL,
    detail TEXT DEFAULT '',
    ip_address TEXT DEFAULT '',
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS task_category (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    display_order INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS task_subcategory (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_id INTEGER NOT NULL REFERENCES task_category(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    display_order INTEGER DEFAULT 0,
    UNIQUE(category_id, name)
);

CREATE TABLE IF NOT EXISTS mail_settings (
    role TEXT PRIMARY KEY,
    to_address TEXT DEFAULT '',
    cc_address TEXT DEFAULT '',
    bcc_address TEXT DEFAULT '',
    subject_template TEXT DEFAULT '',
    body_template TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS project_task (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_id INTEGER REFERENCES task_category(id) ON DELETE SET NULL,
    subcategory_id INTEGER REFERENCES task_subcategory(id) ON DELETE SET NULL,
    task_name TEXT NOT NULL,
    description TEXT DEFAULT '',
    start_date TEXT NOT NULL,
    end_date TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT '未着手',
    delay_days INTEGER DEFAULT 0,
    progress INTEGER DEFAULT 0,
    display_order INTEGER DEFAULT 0,
    created_by INTEGER REFERENCES users(id),
    created_at TEXT DEFAULT (datetime('now','localtime')),
    updated_at TEXT DEFAULT (datetime('now','localtime')),
    updated_by TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS company_holiday (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    holiday_date TEXT NOT NULL UNIQUE,
    holiday_name TEXT NOT NULL DEFAULT '',
    created_by INTEGER REFERENCES users(id),
    created_at TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS routine_schedule (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id),
    task_name TEXT NOT NULL DEFAULT '',
    subcategory_name TEXT DEFAULT '',
    default_hours REAL DEFAULT 0.0,
    row_number INTEGER NOT NULL,
    UNIQUE(user_id, row_number)
);

CREATE TABLE IF NOT EXISTS user_survey (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL UNIQUE,
    usability INTEGER NOT NULL,
    merit TEXT DEFAULT '',
    demerit TEXT DEFAULT '',
    requested_feature TEXT DEFAULT '',
    recommend TEXT NOT NULL,
    recommended_dept TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
);
"""

# users.json のパス
_USERS_JSON_PATH = pathlib.Path(__file__).parent.parent / "data" / "users.json"


def get_db() -> sqlite3.Connection:
    """
    Flaskアプリコンテキスト内でSQLite接続を取得する。

    flask.g._database にキャッシュし、同一リクエスト内で
    接続を使い回す。WALモード・外部キー制約を有効化する。

    Returns:
        sqlite3.Connection: データベース接続オブジェクト
    """
    from flask import current_app

    if not hasattr(g, "_database"):
        db_path = current_app.config["DATABASE"]
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        g._database = conn
    return g._database


def close_db(e: BaseException | None = None) -> None:
    """
    アプリコンテキスト終了時にDBコネクションをクローズする。

    Args:
        e: teardown時に渡される例外オブジェクト（使用しない）
    """
    db = g.pop("_database", None)
    if db is not None:
        db.close()


def _migrate_schema(db: sqlite3.Connection) -> None:
    """既存DBに不足しているカラムを追加するマイグレーションを実行する。

    Args:
        db: データベース接続オブジェクト
    """
    # task_master に default_hours / category_id / subcategory_id を追加（なければ）
    cols = {row[1] for row in db.execute("PRAGMA table_info(task_master)").fetchall()}
    if "default_hours" not in cols:
        db.execute("ALTER TABLE task_master ADD COLUMN default_hours REAL DEFAULT 0.0")
        logger.info("task_master.default_hours カラムを追加しました。")
    if "category_id" not in cols:
        db.execute("ALTER TABLE task_master ADD COLUMN category_id INTEGER REFERENCES task_category(id) ON DELETE SET NULL")
        logger.info("task_master.category_id カラムを追加しました。")
    if "subcategory_id" not in cols:
        db.execute("ALTER TABLE task_master ADD COLUMN subcategory_id INTEGER REFERENCES task_subcategory(id) ON DELETE SET NULL")
        logger.info("task_master.subcategory_id カラムを追加しました。")

    # users カラム一覧を一括取得（以降の複数チェックで再利用）
    user_cols = {row[1] for row in db.execute("PRAGMA table_info(users)").fetchall()}

    # users に std_hours を追加（なければ）
    if "std_hours" not in user_cols:
        db.execute("ALTER TABLE users ADD COLUMN std_hours REAL DEFAULT 8.0")
        db.execute("UPDATE users SET std_hours = std_hours_am + std_hours_pm")
        user_cols.add("std_hours")
        logger.info("users.std_hours カラムを追加し、既存データを移行しました。")

    # weekly_schedule に updated_by を追加（なければ）
    cols = {row[1] for row in db.execute("PRAGMA table_info(weekly_schedule)").fetchall()}
    if "updated_by" not in cols:
        db.execute("ALTER TABLE weekly_schedule ADD COLUMN updated_by TEXT DEFAULT ''")
        logger.info("weekly_schedule.updated_by カラムを追加しました。")

    # users に password_hash を追加（なければ）
    if "password_hash" not in user_cols:
        db.execute("ALTER TABLE users ADD COLUMN password_hash TEXT DEFAULT ''")
        user_cols.add("password_hash")
        logger.info("users.password_hash カラムを追加しました。")

    # daily_comment に admin_comment を追加（なければ）
    cols = {row[1] for row in db.execute("PRAGMA table_info(daily_comment)").fetchall()}
    if "admin_comment" not in cols:
        db.execute("ALTER TABLE daily_comment ADD COLUMN admin_comment TEXT DEFAULT ''")
        logger.info("daily_comment.admin_comment カラムを追加しました。")

    # daily_result に defer_date を追加（なければ）
    cols = {row[1] for row in db.execute("PRAGMA table_info(daily_result)").fetchall()}
    if "defer_date" not in cols:
        db.execute("ALTER TABLE daily_result ADD COLUMN defer_date TEXT DEFAULT ''")
        logger.info("daily_result.defer_date カラムを追加しました。")

    # daily_result に is_carryover を追加（なければ）
    cols = {row[1] for row in db.execute("PRAGMA table_info(daily_result)").fetchall()}
    if "is_carryover" not in cols:
        db.execute("ALTER TABLE daily_result ADD COLUMN is_carryover INTEGER DEFAULT 0")
        logger.info("daily_result.is_carryover カラムを追加しました。")

    # users に remember_token / remember_token_expiry を追加（なければ）
    if "remember_token" not in user_cols:
        db.execute("ALTER TABLE users ADD COLUMN remember_token TEXT DEFAULT ''")
        logger.info("users.remember_token カラムを追加しました。")
    if "remember_token_expiry" not in user_cols:
        db.execute("ALTER TABLE users ADD COLUMN remember_token_expiry TEXT DEFAULT ''")
        logger.info("users.remember_token_expiry カラムを追加しました。")

    # dept_master テーブルを作成し、既存 users.dept の値を移行（なければ）
    tables = {row[0] for row in db.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
    if "dept_master" not in tables:
        db.execute("""CREATE TABLE IF NOT EXISTS dept_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dept_name TEXT NOT NULL UNIQUE,
            display_order INTEGER DEFAULT 0
        )""")
        db.execute("""
            INSERT OR IGNORE INTO dept_master (dept_name, display_order)
            SELECT DISTINCT dept, rowid FROM users WHERE dept != '' AND dept IS NOT NULL
        """)
        logger.info("dept_master テーブルを作成し、既存部署データを移行しました。")

    # dept_master に enable_dev_summary を追加（なければ）。日報メールの
    # ＜開発状況＞セクション自体を部門ごとに表示/非表示切替できるようにする
    # フラグ。既定は無効（0）＝現状どおり非表示。
    dm_cols = {row[1] for row in db.execute("PRAGMA table_info(dept_master)").fetchall()}
    if "enable_dev_summary" not in dm_cols:
        db.execute("ALTER TABLE dept_master ADD COLUMN enable_dev_summary INTEGER NOT NULL DEFAULT 0")
        logger.info("dept_master.enable_dev_summary カラムを追加しました。")

    # users.role '管理者' → '管理職' にリネーム
    count = db.execute("SELECT COUNT(*) FROM users WHERE role = '管理者'").fetchone()[0]
    if count > 0:
        db.execute("UPDATE users SET role = '管理職' WHERE role = '管理者'")
        logger.info("users.role '管理者' → '管理職' にリネームしました（%d件）。", count)

    # users.role の4段階ロール移行（冪等）: マスタ→システム管理者, ユーザー→一般。
    # ※ mail_settings.role（'マスタ'/'ユーザー_{uid}' 等のDB主キー値）は役職とは別概念のため
    #   絶対に変換しない。この UPDATE は必ず users テーブルのみに限定する。
    cnt_admin = db.execute("SELECT COUNT(*) FROM users WHERE role = 'マスタ'").fetchone()[0]
    if cnt_admin > 0:
        db.execute("UPDATE users SET role = 'システム管理者' WHERE role = 'マスタ'")
        logger.info("users.role 'マスタ' → 'システム管理者' に移行しました（%d件）。", cnt_admin)
    cnt_general = db.execute("SELECT COUNT(*) FROM users WHERE role = 'ユーザー'").fetchone()[0]
    if cnt_general > 0:
        db.execute("UPDATE users SET role = '一般' WHERE role = 'ユーザー'")
        logger.info("users.role 'ユーザー' → '一般' に移行しました（%d件）。", cnt_general)

    # users に display_order を追加（なければ）
    if "display_order" not in user_cols:
        db.execute("ALTER TABLE users ADD COLUMN display_order INTEGER DEFAULT 0")
        db.execute("UPDATE users SET display_order = id")
        logger.info("users.display_order カラムを追加し、既存ユーザーにid順を設定しました。")

    # users に manager_id を追加（なければ）
    if "manager_id" not in user_cols:
        db.execute("ALTER TABLE users ADD COLUMN manager_id INTEGER")
        logger.info("users.manager_id カラムを追加しました。")

    # operation_log テーブルを作成（なければ）
    tables = {row[0] for row in db.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
    if "operation_log" not in tables:
        db.execute("""CREATE TABLE IF NOT EXISTS operation_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            user_name TEXT NOT NULL DEFAULT '',
            action_type TEXT NOT NULL,
            detail TEXT DEFAULT '',
            ip_address TEXT DEFAULT '',
            created_at TEXT NOT NULL
        )""")
        logger.info("operation_log テーブルを作成しました。")

    # daily_result に subcategory_name を追加（なければ）
    cols = {row[1] for row in db.execute("PRAGMA table_info(daily_result)").fetchall()}
    if "subcategory_name" not in cols:
        db.execute("ALTER TABLE daily_result ADD COLUMN subcategory_name TEXT DEFAULT ''")
        logger.info("daily_result.subcategory_name カラムを追加しました。")

    # weekly_schedule に subcategory_name を追加（なければ）
    cols = {row[1] for row in db.execute("PRAGMA table_info(weekly_schedule)").fetchall()}
    if "subcategory_name" not in cols:
        db.execute("ALTER TABLE weekly_schedule ADD COLUMN subcategory_name TEXT DEFAULT ''")
        logger.info("weekly_schedule.subcategory_name カラムを追加しました。")

    # mail_settings テーブルを追加（なければ）
    tables = {row[0] for row in db.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
    if "mail_settings" not in tables:
        db.execute("""
            CREATE TABLE IF NOT EXISTS mail_settings (
                role TEXT PRIMARY KEY,
                to_address TEXT DEFAULT '',
                cc_address TEXT DEFAULT '',
                bcc_address TEXT DEFAULT '',
                subject_template TEXT DEFAULT '',
                body_template TEXT DEFAULT ''
            )
        """)
        logger.info("mail_settings テーブルを作成しました。")

    # mail_settings に bcc_address カラムを追加（なければ）
    ms_cols = {row[1] for row in db.execute("PRAGMA table_info(mail_settings)").fetchall()}
    if "bcc_address" not in ms_cols:
        db.execute("ALTER TABLE mail_settings ADD COLUMN bcc_address TEXT DEFAULT ''")
        logger.info("mail_settings.bcc_address カラムを追加しました。")

    # project_task テーブルを作成（なければ）
    tables = {row[0] for row in db.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
    if "project_task" not in tables:
        db.execute("""CREATE TABLE IF NOT EXISTS project_task (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category_id INTEGER REFERENCES task_category(id) ON DELETE SET NULL,
            subcategory_id INTEGER REFERENCES task_subcategory(id) ON DELETE SET NULL,
            task_name TEXT NOT NULL,
            description TEXT DEFAULT '',
            assigned_to INTEGER REFERENCES users(id) ON DELETE SET NULL,
            is_milestone INTEGER DEFAULT 0,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT '未着手',
            progress INTEGER DEFAULT 0,
            delay_days INTEGER DEFAULT 0,
            display_order INTEGER DEFAULT 0,
            created_by INTEGER REFERENCES users(id),
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime')),
            updated_by TEXT DEFAULT ''
        )""")
        logger.info("project_task テーブルを作成しました。")

    # project_task にカラムを追加（なければ）
    if "project_task" in tables:
        pt_cols = {row[1] for row in db.execute("PRAGMA table_info(project_task)").fetchall()}
        if "delay_days" not in pt_cols:
            db.execute("ALTER TABLE project_task ADD COLUMN delay_days INTEGER DEFAULT 0")
            logger.info("project_task.delay_days カラムを追加しました。")
        if "assigned_to" not in pt_cols:
            db.execute("ALTER TABLE project_task ADD COLUMN assigned_to INTEGER REFERENCES users(id) ON DELETE SET NULL")
            logger.info("project_task.assigned_to カラムを追加しました。")
        if "is_milestone" not in pt_cols:
            db.execute("ALTER TABLE project_task ADD COLUMN is_milestone INTEGER DEFAULT 0")
            logger.info("project_task.is_milestone カラムを追加しました。")
        if "assigned_to_2" not in pt_cols:
            db.execute("ALTER TABLE project_task ADD COLUMN assigned_to_2 INTEGER REFERENCES users(id) ON DELETE SET NULL")
            logger.info("project_task.assigned_to_2 カラムを追加しました。")
        # 担当者ごとの「週間予定タスク管理反映」対象フラグ（デフォルト=ONで後方互換）
        if "import_to_schedule_1" not in pt_cols:
            db.execute("ALTER TABLE project_task ADD COLUMN import_to_schedule_1 INTEGER NOT NULL DEFAULT 1")
            logger.info("project_task.import_to_schedule_1 カラムを追加しました。")
        if "import_to_schedule_2" not in pt_cols:
            db.execute("ALTER TABLE project_task ADD COLUMN import_to_schedule_2 INTEGER NOT NULL DEFAULT 1")
            logger.info("project_task.import_to_schedule_2 カラムを追加しました。")
        # イベント参加者を複数保存するためのカンマ区切りIDリスト（"5,10,7" など）
        if "event_member_ids" not in pt_cols:
            db.execute("ALTER TABLE project_task ADD COLUMN event_member_ids TEXT NOT NULL DEFAULT ''")
            logger.info("project_task.event_member_ids カラムを追加しました。")

    # project_task にイベント関連カラムを追加（なければ）
    if "project_task" in tables:
        pt_cols2 = {row[1] for row in db.execute("PRAGMA table_info(project_task)").fetchall()}
        if "is_event" not in pt_cols2:
            db.execute("ALTER TABLE project_task ADD COLUMN is_event INTEGER DEFAULT 0")
            logger.info("project_task.is_event カラムを追加しました。")
        if "event_start_time" not in pt_cols2:
            db.execute("ALTER TABLE project_task ADD COLUMN event_start_time TEXT DEFAULT ''")
            logger.info("project_task.event_start_time カラムを追加しました。")
        if "event_end_time" not in pt_cols2:
            db.execute("ALTER TABLE project_task ADD COLUMN event_end_time TEXT DEFAULT ''")
            logger.info("project_task.event_end_time カラムを追加しました。")
        if "planned_hours" not in pt_cols2:
            db.execute("ALTER TABLE project_task ADD COLUMN planned_hours REAL DEFAULT 0.0")
            logger.info("project_task.planned_hours カラムを追加しました。")
        # 親タスクID（ツリー階層。NULL=最上位）。既存機能は本列を参照しないため後方互換。
        if "parent_task_id" not in pt_cols2:
            db.execute("ALTER TABLE project_task ADD COLUMN parent_task_id INTEGER")
            logger.info("project_task.parent_task_id カラムを追加しました。")
        # 親タスクの関係ユーザー（メンバー）。カンマ区切りのユーザーID。ガント入力の権限判定に使用。
        if "member_ids" not in pt_cols2:
            db.execute("ALTER TABLE project_task ADD COLUMN member_ids TEXT NOT NULL DEFAULT ''")
            logger.info("project_task.member_ids カラムを追加しました。")
        # 子タスクの担当者（2名以上に対応）。カンマ区切りのユーザーID。
        # assigned_to/assigned_to_2（先頭2名の互換フィールド）は当面残し、
        # 3名以上が設定された場合のみ本列を使う。
        if "assigned_ids" not in pt_cols2:
            db.execute("ALTER TABLE project_task ADD COLUMN assigned_ids TEXT NOT NULL DEFAULT ''")
            logger.info("project_task.assigned_ids カラムを追加しました。")
        # 親タスク（見出し行）を日報メールの＜開発状況＞集計に含めるかどうかの
        # フラグ。既定は含めない（0）。部門ごとに機能自体の表示/非表示を切替
        # できるようにするため、対応する dept_master.enable_dev_summary も参照する。
        if "include_in_dev_summary" not in pt_cols2:
            db.execute("ALTER TABLE project_task ADD COLUMN include_in_dev_summary INTEGER NOT NULL DEFAULT 0")
            logger.info("project_task.include_in_dev_summary カラムを追加しました。")

    # weekly_schedule に project_task_id を追加（なければ）
    ws_cols = {row[1] for row in db.execute("PRAGMA table_info(weekly_schedule)").fetchall()}
    if "project_task_id" not in ws_cols:
        db.execute("ALTER TABLE weekly_schedule ADD COLUMN project_task_id INTEGER")
        logger.info("weekly_schedule.project_task_id カラムを追加しました。")

    # daily_result に project_task_id を追加（なければ）
    dr_cols = {row[1] for row in db.execute("PRAGMA table_info(daily_result)").fetchall()}
    if "project_task_id" not in dr_cols:
        db.execute("ALTER TABLE daily_result ADD COLUMN project_task_id INTEGER")
        logger.info("daily_result.project_task_id カラムを追加しました。")

    # users に姓・名カラムを追加（なければ）
    u_cols = {row[1] for row in db.execute("PRAGMA table_info(users)").fetchall()}
    if "last_name" not in u_cols:
        db.execute("ALTER TABLE users ADD COLUMN last_name TEXT DEFAULT ''")
        db.execute("ALTER TABLE users ADD COLUMN first_name TEXT DEFAULT ''")
        # 既存ユーザーの name を姓として初期設定
        db.execute("UPDATE users SET last_name = name WHERE last_name = '' OR last_name IS NULL")
        logger.info("users.last_name / first_name カラムを追加しました。")

    # users に login_id を追加（なければ）。ログインID方式（Phase1）で本人特定に使う。
    # 既存ユーザーには 'user{id}' 形式の初期IDを付与し、管理画面で編集できるようにする。
    u_cols = {row[1] for row in db.execute("PRAGMA table_info(users)").fetchall()}
    if "login_id" not in u_cols:
        db.execute("ALTER TABLE users ADD COLUMN login_id TEXT DEFAULT ''")
        db.execute(
            "UPDATE users SET login_id = 'user' || id "
            "WHERE login_id IS NULL OR login_id = ''"
        )
        # login_id の一意性を担保するインデックス（空文字は許容しないため部分インデックス）
        db.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_users_login_id "
            "ON users(login_id) WHERE login_id != ''"
        )
        logger.info("users.login_id カラムを追加し、初期IDを付与しました。")

    # dept_master が空で users.dept に実値がある場合、そこから部署マスタを補完する（冪等）。
    # 所属切替の候補や管理画面の部署選択に dept_master を使うため、実データと同期させる。
    dm_count = db.execute("SELECT COUNT(*) FROM dept_master").fetchone()[0]
    if dm_count == 0:
        db.execute("""
            INSERT OR IGNORE INTO dept_master (dept_name, display_order)
            SELECT dept, MIN(id) FROM users
            WHERE dept IS NOT NULL AND dept != ''
            GROUP BY dept
        """)
        added = db.execute("SELECT COUNT(*) FROM dept_master").fetchone()[0]
        if added:
            logger.info("dept_master を users.dept から補完しました（%d件）。", added)

    # user_affiliation テーブル（所属長の担当所属集合。複数所属対応）を作成（なければ）。
    # users.dept（主所属・単一）は温存し、所属長の複数所属をこの中間テーブルで表す。
    tables = {row[0] for row in db.execute(
        "SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
    if "user_affiliation" not in tables:
        db.execute("""CREATE TABLE user_affiliation (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            dept_name TEXT NOT NULL,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            UNIQUE(user_id, dept_name),
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        )""")
        db.execute(
            "CREATE INDEX IF NOT EXISTS idx_user_affiliation_user "
            "ON user_affiliation(user_id)"
        )
        logger.info("user_affiliation テーブルを作成しました。")

    # デフォルトデータ挿入（なければ）
    db.execute("""
        INSERT OR IGNORE INTO mail_settings (role, subject_template, body_template)
        VALUES ('管理職',
            '【日報】{date}（{day_of_week}） {dept}',
            'お疲れ様です。{sender_name}です。\n本日{date}（{day_of_week}）のチーム日報をご報告します。\n\n■ メンバー実績\n\n{member_reports}\n\n以上、よろしくお願いいたします。\n{sender_name}')
    """)
    db.execute("""
        INSERT OR IGNORE INTO mail_settings (role, subject_template, body_template)
        VALUES ('マスタ', '', '矢野会長様、志保社長様')
    """)

    # マスタの body_template が旧デフォルト値の場合、新しい挨拶文に更新する
    old_master_body = (
        "お疲れ様です。{sender_name}です。\n本日{date}（{day_of_week}）の{dept}日報をご報告します。"
        "\n\n■ 部署実績\n\n{member_reports}\n\n以上、よろしくお願いいたします。\n{sender_name}"
    )
    row = db.execute(
        "SELECT body_template FROM mail_settings WHERE role = 'マスタ'"
    ).fetchone()
    if row and row["body_template"] == old_master_body:
        db.execute(
            "UPDATE mail_settings SET body_template = '矢野会長様、志保社長様', subject_template = '' WHERE role = 'マスタ'"
        )
        logger.info("マスタ用メール設定を新しい挨拶文テンプレートに更新しました。")

    # routine_schedule に days（曜日フラグ）を追加（なければ）。
    rs_cols = {row[1] for row in db.execute("PRAGMA table_info(routine_schedule)").fetchall()}
    if "days" not in rs_cols:
        db.execute("ALTER TABLE routine_schedule ADD COLUMN days TEXT DEFAULT '1,1,1,1,1'")
        logger.info("routine_schedule.days カラムを追加しました。")
        rs_cols.add("days")

    # routine_schedule に fill_direction（空きスロットへの詰め方向）を追加（なければ）。
    # 'top'=上から詰める（既定） / 'bottom'=下から詰める。
    if "fill_direction" not in rs_cols:
        db.execute("ALTER TABLE routine_schedule ADD COLUMN fill_direction TEXT DEFAULT 'top'")
        logger.info("routine_schedule.fill_direction カラムを追加しました。")

    # routine_schedule に頻度種別（daily=曜日ベース／weekly_spot=週次スポット日付／
    # weekly_pattern=週次の週番号+曜日／yearly=年次の対象月日）を追加（なければ）。
    if "freq_type" not in rs_cols:
        db.execute("ALTER TABLE routine_schedule ADD COLUMN freq_type TEXT DEFAULT 'daily'")
        logger.info("routine_schedule.freq_type カラムを追加しました。")
    if "spot_date" not in rs_cols:
        db.execute("ALTER TABLE routine_schedule ADD COLUMN spot_date TEXT DEFAULT ''")
        logger.info("routine_schedule.spot_date カラムを追加しました。")
    if "week_numbers" not in rs_cols:
        db.execute("ALTER TABLE routine_schedule ADD COLUMN week_numbers TEXT DEFAULT ''")
        logger.info("routine_schedule.week_numbers カラムを追加しました。")
    if "yearly_month" not in rs_cols:
        db.execute("ALTER TABLE routine_schedule ADD COLUMN yearly_month INTEGER DEFAULT 0")
        logger.info("routine_schedule.yearly_month カラムを追加しました。")
    if "yearly_day" not in rs_cols:
        db.execute("ALTER TABLE routine_schedule ADD COLUMN yearly_day INTEGER DEFAULT 0")
        logger.info("routine_schedule.yearly_day カラムを追加しました。")
    # weekly_spot / weekly_pattern / yearly は row_number（11以降）だけでは
    # AM/PM区分を判定できない（daily は row_number<=5=AM の暗黙ルールに依存）ため、
    # 明示的な区分カラムを追加する。
    if "period" not in rs_cols:
        db.execute("ALTER TABLE routine_schedule ADD COLUMN period TEXT DEFAULT ''")
        logger.info("routine_schedule.period カラムを追加しました。")

    # user_survey テーブル（ログイン時アンケート、ユーザー1人につき1回のみ回答）を作成（なければ）。
    if "user_survey" not in tables:
        db.execute("""CREATE TABLE user_survey (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL UNIQUE,
            usability INTEGER NOT NULL,
            merit TEXT DEFAULT '',
            demerit TEXT DEFAULT '',
            requested_feature TEXT DEFAULT '',
            recommend TEXT NOT NULL,
            recommended_dept TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        )""")
        logger.info("user_survey テーブルを作成しました。")
    else:
        # recommended_dept（おすすめしたい部門の自由記述）を追加（なければ）。
        survey_cols = {row[1] for row in db.execute("PRAGMA table_info(user_survey)").fetchall()}
        if "recommended_dept" not in survey_cols:
            db.execute("ALTER TABLE user_survey ADD COLUMN recommended_dept TEXT DEFAULT ''")
            logger.info("user_survey.recommended_dept カラムを追加しました。")

    db.commit()


def _migrate_users_from_json(db: sqlite3.Connection) -> None:
    """
    users.json が存在する場合、usersテーブルが空のときのみINSERTで移行する。

    Args:
        db: データベース接続オブジェクト
    """
    if not _USERS_JSON_PATH.exists():
        logger.info("users.json が見つかりません。ユーザー移行をスキップします。")
        return

    row_count = db.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    if row_count > 0:
        logger.debug("usersテーブルにデータが存在するため、移行をスキップします。")
        return

    try:
        with _USERS_JSON_PATH.open(encoding="utf-8-sig") as f:
            users_data = json.load(f)

        for user in users_data:
            name = user.get("name", "").strip()
            if not name:
                continue
            role = user.get("role", "ユーザー")
            dept = user.get("dept", "")
            db.execute(
                "INSERT OR IGNORE INTO users (name, role, dept, std_hours_am, std_hours_pm) "
                "VALUES (?, ?, ?, 4.0, 4.0)",
                (name, role, dept),
            )
        db.commit()
        logger.info("users.json からユーザーデータを移行しました。")
    except (json.JSONDecodeError, KeyError, OSError) as exc:
        logger.warning("users.json の移行中にエラーが発生しました: %s", exc)


def init_db(app: Flask) -> None:
    """
    DBスキーマを作成し、users.json からユーザーデータを移行する。

    Flaskアプリコンテキスト内で呼び出すこと。

    Args:
        app: Flaskアプリケーションインスタンス
    """
    db_path = pathlib.Path(app.config["DATABASE"])
    db_path.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")

    try:
        conn.executescript(_SCHEMA_SQL)
        conn.commit()
        logger.debug("DBスキーマの作成/確認が完了しました。")
        _migrate_schema(conn)
        _migrate_users_from_json(conn)
    finally:
        conn.close()

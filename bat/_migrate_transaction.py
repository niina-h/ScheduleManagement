"""トランザクションテーブルを開発DBから本番DBへ移行するスクリプト。

マスタテーブル（users, dept_master, task_master, task_category,
task_subcategory, mail_settings）は本番データを維持し、
トランザクションテーブルのみ開発データで上書きする。

Usage:
    python _migrate_transaction.py <開発DB> <本番DB>
"""
from __future__ import annotations

import sqlite3
import sys

# 移行対象のトランザクションテーブル
TRANSACTION_TABLES: list[str] = [
    "weekly_schedule",
    "daily_result",
    "daily_comment",
    "carryover",
    "weekly_leave",
    "operation_log",
]


def migrate(dev_db: str, prod_db: str) -> None:
    """開発DBのトランザクションデータを本番DBへ移行する。

    Args:
        dev_db: 開発DBのパス。
        prod_db: 本番DBのパス。
    """
    # 開発DBからデータ読み取り
    dev_conn = sqlite3.connect(dev_db)
    dev_conn.row_factory = sqlite3.Row
    data: dict[str, dict] = {}
    for table in TRANSACTION_TABLES:
        rows = dev_conn.execute(f"SELECT * FROM [{table}]").fetchall()
        if rows:
            cols = rows[0].keys()
            data[table] = {"cols": cols, "rows": [tuple(r) for r in rows]}
        else:
            data[table] = {"cols": [], "rows": []}
        print(f"  読取: {table} → {len(data[table]['rows'])} 件")
    dev_conn.close()

    # 本番DBへ書き込み
    prod_conn = sqlite3.connect(prod_db)
    prod_conn.execute("PRAGMA foreign_keys = OFF")
    for table in TRANSACTION_TABLES:
        prod_conn.execute(f"DELETE FROM [{table}]")
        info = data[table]
        if info["rows"]:
            placeholders = ",".join(["?" for _ in info["cols"]])
            col_names = ",".join(info["cols"])
            prod_conn.executemany(
                f"INSERT INTO [{table}] ({col_names}) VALUES ({placeholders})",
                info["rows"],
            )
        # AUTOINCREMENTシーケンス値を同期
        try:
            max_row = prod_conn.execute(f"SELECT MAX(id) FROM [{table}]").fetchone()
            if max_row and max_row[0]:
                prod_conn.execute(
                    "UPDATE sqlite_sequence SET seq = ? WHERE name = ?",
                    (max_row[0], table),
                )
        except sqlite3.OperationalError:
            pass
        print(f"  書込: {table} → {len(info['rows'])} 件")

    prod_conn.execute("PRAGMA foreign_keys = ON")
    prod_conn.commit()
    prod_conn.close()
    print("  → トランザクション移行完了")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python _migrate_transaction.py <開発DB> <本番DB>", file=sys.stderr)
        sys.exit(1)
    migrate(sys.argv[1], sys.argv[2])

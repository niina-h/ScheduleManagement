"""診断スクリプト: 指定ユーザーの週間予定データを確認する。

使い方:
    python check_schedule.py
"""
from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

DB_PATH = Path(__file__).parent / "db" / "web_app.db"


def main() -> None:
    """髙陽の週間予定データを診断する。"""
    if not DB_PATH.exists():
        print(f"DB not found: {DB_PATH}")
        sys.exit(1)

    db = sqlite3.connect(str(DB_PATH))
    db.row_factory = sqlite3.Row

    # ユーザー一覧
    print("=== ユーザー一覧 ===")
    for u in db.execute("SELECT id, name, dept, role FROM users ORDER BY id").fetchall():
        print(f"  id={u['id']}  name={u['name']}  dept={u['dept']}  role={u['role']}")

    # 髙陽を探す（部分一致）
    target_name = "髙陽"
    user = db.execute(
        "SELECT id, name FROM users WHERE name LIKE ?", (f"%{target_name}%",)
    ).fetchone()

    if not user:
        # 全ユーザー名をバイト列で表示して確認
        print(f"\n'{target_name}' が見つかりません。全ユーザー名:")
        for u in db.execute("SELECT id, name FROM users").fetchall():
            print(f"  id={u['id']}  name={u['name']!r}  bytes={u['name'].encode('utf-8')!r}")
        # IDで直接指定
        target_id = input("\n確認するユーザーIDを入力: ").strip()
        if not target_id.isdigit():
            print("無効なIDです")
            return
        uid = int(target_id)
    else:
        uid = user["id"]
        print(f"\n対象ユーザー: id={uid}, name={user['name']}")

    # 3/23週 (2026-03-23) の週間予定を確認
    week_starts = ["2026-03-23", "2026-03-16"]
    for ws in week_starts:
        print(f"\n=== 週間予定 week_start={ws} user_id={uid} ===")
        rows = db.execute(
            "SELECT day_of_week, time_slot, slot_index, task_name, hours, updated_at "
            "FROM weekly_schedule "
            "WHERE user_id = ? AND week_start = ? "
            "ORDER BY day_of_week, time_slot, slot_index",
            (uid, ws),
        ).fetchall()
        if not rows:
            print("  データなし（予定が登録されていません）")
        else:
            print(f"  レコード数: {len(rows)}")
            for r in rows:
                day_names = ["月", "火", "水", "木", "金"]
                day = day_names[r["day_of_week"]] if r["day_of_week"] < 5 else "?"
                tn = r["task_name"] or "(空)"
                h = r["hours"] or 0
                if tn != "(空)" or h > 0:
                    print(f"  {day}曜 {r['time_slot']} [{r['slot_index']}] {tn}  {h}h  更新={r['updated_at']}")

    # 日次実績も確認
    dates = ["2026-03-23", "2026-03-24", "2026-03-25"]
    for d in dates:
        print(f"\n=== 日次実績 date={d} user_id={uid} ===")
        rows = db.execute(
            "SELECT time_slot, slot_index, task_name, hours, updated_at "
            "FROM daily_result "
            "WHERE user_id = ? AND date = ? "
            "ORDER BY time_slot, slot_index",
            (uid, d),
        ).fetchall()
        if not rows:
            print("  データなし")
        else:
            for r in rows:
                tn = r["task_name"] or "(空)"
                h = r["hours"] or 0
                if tn != "(空)" or h > 0:
                    print(f"  {r['time_slot']} [{r['slot_index']}] {tn}  {h}h  更新={r['updated_at']}")

    db.close()


if __name__ == "__main__":
    main()

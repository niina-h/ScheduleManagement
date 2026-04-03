import sqlite3
conn = sqlite3.connect(r"db/web_app.db")
conn.row_factory = sqlite3.Row
rows = conn.execute("SELECT id,name,std_hours_am,std_hours_pm,std_hours FROM users ORDER BY display_order").fetchall()
for r in rows:
    print(dict(r))
conn.close()

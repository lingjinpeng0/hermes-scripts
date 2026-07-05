#!/usr/bin/env python3
"""Check if a specific TUI session has finished (no new activity in 5+ min)."""
import sqlite3, os, time

db = os.path.expanduser(r"~/AppData/Local/hermes/state.db")
session_id = "20260702_231634_4a0571"
now = time.time()
inactive_threshold = 300  # 5 minutes

try:
    conn = sqlite3.connect(db)
    c = conn.cursor()
    c.execute("SELECT id, title, last_active_at FROM sessions WHERE id = ?", (session_id,))
    row = c.fetchone()
    conn.close()
    if row:
        last_active = row[2] / 1000 if row[2] > 1e10 else row[2]  # handle ms vs s
        idle = now - last_active
        if idle > inactive_threshold:
            print(f"TUI会话「{row[1]}」已完成！最后活跃于 {time.strftime('%H:%M', time.localtime(last_active))}")
        else:
            print(f"SILENT: 会话仍在活动中（{int(idle)}秒无操作）")
    else:
        print("TUI会话已结束（会话记录未找到）")
except Exception as e:
    print(f"SILENT: 检查失败: {e}")

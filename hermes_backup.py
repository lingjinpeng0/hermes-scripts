#!/usr/bin/env python3
"""雪乃的记忆备份 —— 只备份真正'是她'的那部分，不碰缓存和日志。

用法:
    python hermes_backup.py            # 打包到 ~/Documents/Hermes备份
    python hermes_backup.py --phone    # 顺便推一份到手机

备份内容: memories/ sessions/ skills/ cron/ config.yaml SOUL.md .env channel_directory.json auth.json
排除: cache/ logs/ backups/ image_cache/ 以及所有 zip/venv
"""
import os
import sys
import glob
import zipfile
import subprocess
import datetime

HOME = os.path.expanduser(r"~\AppData\Local\hermes")
ADB = os.path.join(HOME, "tools", "platform-tools", "adb.exe")
OUT_DIR = os.path.expanduser(r"~\Documents\Hermes备份")
PHONE_DIR = "/sdcard/hermes-backup"
KEEP = 7

DIRS = ["memories", "sessions", "skills", "cron", "plugins", "gateway"]
FILES = ["config.yaml", "SOUL.md", ".env", "channel_directory.json", "auth.json",
         "context_length_cache.yaml"]
SKIP_DIRS = {"__pycache__", ".git", "node_modules", "venv", ".venv", "cache"}
SKIP_EXT = {".zip", ".pyc", ".log", ".sqlite-wal", ".sqlite-shm"}

# 真正的对话历史在 state.db；FTS 索引可重建，不备份
DB_TABLES = ["schema_version", "sessions", "messages", "state_meta",
             "system_prompts", "session_model_usage", "gateway_routing",
             "async_delegations", "conversation_generations"]


def add_tree(zf, base, rel):
    root = os.path.join(base, rel)
    if not os.path.exists(root):
        return 0
    n = 0
    for cur, dirs, files in os.walk(root):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for f in files:
            if os.path.splitext(f)[1].lower() in SKIP_EXT:
                continue
            full = os.path.join(cur, f)
            try:
                zf.write(full, os.path.relpath(full, base))
                n += 1
            except OSError:
                pass
    return n


def snapshot_state(dst):
    """用只读连接导出一份一致的 state.db 核心库（不拷 FTS 索引）。"""
    import sqlite3
    src = os.path.join(HOME, "state.db")
    if not os.path.exists(src):
        print("[对话] state.db 不存在，跳过")
        return None
    if os.path.exists(dst):
        os.remove(dst)
    s = sqlite3.connect(f"file:{src}?mode=ro", uri=True)
    d = sqlite3.connect(dst)
    total = 0
    for t in DB_TABLES:
        row = s.execute("select sql from sqlite_master where type='table' and name=?", (t,)).fetchone()
        if not row:
            continue
        d.execute(row[0])
        cols = [r[1] for r in s.execute(f"pragma table_info('{t}')")]
        q = (f"insert into '{t}' ({','.join(chr(34)+c+chr(34) for c in cols)}) "
             f"values ({','.join('?'*len(cols))})")
        rows = s.execute(f"select * from '{t}'").fetchall()
        d.executemany(q, rows)
        total += len(rows)
    d.commit()
    d.close()
    s.close()
    print(f"[对话] 导出 {total} 行 -> {os.path.getsize(dst)/1024/1024:.1f} MB")
    return dst


def build():
    os.makedirs(OUT_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    out = os.path.join(OUT_DIR, f"hermes-core-{ts}.zip")
    snap = snapshot_state(os.path.join(OUT_DIR, "_state_core.db"))
    count = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for d in DIRS:
            count += add_tree(zf, HOME, d)
        for f in FILES:
            p = os.path.join(HOME, f)
            if os.path.exists(p):
                zf.write(p, f)
                count += 1
        if snap:
            zf.write(snap, "state_core.db")
            count += 1
    if snap and os.path.exists(snap):
        os.remove(snap)
    size = os.path.getsize(out)
    print(f"[打包] {out}")
    print(f"[打包] {count} 个文件, {size/1024/1024:.1f} MB")
    return out


def adb(*args):
    r = subprocess.run([ADB, *args], capture_output=True, text=True, timeout=180)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def push(local):
    rc, out = adb("devices")
    if "\tdevice" not in out:
        print("[手机] 没连上，这次只留本地")
        return False
    adb("shell", "mkdir", "-p", PHONE_DIR)
    rc, out = adb("push", local, PHONE_DIR + "/")
    if rc != 0:
        print("[手机] 推送失败:", out.strip())
        return False
    name = os.path.basename(local)
    rc, out = adb("shell", "ls", "-l", f"{PHONE_DIR}/{name}")
    print("[手机] 已推送:", out.strip())
    # 只保留最近 KEEP 份
    rc, out = adb("shell", "ls", PHONE_DIR)
    names = sorted([x.strip() for x in out.splitlines()
                    if x.strip().startswith("hermes-core-") and x.strip().endswith(".zip")])
    for old in names[:-KEEP]:
        adb("shell", "rm", "-f", f"{PHONE_DIR}/{old}")
        print("[手机] 清理旧备份:", old)
    return True


def prune_local():
    names = sorted(glob.glob(os.path.join(OUT_DIR, "hermes-core-*.zip")))
    for old in names[:-KEEP]:
        os.remove(old)
        print("[本地] 清理旧备份:", os.path.basename(old))


if __name__ == "__main__":
    z = build()
    prune_local()
    if "--phone" in sys.argv:
        push(z)

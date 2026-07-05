#!/usr/bin/env python3
"""Background watchdog: watches for blackout signal file and toggles blackout.exe."""
import subprocess
import time
import os
import sys
import signal

BLACKOUT_EXE = r"C:\Users\Rei\AppData\Local\hermes\skills\productivity\screen-blackout\scripts\blackout.exe"
SIGNAL_FILE = os.path.expanduser(r"~/.hermes/blackout_signal.tmp")

def is_running():
    result = subprocess.run(
        ["tasklist", "/fi", "imagename eq blackout.exe"],
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    return "blackout.exe" in (result.stdout or "")

def main():
    print(f"🖤 黑屏监听已启动")
    print(f"   信号文件: {SIGNAL_FILE}")
    print(f"   Ctrl+C 退出")

    notified = False
    while True:
        if os.path.exists(SIGNAL_FILE):
            try:
                with open(SIGNAL_FILE, "r") as f:
                    cmd = f.read().strip()
                os.remove(SIGNAL_FILE)

                if cmd == "on":
                    if not is_running():
                        subprocess.Popen([BLACKOUT_EXE], shell=True)
                        print(f"[{time.strftime('%H:%M:%S')}] ✅ 已开启黑屏")
                    else:
                        print(f"[{time.strftime('%H:%M:%S')}] ⚠️ 已处于黑屏状态")
                elif cmd == "off":
                    if is_running():
                        subprocess.run(["taskkill", "/f", "/im", "blackout.exe"],
                                     capture_output=True)
                        print(f"[{time.strftime('%H:%M:%S')}] ✅ 已退出黑屏")
                    else:
                        print(f"[{time.strftime('%H:%M:%S')}] ⚠️ 当前没有黑屏")
            except Exception as e:
                print(f"[{time.strftime('%H:%M:%S')}] ❌ 错误: {e}")

        if not notified:
            print(f"[{time.strftime('%H:%M:%S')}] 监听中...")
            notified = True

        time.sleep(1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n🖤 监听已停止")

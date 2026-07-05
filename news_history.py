#!/usr/bin/env python3
import urllib.request
url = "https://60s.viki.moe/v2/today-in-history?encoding=text"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
try:
    with urllib.request.urlopen(req, timeout=15) as r:
        print(r.read().decode('utf-8'))
except Exception as e:
    print(f"[历史上的今天] 请求失败: {e}")

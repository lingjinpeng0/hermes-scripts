#!/usr/bin/env python3
import urllib.request
url = "https://60s.viki.moe/v2/60s?encoding=text"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
try:
    with urllib.request.urlopen(req, timeout=15) as r:
        text = r.read().decode('utf-8')
        # 去掉微语行
        lines = [l for l in text.split('\n') if not l.strip().startswith('【微语】')]
        print('\n'.join(lines).strip())
except Exception as e:
    print(f"[60s每日新闻] 请求失败: {e}")

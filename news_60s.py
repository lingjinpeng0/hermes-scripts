#!/usr/bin/env python3
"""60s每日新闻 — cron 版（多实例 fallback + HTTP 状态校验 + 重试）

公共逻辑见 news_common.py。旧备用镜像 60s.7se.cn 忽略 encoding 参数、永远返回
JSON，这里的 fmt 会把那种 JSON 拆成正常文本，避免推出一坨原始 JSON。
"""
import json
import sys

from news_common import format_list, run


def fmt(body):
    if body.lstrip().startswith("{"):
        try:
            d = json.loads(body)
            items = d.get("data") or []
            items = [i if isinstance(i, str) else i.get("title", "") for i in items]
            title = d.get("title", "每天60秒知天下")
            return title + "\n\n" + "\n".join(items)
        except Exception:
            return format_list(body, "每天60秒知天下")
    lines = [l for l in body.split("\n") if not l.strip().startswith("【微语】")]
    return "\n".join(lines).strip()


if __name__ == "__main__":
    sys.exit(run("60s每日新闻", "/v2/60s?encoding=text", fmt))

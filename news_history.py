#!/usr/bin/env python3
"""历史上的今天 — cron 版（多实例 fallback + HTTP 状态校验 + 重试）

公共逻辑见 news_common.py：主站超时、备用镜像返回 404 页面都靠那里的
HTTP 状态码 + HTML 校验兜住，绝不会再把错误页当资讯推送。
"""
import sys

from news_common import format_list, run


def fmt(body):
    return format_list(body, "历史上的今天")


if __name__ == "__main__":
    sys.exit(run("历史上的今天", "/v2/today-in-history?encoding=text", fmt))

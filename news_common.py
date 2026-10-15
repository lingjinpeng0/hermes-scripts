#!/usr/bin/env python3
"""60s 新闻抓取的公共部分：多实例 fallback + HTTP 状态校验 + 重试。

踩坑记录（2026-10-14 / 10-15 连续两天推送抽风）：
- 主站 `60s.viki.moe` 会间歇性整体超时（Cloudflare 线路，实测 3 轮里有一轮三个端点
  全部 timeout 15s），一旦失败就会整体落到备用实例。
- 备用 `60s.7se.cn` 是旧版镜像：只镜像了 `/v2/60s`（而且忽略 `encoding` 参数，永远返回
  JSON），其他路径返回 nginx 的 `404 Not Found` HTML 页面。
- 旧脚本的判据只有「body 非空」，于是 404 HTML 页被当成资讯原样推送到 QQ；
  `/v2/60s` 则推出一坨原始 JSON。
- 结论：**必须校验 HTTP 状态码并拒绝 HTML**，否则备用实例的 404 会被静默当成成功。
"""
import json
import re
import subprocess
import sys
import time

# 顺序 = 尝试顺序。60s.viki.moe 最快但会间歇性超时；crystelf.top 实测三个端点齐全且稳定；
# 7se.cn 只留作最后兜底（仅 /v2/60s 可用）。
MIRRORS = [
    "https://60s.viki.moe",
    "https://60s.crystelf.top",
    "https://60s.7se.cn",
]
MAX_ATTEMPTS = 2
BACKOFF = 5
TIMEOUT = 8

_EMOJI = re.compile(
    r'[\U0001F300-\U0001F9FF\u2600-\u27BF\u2700-\u27BF\uFE00-\uFE0F]')


def rm_emoji(s):
    return _EMOJI.sub('', s)


def fetch(url, timeout=TIMEOUT):
    """返回 (body, None) 或 (None, 失败原因)。非 200 / 空 body / HTML 都算失败。"""
    r = subprocess.run(
        ["curl", "-s", "--max-time", str(timeout),
         "-H", "User-Agent: Mozilla/5.0",
         "-w", "\n%{http_code}", url],
        capture_output=True, text=True, timeout=timeout + 10
    )
    if r.returncode != 0:
        return None, f"curl exit={r.returncode}"
    body, _, code = r.stdout.rpartition("\n")
    code = code.strip()
    if code != "200":
        return None, f"HTTP {code or '???'}"
    body = body.strip()
    if not body:
        return None, "空响应"
    head = body[:200].lower()
    if "<html" in head or "<!doctype" in head or body.startswith("error code"):
        return None, "返回 HTML/错误页而非数据"
    return body, None


def format_list(body, fallback_title):
    """兼容旧镜像返回的 JSON（{status,title,data:[...]}）和正常的 text 格式。"""
    if body.lstrip().startswith("{"):
        try:
            d = json.loads(body)
        except Exception:
            return body
        items = d.get("data") or d.get("news") or []
        items = [i if isinstance(i, str) else i.get("title", "") for i in items]
        title = d.get("title") or fallback_title
        return title + "\n\n" + "\n".join(items)
    return body


def run(name, path, formatter=None, mirrors=None, attempts=MAX_ATTEMPTS):
    mirrors = mirrors or MIRRORS
    last_err = ""
    for attempt in range(1, attempts + 1):
        for base in mirrors:
            body, err = fetch(base + path)
            if err is None:
                out = formatter(body) if formatter else body
                print(rm_emoji(out).strip())
                return 0
            last_err = f"{base} -> {err}"
        if attempt < attempts:
            time.sleep(BACKOFF * attempt)
    print(f"⚠️ [{name}] 所有 60s 实例均不可用: {last_err}（已重试{attempts}轮）")
    return 0


if __name__ == "__main__":
    sys.exit(run("自检", "/v2/60s?encoding=text", format_list))

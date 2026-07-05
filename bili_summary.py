#!/usr/bin/env python3
"""
Bilibili视频字幕提取+AI总结工具
用法: python bili_summary.py <BV号或B站URL>
Cookie自动从 ~/.hermes/bili_cred.json 读取
"""
import sys
import json
import re
import os
import requests

VENV_SITE = r"C:\Users\Rei\AppData\Local\hermes\hermes-agent\venv\Lib\site-packages"
if VENV_SITE not in sys.path:
    sys.path.insert(0, VENV_SITE)

from bilibili_api import video, sync, Credential

CRED_FILE = os.path.expanduser("~/.hermes/bili_cred.json")


def load_cred():
    """加载B站凭证"""
    if os.path.exists(CRED_FILE):
        with open(CRED_FILE) as f:
            d = json.load(f)
            return d.get("sessdata", ""), d.get("bili_jct", "")
    return "", ""


def extract_bv(url_or_bv: str) -> str:
    # 纯BV号
    if url_or_bv.startswith('BV') and len(url_or_bv) == 12:
        return url_or_bv
    
    # 尝试从URL提取BV
    m = re.search(r'BV\w{10}', url_or_bv)
    if m:
        return m.group()
    
    # b23.tv短链接 → 解析重定向
    if 'b23.tv' in url_or_bv or 'b23' in url_or_bv:
        try:
            resp = requests.get(url_or_bv, allow_redirects=True, timeout=10)
            final_url = resp.url
            m2 = re.search(r'BV\w{10}', final_url)
            if m2:
                return m2.group()
        except Exception:
            pass
    
    # AV号
    m_aid = re.search(r'av(\d+)', url_or_bv, re.I)
    if m_aid:
        return m_aid.group()
    
    raise ValueError(f"无法解析BV号: {url_or_bv}")


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"error": "用法: python bili_summary.py <BV号或URL>"}, ensure_ascii=False))
        sys.exit(1)

    url_or_bv = sys.argv[1]
    bv = extract_bv(url_or_bv)
    sessdata, bili_jct = load_cred()

    cred = None
    if sessdata and bili_jct:
        cred = Credential(sessdata=sessdata, bili_jct=bili_jct)

    v = video.Video(bvid=bv, credential=cred)
    info = sync(v.get_info())

    result = {
        "title": info.get("title", ""),
        "author": info.get("owner", {}).get("name", ""),
        "duration": info.get("duration", 0),
        "views": info.get("stat", {}).get("view", 0),
        "likes": info.get("stat", {}).get("like", 0),
        "danmaku": info.get("stat", {}).get("danmaku", 0),
        "subtitle_text": "",
        "ai_summary": "",
        "ai_outline": [],
    }

    cid = info.get("cid", 0)

    if cred and cid:
        # 1) AI结论（含总结 + 提纲 + 字幕）
        try:
            conc = sync(v.get_ai_conclusion(cid=cid))
            mr = conc.get("model_result", {})
            result["ai_summary"] = mr.get("summary", "")
            result["ai_outline"] = mr.get("outline", [])

            # 提取字幕文本
            subtitle_parts = mr.get("subtitle", [])
            texts = []
            for sp in subtitle_parts:
                for ps in sp.get("part_subtitle", []):
                    texts.append(ps.get("content", ""))
            result["subtitle_text"] = "\n".join(texts)

        except Exception as e:
            result["subtitle_text"] = f"[字幕/总结提取失败: {e}]"

        # 2) 如果AI结论里没字幕，尝试从subtitle_url拿
        if not result["subtitle_text"].strip():
            try:
                subs = sync(v.get_subtitle(cid=cid))
                sub_list = subs.get("subtitles", []) if isinstance(subs, dict) else []
                for s in sub_list:
                    if "ai-zh" in s.get("lan", "") or "中文" in s.get("lan_doc", ""):
                        sub_url = s.get("subtitle_url", "")
                        if sub_url:
                            if sub_url.startswith("//"):
                                sub_url = "https:" + sub_url
                            resp = requests.get(sub_url, timeout=10)
                            data = resp.json()
                            texts = [item.get("content", "") for item in data.get("body", [])]
                            result["subtitle_text"] = "\n".join(texts)
                            break
            except Exception:
                pass

    # 没字幕时补充描述信息
    if not result["subtitle_text"] and not result["ai_summary"]:
        result["desc"] = info.get("desc", "")[:300]

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

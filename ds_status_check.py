#!/usr/bin/env python3
"""Monitor DeepSeek API status via RSS feed. Reports changes only."""
import urllib.request, ssl, json, os, xml.etree.ElementTree as ET
from datetime import datetime

STATE_FILE = os.path.expanduser("~/.hermes/scripts/ds_status_state.json")
RSS_URL = "https://status.deepseek.com/feed.rss"

def fetch_rss():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(RSS_URL, headers={"User-Agent": "Mozilla/5.0"})
    resp = urllib.request.urlopen(req, timeout=20, context=ctx)
    return resp.read().decode()

def parse_rss(xml_text):
    root = ET.fromstring(xml_text)
    items = root.findall(".//item")
    statuses = []
    for item in items[:10]:
        title = item.findtext("title", "")
        link = item.findtext("link", "")
        pub_date = item.findtext("pubDate", "")
        statuses.append({"title": title, "link": link, "date": pub_date})
    return statuses

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_titles": [], "last_check": None}

def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def main():
    try:
        xml_text = fetch_rss()
    except Exception as e:
        print(f"❌ DeepSeek状态页RSS拉取失败: {e}")
        return

    statuses = parse_rss(xml_text)
    old_state = load_state()
    old_titles = old_state.get("last_titles", [])

    # Check for new entries
    new_entries = [s for s in statuses if s["title"] not in old_titles]

    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    if not old_titles:
        # First run, just save state
        save_state({"last_titles": [s["title"] for s in statuses], "last_check": now})
        print(f"✅ DeepSeek状态监控已初始化 ({len(statuses)}条记录)")
        return

    if new_entries:
        # Report changes
        lines = [f"🔔 DeepSeek API 状态更新 ({now})"]
        for entry in new_entries:
            lines.append(f"📌 {entry['title']}")
            if entry["link"]:
                lines.append(f"   {entry['link']}")
            if entry["date"]:
                lines.append(f"   {entry['date']}")
        print("\n".join(lines))
    else:
        # No changes - silent (no output = cron won't send)
        pass

    # Update state
    save_state({"last_titles": [s["title"] for s in statuses], "last_check": now})

if __name__ == "__main__":
    main()

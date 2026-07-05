#!/usr/bin/env python3
import sys, os, json, base64, urllib.request, re

# 直接从 .env 读 key
env_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env")
api_key = ""
try:
    with open(env_path, "r") as f:
        for line in f:
            m = re.match(r'^\s*XIAOMI_API_KEY\s*=\s*(.+?)\s*$', line)
            if m:
                api_key = m.group(1).strip("\"'")
except:
    pass

BASE_URL = "https://api.xiaomimimo.com/v1"
MODEL = "mimo-v2.5-tts-voicedesign"

def tts_from_file(input_path, output_path, voice_desc=""):
    global api_key
    if not api_key:
        print("need XIAOMI_API_KEY in .env", file=sys.stderr)
        sys.exit(1)
    with open(input_path, "r", encoding="utf-8") as f:
        text = f.read()
    payload = json.dumps({
        "model": MODEL,
        "messages": [
            {"role": "user", "content": voice_desc},
            {"role": "assistant", "content": text}
        ],
        "audio": {"format": "wav", "optimize_text_preview": True}
    }).encode()
    req = urllib.request.Request(
        f"{BASE_URL}/chat/completions",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        resp = json.loads(r.read())
        audio = base64.b64decode(resp["choices"][0]["message"]["audio"]["data"])
        with open(output_path, "wb") as f:
            f.write(audio)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("usage: script <input> <output> [voice_desc]", file=sys.stderr)
        sys.exit(1)
    desc = sys.argv[3] if len(sys.argv) > 3 else "一位年轻女性，声音清澈温柔"
    tts_from_file(sys.argv[1], sys.argv[2], desc)

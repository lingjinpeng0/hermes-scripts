#!/usr/bin/env python3
"""MiMo TTS - Hermes 自定义 TTS provider
Hermes 调用: 配在 config.yaml 的 tts.providers.mimo 下
脚本读取 {input_path} 文本，输出 {output_path} 音频
支持 {voice} 参数切换音色
"""
import sys, os, json, base64, urllib.request

API_KEY = os.environ.get("XIAOMI_API_KEY", "")
BASE_URL = os.environ.get("XIAOMI_BASE_URL", "https://api.xiaomimimo.com/v1")
VOICE = os.environ.get("MIMO_TTS_VOICE", "冰糖")
MODEL = os.environ.get("MIMO_TTS_MODEL", "mimo-v2.5-tts")

def tts_from_file(input_path: str, output_path: str, voice: str = ""):
    if not API_KEY:
        print("请设置 XIAOMI_API_KEY 环境变量", file=sys.stderr)
        sys.exit(1)
    
    with open(input_path, "r", encoding="utf-8") as f:
        text = f.read()
    
    active_voice = voice or VOICE
    
    payload = json.dumps({
        "model": MODEL,
        "messages": [
            {"role": "user", "content": "自然流畅，语调适中"},
            {"role": "assistant", "content": text}
        ],
        "audio": {"format": "wav", "voice": active_voice}
    }).encode()
    
    req = urllib.request.Request(
        f"{BASE_URL}/chat/completions",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {API_KEY}"
        }
    )
    
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            resp = json.loads(r.read())
            audio_b64 = resp["choices"][0]["message"]["audio"]["data"]
            audio_bytes = base64.b64decode(audio_b64)
            with open(output_path, "wb") as f:
                f.write(audio_bytes)
    except Exception as e:
        print(f"MiMo TTS 失败: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法: python mimo_tts.py <input_path> <output_path> [voice]", file=sys.stderr)
        sys.exit(1)
    voice = sys.argv[3] if len(sys.argv) > 3 else ""
    tts_from_file(sys.argv[1], sys.argv[2], voice)

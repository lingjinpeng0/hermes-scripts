#!/usr/bin/env python3
"""森空岛自动签到 — 仅终末地 — cron 版"""
import hashlib, hmac, json, os, sys, time
from urllib import parse
import requests

TOKEN = "KzZ8jf0FeAF+9ZKRuUwCgCKx"
REQUEST_TIMEOUT = 9
APP_CODE = "4ca99fa6b56cc2ba"
UA = {"User-Agent": "Skland/1.0.1 (com.hypergryph.skland; build:100001014; Android 31; ) Okhttp/4.11.0"}
SIGN_TPL = {"platform": "", "timestamp": "", "dId": "", "vName": ""}
BINDING = "https://zonai.skland.com/api/v1/game/player/binding"
GRANT = "https://as.hypergryph.com/user/oauth2/v2/grant"
CRED = "https://zonai.skland.com/api/v1/user/auth/generate_cred_by_code"
CHECKIN = "https://zonai.skland.com/api/v1/game/endfield/attendance"
ses = requests.Session()
_sign_token = ""

def gsign(path, body):
    t = str(int(time.time()) - 2)
    h = json.loads(json.dumps(SIGN_TPL))
    h["timestamp"] = t
    s = path + body + t + json.dumps(h, separators=(",", ":"))
    return hashlib.md5(hmac.new(_sign_token.encode(), s.encode(), hashlib.sha256).hexdigest().encode()).hexdigest(), h

def sh(url, method, body, headers):
    h = json.loads(json.dumps(headers))
    p = parse.urlparse(url)
    sign, shd = gsign(p.path, json.dumps(body) if method.lower()=="post" else p.query)
    h["sign"] = sign; h.update(shd); return h

def login(tok):
    global _sign_token
    r = ses.post(GRANT, json={"appCode":APP_CODE,"token":tok,"type":0}, headers=UA, timeout=REQUEST_TIMEOUT).json()
    if r.get("status")!=0: return None, f"grant: {r.get('msg','?')}"
    r2 = ses.post(CRED, json={"code":r["data"]["code"],"kind":1}, headers=UA, timeout=REQUEST_TIMEOUT).json()
    if r2.get("code")!=0: return None, f"cred: {r2.get('message','?')}"
    _sign_token = r2["data"]["token"]
    return r2["data"]["cred"], None

def get_roles(cred, app_code):
    h = sh(BINDING, "get", None, UA); h["cred"] = cred
    r = ses.get(BINDING, headers=h, timeout=REQUEST_TIMEOUT).json()
    if r.get("code")!=0: raise Exception(r.get("message","?"))
    for a in r["data"]["list"]:
        if a.get("appCode")==app_code: return a.get("bindingList",[])
    return []

def main():
    cred, err = login(TOKEN)
    if err: print(f"❌ {err}"); return 1
    try: roles = get_roles(cred, "endfield")
    except Exception as e: print(f"❌ 角色查询: {e}"); return 1
    if not roles: print("⚠️ 无终末地角色"); return 0
    for rl in roles:
        d = rl.get("defaultRole",{})
        nick, uid, rid, sid = d.get("nickname","?"), rl.get("uid",""), d.get("roleId",""), d.get("serverId","")
        if not all([rid,sid]): print(f"⚠️ {nick}: 缺参数"); continue
        body = {"uid":uid,"gameId":3,"roleId":rid,"serverId":sid}
        h = sh(CHECKIN, "post", body, UA); h["cred"] = cred
        r = ses.post(CHECKIN, json=body, headers=h, timeout=REQUEST_TIMEOUT).json()
        c, m = r.get("code"), r.get("message","")
        if c==0:
            aids = r.get("data",{}).get("awardIds",[])
            rmap = r.get("data",{}).get("resourceInfoMap",{})
            if aids and rmap:
                txt = [f'{rmap[a["id"]]["name"]}×{rmap[a["id"]].get("count",1)}' for a in aids if a.get("id") in rmap]
                print(f"✅ {nick} {'、'.join(txt)}" if txt else f"✅ {nick} 签到成功")
            else: print(f"✅ {nick} 签到成功")
        elif "重复" in m or c==10001: print(f"👍 {nick} 今日已签到")
        else: print(f"❌ {nick} (code={c}): {m}")
    return 0

if __name__=="__main__":
    sys.exit(main())

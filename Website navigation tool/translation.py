# -*- coding: utf-8 -*-
"""
翻译模块：百度翻译 + 腾讯翻译君（OpenAI 兼容之外的国内服务商）
- 不依赖外部第三方 SDK，requests 即可
- 百度：通用翻译 API，appid + key（MD5 签名），q 自动按 utf-8 处理
- 腾讯：tmt 通用机器翻译，密钥签名 TC3-HMAC-SHA256（v3 签名）

API 配置存到 json/_translation_config.json（不入 git），字段：
{
  "provider": "baidu" | "tencent",
  "baidu": {"appid": "", "key": ""},
  "tencent": {"secret_id": "", "secret_key": ""},
  "qps": 1,                  # 限速，避免被 ban
  "retry": 3,                # 失败重试次数
  "last_error": ""           # 上次错误信息（自检用）
}
"""

import hashlib
import json
import os
import random
import string
import time
import urllib.parse

import requests


# ----- 配置读写 -----

def _config_path():
    """翻译配置：固定放在 json/_translation_config.json（与 pintree.json 同目录）。"""
    base = os.path.dirname(os.path.abspath(__file__))              # .../网站导航工具
    return os.path.normpath(os.path.join(base, "..", "json", "_translation_config.json"))


_DEFAULT_CFG = {
    "provider": "baidu",
    "baidu": {"appid": "", "key": ""},
    "tencent": {"secret_id": "", "secret_key": ""},
    "qps": 1,
    "retry": 3,
    "last_error": "",
}


def load_config():
    p = _config_path()
    if not os.path.isfile(p):
        return dict(_DEFAULT_CFG)
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        # 缺字段兜底
        for k, v in _DEFAULT_CFG.items():
            data.setdefault(k, v)
        for sub in ("baidu", "tencent"):
            data[sub].setdefault(*({} if False else (_DEFAULT_CFG[sub].copy().popitem()[0], "")))
        # 简单补全（setdefault 上面那种写法容易出问题，重新兜底）
        for sub in ("baidu", "tencent"):
            for k, v in _DEFAULT_CFG[sub].items():
                data[sub].setdefault(k, v)
        return data
    except Exception:
        return dict(_DEFAULT_CFG)


def save_config(cfg):
    p = _config_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return p


# ----- 百度翻译 -----

_BAIDU_ENDPOINT = "https://fanyi-api.baidu.com/api/trans/vip/translate"


def _baidu_sign(appid, q, salt, key):
    s = appid + q + salt + key
    m = hashlib.md5()
    m.update(s.encode("utf-8"))
    return m.hexdigest()


def _baidu_translate(text, appid, key, src="zh", dst="en", timeout=10):
    if not appid or not key:
        raise RuntimeError("百度翻译 appid / key 未配置")
    salt = "".join(random.choices(string.ascii_letters + string.digits, k=10))
    sign = _baidu_sign(appid, text, salt, key)
    params = {
        "q": text,
        "from": src,
        "to": dst,
        "appid": appid,
        "salt": salt,
        "sign": sign,
    }
    r = requests.get(_BAIDU_ENDPOINT, params=params, timeout=timeout)
    r.raise_for_status()
    data = r.json()
    # 错误码：https://api.fanyi.baidu.com/product/113
    if "error_code" in data:
        # 52001/52002：系统错误；54000/54001：必填参数为空；58000：客户端 IP 非法；
        # 58001：译文语言不支持；58002：服务异常；90107：认证未通过或未生效
        raise RuntimeError(f"百度翻译错误 [{data.get('error_code')}]: {data.get('error_msg', '')}")
    items = data.get("trans_result") or []
    if not items:
        raise RuntimeError("百度翻译返回为空")
    return "".join(it.get("dst", "") for it in items)


# ----- 腾讯翻译君（v3 签名） -----

_TENCENT_HOST = "tmt.tencentcloudapi.com"
_TENCENT_ENDPOINT = "https://" + _TENCENT_HOST
_TENCENT_ACTION = "TextTranslate"
_TENCENT_VERSION = "2018-03-21"


def _tencent_sign_v3(secret_id, secret_key, payload, timestamp):
    """TC3-HMAC-SHA256。详见 https://cloud.tencent.com/document/api/551/15619"""
    import datetime
    # 1. 拼接规范请求串
    http_request_method = "POST"
    canonical_uri = "/"
    canonical_querystring = ""
    ct = "application/json; charset=utf-8"
    canonical_headers = "content-type:%s\nhost:%s\n" % (ct, _TENCENT_HOST)
    signed_headers = "content-type;host"
    payload_hash = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    canonical_request = (
        http_request_method + "\n" +
        canonical_uri + "\n" +
        canonical_querystring + "\n" +
        canonical_headers + "\n" +
        signed_headers + "\n" +
        payload_hash
    )
    # 2. 拼接待签名字符串
    date = datetime.datetime.utcfromtimestamp(timestamp).strftime("%Y-%m-%d")
    credential_scope = f"{date}/tmt/tc3_request"
    string_to_sign = (
        "TC3-HMAC-SHA256\n" +
        str(timestamp) + "\n" +
        credential_scope + "\n" +
        hashlib.sha256(canonical_request.encode("utf-8")).hexdigest()
    )
    # 3. 计算签名
    secret_date = _hmac_sha256(f"TC3{secret_key}".encode("utf-8"), date)
    secret_service = _hmac_sha256(secret_date, "tmt")
    secret_signing = _hmac_sha256(secret_service, "tc3_request")
    signature = _hmac_sha256_hex(secret_signing, string_to_sign)
    # 4. 拼装 Authorization
    authorization = (
        f"TC3-HMAC-SHA256 Credential={secret_id}/{credential_scope}, "
        f"SignedHeaders={signed_headers}, Signature={signature}"
    )
    return ct, authorization


def _hmac_sha256(key, msg):
    import hmac
    return hmac.new(key, msg.encode("utf-8"), hashlib.sha256).digest()


def _hmac_sha256_hex(key, msg):
    import hmac
    return hmac.new(key, msg.encode("utf-8"), hashlib.sha256).hexdigest()


def _tencent_translate(text, secret_id, secret_key, src="zh", dst="en", timeout=10):
    if not secret_id or not secret_key:
        raise RuntimeError("腾讯翻译 secret_id / secret_key 未配置")
    # 腾讯 API 要求 Source/Target 各 <= 5 字符，"zh" / "en" 即可
    payload_obj = {
        "SourceText": text,
        "Source": src,
        "Target": dst,
        "ProjectId": 0,
    }
    payload = json.dumps(payload_obj, ensure_ascii=False)
    timestamp = int(time.time())
    ct, authorization = _tencent_sign_v3(secret_id, secret_key, payload, timestamp)
    headers = {
        "Authorization": authorization,
        "Content-Type": ct,
        "Host": _TENCENT_HOST,
        "X-TC-Action": _TENCENT_ACTION,
        "X-TC-Timestamp": str(timestamp),
        "X-TC-Version": _TENCENT_VERSION,
    }
    r = requests.post(_TENCENT_ENDPOINT, data=payload.encode("utf-8"), headers=headers, timeout=timeout)
    r.raise_for_status()
    data = r.json()
    if "Response" not in data:
        raise RuntimeError(f"腾讯翻译返回异常: {data}")
    resp = data["Response"]
    if "Error" in resp:
        e = resp["Error"]
        raise RuntimeError(f"腾讯翻译错误 [{e.get('Code')}]: {e.get('Message', '')}")
    return resp.get("TargetText", "")


# ----- 对外统一入口 -----

def translate_one(text, cfg, src="zh", dst="en"):
    """根据 cfg["provider"] 选百度/腾讯，返回译文。空字符串直接返回。"""
    if not text or not text.strip():
        return ""
    provider = cfg.get("provider", "baidu")
    last_err = None
    retry = max(1, int(cfg.get("retry", 3)))
    for i in range(retry):
        try:
            if provider == "tencent":
                return _tencent_translate(
                    text,
                    cfg["tencent"]["secret_id"],
                    cfg["tencent"]["secret_key"],
                    src=src, dst=dst,
                )
            # 默认百度
            return _baidu_translate(
                text,
                cfg["baidu"]["appid"],
                cfg["baidu"]["key"],
                src=src, dst=dst,
            )
        except Exception as e:
            last_err = e
            time.sleep(0.5 * (i + 1))
    raise last_err


def translate_many(texts, cfg, src="zh", dst="en", qps=1, log_callback=None, stop_flag=None):
    """
    批量翻译，带 QPS 限速。
    :param texts: 待翻译字符串列表
    :param qps: 每秒请求数（默认 1，免费档够用）
    :param log_callback: 日志回调 fn(str)
    :param stop_flag: 函数，返回 True 时中断（线程安全：调用方自己控制）
    :return: 与输入等长的译文列表（失败位置保留空字符串）
    """
    qps = max(0.1, float(qps or 1))
    interval = 1.0 / qps
    out = []
    total = len(texts)
    for i, t in enumerate(texts, 1):
        if stop_flag and stop_flag():
            out.append("")
            continue
        try:
            out.append(translate_one(t, cfg, src=src, dst=dst))
        except Exception as e:
            out.append("")           # 失败留空，方便人工补
            if log_callback:
                log_callback(f"[{i}/{total}] 翻译失败: {e}")
        if log_callback and i % 10 == 0:
            log_callback(f"已翻译 {i}/{total}")
        time.sleep(interval)
    return out
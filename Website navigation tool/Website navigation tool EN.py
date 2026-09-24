# -*- coding: utf-8 -*-
"""
Visual Website Navigation Management Tool
Features: Import Excel / Import JSON / Edit data / Search & filter / Category sorting / Link sorting / Export pintree.json
"""

import json
import os
import sys
import time
import threading
import webbrowser
from collections import OrderedDict
from pathlib import Path
from urllib.parse import urlparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

try:
    import openpyxl
except ImportError:
    print("Missing dependency: openpyxl. Please run: pip install openpyxl")
    sys.exit(1)

try:
    import requests
except ImportError:
    requests = None  # Icon download needs requests; if missing, show a friendly hint without affecting other features

try:
    import io
    from PIL import Image as PILImage
except ImportError:
    io = None
    PILImage = None  # Unified PNG conversion needs Pillow; if missing, save in the original format


# ============== Config ==============
HEADERS = ["Website Name", "Category Path", "URL", "Description"]
SEPARATOR = ">"          # category-path separator
ICON_SHEET = "Category Icons"   # Auxiliary sheet mapping category path -> emoji (optional)
SAVE_JSON_NAME = "pintree.json"
APP_TITLE = "Visual Website Navigation Management Tool"
FAVICON_IM = "https://favicon.im/{d}"

# Local icon directory: assets/logo under the project root (the dir that contains index.html).
# Supports multiple run modes: run .py directly, .pyc from __pycache__, or a PyInstaller-built exe.
# Locate the project root by walking up from the actual script/exe location, so icons never land in the wrong dir.
def _detect_logo_dir():
    if getattr(sys, "frozen", False):          # PyInstaller build: use the exe's directory as the base
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:                                       # Script mode: use the script's directory as the base
        base = os.path.dirname(os.path.abspath(__file__))
    d = base
    for _ in range(6):                          # Walk up at most 6 levels to find the project root
        if os.path.isfile(os.path.join(d, "index.html")):
            return os.path.join(d, "assets", "logo")
        nxt = os.path.dirname(d)
        if nxt == d:
            break
        d = nxt
    # Fallback: if index.html isn't found, go up two levels assuming the common 'root/subdir/' layout
    return os.path.join(os.path.dirname(base), "assets", "logo")


LOGO_DIR = _detect_logo_dir()
try:
    os.makedirs(LOGO_DIR, exist_ok=True)
except Exception:
    pass
ICON_EXTS = (".png", ".jpg", ".gif", ".webp", ".ico")


# ============== Translation / Progress Window Helpers ==============
def _make_progress_window(parent, title):
    """返回进度窗。返回 (state, log_box)：
    state = {"stop": bool, "running": bool, "done": bool, "win": win,
             "bar": ttk.Progressbar, "pct_label": ttk.Label}
    - stop：置 True 请求停止（worker 自行检测）
    - running：worker 运行中置 True，结束置 False
    - done：worker 是否已正常结束
    - win：进度窗
    - bar/pct_label：进度条与百分比标签（配合 _set_progress 使用）
    进度窗在翻译进行中时若被点 X 关闭，会弹窗询问：已翻部分已实时写入
    增量缓存，安全关闭（后台线程可能被终止，但进度已落盘，不丢已翻成果）。
    """
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("560x400")
    win.transient(parent)

    # 顶部：进度条 + 百分比
    top = ttk.Frame(win, padding=(8, 8, 8, 0))
    top.pack(fill=tk.X)
    pct_label = ttk.Label(top, text="0%", anchor="e", width=5)
    pct_label.pack(side=tk.RIGHT)
    bar = ttk.Progressbar(top, mode="determinate", maximum=100, value=0)
    bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

    log_box = tk.Text(win, wrap="word", height=11)
    log_box.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    state = {"stop": False, "running": False, "done": False, "win": win,
             "log": log_box, "bar": bar, "pct_label": pct_label}
    btns = ttk.Frame(win, padding=4)
    btns.pack(fill=tk.X)
    ttk.Button(btns, text="Stop", command=lambda: state.__setitem__("stop", True)).pack(side=tk.RIGHT, padx=4)
    ttk.Button(btns, text="Hide (keep running)", command=win.iconify).pack(side=tk.RIGHT, padx=4)
    hint = ttk.Label(btns, text="Translated items are cached live; interruption won't lose progress", foreground="#888")
    hint.pack(side=tk.LEFT, padx=4)

    def _really_close():
        state["running"] = False
        state["done"] = True
        try:
            win.destroy()
        except Exception:
            pass

    def _on_close():
        if state["running"] and not state["done"]:
            if messagebox.askyesno(
                "Translation in progress",
                "Translation is not finished.\nThe translated part has been written to the cache file in real time; "
                "the next 'Translate Now' will resume automatically.\n\nClose now?",
                parent=win,
            ):
                _really_close()
        else:
            _really_close()

    win.protocol("WM_DELETE_WINDOW", _on_close)
    return state, log_box


def _log(win, log_box, msg):
    def _do():
        try:
            log_box.insert("end", msg + "\n")
            log_box.see("end")
        except Exception:
            pass
    try:
        win.after(0, _do)
    except Exception:
        pass


def _finish_progress(win, success=True):
    try:
        win.destroy()
    except Exception:
        pass


def _finish_when_idle(win, log_cb=None, delay=400):
    """延时安全关闭进度窗：先确保窗口仍存在再 destroy，避免线程与 UI 竞态。"""
    def _do():
        try:
            win.update_idletasks()
        except Exception:
            return
        try:
            win.destroy()
        except Exception:
            pass
    try:
        win.after(delay, _do)
    except Exception:
        pass


def _set_progress(state, done, total):
    """更新进度窗里的进度条与百分比。done/total 为翻译条数。
    在线程里调用即可，内部经 win.after 切回 UI 线程更新。
    total<=0（进度未知）或 done>=total 时直接把进度条拉满到 100%。"""
    if total <= 0:
        done, total = 1, 1   # 用满格表示『已完成/未知总量』
    pct = min(100.0, 100.0 * done / total)

    def _do():
        try:
            state["bar"]["value"] = pct
            state["pct_label"]["text"] = "%d%%" % round(pct)
        except Exception:
            pass
    try:
        state["win"].after(0, _do)
    except Exception:
        pass


def _build_cache_from_en_json(nodes):
    """从 pintree.en.json 反向提取『原文 → 译文』缓存（folder/link 的 title 与 link.description）。
    注：这里的『原文』是 en 文件里当前值，把它当 key 是错的——
    因为 en 文件里存的是译文。所以此函数实际上拿到的是『译文 → 译文』，并不能用作中文缓存。
    真正的中文缓存要从源 JSON（pintree.json）按 title 对应关系推断。
    """
    # 实际上无法仅凭 en.json 反推：因此缓存机制改为"读 pintree.json 的结构 + 已有 en.json 的译文匹配"。
    # 调用方应使用 _build_cache_from_pair(zh_json, en_json)。
    return {}


def _build_cache_from_pair(zh_nodes, en_nodes):
    """从中文 JSON 与已有的 en JSON 中，按『位置对应』抽取 title/desc 映射。
    走的是结构对齐（同位置、同 type、同层级顺序）而非字符串匹配，更稳。"""
    cache = {}

    def walk(zh_list, en_list):
        zh_list = list(zh_list or [])
        en_list = list(en_list or [])
        # 按出现顺序对齐（两边都按同一逻辑生成）
        i = j = 0
        while i < len(zh_list) and j < len(en_list):
            z = zh_list[i]
            e = en_list[j]
            if z.get("type") == e.get("type"):
                if z.get("type") == "folder":
                    if z.get("title"):
                        cache[z["title"]] = e.get("title", "")
                    walk(z.get("children", []), e.get("children", []))
                else:  # link
                    if z.get("title"):
                        cache[z["title"]] = e.get("title", "")
                    if z.get("description"):
                        cache[z["description"]] = e.get("description", "")
                i += 1; j += 1
            else:
                # 类型对不齐就跳过这一对（理论上不应发生）
                i += 1; j += 1

    walk(zh_nodes, en_nodes)
    return cache


def _build_en_json(store, mapping):
    """复用 store.to_nested_tree() 的结构，把 title/description 换成 mapping 里的译文。"""
    base = store.to_nested_tree()

    def apply(nodes):
        out = []
        for n in nodes:
            if n.get("type") == "folder":
                n = dict(n)
                t = n.get("title")
                if t and t in mapping and mapping[t]:
                    n["title"] = mapping[t]
                n["children"] = apply(n.get("children", []))
                out.append(n)
            elif n.get("type") == "link":
                n = dict(n)
                t = n.get("title")
                d = n.get("description")
                if t and t in mapping and mapping[t]:
                    n["title"] = mapping[t]
                if d and d in mapping and mapping[d]:
                    n["description"] = mapping[d]
                out.append(n)
        return out
    return apply(base)


# ----- On-disk JSON path conventions (zh/en pintree both live under json/) -----
def _project_root_dir():
    """网站导航工具/ 的上两级就是项目根（含 index.html）。"""
    return os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))


def _zh_json_path():
    """中文源文件：项目根/json/pintree.json"""
    return os.path.normpath(os.path.join(_project_root_dir(), "json", "pintree.json"))


def _en_json_path():
    """英文目标文件：项目根/json/pintree.en.json"""
    return os.path.normpath(os.path.join(_project_root_dir(), "json", "pintree.en.json"))


def _load_json_or_none(path):
    """读 json；不存在或解析失败返回 None。"""
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


# ----- Incremental translation cache (in tool dir; does not pollute the json/ output) -----
# 翻译过程中实时把『已确认的中文→英文』写进这个缓存文件；
# 中断/停止/关窗后已翻的不丢，下次「立刻翻译」自动读缓存续翻。
# 只有整批全部翻完，才据此生成干净的 pintree.en.json。
def _trans_cache_path():
    """增量翻译缓存文件：位于 网站导航工具/_translation_cache.json"""
    return os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "_translation_cache.json"))


def _load_trans_cache():
    """读增量翻译缓存，返回 {中文: 英文} dict（可空）。"""
    data = _load_json_or_none(_trans_cache_path())
    if not isinstance(data, dict):
        return {}
    out = {}
    for k, v in data.items():
        if isinstance(k, str) and isinstance(v, str) and v:
            out[k] = v
    return out


def _save_trans_cache(cache):
    """原子写增量翻译缓存。cache 为 {中文: 英文} dict。"""
    p = _trans_cache_path()
    try:
        tmp = p + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
        os.replace(tmp, p)          # 原子替换，避免写一半损坏
        return True
    except Exception:
        return False


def _walk_texts(zh_nodes, collect):
    """深度遍历 zh 树，把需要翻译的中文字符串交给 collect()。folder.title / link.title / link.description"""
    for n in zh_nodes or []:
        if n.get("type") == "folder":
            t = n.get("title", "")
            if t:
                collect(t)
            _walk_texts(n.get("children", []), collect)
        elif n.get("type") == "link":
            for k in ("title", "description"):
                v = n.get(k, "")
                if v:
                    collect(v)


def _collect_texts_unique(zh_nodes):
    """按出现顺序去重收集 zh 树中全部需翻译字符串。"""
    seen = set()
    order = []
    def collect(t):
        if t not in seen:
            seen.add(t)
            order.append(t)
    _walk_texts(zh_nodes, collect)
    return order


def _remap_zh_tree_to_en(zh_nodes, mapping):
    """以中文树为骨架，仅把 folder/link 的 title/description 换成 mapping 译文，
    其余字段（emoji/icon/addDate/url 等）原样保留。返回新树。"""
    out = []
    for n in zh_nodes or []:
        if n.get("type") == "folder":
            nn = dict(n)
            t = nn.get("title", "")
            if t and mapping.get(t):
                nn["title"] = mapping[t]
            nn["children"] = _remap_zh_tree_to_en(nn.get("children", []), mapping)
            out.append(nn)
        elif n.get("type") == "link":
            nn = dict(n)
            t = nn.get("title", "")
            if t and mapping.get(t):
                nn["title"] = mapping[t]
            d = nn.get("description", "")
            if d and mapping.get(d):
                nn["description"] = mapping[d]
            out.append(nn)
        else:
            out.append(n)
    return out


# ----- Batched translation + incremental cache persistence (worker thread) -----
# 翻译结果不再攒到最后一次性写 en.json，而是每翻完一小批就并入增量缓存并原子落盘。
# 这样中断/停止/关窗都不会丢已翻成果；en.json 只在『整批全部翻完』时才生成干净版本。
#
# 约定：返回的 mapping 形如 {中文: 译文或原文}，其中：
#   - 已确认译文（翻成功或命中缓存/旧译文）→ 译文/缓存值；
#   - 停止或失败时未能翻译的 → 用 src 原文占位（保证长度对齐，但不落盘成译文）。
# 调用方据此区分：只有 is_done 为真时才生成 en.json。
def _translate_batches_persist(
    order,            # 需翻译的中文列表（去重保序）
    cache,            # {中文: 英文}，函数内会被就地补充
    cfg,              # 翻译配置
    src="zh", dst="en",
    batch=20,         # 每批条数
    log_callback=None,
    stop_flag=None,
    progress_cb=None, # 每翻完一批回调 progress_cb(confirmed_total, len(need))
    cache_step=20,    # 每累计翻满多少条就原子写一次缓存
):
    """分批翻译 order，边翻边并入 cache 并周期落盘。
    返回 (mapping, cache, n_failed, stopped)：
      mapping = {src: 译文}（已确认的用译文；未翻成功的占位为 src 原文）
      cache   = 就地补充后的增量缓存 {中文: 确认译文}
      n_failed = 空译文条数（失败/未确认，非停止引起）
      stopped  = 是否被中途停止
    调用方据此决定：仅当 n_failed==0 and not stopped 时才生成干净的 en.json。
    注意：本函数不清空也不覆盖 cache 里已有的其它中文——只往里补新译文。
    """
    mapping = dict(cache)              # 先带上缓存里已有的（含缓存命中的）
    from translation import translate_many
    need = [t for t in order if not cache.get(t)]
    if not need:
        if progress_cb:
            progress_cb(len(order), len(order))   # 全部命中缓存 → 进度直接拉满
        return mapping, cache, 0, False    # 全部已在缓存命中，无失败、未停止
    n_failed = 0
    stopped = False
    done_cnt = 0          # 距上次缓存落盘的增量计数
    confirmed_total = 0   # 本次已确认译文的累计（含缓存命中部分，用于进度显示）
    qps = max(0.1, float(cfg.get("qps", 1) or 1))
    # 按 batch 切片逐批翻译，避免一次性把几百条全塞给 translate_many（也便于边翻边落盘）
    for start in range(0, len(need), batch):
        if stop_flag and stop_flag():
            stopped = True
            break
        chunk = need[start:start + batch]
        chunk_dst = translate_many(chunk, cfg, src=src, dst=dst,
                                   qps=qps,
                                   log_callback=log_callback,
                                   stop_flag=stop_flag)
        # 若中途停止：translate_many 会提前返回，长度可能不足 chunk，需兜底
        for k, v in zip(chunk, chunk_dst):
            if v:                                   # 翻成功 → 入缓存（真译文）
                cache[k] = v
                mapping[k] = v
                done_cnt += 1
                confirmed_total += 1
            else:                                   # 空串：失败或未翻到 → 计数并占位原文
                n_failed += 1
                mapping[k] = k
        if progress_cb:
            progress_cb(confirmed_total, len(need))
        # 本批内停止
        if stop_flag and stop_flag():
            stopped = True
            break
        # 周期落盘缓存（含中途停止前已确认部分）
        if done_cnt >= cache_step or (start + batch) >= len(need):
            _save_trans_cache(cache)
            done_cnt = 0
            if log_callback:
                log_callback("[Cache Saved] Translated %d/%d items, progress written to cache file" % (start + batch, len(need)))
    # 收尾再落一次，确保最后一批也进缓存
    if cache:
        _save_trans_cache(cache)
    return mapping, cache, n_failed, stopped


# ============== Utility Functions ==============
def domain_from_url(url):
    if not url:
        return ""
    try:
        netloc = urlparse(url).netloc
        if not netloc:
            return ""
        if netloc.startswith("www."):
            netloc = netloc[4:]
        return netloc.split(":")[0]
    except Exception:
        return ""


def logo_key(url):
    """从链接提取用于图标文件名的稳定 key：保留 www、去掉端口、转小写。"""
    if not url:
        return ""
    try:
        netloc = urlparse(url).netloc
        if not netloc:
            return ""
        return netloc.split(":")[0].lower()
    except Exception:
        return ""


def _logo_local_rel(url):
    """若本地已存在该站点的图标，返回相对路径（assets/logo/...），否则 None。"""
    key = logo_key(url)
    if not key:
        return None
    for ext in ICON_EXTS:
        p = os.path.join(LOGO_DIR, key + ext)
        if os.path.isfile(p):
            return "assets/logo/" + key + ext
    return None


def _logo_key_exists(key):
    """该 key（logo_key 结果）是否已有本地图标（任意扩展名）。"""
    if not key:
        return False
    return any(os.path.isfile(os.path.join(LOGO_DIR, key + ext)) for ext in ICON_EXTS)


def _cached_icon_usable(icon):
    """校验缓存 icon：指向 assets/logo/ 的本地路径若文件已不存在，视为失效。"""
    if not icon:
        return False
    if icon.startswith("assets/logo/"):
        # 相对路径基于项目根（= LOGO_DIR 向上两级，即 assets 的上一级）
        p = os.path.join(os.path.dirname(os.path.dirname(LOGO_DIR)), icon)
        return os.path.isfile(p)
    return True  # 远程 URL 或其它自定义路径，原样保留


def icon_for(url):
    """图标优先用本地下载的文件，否则回退远程图床，最后回退默认图标。"""
    local = _logo_local_rel(url)
    if local:
        return local
    d = domain_from_url(url)
    return FAVICON_IM.format(d=d) if d else "assets/default-icon.svg"


# ============== Icon Download Service (merged from website-favicon-downloader) ==============
class FaviconService:
    """批量/单个下载网站 favicon，存到 LOGO_DIR；多公共 API 兜底。"""

    def __init__(self, logo_dir=LOGO_DIR, log_callback=None):
        self.logo_dir = Path(logo_dir)
        self.logo_dir.mkdir(parents=True, exist_ok=True)
        self.session = requests.Session() if requests else None
        if self.session is not None:
            self.session.headers.update({
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                               "(KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            })
        self.api_endpoints = [
            "https://www.google.com/s2/favicons?domain={}&sz=256",
            "https://api.faviconkit.com/{}/256",
            "https://favicon.yandex.net/favicon/{}/256",
            "https://icons.duckduckgo.com/ip3/{}.ico",
            "https://icon.horse/icon/{}",
        ]
        self.log_callback = log_callback

    def log(self, msg):
        if self.log_callback:
            self.log_callback(msg)
        else:
            print(msg)

    def get_favicon_url(self, domain):
        """尝试多个公共 API 获取图标二进制内容，成功返回 (url, content)。"""
        if self.session is None:
            return None, None
        for api_url in self.api_endpoints:
            try:
                url = api_url.format(domain)
                resp = self.session.get(url, timeout=10)
                if resp.status_code == 200 and len(resp.content) > 0:
                    content = resp.content
                    ctype = resp.headers.get("content-type", "").lower()
                    # 大小不作为过滤条件：小于 70KB 的图标一律正常下载；
                    # 仅当响应既不是图片类型、内容也不像图片二进制时才跳过（防错误页）
                    if "image" in ctype or "octet-stream" in ctype or self._is_image_bytes(content):
                        return url, content
                    self.log(f"Skipped non-image content: {domain} ({ctype or 'unknown type'}, {len(content)} bytes)")
            except Exception as e:
                self.log(f"API failed {api_url.format(domain)}: {e}")
                continue
        return None, None

    def _is_image_bytes(self, content):
        """按文件头魔数判断内容是否为常见图片格式（含 SVG）。"""
        if not content:
            return False
        if content[:4] == b"\x89PNG":
            return True
        if content[:3] == b"\xff\xd8\xff":
            return True
        if content[:6] in (b"GIF87a", b"GIF89a"):
            return True
        if content[:4] in (b"\x00\x00\x01\x00", b"\x00\x00\x02\x00"):
            return True
        if content[:4] == b"RIFF" and content[8:12] == b"WEBP":
            return True
        if content[:2] == b"BM":
            return True
        head = content[:256].lstrip().lower()
        if head.startswith(b"<?xml") or head.startswith(b"<svg") or b"<svg" in head:
            return True
        return False

    def ext_for_data(self, content):
        """按文件头判断图像扩展名（增强：含 .ico 识别）。"""
        if content[:4] == b"\x89PNG":
            return ".png"
        if content[:3] == b"\xff\xd8\xff":
            return ".jpg"
        if content[:6] in (b"GIF87a", b"GIF89a"):
            return ".gif"
        if content[:4] in (b"\x00\x00\x01\x00", b"\x00\x00\x02\x00"):
            return ".ico"
        if content[:4] == b"RIFF" and content[8:12] == b"WEBP":
            return ".webp"
        return ".png"

    def save_as_png(self, key, content):
        """将图标内容统一转换为 PNG 保存为 {key}.png，返回最终文件名。
        Pillow 缺失或无法识别（如 SVG）时按原始格式回退保存。"""
        # 已是 PNG 直接落盘
        if content[:4] == b"\x89PNG":
            path = self.logo_dir / f"{key}.png"
            with open(path, "wb") as f:
                f.write(content)
            return path.name
        # 尝试用 Pillow 转成 PNG
        if PILImage is not None:
            try:
                img = PILImage.open(io.BytesIO(content))
                if img.mode not in ("RGBA", "RGB"):
                    img = img.convert("RGBA")
                path = self.logo_dir / f"{key}.png"
                img.save(path, "PNG")
                return path.name
            except Exception:
                pass
        # 回退：按原始格式保存（SVG 等无法转换的情况）
        ext = self.ext_for_data(content)
        path = self.logo_dir / f"{key}{ext}"
        with open(path, "wb") as f:
            f.write(content)
        return path.name

    def download_one(self, key):
        """下载单个站点图标（key = logo_key(url)），统一转为 PNG，成功返回 (True, url)。
        转换后的 PNG 若 ≤ 70 字节（无效/空图标），不保留，视为下载失败。"""
        favicon_url, content = self.get_favicon_url(key)
        if not content:
            self.log(f"No icon: {key}")
            return False, None
        try:
            self.remove_existing(key)  # 先清除旧文件，避免多扩展名并存
            name = self.save_as_png(key, content)
            path = self.logo_dir / name
            try:
                size = path.stat().st_size if path.is_file() else len(content)
            except OSError:
                size = len(content)
            if size <= 70:
                try:
                    if path.is_file():
                        path.unlink()
                except OSError:
                    pass
                self.log(f"Discarded: {key} icon too small after conversion ({size} bytes <= 70), not saved")
                return False, favicon_url
            self.log(f"Success: {key} -> {name}")
            return True, favicon_url
        except Exception as e:
            self.log(f"Save failed {key}: {e}")
            return False, favicon_url

    def remove_existing(self, key):
        """删除该 key 已有的任意扩展名图标，避免多个文件并存。"""
        for ext in ICON_EXTS:
            p = self.logo_dir / f"{key}{ext}"
            try:
                if p.is_file():
                    p.unlink()
            except Exception:
                pass


# ============== Data Model ==============
class DataStore:
    def __init__(self):
        self.items = []              # [dict{name, category, url, desc}]   # 顺序 = 显示顺序
        self.category_icons = {}     # 分类完整路径(用 SEPARATOR 连接) -> emoji
        self.current_file = None     # 已打开或保存的 xlsx 文件路径
        self.modified = False

    # --- Excel read/write ---
    def load_excel(self, path):
        wb = openpyxl.load_workbook(path)
        # 读「分类图标」辅助 sheet（若存在）
        self.category_icons = {}
        if ICON_SHEET in wb.sheetnames:
            ws_icon = wb[ICON_SHEET]
            for r in range(2, ws_icon.max_row + 1):
                cat = ws_icon.cell(row=r, column=1).value
                emoji = ws_icon.cell(row=r, column=2).value
                if cat and emoji:
                    self.category_icons[str(cat).strip()] = str(emoji).strip()
        ws = wb.active
        items = []
        for r in range(2, ws.max_row + 1):
            name = ws.cell(row=r, column=1).value or ""
            cat = ws.cell(row=r, column=2).value or ""
            url = ws.cell(row=r, column=3).value or ""
            desc = ws.cell(row=r, column=4).value or ""
            if not name and not url and not cat:
                continue
            items.append({
                "name": str(name).strip(),
                "category": str(cat).strip(),
                "url": str(url).strip(),
                "desc": str(desc).strip(),
            })
        self.items = items
        self.current_file = path
        self.modified = False
        return len(items)

    # --- JSON reading ---
    def load_json(self, path):
        """读取 pintree.json 嵌套树结构，展开为扁平 items 列表"""
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        items = []
        self.category_icons = {}

        def collect_icons(nodes, path_parts):
            for node in nodes:
                if node.get("type") == "folder":
                    title = node.get("title", "")
                    new_path = path_parts + [title] if title else path_parts
                    emoji = node.get("emoji")
                    if emoji:
                        self.category_icons[SEPARATOR.join(new_path)] = emoji
                    collect_icons(node.get("children", []), new_path)

        def walk(nodes, path_parts):
            for node in nodes:
                if node.get("type") == "folder":
                    title = node.get("title", "")
                    new_path = path_parts + [title] if title else path_parts
                    walk(node.get("children", []), new_path)
                elif node.get("type") == "link":
                    category = SEPARATOR.join(path_parts) if path_parts else ""
                    items.append({
                        "name": node.get("title", ""),
                        "category": category,
                        "url": node.get("url", ""),
                        "desc": node.get("description", ""),
                        "icon": node.get("icon", ""),  # 保留原图标地址，下载成功后会改写为本地路径
                    })

        collect_icons(data, [])
        walk(data, [])
        self.items = items
        self.current_file = None  # JSON 不绑定 xlsx 路径，保存时需另存
        self.modified = False
        return len(items)

    def save_excel(self, path=None):
        path = path or self.current_file
        if not path:
            raise ValueError("未指定保存路径")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Website Nav"
        for col_idx, h in enumerate(HEADERS, start=1):
            ws.cell(row=1, column=col_idx, value=h)
        for i, it in enumerate(self.items, start=2):
            ws.cell(row=i, column=1, value=it["name"])
            ws.cell(row=i, column=2, value=it["category"])
            ws.cell(row=i, column=3, value=it["url"])
            ws.cell(row=i, column=4, value=it["desc"])
        # 「分类图标」辅助 sheet：分类路径 -> emoji，便于在 Excel 里直接编辑一级/二级分类图标
        ws_icon = wb.create_sheet(ICON_SHEET)
        ws_icon.cell(row=1, column=1, value="Category Path")
        ws_icon.cell(row=1, column=2, value="Emoji")
        for i, (cat, emoji) in enumerate(self.category_icons.items(), start=2):
            ws_icon.cell(row=i, column=1, value=cat)
            ws_icon.cell(row=i, column=2, value=emoji)
        # 若图标表为空，则把现有 items 里出现的一级分类补进去（emoji 留空待用户填）
        if not self.category_icons:
            filled = set()
            for it in self.items:
                parts = [p.strip() for p in (it["category"] or "").split(SEPARATOR) if p.strip()]
                if parts and parts[0] not in filled:
                    filled.add(parts[0])
                    rr = ws_icon.max_row + 1
                    ws_icon.cell(row=rr, column=1, value=parts[0])
        wb.save(path)
        self.current_file = path
        self.modified = False
        return path

    # --- Data operations ---
    def add(self, item, index=None):
        if index is None or index >= len(self.items):
            self.items.append(item)
        else:
            self.items.insert(index, item)
        self.modified = True

    def update(self, index, item):
        self.items[index] = item
        self.modified = True

    def delete(self, indices):
        for i in sorted(indices, reverse=True):
            if 0 <= i < len(self.items):
                self.items.pop(i)
        self.modified = True

    def move_up(self, indices):
        indices = sorted(set(indices))
        for i in indices:
            if i > 0:
                self.items[i - 1], self.items[i] = self.items[i], self.items[i - 1]
        self.modified = True

    def move_down(self, indices):
        indices = sorted(set(indices), reverse=True)
        for i in indices:
            if i < len(self.items) - 1:
                self.items[i + 1], self.items[i] = self.items[i], self.items[i + 1]
        self.modified = True

    def move_top(self, indices):
        indices = sorted(set(indices))
        new_list = [self.items[i] for i in indices] + [self.items[i] for i in range(len(self.items)) if i not in indices]
        self.items = new_list
        self.modified = True

    def move_bottom(self, indices):
        indices = sorted(set(indices))
        new_list = [self.items[i] for i in range(len(self.items)) if i not in indices] + [self.items[i] for i in indices]
        self.items = new_list
        self.modified = True

    # --- Category stats ---
    def category_counts(self):
        """返回 OrderedDict { 分类完整路径字符串 -> 数量 }，按首次出现顺序"""
        counts = OrderedDict()
        for it in self.items:
            cat = it["category"] or "Uncategorized"
            counts[cat] = counts.get(cat, 0) + 1
        # 聚合父分类计数: "a>b" 的每个前缀都要计入
        agg = OrderedDict()

        def incr(key):
            agg[key] = agg.get(key, 0) + 1

        for it in self.items:
            cat = it["category"] or "Uncategorized"
            parts = [p.strip() for p in cat.split(SEPARATOR)]
            key = ""
            for j, p in enumerate(parts):
                key = p if j == 0 else key + SEPARATOR + p
                incr(key)
        return counts, agg

    # --- Category tree structure (for exporting JSON) ---
    def to_nested_tree(self):
        """按 items 顺序构建嵌套 folder/link 树结构，返回顶级 folder 列表"""
        base_ts = 1718526477999
        counter = [0]

        def nxt():
            counter[0] += 1
            return base_ts + counter[0]

        def make_folder(t, ch, emoji=None):
            node = {"type": "folder", "addDate": nxt(), "title": t, "children": ch}
            if emoji:
                node["emoji"] = emoji
            return node

        def make_link(it):
            # 优先本地图标（磁盘存在）；其次导入时缓存的 icon（本地路径需校验文件仍存在，
            # 已删除的视为失效并丢弃，回退图床）；最后回退图床
            cached = it.get("icon") or ""
            if not _cached_icon_usable(cached):
                cached = ""
            icon = _logo_local_rel(it.get("url", "")) or cached or icon_for(it["url"])
            return {
                "type": "link",
                "addDate": nxt(),
                "title": it["name"] or it["url"],
                "icon": icon,
                "url": it["url"],
                "description": it["desc"],
            }

        def icons_for_path(path_parts):
            """按「最长前缀」在分类图标表里找 emoji：先试完整路径，再逐级缩短到一级，都没有返回 None。
            例：['实用工具','AI','写作'] 命中顺序：'实用工具>AI>写作' → '实用工具>AI' → '实用工具'。"""
            for end in range(len(path_parts), 0, -1):
                key = SEPARATOR.join(path_parts[:end])
                v = self.category_icons.get(key)
                if v:
                    return v
            return None

        def attach_icons(node, path_parts):
            """给 folder 及其子树挂 emoji：本层优先（精确路径），子树逐层向下继承（若精确无则取父辈最长前缀）。
            这样一级分类能单独配，二级若没单独配就沿用父级 emoji，层级展示有图标可看。"""
            own = self.category_icons.get(SEPARATOR.join(path_parts)) or icons_for_path(path_parts)
            if own:
                node["emoji"] = own
            for ch in node.get("children", []):
                if ch.get("type") == "folder":
                    attach_icons(ch, path_parts + [ch["title"]])

        # 顶层字典：一级分类名 -> (folder_node, top_order_index)
        top_order = []
        top_map = {}

        for it in self.items:
            cat = it["category"] or "Uncategorized"
            parts = [p.strip() for p in cat.split(SEPARATOR) if p.strip()]
            if not parts:
                parts = ["Uncategorized"]

            top_name = parts[0]
            if top_name not in top_map:
                f = make_folder(top_name, [])
                top_map[top_name] = f
                top_order.append(f)
            folder_node = top_map[top_name]

            # 下钻中间级别
            for i in range(1, len(parts)):
                sub_name = parts[i]
                sub_node = None
                for ch in folder_node["children"]:
                    if ch.get("type") == "folder" and ch.get("title") == sub_name:
                        sub_node = ch
                        break
                if sub_node is None:
                    sub_node = make_folder(sub_name, [])
                    folder_node["children"].append(sub_node)
                folder_node = sub_node

            folder_node["children"].append(make_link(it))

        # 最后统一按分类图标表给所有 folder 挂 emoji
        for f in top_order:
            attach_icons(f, [f["title"]])

        return top_order


# ============== Main Window ==============
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1450x800")
        self.minsize(1200, 660)
        self.store = DataStore()
        self._trans_states = []          # 进行中的翻译进度窗 state 列表（用于关窗拦截）
        self.protocol("WM_DELETE_WINDOW", self._on_main_close)
        self._build_style()
        self._build_ui()
        self._refresh_all()

    # -------- Main-window close guard (warn while translating) --------
    def _register_trans_state(self, state):
        """登记一个翻译进度窗 state；结束时由 _trans_state_done 移除。"""
        self._trans_states.append(state)

    def _trans_state_done(self, state):
        try:
            if state in self._trans_states:
                self._trans_states.remove(state)
        except Exception:
            pass

    def _trans_active(self):
        """是否有翻译 worker 仍在运行（未 done 且 running）。"""
        for st in list(self._trans_states):
            if st.get("running") and not st.get("done"):
                return True
        return False

    def _on_main_close(self):
        if self._trans_active():
            if not messagebox.askyesno(
                "Translation in progress",
                "A translation task is still running in the background.\nProgress is cached live, but closing the window now may lose "
                "the current small batch being translated (up to ~20 items).\n\nIt is recommended to click 'Stop' in the progress window or wait for it to finish."
                "\nExit anyway?",
                parent=self,
            ):
                return
        self.destroy()

    # -------- Styles --------
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TButton", padding=6)
        style.configure("Toolbar.TButton", padding=(10, 6))
        style.configure("Sort.TButton", padding=(8, 4))
        style.configure("TLabelframe", padding=6)
        style.configure("Treeview", rowheight=24)
        style.configure("Treeview.Heading", font=("Microsoft YaHei", 9, "bold"))
        style.configure("TLabelframe.Label", font=("Microsoft YaHei", 9))

    # -------- UI --------
    def _build_ui(self):
        def btn(parent, text, cmd, style="TButton", width=None):
            kw = {"text": text, "command": cmd, "style": style}
            if width is not None:
                kw["width"] = width
            return ttk.Button(parent, **kw)

        # ========== Row 1: File/Edit main toolbar ==========
        toolbar = ttk.Frame(self, padding=6)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        file_box = ttk.LabelFrame(toolbar, text="File", padding=4)
        file_box.pack(side=tk.LEFT, padx=(0, 6))
        btn(file_box, "Import Excel", self.on_import_excel, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(file_box, "Import JSON", self.on_import_json, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(file_box, "Save", self.on_save, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(file_box, "Save As...", self.on_save_as, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        self.btn_export = btn(file_box, "Export JSON", self.on_export_json, style="Toolbar.TButton")
        self.btn_export.pack(side=tk.LEFT, padx=2)

        # ========== Translate / English ==========
        i18n_box = ttk.LabelFrame(toolbar, text="Translate", padding=4)
        i18n_box.pack(side=tk.LEFT, padx=6)
        btn(i18n_box, "Translate Settings", self.on_translation_settings, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(i18n_box, "Translate Now", self.on_translate_incremental, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(i18n_box, "Re-translate", self.on_retranslate_old, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(i18n_box, "Export EN JSON", self.on_export_english, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)

        edit_box = ttk.LabelFrame(toolbar, text="Edit", padding=4)
        edit_box.pack(side=tk.LEFT, padx=6)
        btn(edit_box, "Add", self.on_add, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(edit_box, "Edit", self.on_edit, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(edit_box, "Delete", self.on_delete, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(edit_box, "Open Link", self.on_open_link, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)

        # ========== 图标操作（Row 2，放在分类排序后的空白处） ==========

        # ========== Row 2: Sort/Search toolbar ==========
        sort_bar = ttk.Frame(self, padding=6)
        sort_bar.pack(side=tk.TOP, fill=tk.X)

        # 网址排序
        link_sort_box = ttk.LabelFrame(sort_bar, text="Link Sorting", padding=6)
        link_sort_box.pack(side=tk.LEFT, padx=(0, 8))
        self.sort_scope = tk.StringVar(value="cur")        # cur=当前筛选 / same=同分类 / all=全部
        ttk.Radiobutton(link_sort_box, text="Current Filter", variable=self.sort_scope, value="cur").pack(side=tk.LEFT)
        ttk.Radiobutton(link_sort_box, text="Same Category", variable=self.sort_scope, value="same").pack(side=tk.LEFT, padx=6)
        ttk.Radiobutton(link_sort_box, text="All", variable=self.sort_scope, value="all").pack(side=tk.LEFT)
        ttk.Separator(link_sort_box, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6)
        btn(link_sort_box, "Top", self.on_link_top, style="Sort.TButton", width=6).pack(side=tk.LEFT, padx=2)
        btn(link_sort_box, "Up", self.on_link_up, style="Sort.TButton", width=6).pack(side=tk.LEFT, padx=2)
        btn(link_sort_box, "Down", self.on_link_down, style="Sort.TButton", width=6).pack(side=tk.LEFT, padx=2)
        btn(link_sort_box, "Bottom", self.on_link_bottom, style="Sort.TButton", width=6).pack(side=tk.LEFT, padx=2)

        # 分类排序
        cat_sort_box = ttk.LabelFrame(sort_bar, text="Category Sorting", padding=6)
        cat_sort_box.pack(side=tk.LEFT, padx=(0, 8))
        btn(cat_sort_box, "Top", self.on_cat_top, style="Sort.TButton", width=6).grid(row=0, column=0, padx=2)
        btn(cat_sort_box, "Up", self.on_cat_up, style="Sort.TButton", width=6).grid(row=0, column=1, padx=2)
        btn(cat_sort_box, "Down", self.on_cat_down, style="Sort.TButton", width=6).grid(row=0, column=2, padx=2)
        btn(cat_sort_box, "Bottom", self.on_cat_bottom, style="Sort.TButton", width=6).grid(row=0, column=3, padx=2)
        more_btn = ttk.Menubutton(cat_sort_box, text="More ▾", direction="below", width=8)
        more_btn.grid(row=0, column=4, padx=(6, 2))
        more_menu = tk.Menu(more_btn, tearoff=False)
        more_btn.config(menu=more_menu)
        more_menu.add_command(label="By Name A→Z (Asc)", command=lambda: self.on_cat_sort("name", 1))
        more_menu.add_command(label="By Name Z→A (Desc)", command=lambda: self.on_cat_sort("name", -1))
        more_menu.add_command(label="By Count Low→High (Asc)", command=lambda: self.on_cat_sort("count", 1))
        more_menu.add_command(label="By Count High→Low (Desc)", command=lambda: self.on_cat_sort("count", -1))
        more_menu.add_separator()
        more_menu.add_command(label="Links in Category by Name (Asc)", command=lambda: self.on_links_in_cat_sort("name", 1))
        more_menu.add_command(label="Links in Category by Name (Desc)", command=lambda: self.on_links_in_cat_sort("name", -1))
        more_menu.add_command(label="Links in Category by URL (Asc)", command=lambda: self.on_links_in_cat_sort("url", 1))
        more_menu.add_command(label="Links in Category by URL (Desc)", command=lambda: self.on_links_in_cat_sort("url", -1))
        more_menu.add_separator()
        more_menu.add_command(label="Restore Table Order (Original)", command=self.on_cat_original)

        # 图标操作（占用分类排序右侧的空白）
        icon_box = ttk.LabelFrame(sort_bar, text="Icons", padding=6)
        icon_box.pack(side=tk.LEFT, padx=(0, 8))
        btn(icon_box, "Download Icons", self.on_download_icons, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)
        btn(icon_box, "Update Icons", self.on_update_icon, style="Toolbar.TButton").pack(side=tk.LEFT, padx=2)

        # 搜索框
        search_frame = ttk.LabelFrame(sort_bar, text="Search", padding=6)
        search_frame.pack(side=tk.RIGHT, padx=(8, 0), fill=tk.Y)
        ttk.Label(search_frame, text="Keyword:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._refresh_table())
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=26)
        self.search_entry.pack(side=tk.LEFT, padx=6)
        btn(search_frame, "Reset", self.on_reset_filter, style="Sort.TButton", width=6).pack(side=tk.LEFT)

        # ========== Main content: PanedWindow, left=category tree, right=table ==========
        # 使用 tk.PanedWindow 而非 ttk.PanedWindow，因为前者支持 minsize 且 sash 更稳定
        self.body = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=4)
        self.body.pack(fill=tk.BOTH, expand=True, padx=6, pady=(2, 0))

        # --- Left: category tree ---
        left_frame = tk.Frame(self.body, width=280, bg="#f5f5f5")
        left_head = ttk.Frame(left_frame)
        left_head.pack(fill=tk.X, pady=(4, 2))
        ttk.Label(left_head, text="Categories (with counts)", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(left_head, text="(Use the buttons on the right to reorder)", foreground="#666").pack(side=tk.LEFT, padx=6)

        self.cat_tree = ttk.Treeview(left_frame, show="tree")
        cat_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.cat_tree.yview)
        self.cat_tree.configure(yscrollcommand=cat_scroll.set)
        self.cat_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=2)
        cat_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.cat_tree.bind("<<TreeviewSelect>>", self.on_cat_select)
        self.cat_tree.bind("<Button-3>", self._cat_tree_on_right_click)

        self.body.add(left_frame, minsize=240)

        # --- Right: data table ---
        right_frame2 = tk.Frame(self.body, bg="#f5f5f5")
        # 统计条
        self.count_label = ttk.Label(right_frame2, text="Showing 0 / 0 items", foreground="#555")
        self.count_label.pack(anchor=tk.W, pady=(4, 2))
        # 表格
        table_frame = ttk.Frame(right_frame2)
        table_frame.pack(fill=tk.BOTH, expand=True)
        self.table = ttk.Treeview(table_frame, columns=("name", "category", "url", "desc"),
                                  show="headings", selectmode="extended")
        for col, h, w in [("name", "Website Name", 220), ("category", "Category Path", 180),
                          ("url", "URL", 320), ("desc", "Description", 520)]:
            self.table.heading(col, text=h)
            self.table.column(col, width=w, minwidth=100, anchor=tk.W, stretch=True)
        vscr = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.table.yview)
        hscr = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.table.xview)
        self.table.configure(yscrollcommand=vscr.set, xscrollcommand=hscr.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        vscr.grid(row=0, column=1, sticky="ns")
        hscr.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.table.bind("<Double-1>", lambda e: self.on_edit())
        self.body.add(right_frame2, minsize=500)

        # --- Status bar ---
        self.status_var = tk.StringVar(value="Ready. Click 'Import Excel' to start, or 'Export JSON' to generate pintree.json.")
        status = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=(8, 3))
        status.pack(side=tk.BOTTOM, fill=tk.X)

        self._data_rowids = []  # 当前显示项（筛选后）对应 store.items 的下标列表

    # ============== Refresh ==============
    def _refresh_all(self):
        self._refresh_categories()
        self._refresh_table()

    def _cat_root_items(self):
        """返回顶级分类在 store 中首次出现的顺序"""
        seen = OrderedDict()
        for i, it in enumerate(self.store.items):
            cat = it["category"] or "Uncategorized"
            parts = [p.strip() for p in cat.split(SEPARATOR) if p.strip()]
            root = parts[0] if parts else "Uncategorized"
            if root not in seen:
                seen[root] = True
        return list(seen.keys())

    def _save_layout(self):
        """保存窗口几何与分栏位置，供刷新后恢复，防止界面收缩"""
        geo = self.geometry()
        sash_pos = None
        try:
            sash_pos = self.body.sash_coord(0)[0]
        except Exception:
            pass
        return geo, sash_pos

    def _restore_layout(self, geo=None, sash_pos=None):
        """刷新后恢复窗口几何与分栏位置。
        用 after_idle 延迟执行，等待本次刷新触发的布局全部完成后再恢复，
        否则 sash 位置会被后续布局阶段覆盖（表现为左侧面板收缩）。
        """
        def _apply():
            if geo:
                try:
                    self.geometry(geo)
                except Exception:
                    pass
            if sash_pos is not None:
                try:
                    self.update_idletasks()
                    self.body.sash_place(0, sash_pos, 0)
                except Exception:
                    pass
        self.after_idle(_apply)
        # 再追加一次延迟恢复，双保险：有些场景下两轮 idle 才稳定
        self.after(120, _apply)

    def _refresh_categories(self):
        # 保存当前选中
        sel = self.cat_tree.selection()
        sel_val = self.cat_tree.item(sel[0], "values")[0] if sel else None

        # 保存展开状态（用分类完整路径作为稳定键）
        open_set = set()
        def collect_open(node_iid):
            for child in self.cat_tree.get_children(node_iid):
                vals = self.cat_tree.item(child, "values")
                if vals and self.cat_tree.item(child, "open"):
                    open_set.add(vals[0])
                collect_open(child)
        collect_open("")

        # 保存 sash 位置，防止刷新后左侧面板收缩
        sash_pos = self._save_layout()[1]

        self.cat_tree.delete(*self.cat_tree.get_children())

        # 全部站点
        total = len(self.store.items)
        root_iid = self.cat_tree.insert(
            "", tk.END, text=f"All Sites ({total})",
            values=("__ALL__",), open=True, tags=("cat",)
        )

        # 聚合计数
        _, agg = self.store.category_counts()

        # 顶级分类顺序
        top_roots = self._cat_root_items()

        def insert_node(parent_iid, parent_prefix, level_names_left, order_list):
            for name in order_list:
                full = name if not parent_prefix else parent_prefix + SEPARATOR + name
                cnt = agg.get(full, 0)
                # 判断该 full 是否拥有子项(直接属于它+孙子)
                # 存在后代前缀 full+SEPARATOR 就说明有子文件夹
                has_children = any(k.startswith(full + SEPARATOR) for k in agg.keys())
                emoji = self._emoji_for_full(full)
                emoji_txt = (emoji + " ") if emoji else ""
                node_iid = self.cat_tree.insert(
                    parent_iid, tk.END,
                    text=f"{'  ' if parent_iid != '' else ''}  {emoji_txt}{name} ({cnt})",
                    values=(full,)
                )
                if has_children:
                    # 取属于该 full 的直接下一级名字
                    children_names = OrderedDict()
                    for k in agg.keys():
                        if k.startswith(full + SEPARATOR):
                            rest = k[len(full + SEPARATOR):]
                            first = rest.split(SEPARATOR)[0]
                            children_names[first] = True
                    insert_node(node_iid, full, [], list(children_names.keys()))

        insert_node(root_iid, "", [], top_roots)

        # 展开 root
        if root_iid:
            self.cat_tree.item(root_iid, open=True)

        # 恢复展开状态
        def apply_open(node_iid):
            for child in self.cat_tree.get_children(node_iid):
                vals = self.cat_tree.item(child, "values")
                if vals and vals[0] in open_set:
                    self.cat_tree.item(child, open=True)
                apply_open(child)
        apply_open("")

        # 恢复选中（递归搜索任意层级，并展开祖先节点保证选中项可见）
        if sel_val:
            def find_and_select(node_iid):
                for child in self.cat_tree.get_children(node_iid):
                    try:
                        vals = self.cat_tree.item(child, "values")
                    except Exception:
                        continue
                    if vals and vals[0] == sel_val:
                        return [child]
                    found = find_and_select(child)
                    if found:
                        # 展开当前祖先节点，让深层选中项保持可见
                        try:
                            self.cat_tree.item(child, open=True)
                        except Exception:
                            pass
                        return [child] + found
                return None
            path = find_and_select("")
            if path:
                target = path[-1]
                self.cat_tree.selection_set(target)
                self.cat_tree.see(target)

        # 恢复 sash 位置，保持左侧面板宽度
        if sash_pos is not None:
            self._restore_layout(sash_pos=sash_pos)

    def _refresh_table(self):
        # 筛选：分类 + 搜索关键词
        cat_sel = self.cat_tree.selection()
        cat_filter = None
        if cat_sel:
            v = self.cat_tree.item(cat_sel[0], "values")
            if v and v[0] != "__ALL__":
                cat_filter = v[0]

        kw = self.search_var.get().strip().lower()

        self.table.delete(*self.table.get_children())
        self._data_rowids = []

        for idx, it in enumerate(self.store.items):
            cat = it["category"] or "Uncategorized"
            if cat_filter:
                # cat= "a>b"，筛选 a 时应命中；筛选 a>b 时应命中；筛选 a>b>c 时仅 a>b>c 命中
                if not (cat == cat_filter or cat.startswith(cat_filter + SEPARATOR)):
                    continue
            if kw:
                if (kw not in it["name"].lower()
                        and kw not in it["url"].lower()
                        and kw not in it["desc"].lower()
                        and kw not in it["category"].lower()):
                    continue
            rowid = self.table.insert(
                "", tk.END,
                values=(it["name"], it["category"], it["url"], it["desc"])
            )
            self._data_rowids.append((idx, rowid))

        total = len(self.store.items)
        shown = len(self._data_rowids)
        self.count_label.config(text=f"Showing {shown} / {total} items")

    # ============== Toolbar actions ==============
    def on_import_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            cnt = self.store.load_excel(path)
        except Exception as e:
            messagebox.showerror("Import Failed", str(e))
            return
        self._refresh_all()
        self.set_status(f"Imported {path} ({cnt} items)")
        # 默认选中"全部站点"
        root = self.cat_tree.get_children()
        if root:
            self.cat_tree.selection_set(root[0])

    def on_import_json(self):
        path = filedialog.askopenfilename(
            title="Select JSON File",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            cnt = self.store.load_json(path)
        except Exception as e:
            messagebox.showerror("Import Failed", str(e))
            return
        self._refresh_all()
        self.set_status(f"Imported {path} ({cnt} items)")
        # 默认选中"全部站点"
        root = self.cat_tree.get_children()
        if root:
            self.cat_tree.selection_set(root[0])

    def on_save(self):
        if not self.store.current_file:
            self.on_save_as()
            return
        try:
            p = self.store.save_excel()
            self.set_status(f"Saved to {p}")
        except Exception as e:
            messagebox.showerror("Save Failed", str(e))

    def on_save_as(self):
        init = self.store.current_file or "Website Nav.xlsx"
        init_dir = os.path.dirname(init) if os.path.dirname(init) else os.getcwd()
        init_name = os.path.basename(init) or "Website Nav.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save Excel As",
            defaultextension=".xlsx",
            initialdir=init_dir,
            initialfile=init_name,
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if not path:
            return
        try:
            p = self.store.save_excel(path)
            self.set_status(f"Saved as {p}")
        except Exception as e:
            messagebox.showerror("Save Failed", str(e))

    # ---- Export JSON ----
    def on_export_json(self):
        if not self.store.items:
            messagebox.showwarning("No Data", "Please import an Excel file or add data first.")
            return
        # 默认输出路径:与 xlsx 同目录下的 json/pintree.json；若无则当前程序目录下 json/pintree.json
        default_dir = ""
        if self.store.current_file:
            default_dir = os.path.dirname(self.store.current_file)
        if not default_dir:
            default_dir = os.getcwd()
        default_path = os.path.join(default_dir, "json", SAVE_JSON_NAME)

        path = filedialog.asksaveasfilename(
            title="Export pintree.json",
            defaultextension=".json",
            initialdir=os.path.dirname(default_path),
            initialfile=SAVE_JSON_NAME,
            filetypes=[("JSON Files", "*.json")],
        )
        if not path:
            return
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            data = self.store.to_nested_tree()
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))
            return

        # 统计（含本地图标引用情况）
        total_links = 0
        total_folders = 0
        local_icons = 0

        def walk(nodes):
            nonlocal total_links, total_folders, local_icons
            for n in nodes:
                if n.get("type") == "link":
                    total_links += 1
                    if str(n.get("icon", "")).startswith("assets/logo/"):
                        local_icons += 1
                elif n.get("type") == "folder":
                    total_folders += 1
                    walk(n.get("children", []))

        walk(data)
        self.set_status(
            f"Exported {path} ({len(data)} top-level folders / {total_folders} total folders / {total_links} links / "
            f"{local_icons} updated to local icons)"
        )
        messagebox.showinfo(
            "Export Successful",
            f"Successfully generated:\n{path}\n\nTop-level folders: {len(data)}  Total folders: {total_folders}  Links: {total_links}\n"
            f"Icon URLs updated: {local_icons} now reference local PNGs (assets/logo/),"
            f"the rest remain remote.",
        )

    # ---- Translate / Export English ----
    def on_translation_settings(self):
        """打开翻译设置对话框：选择服务商、填 Key、调速等。"""
        try:
            from translation import load_config, save_config
        except ImportError:
            messagebox.showerror("Missing Dependency", "The translation.py module is missing; please ensure it is in the same directory.")
            return
        cfg = load_config()

        win = tk.Toplevel(self)
        win.title("Translation Settings")
        win.geometry("520x420")
        win.transient(self)
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        provider_var = tk.StringVar(value=cfg.get("provider", "baidu"))
        baidu_appid = tk.StringVar(value=cfg["baidu"].get("appid", ""))
        baidu_key = tk.StringVar(value=cfg["baidu"].get("key", ""))
        tencent_sid = tk.StringVar(value=cfg["tencent"].get("secret_id", ""))
        tencent_skey = tk.StringVar(value=cfg["tencent"].get("secret_key", ""))
        qps_var = tk.StringVar(value=str(cfg.get("qps", 1)))

        ttk.Label(frm, text="Provider:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frm, textvariable=provider_var, state="readonly",
            values=("baidu", "tencent"), width=14,
        ).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Separator(frm, orient="horizontal").grid(row=1, column=0, columnspan=2, sticky="we", pady=8)

        ttk.Label(frm, text="Baidu APPID:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=baidu_appid, width=42).grid(row=2, column=1, sticky="we", pady=4)
        ttk.Label(frm, text="Baidu Key:").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=baidu_key, width=42, show="*").grid(row=3, column=1, sticky="we", pady=4)

        ttk.Separator(frm, orient="horizontal").grid(row=4, column=0, columnspan=2, sticky="we", pady=8)

        ttk.Label(frm, text="Tencent SecretId:").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=tencent_sid, width=42).grid(row=5, column=1, sticky="we", pady=4)
        ttk.Label(frm, text="Tencent SecretKey:").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=tencent_skey, width=42, show="*").grid(row=6, column=1, sticky="we", pady=4)

        ttk.Separator(frm, orient="horizontal").grid(row=7, column=0, columnspan=2, sticky="we", pady=8)

        ttk.Label(frm, text="QPS (requests per second):").grid(row=8, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=qps_var, width=10).grid(row=8, column=1, sticky="w", pady=4)

        info = ttk.LabelFrame(frm, text="Notes", padding=6)
        info.grid(row=9, column=0, columnspan=2, sticky="we", pady=8)
        ttk.Label(
            info,
            text="Baidu Translate API: https://api.fanyi.baidu.com/api/trans/vip/translate\n"
                 "Tencent TMT: https://cloud.tencent.com/product/tmt\n"
                 "Baidu gives ~2M chars free for new users; Tencent gives 5M chars/month free for new users.",
            foreground="#374151", justify="left",
        ).pack(anchor="w")

        btns = ttk.Frame(frm)
        btns.grid(row=10, column=0, columnspan=2, pady=8)

        def on_save():
            try:
                qps = float(qps_var.get())
            except ValueError:
                messagebox.showerror("Error", "QPS must be a number")
                return
            new_cfg = {
                "provider": provider_var.get(),
                "baidu": {"appid": baidu_appid.get().strip(), "key": baidu_key.get().strip()},
                "tencent": {"secret_id": tencent_sid.get().strip(), "secret_key": tencent_skey.get().strip()},
                "qps": qps,
                "retry": cfg.get("retry", 3),
                "last_error": cfg.get("last_error", ""),
            }
            save_config(new_cfg)
            messagebox.showinfo("Saved", "Translation config saved to json/_translation_config.json.", parent=win)
            win.destroy()

        def on_test():
            """快速测一下 key 是否有效。"""
            try:
                from translation import translate_one
                test_cfg = {
                    "provider": provider_var.get(),
                    "baidu": {"appid": baidu_appid.get().strip(), "key": baidu_key.get().strip()},
                    "tencent": {"secret_id": tencent_sid.get().strip(), "secret_key": tencent_skey.get().strip()},
                    "retry": 1,
                }
                result = translate_one("你好", test_cfg)
                messagebox.showinfo("Test Succeeded", f"Source 'Hello' -> Translation '{result}'", parent=win)
            except Exception as e:
                messagebox.showerror("Test Failed", str(e), parent=win)

        ttk.Button(btns, text="Test Connection", command=on_test).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Save", command=on_save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side=tk.LEFT, padx=6)

        frm.columnconfigure(1, weight=1)

    # ============ Translation common ============
    def _translation_ready(self, cfg):
        """检查 key 是否齐全，缺则弹窗并返回 False。"""
        if cfg["provider"] == "baidu" and (not cfg["baidu"]["appid"] or not cfg["baidu"]["key"]):
            messagebox.showwarning("Not Configured", "Baidu translation key is empty. Please click 'Translate Settings' to configure it first.")
            return False
        if cfg["provider"] == "tencent" and (not cfg["tencent"]["secret_id"] or not cfg["tencent"]["secret_key"]):
            messagebox.showwarning("Not Configured", "Tencent translation key is empty. Please click 'Translate Settings' to configure it first.")
            return False
        return True

    def _run_zh_file_translate(self, force_all, win_title):
        """文件级翻译核心：以磁盘 json/pintree.json 为中文基准。

        - force_all=False：增量——跳过已确认译文（含增量缓存 _translation_cache.json
          与已有 en.json 的旧译文），只翻缺失/新增的中文。
        - force_all=True：全量重译——忽略旧译文，清空增量缓存后全部重翻覆盖。
        翻译过程实时把已确认译文写入 网站导航工具/_translation_cache.json 增量缓存；
        **只有整批全部翻完且无失败/停止时**，才据此生成干净的 json/pintree.en.json。
        中途停止/中断只保留缓存进度，不产出半成品 en.json（避免中文残留）。
        中文源文件缺失时返回 False 并提示。
        """
        try:
            from translation import load_config
        except ImportError:
            messagebox.showerror("Missing Dependency", "The translation.py module is missing; please ensure it is in the same directory.")
            return False
        cfg = load_config()
        if not self._translation_ready(cfg):
            return False

        zh_path = _zh_json_path()
        if not os.path.isfile(zh_path):
            messagebox.showwarning("Missing Chinese Source", "Chinese source file not found:\n" + zh_path +
                                   "\n\nPlease import data in the tool and 'Export JSON' to this path first.")
            return False
        try:
            with open(zh_path, "r", encoding="utf-8") as f:
                zh_data = json.load(f)
        except Exception as e:
            messagebox.showerror("Read Failed", "Failed to read " + zh_path + " error:\n" + str(e))
            return False

        # 全部需翻译的中文（去重保序）
        all_texts = _collect_texts_unique(zh_data)
        if not all_texts:
            messagebox.showinfo("No Content", "The Chinese file contains no text to translate.")
            return False

        en_path = _en_json_path()

        # 构建『已确认译文』缓存：增量缓存 优先，其次已有 en.json 的旧译文。
        cache = _load_trans_cache()
        if force_all:
            # 全量重译：丢弃旧译文与旧缓存，从零翻
            cache = {}
            _save_trans_cache(cache)
        else:
            # 增量：把旧 en.json 里能对齐的译文并入缓存，避免重复翻
            if os.path.isfile(en_path):
                en_data = _load_json_or_none(en_path)
                if en_data is not None:
                    old = _build_cache_from_pair(zh_data, en_data)
                    for k, v in old.items():
                        cache.setdefault(k, v)   # 增量缓存优先，不覆盖更新过的译文
        need = [t for t in all_texts if not cache.get(t)]

        progress_win, log_box = _make_progress_window(self, win_title)
        pwin = progress_win["win"]
        progress_win["running"] = True   # worker 启动前标记运行中
        self._register_trans_state(progress_win)

        def _logm(m):
            _log(pwin, log_box, m)

        def _prog(done, total):
            _set_progress(progress_win, done, total)

        def _worker():
            try:
                if need:
                    _logm("To translate: %d items (cache hit: %d)" % (len(need), len(all_texts) - len(need)))
                    mapping, _cache, n_failed, stopped = _translate_batches_persist(
                        need, cache, cfg, src="zh", dst="en",
                        batch=20, log_callback=_logm,
                        progress_cb=_prog,
                        stop_flag=lambda: progress_win["stop"],
                    )
                    # cache 已在函数内就地补充，统一为最新
                    mapping = {t: cache.get(t) or t for t in all_texts}
                    if stopped:
                        _logm("⏹ Stopped. Confirmed translations saved to the incremental cache; en.json not generated.")
                        _logm("The next 'Translate Now' will resume from the cache automatically.")
                        return
                    if n_failed:
                        _logm("⚠ %d items failed to translate; en.json not generated (click 'Translate Now' again to retry)." % n_failed)
                        return
                else:
                    _logm("All items hit the cache; no API call needed.")
                    mapping = {t: cache.get(t) or t for t in all_texts}
                    _prog(1, 0)   # 拉满进度条

                # 全部翻完且无失败 → 生成干净 en.json
                en_new = _remap_zh_tree_to_en(zh_data, mapping)
                os.makedirs(os.path.dirname(en_path), exist_ok=True)
                with open(en_path, "w", encoding="utf-8") as f:
                    json.dump(en_new, f, ensure_ascii=False, indent=2)
                _logm("Written: " + en_path)
                _logm("✅ Done!")
            except Exception as e:
                _logm("❌ Failed: " + str(e))
                _save_trans_cache(cache)   # 无论如何尽量保存已确认译文
            finally:
                progress_win["running"] = False
                progress_win["done"] = True
                self._trans_state_done(progress_win)
                _finish_when_idle(pwin, _logm, delay=600)

        threading.Thread(target=_worker, daemon=True).start()
        return True

    def on_translate_incremental(self):
        """立刻翻译：以磁盘 pintree.json 为中文基准，对比 en.json，只翻新增/缺失部分。"""
        self._run_zh_file_translate(force_all=False, win_title="Translate Now")

    def on_retranslate_old(self):
        """重译旧内容：以磁盘 pintree.json 为中文基准，忽略旧译文全量重新翻译并覆盖 en.json。"""
        if not messagebox.askyesno(
            "Confirm Re-translation",
            "This will use disk json/pintree.json as the Chinese base, clear the existing translation cache and old translations, "
            "re-call the API for all content, and overwrite json/pintree.en.json.\n"
            "Existing translations and cache progress will be lost. Continue?",
        ):
            return
        self._run_zh_file_translate(force_all=True, win_title="Re-translate")

    def on_export_english(self):
        """导出英文 json：基于内存当前数据（含刚改动/新增条目）翻译后，
        只生成 json/pintree.en.json（不再连带生成 en.html）。"""
        if not self.store.items:
            messagebox.showwarning("No Data", "Please import an Excel file or add data first.")
            return
        try:
            from translation import load_config, translate_many
        except ImportError:
            messagebox.showerror("Missing Dependency", "The translation.py module is missing; please ensure it is in the same directory.")
            return
        cfg = load_config()
        if not self._translation_ready(cfg):
            return

        # 收集内存数据里所有需翻译字符串（去重保序）
        order = []
        _seen = set()
        for it in self.store.items:
            cat = it.get("category") or ""
            for part in [p.strip() for p in cat.split(SEPARATOR) if p.strip()]:
                if part not in _seen:
                    _seen.add(part)
                    order.append(part)
            for fld in ("name", "desc"):
                v = it.get(fld) or ""
                if v and v not in _seen:
                    _seen.add(v)
                    order.append(v)

        # 已确认译文缓存：增量缓存 优先，其次磁盘 en.json 的旧译文
        cache = _load_trans_cache()
        en_path = _en_json_path()
        if os.path.isfile(en_path) and os.path.isfile(_zh_json_path()):
            try:
                with open(en_path, "r", encoding="utf-8") as f:
                    en_data = json.load(f)
                with open(_zh_json_path(), "r", encoding="utf-8") as f:
                    zh_data = json.load(f)
                old = _build_cache_from_pair(zh_data, en_data)
                for k, v in old.items():
                    cache.setdefault(k, v)
            except Exception:
                pass

        need = [t for t in order if not cache.get(t)]

        progress_win, log_box = _make_progress_window(self, "Export EN JSON")
        pwin = progress_win["win"]
        progress_win["running"] = True
        self._register_trans_state(progress_win)

        def _logm(m):
            _log(pwin, log_box, m)

        def _prog(done, total):
            _set_progress(progress_win, done, total)

        def _worker():
            try:
                if need:
                    _logm("To translate: %d items (cache hit: %d)" % (len(need), len(order) - len(need)))
                    _mapping, _c, n_failed, stopped = _translate_batches_persist(
                        need, cache, cfg, src="zh", dst="en",
                        batch=20, log_callback=_logm,
                        progress_cb=_prog,
                        stop_flag=lambda: progress_win["stop"],
                    )
                    if stopped:
                        _logm("⏹ Stopped. Confirmed translations saved to the incremental cache; en.json not generated.")
                        _logm("You can click 'Export EN JSON' again to resume from the cache.")
                        return
                    if n_failed:
                        _logm("⚠ %d items failed; en.json not generated (click 'Export EN JSON' again to retry)." % n_failed)
                        return
                else:
                    _logm("All items hit the cache; no API call needed.")
                    _prog(1, 0)   # 拉满进度条

                # 全部确认 → 生成干净 en.json（mapping 覆盖全部 order）
                mapping = {t: cache.get(t) or t for t in order}
                data_en = _build_en_json(self.store, mapping)
                os.makedirs(os.path.dirname(en_path), exist_ok=True)
                with open(en_path, "w", encoding="utf-8") as f:
                    json.dump(data_en, f, ensure_ascii=False, indent=2)
                _logm("Written: " + en_path)
                _logm("✅ Done!")
            except Exception as e:
                _logm("❌ Failed: " + str(e))
                _save_trans_cache(cache)
            finally:
                progress_win["running"] = False
                progress_win["done"] = True
                self._trans_state_done(progress_win)
                _finish_when_idle(pwin, delay=600)

        threading.Thread(target=_worker, daemon=True).start()


    # ---- Add/Edit/Delete ----
    def _sel_store_indices(self, scope=None):
        """根据 scope(cur/same/all) 返回 store 中选中项的 index 列表"""
        scope = scope or self.sort_scope.get()
        sel_rows = self.table.selection()
        rowid_to_idx = {r: i for i, r in self._data_rowids}
        if scope == "cur":
            # 当前显示项中的选中
            return [rowid_to_idx[r] for r in sel_rows if r in rowid_to_idx]
        if scope == "all":
            # 全部
            return [rowid_to_idx[r] for r in sel_rows if r in rowid_to_idx]
        if scope == "same":
            # 同分类：取选中项的分类，返回该分类在 store 中所有选中的条目
            idxs = [rowid_to_idx[r] for r in sel_rows if r in rowid_to_idx]
            if not idxs:
                return []
            cat = self.store.items[idxs[0]]["category"]
            return [i for i in idxs if self.store.items[i]["category"] == cat]
        return []

    def on_add(self):
        dlg = ItemDialog(self, title="Add Website")
        if dlg.result:
            self.store.add(dlg.result)
            self._refresh_all()
            # 新分类可能改变树结构，保持选中全部站点便于查看
            root = self.cat_tree.get_children()
            if root:
                self.cat_tree.selection_set(root[0])
            self.set_status("1 item added.")

    def on_edit(self):
        idxs = self._sel_store_indices()
        if len(idxs) != 1:
            messagebox.showinfo("Hint", "Please select a row in the table on the right to edit.")
            return
        dlg = ItemDialog(self, title="Edit Website", item=self.store.items[idxs[0]])
        if dlg.result:
            self.store.update(idxs[0], dlg.result)
            self._refresh_all()
            self.set_status("Edited.")

    def on_delete(self):
        idxs = self._sel_store_indices()
        if not idxs:
            messagebox.showinfo("Hint", "Please select the row(s) to delete first.")
            return
        if not messagebox.askyesno("Confirm", f"Delete the selected {len(idxs)} record(s)?"):
            return
        self.store.delete(idxs)
        self._refresh_all()
        self.set_status(f"Deleted {len(idxs)} item(s).")

    def on_open_link(self):
        idxs = self._sel_store_indices()
        if not idxs:
            messagebox.showinfo("Hint", "Please select the row(s) to open first.")
            return
        ok = 0
        for i in idxs:
            url = self.store.items[i]["url"]
            if url and (url.startswith("http://") or url.startswith("https://")):
                webbrowser.open(url)
                ok += 1
        self.set_status(f"Attempted to open {ok} link(s) in the browser.")

    # ---- Link sorting ----
    def _apply_link_sort(self, func):
        idxs = self._sel_store_indices()
        if not idxs:
            messagebox.showinfo("Hint", "Please select the item(s) to sort in the table on the right first.")
            return
        # 保存窗口几何与 sash 位置
        geo, sash_pos = self._save_layout()
        func(idxs)
        self._refresh_all()
        # 恢复窗口几何与 sash 位置
        self._restore_layout(geo, sash_pos)
        self.set_status("Order adjusted.")

    def on_link_top(self):
        self._apply_link_sort(self.store.move_top)

    def on_link_bottom(self):
        self._apply_link_sort(self.store.move_bottom)

    def on_link_up(self):
        self._apply_link_sort(self.store.move_up)

    def on_link_down(self):
        self._apply_link_sort(self.store.move_down)

    # ---- Category sorting (any level: top-level / child) ----
    def _selected_cat_fullpath(self):
        sel = self.cat_tree.selection()
        if not sel:
            return None
        v = self.cat_tree.item(sel[0], "values")
        if not v or v[0] == "__ALL__":
            return None
        return v[0]

    def _sibling_categories(self, target_full):
        """
        返回 target_full 所属的同级分类列表（完整路径字符串），以及它自身在列表中的下标。
        如 target='实用工具>图片处理' → 返回属于父'实用工具'下所有直接子分类（完整路径）的有序列表，以及 idx。
        顶级时，parent_prefix=''，返回顶级分类完整路径(=分类名)列表。
        """
        if SEPARATOR in target_full:
            parent_prefix = target_full.rsplit(SEPARATOR, 1)[0]
            target_short = target_full[len(parent_prefix + SEPARATOR):]
        else:
            parent_prefix = ""
            target_short = target_full

        # 收集所有分类出现顺序中，同级兄弟（去重）
        siblings = OrderedDict()
        for it in self.store.items:
            cat = it["category"] or "Uncategorized"
            if parent_prefix:
                if not cat.startswith(parent_prefix + SEPARATOR):
                    continue
                rest = cat[len(parent_prefix + SEPARATOR):]
                if not rest:
                    continue
                short = rest.split(SEPARATOR)[0].strip()
            else:
                root = cat.split(SEPARATOR)[0].strip()
                short = root
            siblings[short] = True

        sibling_fulls = [
            (short if not parent_prefix else parent_prefix + SEPARATOR + short)
            for short in siblings.keys()
        ]
        try:
            idx = sibling_fulls.index(target_full)
        except ValueError:
            idx = -1
        return sibling_fulls, idx, parent_prefix, target_short

    def _reorder_sibling_categories(self, sibling_fulls_new, parent_prefix):
        """
        按 sibling_fulls_new 中同级分类顺序重排 store.items 中兄弟项。
        parent_prefix='' 时处理顶级；='a>b' 时处理 'a>b>xxx'。
        组内其它条目的相对顺序保持不变；非同级其它分类的原有整体顺序保持。
        """
        # 保存窗口几何，防止刷新触发收缩
        geo, sash_pos = self._save_layout()

        if parent_prefix:
            prefix_filter = parent_prefix + SEPARATOR
            sibling_shorts_new = [f[len(prefix_filter):] for f in sibling_fulls_new]
            # key = 同级短名，value=所属 bucket
            buckets = OrderedDict()
            for s in sibling_shorts_new:
                buckets[s] = []
            buckets["__others__"] = []
            unaffected = []  # 不属于该 parent_prefix 的条目（其他同级或无关），整体保持原有相对顺序
            for it in self.store.items:
                cat = it["category"] or "Uncategorized"
                if cat.startswith(prefix_filter):
                    rest = cat[len(prefix_filter):]
                    if rest:
                        short = rest.split(SEPARATOR)[0].strip()
                        if short in buckets:
                            buckets[short].append(it)
                            continue
                unaffected.append(it)
            # 构造新前缀下的条目顺序
            affected = []
            for s in sibling_shorts_new:
                affected.extend(buckets[s])
            affected.extend(buckets["__others__"])
            # 回写：把 unaffected 中属于原范围的那些条目（即之前匹配 parent_prefix+SEPARATOR 的）替换为 affected；
            # 其他 unaffected 中原本不匹配的保留原有位置。
            # 简便做法：重新遍历原 items，凡命中 parent_prefix+SEPARATOR 的从 affected 依次取；否则保留。
            it_affected = iter(affected)
            new_items = []
            for it in self.store.items:
                cat = it["category"] or "Uncategorized"
                if cat.startswith(prefix_filter):
                    rest = cat[len(prefix_filter):]
                    if rest:
                        short = rest.split(SEPARATOR)[0].strip()
                        if short in set(sibling_shorts_new):
                            try:
                                new_items.append(next(it_affected))
                                continue
                            except StopIteration:
                                pass
                new_items.append(it)
            # 若还有剩余 affected（理论上不会），追加
            for it in it_affected:
                new_items.append(it)
            self.store.items = new_items
        else:
            # 顶级：以 sibling_fulls_new (=顶级分类名) 作为顶级顺序
            top_names = list(sibling_fulls_new)
            buckets = OrderedDict((n, []) for n in top_names)
            others = []
            for it in self.store.items:
                cat = it["category"] or "Uncategorized"
                root = cat.split(SEPARATOR)[0].strip()
                if root in buckets:
                    buckets[root].append(it)
                else:
                    others.append(it)
            new_items = []
            for n in top_names:
                new_items.extend(buckets[n])
            new_items.extend(others)
            self.store.items = new_items

        self.store.modified = True
        self._refresh_all()

        # 恢复窗口几何与 sash 位置，避免刷新后界面收缩
        self._restore_layout(geo, sash_pos)

    def on_cat_up(self):
        sel = self._selected_cat_fullpath()
        if not sel:
            return
        siblings, idx, parent, short = self._sibling_categories(sel)
        if idx <= 0:
            messagebox.showinfo("Hint", f"'{short}' is already the first among its siblings.")
            return
        siblings[idx - 1], siblings[idx] = siblings[idx], siblings[idx - 1]
        self._reorder_sibling_categories(siblings, parent)
        self.set_status(f"Category '{short}' moved up.")

    def on_cat_down(self):
        sel = self._selected_cat_fullpath()
        if not sel:
            return
        siblings, idx, parent, short = self._sibling_categories(sel)
        if idx < 0 or idx >= len(siblings) - 1:
            messagebox.showinfo("Hint", f"'{short}' is already the last among its siblings.")
            return
        siblings[idx + 1], siblings[idx] = siblings[idx], siblings[idx + 1]
        self._reorder_sibling_categories(siblings, parent)
        self.set_status(f"Category '{short}' moved down.")

    def on_cat_top(self):
        sel = self._selected_cat_fullpath()
        if not sel:
            return
        siblings, idx, parent, short = self._sibling_categories(sel)
        siblings = [sel] + [s for s in siblings if s != sel]
        self._reorder_sibling_categories(siblings, parent)
        self.set_status(f"Category '{short}' moved to top.")

    def on_cat_bottom(self):
        sel = self._selected_cat_fullpath()
        if not sel:
            return
        siblings, idx, parent, short = self._sibling_categories(sel)
        siblings = [s for s in siblings if s != sel] + [sel]
        self._reorder_sibling_categories(siblings, parent)
        self.set_status(f"Category '{short}' moved to bottom.")

    # 更多：顶级分类按名称 / 数量 排序；或按选中分类的同级批量排序
    def on_cat_sort(self, mode="name", direction=1):
        # 若有选中且非顶级，则按同级范围排序；否则按顶级排序
        sel = self._selected_cat_fullpath()
        parent = ""
        if sel and SEPARATOR in sel:
            # 对选中分类的同级进行排序
            siblings, idx, parent, short = self._sibling_categories(sel)
        else:
            # 顶级
            siblings = list(self._cat_root_items())
        # 构建辅助信息：{full: {count, name}}
        _, agg = self.store.category_counts()
        if parent:
            # 同 parent_prefix 下的 agg key 就是这些 siblings 的 full
            info = {s: {"name": s.rsplit(SEPARATOR, 1)[-1], "count": agg.get(s, 0)} for s in siblings}
        else:
            info = {s: {"name": s, "count": agg.get(s, 0)} for s in siblings}

        def key_fn(full):
            d = info.get(full, {})
            if mode == "name":
                return d.get("name", full)
            if mode == "count":
                return (d.get("count", 0), d.get("name", full))
            return full

        reverse = direction < 0
        siblings_sorted = sorted(siblings, key=key_fn, reverse=reverse)
        self._reorder_sibling_categories(siblings_sorted, parent)
        scope = f"Top-level categories" if not parent else f"Sibling categories (under {parent})"
        mode_cn = {"name": "Name", "count": "Count"}.get(mode, mode)
        dir_cn = "Descending" if reverse else "Ascending"
        self.set_status(f"Re-sorted by {mode_cn} {dir_cn} ({scope}).")

    def on_links_in_cat_sort(self, field="name", direction=1):
        """把 items 按原分组稳定性保持，仅对同「网址分类目录」= 完全相同 的条目内部按 field 排序。"""
        import functools
        groups = OrderedDict()
        other = []
        for idx, it in enumerate(self.store.items):
            cat = it["category"] or "Uncategorized"
            if cat not in groups:
                groups[cat] = []
            groups[cat].append(it)

        # 若用户选中了某个分类，则只对该分类（含其子孙）生效
        sel = self._selected_cat_fullpath()

        def need_sort(cat):
            if not sel:
                return True
            return cat == sel or cat.startswith(sel + SEPARATOR)

        reverse = direction < 0
        for cat, arr in groups.items():
            if need_sort(cat):
                if field == "name":
                    arr.sort(key=lambda x: (x["name"] or "").lower(), reverse=reverse)
                elif field == "url":
                    arr.sort(key=lambda x: (x["url"] or "").lower(), reverse=reverse)
                elif field == "desc":
                    arr.sort(key=lambda x: (x["desc"] or "").lower(), reverse=reverse)

        new_items = []
        for cat, arr in groups.items():
            new_items.extend(arr)

        # 保存窗口几何与 sash 位置
        geo, sash_pos = self._save_layout()

        self.store.items = new_items
        self.store.modified = True
        self._refresh_all()

        # 恢复窗口几何与 sash 位置
        self._restore_layout(geo, sash_pos)

        field_cn = {"name": "Website Name", "url": "URL", "desc": "Description"}.get(field, field)
        dir_cn = "Descending" if reverse else "Ascending"
        scope = f"Under selected category '{sel}'" if sel else "Under all categories globally"
        self.set_status(f"{scope} link items sorted by '{field_cn}' ({dir_cn}).")

    def on_cat_original(self):
        """恢复为表格内按首次出现的原始顶级顺序（基本等同于导入时顺序）"""
        # 直接根据首次出现顺序重建顶级
        roots = self._cat_root_items()
        # _reorder_sibling_categories 用同级（顶级）重新分组即可
        self._reorder_sibling_categories(list(roots), "")
        self.set_status("Category order restored to the original table order at import.")

    # ---- Misc ----
    def _emoji_for_full(self, full):
        """按『最长前缀』在 category_icons 里找该分类路径对应的 emoji；没有返回空串。"""
        if not full:
            return ""
        parts = [p for p in full.split(SEPARATOR) if p]
        for end in range(len(parts), 0, -1):
            key = SEPARATOR.join(parts[:end])
            v = self.store.category_icons.get(key)
            if v:
                return v
        return ""

    def _cat_tree_on_right_click(self, evt):
        """左侧分类树右键：选中一级/二级分类后，提供「设置 / 清除 emoji」。"""
        iid = self.cat_tree.identify_row(evt.y)
        if not iid:
            return
        vals = self.cat_tree.item(iid, "values")
        if not vals or vals[0] == "__ALL__":
            return
        self.cat_tree.selection_set(iid)
        full = vals[0]
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Set emoji…", command=lambda: self.on_set_category_emoji(full))
        menu.add_command(label="Clear emoji", command=lambda: self.on_clear_category_emoji(full))
        try:
            menu.tk_popup(evt.x_root, evt.y_root)
        finally:
            menu.grab_release()

    def _current_cat_full(self):
        sel = self.cat_tree.selection()
        if not sel:
            return None
        v = self.cat_tree.item(sel[0], "values")
        return v[0] if v and v[0] != "__ALL__" else None

    def on_set_category_emoji(self, full=None):
        """给一级/二级分类设置 emoji：输入单个 emoji（可 Win+; 或复制粘贴）。"""
        full = full or self._current_cat_full()
        if not full:
            messagebox.showinfo("Hint", "Please select a category on the left first (not 'All Sites').")
            return
        cur = self.store.category_icons.get(full, "")
        res = simpledialog.askstring(
            "Set emoji",
            f"Category: {full}\nCurrent emoji: {cur or '(none)'}\n\n"
            "Enter 1 emoji (e.g. 🔍, 🎨, 🛠).\nTo clear, enter a single space or click 'Clear emoji'.",
            initialvalue=cur,
        )
        if res is None:
            return
        res = res.strip()
        if not res:
            self.store.category_icons.pop(full, None)
        else:
            # 只取首个完整 emoji（用户可能粘贴了带描述文字的内容）
            self.store.category_icons[full] = res[:4] if res else ""
        self.store.modified = True
        self._refresh_categories()
        self.set_status(f"Category '{full}' emoji updated. Takes effect after saving (Excel) or exporting JSON.")

    def on_clear_category_emoji(self, full=None):
        full = full or self._current_cat_full()
        if not full:
            return
        self.store.category_icons.pop(full, None)
        self.store.modified = True
        self._refresh_categories()
        self.set_status(f"Cleared emoji for category '{full}'.")

    def on_cat_select(self, _evt):
        self._refresh_table()

    def on_reset_filter(self):
        self.search_var.set("")
        root = self.cat_tree.get_children()
        if root:
            self.cat_tree.selection_set(root[0])
        self._refresh_table()
        self.set_status("Filter reset.")

    # ---- Icon download / update ----
    def _table_selected_store_indices(self):
        rowid_to_idx = {r: i for i, r in self._data_rowids}
        return [rowid_to_idx[r] for r in self.table.selection() if r in rowid_to_idx]

    def on_download_icons(self):
        if requests is None:
            messagebox.showerror("Missing Dependency", "The requests library is not installed; icons cannot be downloaded.\nPlease run: pip install requests")
            return
        sel = self._table_selected_store_indices()
        if sel:
            dlg = ScopeDialog(self, all_count=len(self.store.items), sel_count=len(sel))
            choice = dlg.result
            if choice is None:
                return
            idxs = list(range(len(self.store.items))) if choice == "all" else sel
        else:
            idxs = list(range(len(self.store.items)))
        # 「下载图标」跳过已存在的本地图标，「更新图标」强制重新下载覆盖
        self._download_icons(idxs, title="Download Icons", skip_existing=True)

    def _download_icons(self, idxs, title="Download Icons", skip_existing=False):
        # 去重（按 logo_key），保留顺序
        tasks = []
        seen = set()
        for i in idxs:
            it = self.store.items[i]
            key = logo_key(it.get("url", ""))
            if key and key not in seen:
                seen.add(key)
                tasks.append((i, key))
        if not tasks:
            messagebox.showinfo("Hint", "No valid links to download.")
            return
        dlg = DownloadDialog(self, total=len(tasks), title=title)
        threading.Thread(
            target=self._download_worker, args=(tasks, dlg, skip_existing), daemon=True
        ).start()
        dlg.wait_window()
        skipped = getattr(dlg, "skip_count", 0)
        self.set_status(
            f"Icon download complete: Success {dlg.ok_count} / Failed {dlg.fail_count}"
            + (f" / Skipped {skipped} (already exists)" if skipped else "")
            + f"(Total {len(tasks)}). Saved to {LOGO_DIR}"
        )

    def _download_worker(self, tasks, dlg, skip_existing=False):
        service = FaviconService(log_callback=lambda m: dlg.append_log(m))
        ok = fail = skipped = 0
        failed_keys = []
        for n, (i, key) in enumerate(tasks, start=1):
            if dlg.cancelled:
                dlg.append_log("User canceled.")
                break
            if skip_existing and _logo_key_exists(key):
                skipped += 1
                service.log(f"Skipped (local icon exists): {key}")
                dlg.set_progress(n, len(tasks), key, None)
                self._apply_icon_to_cache(key)  # 已有本地图标，同步缓存路径
                continue
            success, _ = service.download_one(key)
            if success:
                ok += 1
                self._apply_icon_to_cache(key)  # 下载成功才改写缓存，失败不改写
            else:
                fail += 1
                failed_keys.append(key)
            dlg.set_progress(n, len(tasks), key, success)
            time.sleep(0.3)
        # 写出失败清单
        if failed_keys:
            try:
                with open(os.path.join(LOGO_DIR, "failed_downloads.txt"), "w", encoding="utf-8") as f:
                    f.write("Domain/key\n")
                    for k in failed_keys:
                        f.write(k + "\n")
            except Exception:
                pass
        dlg.finish(ok, fail, skipped)

    def _apply_icon_to_cache(self, key):
        """下载成功（或本地已有）后，把缓存（store.items）中同域名条目的图标路径改写为本地图标。"""
        for it in self.store.items:
            if logo_key(it.get("url", "")) == key:
                rel = _logo_local_rel(it.get("url", ""))
                if rel:
                    it["icon"] = rel

    def _logo_dir_empty(self):
        """图标文件夹（assets/logo）是否没有任何图标文件。"""
        try:
            return not any(
                f.lower().endswith(ICON_EXTS) for f in os.listdir(LOGO_DIR)
            )
        except OSError:
            return True

    def on_update_icon(self):
        if self._logo_dir_empty():
            messagebox.showinfo(
                "Hint",
                "The icon folder (assets\\logo) is empty; there are no icons to update yet.\n"
                "Please use 'Download Icons' to download them; no update needed.",
            )
            return
        idxs = self._table_selected_store_indices()
        if not idxs:
            messagebox.showinfo("Hint", "Please select the row(s) to update icons for in the table on the right (multi-select supported).")
            return
        if len(idxs) == 1:
            self._icon_update_single(idxs[0])
        else:
            if messagebox.askyesno(
                "Update Icons",
                f"Re-download icons for the {len(idxs)} selected sites from the web (force-overwrite all old icons, including ones already downloaded)?\n\n"
                "Tip: when a single row is selected, this button lets you choose a local image to set the icon manually.",
            ):
                self._download_icons(idxs, title="Update Icons", skip_existing=False)

    def _icon_update_single(self, idx):
        it = self.store.items[idx]
        name = it.get("name") or it.get("url") or "this site"
        dlg = IconUpdateDialog(self, name)
        res = dlg.result
        if res == "web":
            self._download_icons([idx], title="Update Icons")
        elif res == "local":
            self._pick_local_icon(idx)

    def _pick_local_icon(self, idx):
        path = filedialog.askopenfilename(
            title="Select Icon Image",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.webp *.svg *.ico"), ("All Files", "*.*")],
        )
        if not path:
            return
        it = self.store.items[idx]
        key = logo_key(it.get("url", ""))
        if not key:
            messagebox.showerror("Error", "This entry has no valid URL, so an icon filename cannot be generated.")
            return
        try:
            with open(path, "rb") as f:
                data = f.read()
        except Exception as e:
            messagebox.showerror("Read Failed", str(e))
            return
        svc = FaviconService()
        try:
            svc.remove_existing(key)  # 清除旧图标，避免多扩展名并存
            name = svc.save_as_png(key, data)  # 统一转为 PNG
        except Exception as e:
            messagebox.showerror("Save Failed", str(e))
            return
        rel = _logo_local_rel(it.get("url", ""))
        if rel:
            it["icon"] = rel  # 同步改写缓存中的图标路径
        self.set_status(f"Icon updated with local image: {os.path.basename(name)} (converted to PNG)")

    def set_status(self, msg):
        self.status_var.set(msg)


# ============== Add / Edit Dialog ==============
class ItemDialog(tk.Toplevel):
    def __init__(self, master, title, item=None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.transient(master)
        self.grab_set()

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        self.var_name = tk.StringVar(value=item["name"] if item else "")
        self.var_cat = tk.StringVar(value=item["category"] if item else "")
        self.var_url = tk.StringVar(value=item["url"] if item else "")
        self.var_desc = tk.StringVar(value=item["desc"] if item else "")

        def row(r, label, widget):
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky=tk.E, padx=4, pady=4)
            widget.grid(row=r, column=1, sticky=tk.W + tk.E, padx=4, pady=4)

        frm.columnconfigure(1, weight=1)
        row(0, "Website Name:", ttk.Entry(frm, textvariable=self.var_name, width=60))
        row(1, "Category Path:", ttk.Entry(frm, textvariable=self.var_cat, width=60))
        ttk.Label(frm, text="(Use '>' to separate multiple levels, e.g.: Tools>Image)",
                  foreground="#888").grid(row=1, column=2, padx=6, sticky=tk.W)
        row(2, "URL:", ttk.Entry(frm, textvariable=self.var_url, width=60))

        ttk.Label(frm, text="Description:").grid(row=3, column=0, sticky=tk.NE, padx=4, pady=4)
        desc_txt = tk.Text(frm, width=60, height=6)
        desc_txt.grid(row=3, column=1, sticky=tk.W + tk.E, padx=4, pady=4)
        if item and item["desc"]:
            desc_txt.insert("1.0", item["desc"])

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=3, pady=(12, 0), sticky=tk.E)

        def ok():
            name = self.var_name.get().strip()
            cat = self.var_cat.get().strip()
            url = self.var_url.get().strip()
            desc = desc_txt.get("1.0", tk.END).strip()
            if not name and not url:
                messagebox.showwarning("Missing Content", "Please fill in at least the website name or the URL.", parent=self)
                return
            self.result = {"name": name, "category": cat, "url": url, "desc": desc}
            self.destroy()

        ttk.Button(btns, text="OK", command=ok).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT)

        self.bind("<Return>", lambda e: ok())
        self.bind("<Escape>", lambda e: self.destroy())


# ============== Icon Download Dialogs ==============
class ScopeDialog(tk.Toplevel):
    def __init__(self, master, all_count, sel_count):
        super().__init__(master)
        self.title("Download Icon Scope")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.result = None
        frm = ttk.Frame(self, padding=14)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="Select the scope of icons to download:").pack(anchor=tk.W, pady=(0, 8))
        var = tk.StringVar(value="sel" if sel_count else "all")
        ttk.Radiobutton(frm, text=f"Selected rows only ({sel_count})", variable=var, value="sel").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(frm, text=f"All URLs ({all_count})", variable=var, value="all").pack(anchor=tk.W, pady=2)
        btns = ttk.Frame(frm)
        btns.pack(anchor=tk.E, pady=(12, 0))
        ttk.Button(btns, text="Start", command=lambda: self._set(var.get())).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT)

    def _set(self, v):
        self.result = v
        self.destroy()


class IconUpdateDialog(tk.Toplevel):
    def __init__(self, master, name):
        super().__init__(master)
        self.title("Update Icons")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.result = None
        frm = ttk.Frame(self, padding=14)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text=f"How to update the icon for '{name}'?").pack(anchor=tk.W, pady=(0, 8))
        ttk.Button(frm, text="Re-download Web Icon", width=26, command=lambda: self._set("web")).pack(fill=tk.X, pady=3)
        ttk.Button(frm, text="Choose Local Image File", width=26, command=lambda: self._set("local")).pack(fill=tk.X, pady=3)
        ttk.Button(frm, text="Cancel", width=26, command=self.destroy).pack(fill=tk.X, pady=(8, 0))

    def _set(self, v):
        self.result = v
        self.destroy()


class DownloadDialog(tk.Toplevel):
    def __init__(self, master, total, title="Download Icons"):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.total = total
        self.cancelled = False
        self.ok_count = 0
        self.fail_count = 0
        self.skip_count = 0
        self.bind("<Destroy>", lambda e: setattr(self, "cancelled", True))
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)
        self.label = ttk.Label(frm, text=f"Preparing download 0 / {total}")
        self.label.pack(anchor=tk.W, pady=(0, 4))
        self.pb = ttk.Progressbar(frm, length=460, maximum=max(total, 1), mode="determinate")
        self.pb.pack(fill=tk.X, pady=(0, 6))
        self.log = tk.Text(frm, width=64, height=14, state="disabled")
        self.log.pack(fill=tk.BOTH, expand=True)
        btns = ttk.Frame(frm)
        btns.pack(anchor=tk.E, pady=(8, 0))
        self.btn_cancel = ttk.Button(btns, text="Cancel", command=self.on_cancel)
        self.btn_cancel.pack(side=tk.RIGHT, padx=4)
        self.btn_close = ttk.Button(btns, text="Close", command=self.destroy, state="disabled")
        self.btn_close.pack(side=tk.RIGHT)

    def append_log(self, msg):
        if self.winfo_exists():
            self.after(0, self._append, msg)

    def _append(self, msg):
        if not self.winfo_exists():
            return
        self.log.configure(state="normal")
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def set_progress(self, done, total, key, success):
        if self.winfo_exists():
            self.after(0, self._setp, done, total, key, success)

    def _setp(self, done, total, key, success):
        if not self.winfo_exists():
            return
        self.pb["value"] = done
        if success is None:
            tag = "-> Skipped"
        else:
            tag = "✓" if success else "✗"
        self.label.config(text=f"Download {done} / {total}  [{tag} {key}]")

    def finish(self, ok, fail, skipped=0):
        if self.winfo_exists():
            self.after(0, self._finish, ok, fail, skipped)

    def _finish(self, ok, fail, skipped=0):
        if not self.winfo_exists():
            return
        self.ok_count, self.fail_count, self.skip_count = ok, fail, skipped
        skipped_txt = f" / Skipped {skipped}" if skipped else ""
        self.label.config(text=f"Done: Success {ok} / Failed {fail}{skipped_txt} (Total {self.total})")
        self.append_log(f"Download complete: {ok} succeeded, {fail} failed."
                         + (f" Skipped {skipped} (local icon already exists)." if skipped else ""))
        if fail:
            self.append_log("Failed items recorded to assets/logo/failed_downloads.txt")
        self.btn_cancel.config(state="disabled")
        self.btn_close.config(state="normal")

    def on_cancel(self):
        self.cancelled = True
        self.append_log("Cancel requested; waiting for the current task to finish…")


# ============== Startup ==============
def main():
    app = App()

    # 如果命令行给了文件路径，自动导入（按扩展名区分 Excel / JSON）
    if len(sys.argv) >= 2 and os.path.isfile(sys.argv[1]):
        try:
            path = sys.argv[1]
            if path.lower().endswith(".json"):
                cnt = app.store.load_json(path)
            else:
                cnt = app.store.load_excel(path)
            app.set_status(f"Auto-imported {path} ({cnt} items)")
            app._refresh_all()
            root = app.cat_tree.get_children()
            if root:
                app.cat_tree.selection_set(root[0])
                app._refresh_table()
        except Exception as e:
            app.set_status(f"Auto-import failed: {e}")

    app.mainloop()


if __name__ == "__main__":
    main()

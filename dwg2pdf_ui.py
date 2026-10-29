#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dwg2pdf_ui.py —— 图形界面启动器（本地网页界面）

双击「start.bat」后会：
  1. 在本机 127.0.0.1 起一个只给本机用的服务；
  2. 自动打开浏览器显示操作界面；
  3. 界面上点「开始转换」，后台调用 AutoCAD 逐张图纸出 PDF。

不装任何第三方界面库（tkinter/PyQt），只需要 pywin32。
"""

from __future__ import annotations

import ctypes
import json
import os
import re
import socket
import stat
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import unquote, urlparse

import dwg2pdf
import i18n

HERE = os.path.dirname(os.path.abspath(__file__))
# 打包成 exe 后，附带的数据文件被解压到 sys._MEIPASS 里
BUNDLE = getattr(sys, "_MEIPASS", HERE)
UI_FILE = os.path.join(BUNDLE, "ui.html")
DATA_DIR = HERE if not getattr(sys, "frozen", False) else os.path.dirname(
    sys.executable if hasattr(sys, "executable") else HERE)

# 版本号：启动时写进 run.log，/api/ping 也会返回，
# 这样「换了新 exe 但后台还是老进程」的情况能一眼看出来
APP_VERSION = "2.8.0"

# 界面语言：'zh' / 'en'；空串表示「跟随系统」（存盘里记作 auto）
LANG_MODE = ""


# ----------------------------------------------------------------- 全局状态

class Hub:
    """保存一次转换的进度与日志，供页面轮询"""

    def __init__(self):
        self.lock = threading.Lock()
        self.events: list[dict] = []
        self.seq = 0
        self.running = False
        self.cancel_event = threading.Event()
        self.converter: dwg2pdf.Converter | None = None
        self.out_dir = ""
        self.counters = {"ok": 0, "fail": 0, "skip": 0, "total": 0}
        self.finished = False
        self.error = ""

    def push(self, ev: dict):
        with self.lock:
            self.seq += 1
            ev["n"] = self.seq
            self.events.append(ev)
            if len(self.events) > 4000:          # 日志太长就丢掉最早的，避免吃内存
                del self.events[:2000]

    def events_after(self, after: int) -> list[dict]:
        with self.lock:
            return [e for e in self.events if e["n"] > after]

    def snapshot(self) -> dict:
        with self.lock:
            return {
                "seq": self.seq,
                "running": self.running,
                "finished": self.finished,
                "counters": dict(self.counters),
                "out_dir": self.out_dir,
                "error": self.error,
            }

    # -------------------------------------------------- 转换控制
    def start(self, cfg: dict) -> dict:
        with self.lock:
            if self.running:
                return {"ok": False, "error": i18n.t("正在转换中，请先等待或点「停止」",
                                                     "A conversion is already running — wait for it "
                                                     "or press Stop")}
            self.events.clear()
            self.counters = {"ok": 0, "fail": 0, "skip": 0, "total": 0}
            self.finished = False
            self.error = ""

        src = (cfg.get("src") or "").strip()
        single = src.lower().endswith(".dwg") and os.path.isfile(src)
        if not single and (not src or not os.path.isdir(src)):
            return {"ok": False, "error": i18n.t("请先选择 DWG 所在文件夹（路径不存在）",
                                                 "Choose the DWG folder first (the path does not exist)")}
        base_dir = os.path.dirname(src) if single else src
        out = (cfg.get("out") or "").strip() or os.path.join(base_dir, "PDF")
        out = os.path.normpath(out)
        try:
            os.makedirs(out, exist_ok=True)
        except OSError as exc:
            return {"ok": False, "error": i18n.t(f"输出文件夹无法创建：{exc}",
                                                 f"Cannot create the output folder: {exc}")}

        if single:
            files = [src]
        else:
            files = dwg2pdf.find_dwg_files(src, bool(cfg.get("recursive", True)),
                                           exclude_dirs=[out])
        if not files:
            return {"ok": False, "error": i18n.t(
                "这个文件夹里没有找到 .dwg 文件"
                + ("（已包含子文件夹）" if cfg.get("recursive", True) else "")
                + "，请确认路径。",
                "No .dwg file found in this folder"
                + (" (sub-folders included)" if cfg.get("recursive", True) else "")
                + " — please check the path.")}

        # 手动输入的路径也算「上次干活的文件夹」，下次选文件夹直接从这儿开
        remember("src", base_dir)
        remember("out", out)

        self.cancel_event = threading.Event()
        with self.lock:
            self.running = True
            self.out_dir = out

        base = self.seq                         # 这三行日志也要发给页面，别被跳掉
        self.push({"type": "log",
                   "message": i18n.t(f"待处理 DWG：{len(files)} 个",
                                     f"DWG files to process: {len(files)}")})
        self.push({"type": "log",
                   "message": i18n.t(f"输出目录：{out}", f"Output folder: {out}")})
        self.push({"type": "log",
                   "message": i18n.t("正在启动 / 连接 AutoCAD，首次运行约需 30 秒…",
                                     "Starting / connecting to AutoCAD — the first run takes "
                                     "about 30 s…")})

        conv = dwg2pdf.Converter(
            src=src, out=out,
            prefix=cfg.get("prefix") or "",
            include_model=bool(cfg.get("include_model", False)),
            recursive=bool(cfg.get("recursive", True)),
            keep_subdir=bool(cfg.get("keep_subdir", True)),
            skip_existing=bool(cfg.get("skip_existing", False)),
            on_event=self._on_event,
            cancel=self.cancel_event.is_set,
            keep_acad=bool(cfg.get("keep_acad", False)),
            one_pdf=bool(cfg.get("one_pdf", True)),
            order_pages=bool(cfg.get("order_pages", False)),
            paper_size=str(cfg.get("paper") or "A1"),
        )
        self.converter = conv
        threading.Thread(target=self._worker, args=(conv,), daemon=True).start()
        return {"ok": True, "seq": base, "out": out, "count": len(files)}

    def _on_event(self, ev: dict):
        if ev.get("type") == "result":
            r = ev.get("result", {})
            with self.lock:
                self.counters["total"] = self.counters.get("total", 0) + 1
                st = r.get("status")
                if st in ("ok", "fail", "skip"):
                    self.counters[st] = self.counters.get(st, 0) + 1
        self.push({k: v for k, v in ev.items() if k != "type"} | {"type": ev.get("type")})

    def _worker(self, conv: dwg2pdf.Converter):
        try:
            summary = conv.run()
            self.push({"type": "done", "summary": summary})
        except Exception as exc:                     # noqa: BLE001
            msg = str(exc)
            with self.lock:
                self.error = msg
            self.push({"type": "fatal", "message": msg})
        finally:
            with self.lock:
                self.running = False
                self.finished = True

    def stop(self) -> dict:
        if not self.running:
            return {"ok": False, "error": i18n.t("当前没有在转换",
                                                 "Nothing is being converted right now")}
        self.cancel_event.set()
        self.push({"type": "log",
                   "message": i18n.t("已请求停止，等当前这张图纸出完后结束…",
                                     "Stop requested — finishing the current sheet first…")})
        return {"ok": True}


HUB = Hub()


# ----------------------------------------------------------------- 文件夹选择

def pick_folder(title: str = "") -> str:
    """弹出 Windows 原生文件夹选择框（阻塞直到用户关掉）"""
    title = title or i18n.t("选择文件夹", "Choose a folder")
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            folder = shell.BrowseForFolder(0, title, 0x0040 | 0x0001, "")
            if folder is None:
                return ""
            try:
                return folder.Self.Path
            except Exception:
                return folder.Items().Item().Path
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        pass

    # shell COM 不可用时退回 PowerShell
    ps = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$d = New-Object System.Windows.Forms.FolderBrowserDialog;"
        f"$d.Description = '{title}';"
        "if ($d.ShowDialog() -eq 'OK') { $d.SelectedPath }"
    )
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-STA", "-Command", ps],
                             capture_output=True, text=True, timeout=1800)
        return (out.stdout or "").strip()
    except Exception:
        return ""


# ----------------------------------------------------------------- 记住上次位置

_STATE_NAME = "dwg2pdf_state.json"


def _state_paths() -> list[str]:
    """状态文件候选路径：优先 exe 旁边（绿色版），不可写时退到用户目录"""
    out = [os.path.join(DATA_DIR, _STATE_NAME)]
    local = (os.environ.get("LOCALAPPDATA")
             or os.path.join(os.path.expanduser("~"), "AppData", "Local"))
    out.append(os.path.join(local, "DWG2PDF", _STATE_NAME))
    return out


def _load_state() -> dict:
    for p in _state_paths():
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            continue
    return {}


def _save_state(patch: dict) -> None:
    """只合并本次变化的键；写不进去就静默放弃（不能因为记不住位置影响出图）"""
    if not patch:
        return
    data = _load_state()
    changed = False
    for k, v in patch.items():
        if v and data.get(k) != v:
            data[k] = v
            changed = True
    if not changed:
        return
    for p in _state_paths():
        try:
            d = os.path.dirname(p)
            if d and not os.path.isdir(d):
                os.makedirs(d, exist_ok=True)
            tmp = p + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=1)
            os.replace(tmp, p)
            return
        except Exception:
            continue


def remember(kind: str, path: str) -> None:
    """记住这次选定的文件夹（kind 为 src / out）"""
    p = (path or "").strip()
    if not p or not os.path.isdir(p):
        return
    kind = (kind or "").strip().lower()
    patch = {"browse": p}
    if kind in ("src", "out"):
        patch[kind] = p
    _save_state(patch)


def start_dir(kind: str = "") -> str:
    """这次该从哪个文件夹开始：该用途上次选的 → 上次浏览到的 → 用户目录"""
    data = _load_state()
    kind = (kind or "").strip().lower()
    keys = ["browse", "src", "out"]
    if kind in ("src", "out"):
        keys.insert(0, kind)
    for key in keys:
        p = data.get(key) or ""
        if p and os.path.isdir(p):
            return p
    home = os.path.expanduser("~")
    return home if os.path.isdir(home) else ""


# ----------------------------------------------------------------- 目录浏览

_SKIP_DIRS = {"system volume information", "$recycle.bin", "recycler",
              "recovery", "config.msi", "perflogs"}


def _drives() -> list[dict]:
    """本机可用盘符（含网络盘）"""
    out: list[dict] = []
    mask = 0
    try:
        mask = ctypes.windll.kernel32.GetLogicalDrives()
    except Exception:
        mask = 0
    for i, ch in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
        if not (mask >> i) & 1:
            continue
        root = f"{ch}:\\"
        try:
            t = ctypes.windll.kernel32.GetDriveTypeW(ctypes.c_wchar_p(root))
        except Exception:
            t = 3
        if t not in (2, 3, 4, 6):        # 只要可移动 / 固定 / 网络 / 内存盘
            continue
        out.append({"name": f"{ch}:", "path": root, "type": t})
    if not out:                          # 兜底
        for ch in "CDEFGHIJ":
            p = f"{ch}:\\"
            if os.path.isdir(p):
                out.append({"name": f"{ch}:", "path": p, "type": 3})
    return out


def _places() -> list[dict]:
    """常用位置：桌面 / 文档 / 下载 / 图片 / 用户目录"""
    home = os.path.expanduser("~")
    out: list[dict] = []
    for zh, en, sub in (("桌面", "Desktop", "Desktop"),
                        ("文档", "Documents", "Documents"),
                        ("下载", "Downloads", "Downloads"),
                        ("图片", "Pictures", "Pictures")):
        p = os.path.join(home, sub)
        if os.path.isdir(p):
            out.append({"name": i18n.t(zh, en), "path": p})
    if os.path.isdir(home):
        out.append({"name": i18n.t("用户目录", "Home"), "path": home})
    return out


def _clean_path(raw: str) -> str:
    path = (raw or "").strip().strip('"').strip("'")
    path = os.path.expandvars(os.path.expanduser(path))
    if re.fullmatch(r"[A-Za-z]:", path):
        path += "\\"
    if not path:
        return os.path.expanduser("~")
    return os.path.abspath(path)


def list_dir(raw: str, kind: str = "") -> dict:
    """列出某个文件夹下的子文件夹（供自绘选择器用）

    raw 为空时从「上次打开过的文件夹」开始，让程序每次启动后
    点「选择…」都直接落在上次干活的地方。
    """
    if not (raw or "").strip():
        raw = start_dir(kind)
    path = _clean_path(raw)
    if not os.path.isdir(path):
        return {"ok": False,
                "error": i18n.t("路径不存在或无法访问：" + path,
                                "Path does not exist or is not accessible: " + path),
                "path": path, "places": _places(), "drives": _drives()}

    dirs: list[dict] = []
    dwg = 0
    try:
        with os.scandir(path) as it:
            for e in it:
                try:
                    if e.is_dir(follow_symlinks=False):
                        if e.name.startswith(".") or e.name.lower() in _SKIP_DIRS:
                            continue
                        attrs = e.stat(follow_symlinks=False).st_file_attributes
                        if attrs & (stat.FILE_ATTRIBUTE_HIDDEN
                                    | stat.FILE_ATTRIBUTE_SYSTEM):
                            continue
                        dirs.append({"name": e.name, "path": e.path})
                    elif e.name.lower().endswith(".dwg"):
                        dwg += 1
                except OSError:
                    continue
    except PermissionError:
        return {"ok": False,
                "error": i18n.t("没有权限访问这个文件夹",
                                "Access to this folder is denied"),
                "path": path, "places": _places(), "drives": _drives()}
    except OSError as exc:
        return {"ok": False, "error": str(exc),
                "path": path, "places": _places(), "drives": _drives()}

    dirs.sort(key=lambda d: d["name"].lower())
    parent = os.path.dirname(path)
    if os.path.normcase(parent) == os.path.normcase(path):
        parent = ""                       # 已经在盘根，没有上一级
    _save_state({"browse": path})         # 逛到哪儿记住哪儿，下次从这里接着开
    return {"ok": True, "path": path, "parent": parent, "dirs": dirs,
            "dwg_count": dwg, "places": _places(), "drives": _drives()}


# ----------------------------------------------------------------- 语言

def init_lang() -> str:
    """按「上次的选择 → 系统语言」初始化界面语言"""
    global LANG_MODE
    LANG_MODE = i18n.normalize(_load_state().get("lang") or "")
    return i18n.set_lang(LANG_MODE or i18n.system_lang())


def set_lang_mode(mode: str) -> str:
    """切换语言：'zh' / 'en' / 'auto'（跟随系统）；写盘后立即生效"""
    global LANG_MODE
    LANG_MODE = i18n.normalize(mode)
    lang = i18n.set_lang(LANG_MODE or i18n.system_lang())
    _save_state({"lang": LANG_MODE or "auto"})
    return lang


def use_lang(mode: str) -> str:
    """只对本次运行生效、不写盘（启动参数 --lang 用，避免污染用户的选择）"""
    global LANG_MODE
    LANG_MODE = i18n.normalize(mode)
    return i18n.set_lang(LANG_MODE or i18n.system_lang())


# ----------------------------------------------------------------- HTTP 服务

def _logfile(msg: str):
    """把关键动作记到 exe 旁边的「run.log」，出问题时有据可查"""
    try:
        path = os.path.join(DATA_DIR, "run.log")
        if os.path.exists(path) and os.path.getsize(path) > 512 * 1024:
            os.replace(path, path + ".old")
        with open(path, "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S  ") + msg.replace("\n", "\n    ") + "\n")
    except Exception:
        pass


class Handler(BaseHTTPRequestHandler):
    server_version = "dwg2pdf/1.6"
    port = 0                                # main() 里填，页面用它做自检

    # 安静一点，不要每个请求都往控制台打
    def log_message(self, fmt, *args):     # noqa: A003
        pass

    # -------------------------------------------------- 工具
    def _send(self, code: int, body: bytes, ctype: str):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(body)
        except Exception:
            pass

    def _json(self, obj, code: int = 200):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _read_json(self) -> dict:
        try:
            length = int(self.headers.get("Content-Length") or 0)
            raw = self.rfile.read(length) if length else b"{}"
            return json.loads(raw.decode("utf-8") or "{}")
        except Exception:
            return {}

    # -------------------------------------------------- 路由
    def do_GET(self):                        # noqa: N802
        try:
            self._route_get()
        except Exception as exc:             # noqa: BLE001
            self._fail(exc)

    def do_POST(self):                       # noqa: N802
        try:
            self._route_post()
        except Exception as exc:             # noqa: BLE001
            self._fail(exc)

    def _fail(self, exc):
        """任何后台异常都要回给页面，不能悄悄断开连接（否则界面看起来就是卡死）"""
        try:
            detail = traceback.format_exc()
            traceback.print_exc()
            _logfile(i18n.t("后台异常：", "Backend exception: ") + detail)
        except Exception:
            pass
        try:
            self._json({"ok": False,
                        "error": i18n.t(f"后台处理出错：{exc}", f"Backend error: {exc}")}, 500)
        except Exception:
            pass

    def _route_get(self):
        path = urlparse(self.path).path
        query = {}
        if "?" in self.path:
            for part in self.path.split("?", 1)[1].split("&"):
                if "=" in part:
                    k, v = part.split("=", 1)
                    query[k] = v

        if path in ("/", "/index.html"):
            _logfile(i18n.t(f"页面载入（端口 {Handler.port}）",
                            f"page loaded (port {Handler.port})"))
            try:
                with open(UI_FILE, "rb") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            except OSError:
                self._send(500, b"ui.html not found", "text/plain; charset=utf-8")
            return

        if path == "/api/poll":
            after = 0
            try:
                after = int(query.get("after", "0"))
            except ValueError:
                pass
            data = HUB.snapshot()
            data["events"] = HUB.events_after(after)
            self._json(data)
            return

        if path == "/api/ping":
            self._json({"ok": True, "app": "dwg2pdf", "pid": os.getpid(),
                        "version": APP_VERSION,
                        "port": Handler.port, "running": HUB.running,
                        "lang": i18n.get_lang(),
                        "lang_mode": LANG_MODE or "auto",
                        "system_lang": i18n.system_lang()})
            return

        if path == "/api/ls":
            target = unquote(query.get("path", ""))
            kind = unquote(query.get("kind", ""))
            self._json(list_dir(target, kind))
            return

        if path == "/api/lang":
            self._json({"ok": True, "lang": i18n.get_lang(),
                        "lang_mode": LANG_MODE or "auto"})
            return

        if path == "/api/lastdir":
            self._json({"ok": True, "path": start_dir(unquote(query.get("kind", "")))})
            return

        if path == "/api/browse":
            kind = "src" if query.get("kind") == "src" else "out"
            title = (i18n.t("选择存放 DWG 的文件夹", "Choose the folder holding the DWGs")
                     if kind == "src"
                     else i18n.t("选择 PDF 输出文件夹", "Choose the PDF output folder"))
            picked = pick_folder(title)
            if picked:
                remember(kind, picked)
            self._json({"path": picked})
            return

        if path == "/api/openout":
            out = HUB.out_dir
            if out and os.path.isdir(out):
                os.startfile(out)            # noqa: S606
                self._json({"ok": True})
            else:
                self._json({"ok": False,
                            "error": i18n.t("输出文件夹还不存在",
                                            "The output folder does not exist yet")})
            return

        if path == "/api/quit":
            self._json({"ok": True})
            threading.Timer(0.4, _shutdown).start()
            return

        self._send(404, b"not found", "text/plain; charset=utf-8")

    def _route_post(self):
        path = urlparse(self.path).path
        if path == "/api/start":
            cfg = self._read_json()
            _logfile(i18n.t("收到开始请求：src=", "start requested: src=") + str(cfg.get("src"))
                     + " | out=" + str(cfg.get("out")))
            res = HUB.start(cfg)
            _logfile(i18n.t("开始请求返回：", "start returned: ") + json.dumps(res, ensure_ascii=False))
            self._json(res)
        elif path == "/api/stop":
            self._json(HUB.stop())
        elif path == "/api/lang":
            body = self._read_json()
            lang = set_lang_mode(str(body.get("lang") or ""))
            _logfile(i18n.t(f"语言切换为 {lang}（{LANG_MODE or 'auto'}）",
                            f"language switched to {lang} ({LANG_MODE or 'auto'})"))
            self._json({"ok": True, "lang": lang, "lang_mode": LANG_MODE or "auto"})
        elif path == "/api/remember":
            body = self._read_json()
            remember(str(body.get("kind") or ""), str(body.get("path") or ""))
            self._json({"ok": True})
        else:
            self._send(404, b"not found", "text/plain; charset=utf-8")


def _free_port(preferred: int = 8791) -> int:
    for port in (preferred, preferred + 1, preferred + 2, 0):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return s.getsockname()[1]
            except OSError:
                continue
    return 0


def _shutdown():
    os._exit(0)


def _instance_info(port: int) -> dict:
    """问一下这个端口上跑的是不是本程序，是就返回它的信息（含版本号）"""
    import urllib.request
    try:
        with urllib.request.urlopen(
                f"http://127.0.0.1:{port}/api/ping", timeout=1.0) as resp:
            data = json.loads(resp.read(600).decode("utf-8", "replace"))
            if data.get("app") == "dwg2pdf":
                return data
    except Exception:
        pass
    return {}


def _alive_instance(port_pref: int = 8791) -> int:
    """如果已经有本程序在跑，返回它的端口（避免开出第二个实例让页面搞混）"""
    for port in range(port_pref, port_pref + 4):
        if _instance_info(port):
            return port
    return 0


def _close_instance(port: int, pid: int = 0) -> bool:
    """让旧实例退出，返回端口是否已释放"""
    import urllib.request
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{port}/api/quit",
                               timeout=2.0).read()
    except Exception:
        pass
    deadline = time.time() + 8
    while time.time() < deadline:
        if not _instance_info(port):
            return True
        time.sleep(0.4)
    if pid:                      # 兜底：请不动就直接结束那个进程
        try:
            os.kill(pid, 9)
        except Exception:
            pass
        time.sleep(1.0)
        return not _instance_info(port)
    return False


def main():
    if not os.path.isfile(UI_FILE):
        raise SystemExit(i18n.t(f"找不到界面文件：{UI_FILE}",
                                f"UI file not found: {UI_FILE}"))

    args = sys.argv[1:]
    no_open = "--no-open" in args

    # --lang zh|en|auto 只影响本次运行（也给自动化测试用）；
    # 否则用上次在界面上选的语言，没选过就跟系统语言走
    if "--lang" in args:
        try:
            use_lang(args[args.index("--lang") + 1])
        except (ValueError, IndexError):
            init_lang()
    else:
        init_lang()
    port_pref = 8791
    if "--port" in args:
        try:
            port_pref = int(args[args.index("--port") + 1])
        except (ValueError, IndexError):
            pass

    # 已经有实例在跑？直接把浏览器指过去，不再开第二个（否则两个页面容易搞混）
    # 但如果那个实例是旧版本（换了新 exe 却还连着老后台），先把它关掉再起新的，
    # 否则新 exe 只是打开一个老页面，改了半天代码根本没用上。
    if "--new" not in args:
        alive = _alive_instance(port_pref)
        if alive:
            info = _instance_info(alive)
            old_ver = str(info.get("version") or "")
            if old_ver == APP_VERSION:
                url = f"http://127.0.0.1:{alive}/"
                print(i18n.t(f"已有实例在运行：{url}",
                             f"An instance is already running: {url}"))
                _logfile(i18n.t(f"发现同版本实例 v{APP_VERSION}，直接复用：{url}",
                                f"same version v{APP_VERSION} already running, reusing {url}"))
                if not no_open:
                    webbrowser.open(url)
                return
            unknown = i18n.t("未知", "unknown")
            print(i18n.t(f"发现旧版本实例（v{old_ver or unknown}），自动关闭后启动新版本…",
                         f"An older instance is running (v{old_ver or unknown}); "
                         f"closing it and starting the new version…"))
            _logfile(i18n.t(f"发现旧版本实例（v{old_ver or unknown}，PID {info.get('pid')}），"
                            f"已关闭并启动 v{APP_VERSION}",
                            f"older instance v{old_ver or unknown} (PID {info.get('pid')}) closed, "
                            f"starting v{APP_VERSION}"))
            _close_instance(alive, int(info.get("pid") or 0))
            port_pref = alive

    port = _free_port(port_pref)
    Handler.port = port
    _logfile(i18n.t(f"程序启动 v{APP_VERSION}（PID {os.getpid()}，端口 {port}，语言 {i18n.get_lang()}）",
                    f"app started v{APP_VERSION} (PID {os.getpid()}, port {port}, "
                    f"lang {i18n.get_lang()})"))
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    print(i18n.t(f"服务已启动：{url}", f"Server started: {url}"))
    print(i18n.t("关闭这个窗口即可结束程序。",
                 "Close this window to stop the program."))
    if not no_open:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()


if __name__ == "__main__":
    # 打包成静默 exe（--noconsole）时没有标准输出，print 会报错，这里兜住
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    try:
        main()
    except Exception:
        # 没有控制台时把错误写文件并弹窗，避免「点了没反应」
        detail = traceback.format_exc()
        try:
            with open(os.path.join(DATA_DIR, "start_error.log"), "w", encoding="utf-8") as f:
                f.write(detail)
        except Exception:
            pass
        try:
            ctypes.windll.user32.MessageBoxW(
                0, i18n.t(f"启动失败：\n\n{detail[-600:]}",
                          f"Failed to start:\n\n{detail[-600:]}"),
                i18n.t("DWG 批量转 PDF", "DWG Batch to PDF"), 0x10)
        except Exception:
            pass

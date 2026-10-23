#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dwg2pdf.py —— 把 DWG 文件里的每一张图纸（布局 Layout）单独导出为一个 PDF

工作原理
--------
调用本机已安装的 AutoCAD 的 COM 接口（AutoCAD ActiveX），逐个打开 DWG，
把每个图纸布局设为当前布局后执行 PlotToFile，每个布局产出一个 PDF。
因为是 AutoCAD 自己出图，线宽、出图比例、打印样式（CTB/STB）、图框、
外部参照的显示效果都与在 CAD 里手工打印一致。

依赖
----
Windows + 已安装并可用（已授权）的 AutoCAD 2013 及以上 + pywin32

命令行示例
----------
    # 预览：只列出会生成哪些 PDF，不真正出图
    python dwg2pdf.py -s "D:\\图纸" -o "D:\\PDF" --dry-run

    # 输出到 D:\\PDF，文件名加前缀 "A-"，例如 "A-一层平面图.pdf"
    python dwg2pdf.py -s "D:\\图纸" -o "D:\\PDF" -p "A-"

    # 保持源目录的文件夹结构，跳过已存在的 PDF（断点续转）
    python dwg2pdf.py -s "D:\\图纸" -o "D:\\PDF" --keep-subdir --skip-existing
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
import traceback
from dataclasses import dataclass, asdict
from typing import Callable, Optional

# ---------------------------------------------------------------- 常量

DWG_TO_PDF = "DWG To PDF.pc3"          # AutoCAD 自带的 PDF 打印设备
MODEL_LAYOUT = "model"                  # 模型空间布局的内部名称
TEMP_EXT = ".part"                      # 出图先用这个扩展名，避免 AutoCAD 自动打开成品
MONO_CTB = "monochrome.ctb"            # 图纸没设打印样式表时用的黑白样式
INVALID_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')
RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

# 标准图幅（mm，长边 x 短边）。模型空间里多图排布时，图框一般按「图幅 x 出图比例」放大绘制
PAPER_SIZES = {
    "A0": (1189, 841),
    "A1": (841, 594),
    "A2": (594, 420),
    "A3": (420, 297),
    "A4": (297, 210),
}
# 常见出图比例，用于把「模型空间里的图框尺寸」还原成图幅
KNOWN_SCALES = (1, 2, 2.5, 5, 10, 20, 25, 30, 40, 50, 75, 100, 125, 150,
                200, 250, 300, 400, 500, 600, 750, 1000, 1500, 2000)
FRAME_BLOCK = re.compile(r"^A[0-4]$", re.I)          # 图框块命名约定：A0 / A1 / A2 / A3 / A4
SHEET_CODE = re.compile(r"^[A-Za-z]{1,4}[-_ ]?\d{1,4}$")   # 图号，如 FD-09 / A-101
MTEXT_CODES = re.compile(r"\\[A-Za-z][^;\\]*;?|[{}]")       # MText 里的格式控制符


# ---------------------------------------------------------------- 数据结构

@dataclass
class LayoutResult:
    """一个图纸（布局）的导出结果"""
    dwg: str            # 源 DWG 完整路径
    layout: str         # 图纸名（布局名）
    pdf: str = ""       # 输出 PDF 完整路径
    status: str = "ok"  # ok / skip / fail
    message: str = ""

    def to_dict(self):
        d = asdict(self)
        d["dwg"] = os.path.basename(self.dwg)
        d["pdf"] = os.path.basename(self.pdf)
        return d


# ---------------------------------------------------------------- 文件名处理

def _remove_quietly(path: str) -> None:
    """删文件，不存在或删不掉都不报错"""
    try:
        if path and os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass


def clean_text(s: str) -> str:
    """清理 AutoCAD 文字里的格式代码（MText 常见）"""
    out = str(s or "")
    out = out.replace("\\P", " ").replace("\\~", " ")
    out = out.replace("%%c", "Ø").replace("%%d", "°").replace("%%p", "±")
    out = MTEXT_CODES.sub("", out)
    out = re.sub(r"\s+", " ", out)
    return out.strip()


def safe_filename(name: str, maxlen: int = 120) -> str:
    """把任意字符串变成 Windows 合法文件名"""
    s = INVALID_CHARS.sub("_", str(name))
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip(" .")                      # Windows 不允许结尾的点和空格
    if not s:
        s = "未命名图纸"
    if s.upper() in RESERVED_NAMES:        # CON / PRN 之类的保留名
        s = "_" + s
    if len(s) > maxlen:
        s = s[:maxlen].strip(" .")
    return s


def build_filename(tpl: str, prefix: str, layout: str, dwg_stem: str, index: int) -> str:
    """
    按模板拼文件名（不含 .pdf）
    可用占位符：{prefix} 前缀、{layout} 图纸名、{dwg} 源文件名、{n} 序号
    用户模板写错占位符时不报错，退回默认的「前缀+图纸名」
    """
    data = {"prefix": prefix, "layout": layout, "dwg": dwg_stem, "n": index}
    try:
        name = tpl.format(**data)
    except (KeyError, IndexError, ValueError):
        name = "{prefix}{layout}".format(**data)
    return safe_filename(name)


# ---------------------------------------------------------------- 扫描 DWG

def find_dwg_files(src: str, recursive: bool = True, exclude_dirs=()) -> list[str]:
    """找出源文件夹下所有 .dwg；跳过 ~$ 临时文件、隐藏目录和输出目录"""
    src = os.path.abspath(src)
    excluded = {os.path.normcase(os.path.abspath(p)) for p in exclude_dirs if p}
    found: list[str] = []

    if not recursive:
        for name in sorted(os.listdir(src)):
            full = os.path.join(src, name)
            if name.lower().endswith(".dwg") and not name.startswith("~$") and os.path.isfile(full):
                found.append(full)
        return found

    for root, dirs, names in os.walk(src):
        dirs[:] = [
            d for d in dirs
            if not d.startswith(".")
            and os.path.normcase(os.path.abspath(os.path.join(root, d))) not in excluded
        ]
        for name in sorted(names):
            if name.lower().endswith(".dwg") and not name.startswith("~$"):
                found.append(os.path.join(root, name))
    # 先处理上层目录、再处理子目录，输出重名时先到先得的结果更符合直觉
    return sorted(found, key=lambda p: (p.count(os.sep), p.lower()))


# ---------------------------------------------------------------- AutoCAD 引擎

class AcadPlotter:
    """封装 AutoCAD COM 操作：打开图纸、逐布局出图、关闭图纸"""

    def __init__(self, pc3: Optional[str] = None, include_model: bool = False,
                 force_pc3: bool = False, frames_mode: str = "auto",
                 log: Optional[Callable[[str], None]] = None):
        self.pc3 = pc3 or DWG_TO_PDF
        self.include_model = include_model
        self.force_pc3 = force_pc3
        self.frames_mode = frames_mode      # auto / frames / layouts
        self.log = log or (lambda msg: None)
        self.app = None
        self._pythoncom = None
        self._owned = False        # True 表示 AutoCAD 是我们启动的
        self._keeper = None        # 守护文档：防止关掉最后一个图纸后 AutoCAD 自动退出
        self.version = ""

    # -------------------------------------------------- 连接
    def connect(self):
        """连接已在运行的 AutoCAD；没有就启动一个（首次约需 20~60 秒）"""
        try:
            import pythoncom            # noqa: F401
            import win32com.client      # noqa: F401
        except ImportError:
            raise RuntimeError(
                "缺少 pywin32 组件，无法调用 AutoCAD。\n"
                "请在命令行执行：pip install pywin32"
            )

        self._pythoncom = pythoncom
        pythoncom.CoInitialize()         # 工作线程里必须初始化 COM

        # 1) 先尝试复用已经打开的 AutoCAD（用户可能正开着图，绝不强杀）
        try:
            self.app = win32com.client.GetActiveObject("AutoCAD.Application")
            self._owned = False
        except Exception:
            self.app = None

        # 2) 没有运行中的实例，则按版本从新到旧启动
        if self.app is None:
            last_err = None
            for progid in ("AutoCAD.Application.24.3", "AutoCAD.Application.24",
                           "AutoCAD.Application.25", "AutoCAD.Application"):
                try:
                    self.app = win32com.client.Dispatch(progid)
                    self._owned = True
                    break
                except Exception as exc:      # noqa: BLE001
                    last_err = exc
            if self.app is None:
                raise RuntimeError(
                    "无法启动 AutoCAD。请确认已正确安装并已授权（未过试用期/已登录）。\n"
                    f"最后一条错误：{last_err}"
                )

        try:
            self.version = str(self.app.Version)
        except Exception:
            self.version = "未知"

        # 冷启动时 AutoCAD 的 COM 接口要过一会儿才真正可用；不等它，
        # 头几张图会报「Open.Layouts」「<unknown>.Item」这种看不出所以然的错。
        self._wait_ready()

        try:
            self.app.Visible = False      # 后台跑，不弹窗口
        except Exception:
            pass

        # 批量出图要反复「打开 → 出图 → 关闭」，AutoCAD 关掉最后一个文档会自动退出，
        # 所以留一个空的守护文档，保证整个批次中途不断线。
        try:
            if self.app.Documents.Count == 0:
                self._keeper = self.app.Documents.Add()
        except Exception:
            self._keeper = None

        # 打开文件时不弹「文件已存在/只读」之类的对话框
        for var, value in (("FILEDIA", 0), ("CMDDIA", 0), ("BACKGROUNDPLOT", 0),
                           ("EXPERT", 5), ("PROXYNOTICE", 0)):
            try:
                self.app.ActiveDocument.SetVariable(var, value)
            except Exception:
                pass

        self.log(f"已连接 AutoCAD {self.version}"
                 f"（{'复用已打开的实例' if not self._owned else '新启动实例'}）")

    def _wait_ready(self, timeout: float = 180.0):
        """等 AutoCAD 的文档模型真正能应答（冷启动约 20~60 秒）"""
        t0 = time.time()
        waiting = False
        while True:
            try:
                int(self.app.Documents.Count)
                try:
                    self.app.ActiveDocument.Name
                except Exception:
                    pass                       # 没打开图纸时会报错，属正常
                if waiting:
                    self.log(f"  AutoCAD 已就绪（等了 {time.time() - t0:.0f} 秒）")
                return
            except Exception:
                pass
            if time.time() - t0 > timeout:
                self.log("  ⚠ AutoCAD 长时间没有响应，仍继续尝试出图")
                return
            if not waiting:
                waiting = True
                self.log("  AutoCAD 还在启动，等它就绪…")
            time.sleep(1.5)

    # -------------------------------------------------- 内部工具
    @staticmethod
    def _setvar(doc, name, value):
        try:
            doc.SetVariable(name, value)
        except Exception:
            pass

    def _layout_names(self, doc) -> list[str]:
        """取出所有要出图的布局名；默认排除模型空间"""
        names = []
        count = doc.Layouts.Count
        for i in range(count):
            layout = doc.Layouts.Item(i)
            name = str(layout.Name)
            if not self.include_model and name.strip().lower() == MODEL_LAYOUT:
                continue
            names.append(name)

        # 只有模型空间、没有任何图纸布局的图（简单图常见）：按一张图纸输出，
        # 否则这种文件会「什么都没生成」，容易让人以为程序没工作。
        if not names and not self.include_model:
            self.log("  （该文件没有图纸布局，按模型空间输出一张 Model.pdf）")
            names = ["Model"]
        return names

    # -------------------------------------------------- 模型空间多图框识别
    @staticmethod
    def _ss(doc, name: str):
        """取一个干净的选择集（同名先删掉）"""
        try:
            doc.SelectionSets.Item(name).Delete()
        except Exception:
            pass
        return doc.SelectionSets.Add(name)

    @staticmethod
    def _double_array(vals):
        import pythoncom
        from win32com.client import VARIANT
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list(vals))

    @staticmethod
    def _variant_array(vals):
        import pythoncom
        from win32com.client import VARIANT
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, list(vals))

    @staticmethod
    def _short_array(vals):
        import pythoncom
        from win32com.client import VARIANT
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list(vals))

    def _name_may_be_frame(self, name: str) -> bool:
        """块名里是否可能写着图幅（A0-A4，或 1051-594 这种尺寸）"""
        if FRAME_BLOCK.match(name):
            return True
        nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", name)]
        for a, b in zip(nums, nums[1:]):
            if 150 <= b <= 3200 and b <= a <= 3200 and self._is_known_sheet_mm(a, b):
                return True
        return False

    def _block_frame_rect(self, doc, block_name: str):
        """块定义里最大的闭合矩形（就是图框），返回块坐标系下的 (x0,y0,x1,y1)"""
        if not hasattr(self, "_rect_cache"):
            self._rect_cache = {}
        if block_name in self._rect_cache:
            return self._rect_cache[block_name]
        rect = None
        try:
            blk = doc.Blocks.Item(block_name)
            best = None
            for i in range(blk.Count):
                e = blk.Item(i)
                if str(e.EntityName) != "AcDbPolyline":
                    continue
                try:
                    if not e.Closed:
                        continue
                    c = list(e.Coordinates)
                except Exception:
                    continue
                xs, ys = c[0::2], c[1::2]
                w, h = max(xs) - min(xs), max(ys) - min(ys)
                if w > 0 and h > 0 and (best is None or w * h > best[4]):
                    best = (min(xs), min(ys), max(xs), max(ys), w * h)
            if best:
                rect = best[:4]
        except Exception:
            rect = None
        self._rect_cache[block_name] = rect
        return rect

    def _frames_from_blocks(self, doc) -> list[dict]:
        """
        识别图框块：类型一按命名约定 A0/A1/A2/A3/A4；
        类型二（如天线图框块 yct-1051-594-105100-59400-100）从块名里解析出图幅尺寸，
        这样 A1 加长（1051x594）这类非标准图幅也能识别。
        """
        frames = []
        try:
            ss = self._ss(doc, "d2p_ins")
            ss.Select(5, 0, 0, self._short_array([0]), self._variant_array(["INSERT"]))
        except Exception:
            return frames
        import math
        total = com_get(lambda: ss.Count, 0)
        if not total:
            return frames

        def read_one(idx: int):
            """读完一个块参照 → 图框字典；不是图框返回 None"""
            e = ss.Item(idx)
            name = str(e.Name).strip()
            # 先看名字像不像图框，再决定要不要去翻块定义。
            # 建筑图里常有搬来的匿名块（A$Ca5f0b8dd 之类，块内上千个实体），
            # 逐个翻一遍会白白卡上几分钟。
            if not self._name_may_be_frame(name):
                return None
            rect = self._block_frame_rect(doc, name)
            if not rect:
                return None
            p = e.InsertionPoint
            sx, sy = float(e.XScaleFactor), float(e.YScaleFactor)
            rot = float(e.Rotation)
            cos_r, sin_r = math.cos(rot), math.sin(rot)
            xs, ys = [], []
            for bx, by in ((rect[0], rect[1]), (rect[2], rect[1]),
                           (rect[2], rect[3]), (rect[0], rect[3])):
                bx, by = bx * sx, by * sy
                xs.append(p[0] + bx * cos_r - by * sin_r)
                ys.append(p[1] + bx * sin_r + by * cos_r)
            w_units, h_units = abs(max(xs) - min(xs)), abs(max(ys) - min(ys))

            if FRAME_BLOCK.match(name):
                paper, paper_mm, scale = name.upper(), PAPER_SIZES[name.upper()], None
            else:
                parsed = self._paper_from_block_name(name, w_units, h_units)
                if not parsed:
                    return None
                paper_mm, scale = parsed
                paper = self._paper_label(*paper_mm)
            return {"paper": paper, "paper_mm": paper_mm, "scale": scale,
                    "x0": min(xs), "x1": max(xs),
                    "y0": min(ys), "y1": max(ys)}

        for i in range(total):
            if total > 200 and i and i % 100 == 0:
                self.log(f"    · 已检查 {i}/{total} 个块参照…")
            f = com_get(lambda i=i: read_one(i), None)
            if f:
                frames.append(f)
        return frames

    @staticmethod
    def _paper_label(width_mm: float, height_mm: float) -> str:
        """毫米尺寸 → 图幅名：标准的就是 A0-A4，否则写成 1051x594 这种"""
        long_, short_ = max(width_mm, height_mm), min(width_mm, height_mm)
        for name, (pw, ph) in PAPER_SIZES.items():
            if abs(long_ - pw) <= 2 and abs(short_ - ph) <= 2:
                return name
        return f"{round(long_)}x{round(short_)}"

    @staticmethod
    def _is_known_sheet_mm(width_mm: float, height_mm: float) -> bool:
        """短边是标准 A 系列的尺寸，就当成图幅（含 A1 加长这类长边变长的）"""
        short_ = min(width_mm, height_mm)
        long_ = max(width_mm, height_mm)
        return (any(abs(short_ - v) <= 2 for v in (210, 297, 420, 594, 841, 1189))
                and long_ >= short_ - 2)

    def _paper_from_block_name(self, name: str, w_units: float, h_units: float):
        """
        从块名里解析图幅尺寸，用「图纸单位 / 比例 = 毫米」来校验，避免误读。
        yct-1051-594-105100-59400-100 → (1051, 594)
        """
        nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", name)]
        for i in range(len(nums) - 1):
            a, b = nums[i], nums[i + 1]
            if b < 150 or a > 3200 or a < b:
                continue
            if not self._is_known_sheet_mm(a, b):
                continue
            for s in KNOWN_SCALES:
                if abs(w_units / s - a) <= 2 and abs(h_units / s - b) <= 2:
                    return (a, b), s
        return None

    def _frames_from_polylines(self, doc) -> list[dict]:
        """没有约定命名的图框块时，靠「标准图幅 x 常见比例」识别独立矩形图框"""
        frames = []
        try:
            ss = self._ss(doc, "d2p_poly")
            ss.Select(5, 0, 0, self._short_array([0]), self._variant_array(["LWPOLYLINE"]))
        except Exception:
            return frames
        # 数量只问一次：在循环条件里反复访问 ss.Count 就是每个实体一次 COM 调用，
        # 几千条多段线会慢到离谱，而且 AutoCAD 一忙就在这里抛 Call was rejected by callee。
        total = com_get(lambda: ss.Count, 0)
        if not total:
            return frames

        def read_one(idx: int):
            e = ss.Item(idx)
            if not e.Closed:
                return None
            c = list(e.Coordinates)
            xs, ys = c[0::2], c[1::2]
            x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
            long_, short_ = max(x1 - x0, y1 - y0), min(x1 - x0, y1 - y0)
            if short_ <= 0:
                return None
            for scale in KNOWN_SCALES:
                w_mm, h_mm = long_ / scale, short_ / scale
                if not (150 <= h_mm <= 3200):
                    continue
                if abs(w_mm - round(w_mm)) > 0.6 or abs(h_mm - round(h_mm)) > 0.6:
                    continue        # 不是整数毫米，基本不是图框
                if not self._is_known_sheet_mm(w_mm, h_mm):
                    continue
                return {"paper": self._paper_label(w_mm, h_mm),
                        "paper_mm": (round(w_mm), round(h_mm)),
                        "scale": scale,
                        "x0": x0, "x1": x1, "y0": y0, "y1": y1}
            return None

        for i in range(total):
            if total > 500 and i and i % 200 == 0:
                self.log(f"    · 已检查 {i}/{total} 条多段线…")
            f = com_get(lambda i=i: read_one(i), None)
            if f:
                frames.append(f)
        return frames

    @staticmethod
    def _frame_area(f) -> float:
        return (f["x1"] - f["x0"]) * (f["y1"] - f["y0"])

    def _dedupe_frames(self, frames: list[dict]) -> list[dict]:
        """同一张图既有块又有外框时只留一个（块优先，因为能确定图幅）"""
        def overlap(a, b):
            ix = min(a["x1"], b["x1"]) - max(a["x0"], b["x0"])
            iy = min(a["y1"], b["y1"]) - max(a["y0"], b["y0"])
            if ix <= 0 or iy <= 0:
                return 0.0
            return ix * iy / max(1.0, min(self._frame_area(a), self._frame_area(b)))

        ordered = sorted(frames, key=lambda f: ("scale" in f, -self._frame_area(f)))
        kept: list[dict] = []
        for f in ordered:
            if all(overlap(f, k) < 0.8 for k in kept):
                kept.append(f)
        return kept

    def _all_texts(self, doc) -> list[tuple]:
        """
        一次性取出模型空间里所有文字（x, y, 字高, 内容）并缓存。
        比每张图都做一次窗口选择快得多，也不受视图范围/坐标系影响。
        """
        if getattr(self, "_text_cache", None) is not None:
            return self._text_cache
        items: list[tuple] = []
        ok = True
        try:
            ss = self._ss(doc, "d2p_alltxt")
            ss.Select(5, 0, 0, self._short_array([0]), self._variant_array(["TEXT,MTEXT"]))
            total = com_get(lambda: ss.Count, 0)
            failed = 0
            for i in range(total):
                t = com_get(lambda i=i: self._one_text(ss, i), _MISS)
                if t is _MISS:
                    failed += 1          # 读不到（多半是 AutoCAD 忙），别当成空文字
                elif t:
                    items.append(t)
            if failed > max(20, total // 10):
                ok = False               # 缺得太多，这次的结果不能缓存，下次重新取
        except Exception:
            ok = False
        if ok:
            self._text_cache = items
        return items

    @staticmethod
    def _one_text(ss, idx: int):
        """读一条文字 → (x, y, 字高, 内容)；空文字返回 ()"""
        e = ss.Item(idx)
        p = e.InsertionPoint
        s = clean_text(str(e.TextString))
        if not s:
            return ()
        return (p[0], p[1], float(getattr(e, "Height", 0) or 0), s)

    def _read_title_block(self, doc, f: dict) -> tuple[str, str]:
        """
        从图框右下角标题栏里读图名和图号。
        经验：标题栏里图名在上、图号（如 FD-09 / FG-101）在下，所以先找图号，
        再取它上方那一格里的文字（可能是一行也可能两行）。
        """
        w, h = f["x1"] - f["x0"], f["y1"] - f["y0"]
        rx0, ry0 = f["x1"] - w * 0.32, f["y0"]
        rx1, ry1 = f["x1"], f["y0"] + h * 0.25
        items = [t for t in self._all_texts(doc)
                 if rx0 <= t[0] <= rx1 and ry0 <= t[1] <= ry1]
        if not items:
            return "", ""

        codes = [t for t in items if SHEET_CODE.match(t[3]) and not t[3].isdigit()]
        texts = [t for t in items if not SHEET_CODE.match(t[3]) and len(t[3]) >= 2]
        code = ""
        title = ""
        if codes:
            anchor = min(codes, key=lambda t: t[1])          # 图号在最下面
            code = anchor[3]
            # 图名在「图号正上方那一格」里，可能是一行也可能是两行；
            # 用纵向范围和横向范围圈住那一格，再按从上到下拼起来。
            top = anchor[1] + h * 0.10
            x_tol = max(w * 0.10, 8000)
            lines = [t for t in texts
                     if anchor[1] < t[1] <= top
                     and abs(t[0] - anchor[0]) <= x_tol
                     and not t[3].isdigit()]
            lines.sort(key=lambda t: -t[1])                  # 上面的在前
            if lines:
                seen = []
                for t in lines:
                    if t[3] not in seen:
                        seen.append(t[3])
                title = " ".join(seen)
        if not title and texts:
            title = max(texts, key=lambda t: (t[2], len(t[3])))[3]
        return title, code

    def detect_sheets(self, doc) -> list[dict]:
        """
        找出模型空间里排布的所有图纸：
        先按图框块命名（A0-A4），找不到再按「标准图幅 x 比例」找矩形。
        返回按「从上到下、从左到右」排序的图纸列表。
        """
        frames = self._frames_from_blocks(doc)
        source = "块"
        if not frames:
            frames = self._frames_from_polylines(doc)
            source = "矩形"
        if not frames:
            return []
        frames = self._dedupe_frames(frames)
        frames.sort(key=lambda f: (-round(f["y1"] / 100.0), f["x0"]))
        for i, f in enumerate(frames, 1):
            f["index"] = i
            f["title"], f["code"] = self._read_title_block(doc, f)
            if not f["title"]:
                f["title"] = f["code"] or f"图纸{i:02d}"
            f["name"] = f["title"]
        self.log(f"  （按{source}识别出 {len(frames)} 张图纸："
                 + "、".join(f"{f['paper']}" for f in frames) + "）")
        return frames

    @staticmethod
    def _media_landscape(name: str):
        """从纸张名里读出方向：'ISO_A1_(841.00_x_594.00_MM)' → 横版"""
        m = re.search(r"(\d+(?:\.\d+)?)_x_(\d+(?:\.\d+)?)", name)
        if m:
            try:
                return float(m.group(1)) > float(m.group(2))
            except ValueError:
                return None
        return None

    @staticmethod
    def _media_size(name: str):
        """从纸张名里读出毫米尺寸：(841, 594)"""
        m = re.search(r"(\d+(?:\.\d+)?)_x_(\d+(?:\.\d+)?)", name)
        if not m:
            return None
        try:
            return float(m.group(1)), float(m.group(2))
        except ValueError:
            return None

    def _pick_media(self, layout, paper: str, paper_mm, want_landscape: bool):
        """
        按图纸的实际毫米尺寸挑纸张：
        1) 有完全一致的纸张就直接用（A0-A4 都能命中）；
        2) 没有（例如 A1 加长 1051x594）就找短边一致的纸张，按缩放输出；
        3) 实在没有就按图幅名前缀找.
        返回 (纸张名, 是否横版, 是否尺寸完全一致)。
        """
        try:
            names = [str(n) for n in layout.GetCanonicalMediaNames()]
        except Exception:
            return "", None, False
        pool = [n for n in names if "FULL_BLEED" not in n.upper()] or names
        pool = [n for n in pool if self._media_size(n)]

        target = paper_mm
        if not target and paper in PAPER_SIZES:
            target = PAPER_SIZES[paper]

        if target:
            tw, th = float(target[0]), float(target[1])
            t_long, t_short = max(tw, th), min(tw, th)
            exact = [n for n in pool
                     for (mw, mh) in [self._media_size(n)]
                     if abs(max(mw, mh) - t_long) <= 2 and abs(min(mw, mh) - t_short) <= 2]
            if exact:
                same = [n for n in exact if self._media_landscape(n) == want_landscape]
                pick = same[0] if same else exact[0]
                return pick, self._media_landscape(pick), True
            # 短边一致就用它（加长图幅走这里）：长边取最接近原图的纸张
            near = [(n, self._media_size(n)) for n in pool
                    if abs(min(self._media_size(n)) - t_short) <= 2]
            if near:
                same_dir = [x for x in near if self._media_landscape(x[0]) == want_landscape] or near
                same_dir.sort(key=lambda x: (abs(max(x[1]) - t_long), -max(x[1])))
                pick = same_dir[0][0]
                return pick, self._media_landscape(pick), False
        up = str(paper).upper()
        cands = [n for n in pool if n.upper().startswith(f"ISO_{up}_")] \
            or [n for n in pool if n.upper().startswith(up)]
        if cands:
            same = [n for n in cands if self._media_landscape(n) == want_landscape]
            pick = same[0] if same else cands[0]
            return pick, self._media_landscape(pick), False
        return "", None, False

    def _plot_sheet(self, doc, sheet: dict, pdf_path: str) -> None:
        """把模型空间里的一张图（按图框窗口 + 对应图幅）打印成一个 PDF"""
        model = None
        for i in range(doc.Layouts.Count):
            if str(doc.Layouts.Item(i).Name).strip().lower() == MODEL_LAYOUT:
                model = doc.Layouts.Item(i)
                break
        if model is None:
            raise RuntimeError("找不到模型空间布局")

        doc.ActiveLayout = model
        try:
            model.ConfigName = self.pc3
        except Exception:
            pass

        # 图纸没有指定打印样式表时，颜色 7（白/黑）会被原样输出成白色，
        # 在 PDF 白底上看不见（曾遇到整张总说明「空白」）。这种情况改用黑白样式。
        try:
            if not str(model.StyleSheet or "").strip():
                model.StyleSheet = MONO_CTB
                model.PlotWithPlotStyles = True
                self.log(f"    · 图纸未设打印样式表，改用 {MONO_CTB} 输出（否则白字会看不见）")
        except Exception:
            pass
        w, h = sheet["x1"] - sheet["x0"], sheet["y1"] - sheet["y0"]
        want_landscape = w >= h
        media, media_landscape, media_exact = self._pick_media(
            model, sheet["paper"], sheet.get("paper_mm"), want_landscape)
        if media:
            try:
                model.CanonicalMediaName = media
            except Exception:
                pass
            if not media_exact:
                self.log(f"    · 图幅 {sheet['paper']} 没有对应纸张，按 {media} 缩放输出")

        # 注意顺序：必须先设置窗口范围，再切到「按窗口出图」。
        # 反过来时，如果图纸原有的窗口范围是空的（很多图 PlotType=0 显示），
        # AutoCAD 会直接报「输入无效」。
        model.SetWindowToPlot(self._double_array([sheet["x0"], sheet["y0"]]),
                              self._double_array([sheet["x1"], sheet["y1"]]))
        model.PlotType = 4                                  # 4 = 窗口
        model.CenterPlot = True
        model.UseStandardScale = True
        model.StandardScale = 0                             # 0 = 布满图纸
        # 纸张方向和图纸方向不一致时旋转 90°（例如 A0 只有竖版纸张）
        model.PlotRotation = 1 if (media_landscape is not None
                                   and media_landscape != want_landscape) else 0

        try:
            doc.Plot.QuietErrorMode = True
        except Exception:
            pass

        temp_path = pdf_path + TEMP_EXT
        _remove_quietly(temp_path)
        last_err = ""
        for attempt in range(2):
            try:
                doc.Plot.PlotToFile(temp_path, self.pc3)
            except Exception as exc:                        # noqa: BLE001
                last_err = str(exc)
                continue
            if self._wait_file(temp_path, timeout=30):
                try:
                    os.replace(temp_path, pdf_path)
                    return
                except OSError as exc:
                    _remove_quietly(pdf_path)
                    try:
                        os.replace(temp_path, pdf_path)
                        return
                    except OSError:
                        last_err = f"PDF 改名失败：{exc}"
                        continue
            last_err = "AutoCAD 未生成文件"
            time.sleep(1.0)
        _remove_quietly(temp_path)
        raise RuntimeError(f"出图未生成文件（图幅 {sheet['paper']}"
                           + (f"；最后错误：{last_err}" if last_err else "") + "）")

    def _plot_one(self, doc, layout_name: str, pdf_path: str) -> None:
        """
        把一个布局打印成 PDF。

        三个关键经验（都是实测踩出来的）：
        1. PlotToFile **必须显式传入打印设备名**。只调 PlotToFile(文件名) 时，
           某些图纸的页面设置解析不出来，AutoCAD 会「静默失败」——不报错、也不生成文件。
           传入设备名后同样的页面设置就能正常出图。所以这里始终传设备名。
        2. 优先沿用图纸里原有的页面设置（图幅、比例、打印样式都不动）；
           原设备不是 .pc3 或第一轮没出文件时，改用 DWG To PDF.pc3，
           并尽量把原来的纸张尺寸还原回去，避免 A0 大图变成 Letter。
        3. **出图不能直接写 .pdf 文件名**：AutoCAD 前台出图结束会按扩展名去把文件
           「打开」一次（PDF 就弹阅读器）。所以先写到 .part 这种无关联的扩展名，
           写完再改名成 .pdf，这样就完全不会弹窗。
        """
        layout = doc.Layouts.Item(layout_name)
        doc.ActiveLayout = layout

        plot = doc.Plot
        try:
            plot.QuietErrorMode = True     # 出图报错时不弹窗，交给程序处理
        except Exception:
            pass

        # 记住原始设置，便于切换设备后恢复纸张大小
        try:
            orig_cfg = str(layout.ConfigName or "")
        except Exception:
            orig_cfg = ""
        try:
            orig_media = str(layout.CanonicalMediaName or "")
        except Exception:
            orig_media = ""

        def switch_to_pdf_device():
            try:
                layout.ConfigName = self.pc3
            except Exception:
                pass
            # 原来的纸张（如 ISO A0）在 DWG To PDF 里一般也有同名规格，试着还原
            if orig_media:
                try:
                    layout.CanonicalMediaName = orig_media
                except Exception:
                    pass

        # 候选方案：先用图纸原设备（显式传名），不行再用 DWG To PDF.pc3
        plans: list[str] = []
        native_ok = (not self.force_pc3) and orig_cfg.lower().endswith(".pc3")
        if native_ok:
            plans.append(orig_cfg)
        if orig_cfg.strip().lower() != self.pc3.strip().lower():
            plans.append(self.pc3)

        temp_path = pdf_path + TEMP_EXT          # .part：无关联扩展名，不会被自动打开
        last_err = ""
        for cfg in plans:
            if cfg.strip().lower() != orig_cfg.strip().lower():
                switch_to_pdf_device()
            try:
                _remove_quietly(temp_path)
                plot.PlotToFile(temp_path, cfg)     # ← 传设备名，且写临时扩展名
            except Exception as exc:               # noqa: BLE001
                last_err = str(exc)
                continue
            if self._wait_file(temp_path, timeout=20):
                try:
                    os.replace(temp_path, pdf_path)
                    return
                except OSError as exc:
                    # 目标 PDF 被占用（例如正被阅读器打开）：删掉再试一次
                    try:
                        _remove_quietly(pdf_path)
                        os.replace(temp_path, pdf_path)
                        return
                    except OSError:
                        last_err = f"PDF 改名失败：{exc}"
                        continue
            last_err = "AutoCAD 未生成文件"

        _remove_quietly(temp_path)
        raise RuntimeError(
            f"出图未生成文件（已尝试 {len(plans)} 种打印设备："
            + "、".join(plans) + (f"；最后错误：{last_err}" if last_err else "") + "）"
        )

    @staticmethod
    def _wait_file(path: str, timeout: float = 30.0) -> bool:
        """等 PDF 落盘（后台打印时有延迟），并确认不是 0 字节"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            if os.path.isfile(path) and os.path.getsize(path) > 0:
                return True
            time.sleep(0.25)
        return False

    # -------------------------------------------------- 主流程
    def _reset_cache(self):
        self._text_cache = None
        self._rect_cache = {}
        self._fonts_repaired = False
        self._font_cache = None

    def _targets(self, doc) -> list[dict]:
        """
        决定这个文件要输出哪些「图纸」：
        - 模型空间里排了多张图框（A0-A4 块，或标准尺寸矩形）→ 每张图框一个 PDF
        - 否则按图纸布局（Layout）逐个输出
        """
        layouts = self._layout_names(doc)
        if self.frames_mode != "layouts":
            # 默认只在「图纸布局都是空的」时才拆图框——那正是「多张图全排在模型空间」的画法；
            # 布局里已经有内容的图，仍按布局输出，保持原有行为。
            use_frames = self.frames_mode == "frames"
            if not use_frames:
                use_frames = all(self._layout_entity_count(doc, n) == 0 for n in layouts)
            if use_frames:
                sheets = []
                for attempt in range(3):
                    try:
                        sheets = self.detect_sheets(doc)
                        break
                    except Exception as exc:                # noqa: BLE001
                        if attempt >= 2:
                            self.log(f"  × 图框识别出错：{_humanize_com_error(exc)}")
                            break
                        wait = 2 * (attempt + 1)
                        self.log(f"  · 识别图框时 AutoCAD 正忙，{wait} 秒后重试…")
                        time.sleep(wait)
                if sheets:
                    return sheets
                if self.frames_mode == "frames":
                    raise RuntimeError("已指定按图框拆分，但没识别到图框"
                                       "（模型空间里需要有 A0-A4 图框块或标准图幅的矩形边框）")
                self.log("  ! 没识别到图框，改为按图纸布局输出"
                         "（这张图如果多张图纸都排在模型空间里，就会只出一张）")
        return [{"name": n, "paper": "", "x0": None} for n in layouts]

    # -------------------------------------------------- 缺失字体修复
    def _font_dirs(self) -> list[str]:
        """能放字体的目录：AutoCAD 支持路径 + Windows 字体目录"""
        dirs = []
        try:
            support = str(self.app.Preferences.Files.SupportPath or "")
            dirs += [d for d in support.split(";") if d and os.path.isdir(d)]
        except Exception:
            pass
        for d in (r"C://Windows//Fonts",
                  os.path.join(os.environ.get("WINDIR", r"C://Windows"), "Fonts")):
            if d and os.path.isdir(d) and d not in dirs:
                dirs.append(d)
        return dirs

    def _available_fonts(self) -> set[str]:
        if getattr(self, "_font_cache", None) is not None:
            return self._font_cache
        names: set[str] = set()
        for d in self._font_dirs():
            try:
                for f in os.listdir(d):
                    names.add(f.lower())
            except OSError:
                pass
        self._font_cache = names
        return names

    def _font_exists(self, name: str) -> bool:
        if not name:
            return True
        low = os.path.basename(name).lower()
        fonts = self._available_fonts()
        if low in fonts:
            return True
        if not os.path.splitext(low)[1] and (low + ".shx") in fonts:
            return True
        return False

    def repair_missing_fonts(self, doc) -> int:
        """
        图纸引用了本机没有的字体（探索者 tssdeng.shx、hztxt.shx、PKPM 等）时，
        AutoCAD 算不出这些文字的范围，出图时**整段文字直接不打印**（表现为图纸上一片空白）。
        这里把它们临时换成本机有的字体（宋体 + GB 大字体），只在内存里改，不保存原图。
        """
        if getattr(self, "_fonts_repaired", False):
            return 0
        self._fonts_repaired = True

        main_alt = next((p for p in (r"C://Windows//Fonts//simsun.ttc",
                                     r"C://Windows//Fonts//msyh.ttc") if os.path.isfile(p)), "")
        big_alt = next((os.path.join(d, "gbcbig.shx") for d in self._font_dirs()
                        if os.path.isfile(os.path.join(d, "gbcbig.shx"))), "")

        fixed = 0
        try:
            count = doc.TextStyles.Count
        except Exception:
            return 0
        for i in range(count):
            try:
                st = doc.TextStyles.Item(i)
                main = str(st.fontFile or "").strip()
                big = str(st.BigFontFile or "").strip()
            except Exception:
                continue
            changed = False
            if main and not self._font_exists(main) and main_alt:
                try:
                    st.fontFile = main_alt
                    changed = True
                except Exception:
                    pass
            if big and not self._font_exists(big) and big_alt:
                try:
                    st.BigFontFile = big_alt
                    changed = True
                except Exception:
                    pass
            if changed:
                fixed += 1
        if fixed:
            try:
                doc.Regen(1)          # 重生成后文字才有正确范围，才能被打印
            except Exception:
                pass
            self.log(f"    · 图纸有 {fixed} 个文字样式用了本机没有的字体，"
                     f"已临时替换（否则这些字不会出图）")
        return fixed

    @staticmethod
    def _layout_entity_count(doc, name: str) -> int:
        """布局（图纸空间）里有多少实体；0 表示这是个空布局"""
        try:
            return int(doc.Layouts.Item(name).Block.Count)
        except Exception:
            return 0

    def export_drawing(self, dwg_path: str, namer: Callable[[list[dict]], list[tuple[dict, str]]],
                       skip_existing: bool = False) -> list[LayoutResult]:
        """
        处理一个 DWG：只打开一次文件，先确定要出哪些图纸，再逐个出图。
        单个图纸失败不影响同一个文件的其它图纸。
        """
        results: list[LayoutResult] = []
        win_path = os.path.abspath(dwg_path)

        try:
            # True = 只读打开，绝不改动原图；AutoCAD 忙时自动重试
            doc = com_retry(lambda: self.app.Documents.Open(win_path, True))
        except Exception as exc:                            # noqa: BLE001
            msg = _humanize_com_error(exc)
            return [LayoutResult(win_path, "", "", "fail", f"打开图纸失败：{msg}")]

        try:
            self._reset_cache()
            for var, value in (("BACKGROUNDPLOT", 0), ("FILEDIA", 0), ("CMDDIA", 0)):
                self._setvar(doc, var, value)

            self.repair_missing_fonts(doc)
            targets = com_retry(lambda: self._targets(doc))
            plan = namer(targets)

            for target, pdf_path in plan:
                name = target["name"]
                is_sheet = target.get("x0") is not None
                if skip_existing and os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 0:
                    results.append(LayoutResult(win_path, name, pdf_path, "skip",
                                                "PDF 已存在，已跳过"))
                    continue

                os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
                if os.path.isfile(pdf_path):
                    try:
                        os.remove(pdf_path)       # 清掉旧文件，避免把旧 PDF 当成本次成果
                    except OSError:
                        pass

                try:
                    if is_sheet:
                        com_retry(lambda t=target, p=pdf_path: self._plot_sheet(doc, t, p))
                    else:
                        com_retry(lambda n=name, p=pdf_path: self._plot_one(doc, n, p))
                except Exception as exc:          # noqa: BLE001
                    results.append(LayoutResult(win_path, name, pdf_path, "fail",
                                                f"出图失败：{_humanize_com_error(exc)}"))
                    continue

                if self._wait_file(pdf_path):
                    size_kb = os.path.getsize(pdf_path) / 1024
                    extra = f"{target['paper']} · " if is_sheet and target.get("paper") else ""
                    results.append(LayoutResult(win_path, name, pdf_path, "ok",
                                                f"{extra}{size_kb:.0f} KB"))
                else:
                    results.append(LayoutResult(win_path, name, pdf_path, "fail",
                                                "没有生成 PDF 文件（可能是图纸为空或打印设备不可用）"))
        finally:
            try:
                com_retry(lambda: doc.Close(False), attempts=4)   # 不保存，保持原图不动
            except Exception:
                pass
            time.sleep(0.2)          # 给 AutoCAD 一点收尾时间，降低下一个文件被拒的概率
        return results

    def walk_layouts(self, dwg: str,
                     namer: Callable[[list[dict]], list[tuple[dict, str]]]) -> list[LayoutResult]:
        """试运行/预览：只打开图纸列出将来会生成哪些文件，不出图"""
        win_path = os.path.abspath(dwg)
        try:
            doc = com_retry(lambda: self.app.Documents.Open(win_path, True))
        except Exception as exc:                        # noqa: BLE001
            self.log(f"  × 打开失败：{_humanize_com_error(exc)}")
            return []
        try:
            self._reset_cache()
            self.repair_missing_fonts(doc)
            plan = namer(com_retry(lambda: self._targets(doc)))
        finally:
            try:
                com_retry(lambda: doc.Close(False), attempts=4)
            except Exception:
                pass
            time.sleep(0.2)
        return [LayoutResult(win_path, target["name"], pdf_path, "preview",
                             f"图幅 {target['paper']}" if target.get("paper") else "预览")
                for target, pdf_path in plan]

    def close(self, keep_acad: bool = False):
        """
        收尾：
        - 守护文档是我们加的，要关掉，否则会留下一个「Drawing1」开着
        - 如果 AutoCAD 是本程序启动的且没有其它图纸了，关掉它即自动退出，不留后台进程
        - keep_acad=True 时保留进程，下次转换更快
        """
        try:
            if self._keeper is not None:
                keeper, self._keeper = self._keeper, None
                try:
                    keeper.Close(False)
                except Exception:
                    pass
            if keep_acad:
                self.log("AutoCAD 将保留在后台，下次转换启动更快")
        finally:
            try:
                if self._pythoncom:
                    self._pythoncom.CoUninitialize()
            except Exception:
                pass


def _humanize_com_error(exc: Exception) -> str:
    text = str(exc).strip() or exc.__class__.__name__
    low = text.lower()
    if "文件已打开" in text or "already open" in low:
        return "该 DWG 已在 AutoCAD 中打开，请先关闭后重试"
    if "拒绝访问" in text or "access is denied" in low:
        return "文件被占用或没有读取权限"
    if "找不到" in text or "not found" in low or "无法打开" in text:
        return "文件不存在、已损坏，或由更高版本 CAD 保存（需用同版本或更高版本打开）"
    return text.replace("\r", " ").replace("\n", " ")[:200]


# AutoCAD 在出图/关闭图纸的过程中会短暂「忙」，此时 COM 调用被拒绝
# （-2147418111 = Call was rejected by callee）。等一会儿再试即可，不是真失败。
_BUSY_CODES = {-2147418111, -2147417846}
_BUSY_TEXT = ("call was rejected", "被调用方拒绝", "忙", "server busy",
              "rpc_e_call_rejected")
# pywin32 晚绑定时偶发「成员暂时取不到」，本质上也是对方还没就绪，重试即可：
# 报错文本形如 "Open.Layouts"、"<unknown>.Item"
_COM_MISS_TEXT = ("open.layouts", "<unknown>", "layouts.item", "找不到成员",
                  "member not found")


def _is_busy(exc: Exception) -> bool:
    if getattr(exc, "hresult", None) in _BUSY_CODES:      # pywintypes.com_error
        return True
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int) and args[0] in _BUSY_CODES:
        return True
    text = str(exc).lower()
    return any(t in text for t in _BUSY_TEXT) or any(t in text for t in _COM_MISS_TEXT)


_MISS = object()


def com_get(func, default=None, attempts: int = 6, delay: float = 0.6):
    """
    读单个实体的属性：AutoCAD 忙（Call was rejected by callee）就等一会儿重试，
    其它错误（实体坏了之类）直接当没有，绝不让它冒出去打断整个扫描。
    """
    for i in range(attempts):
        try:
            return func()
        except Exception as exc:                          # noqa: BLE001
            if not _is_busy(exc) or i == attempts - 1:
                return default
            time.sleep(delay * (i + 1))
    return default


def com_retry(func, attempts: int = 8, delay: float = 0.5):
    """COM 调用重试：只在「对方忙」时重试，其它错误立刻抛出"""
    last = None
    for i in range(attempts):
        try:
            return func()
        except Exception as exc:                          # noqa: BLE001
            last = exc
            if not _is_busy(exc) or i == attempts - 1:
                raise
            time.sleep(delay * (i + 1))
    raise last  # pragma: no cover


# ---------------------------------------------------------------- 调度器

class Converter:
    """扫描 → 规划命名 → 逐文件出图 → 汇总"""

    def __init__(self, src: str, out: str, prefix: str = "",
                 name_tpl: str = "{prefix}{layout}", include_model: bool = False,
                 recursive: bool = True, keep_subdir: bool = False,
                 skip_existing: bool = False, pc3: Optional[str] = None,
                 on_event: Optional[Callable[[dict], None]] = None,
                 cancel: Optional[Callable[[], bool]] = None,
                 dry_run: bool = False, keep_acad: bool = False,
                 frames_mode: str = "auto"):
        self.src = os.path.abspath(src)
        # 允许直接把单个 .dwg 文件当作输入
        self.single_file = None
        if os.path.isfile(self.src) and self.src.lower().endswith(".dwg"):
            self.single_file = self.src
            self.src = os.path.dirname(self.src)
        self.out = os.path.abspath(out)
        self.prefix = prefix or ""
        self.name_tpl = name_tpl or "{prefix}{layout}"
        self.include_model = include_model
        self.recursive = recursive
        self.keep_subdir = keep_subdir
        self.skip_existing = skip_existing
        self.pc3 = pc3
        self.keep_acad = keep_acad
        self.frames_mode = frames_mode
        self.on_event = on_event or (lambda e: None)
        self.cancel = cancel or (lambda: False)
        self.dry_run = dry_run
        self.results: list[LayoutResult] = []
        self._used_names: dict[str, int] = {}

    # -------------------------------------------------- 事件
    def _emit(self, kind: str, **kw):
        kw["type"] = kind
        self.on_event(kw)

    def _log(self, msg: str):
        self._emit("log", message=msg)

    # -------------------------------------------------- 命名
    def _unique_name(self, base: str, folder: str, stem: str = "") -> str:
        """
        避免不同 DWG 里的同名布局互相覆盖：
        第一次直接用图纸名；重名时先补上源文件名（一眼看出是哪张图），
        仍然重名才加序号。
        """
        def taken(name: str) -> bool:
            key = os.path.normcase(os.path.join(folder, name)).lower()
            return key in self._used_names

        def mark(name: str) -> str:
            key = os.path.normcase(os.path.join(folder, name)).lower()
            self._used_names[key] = 1
            return name

        if not taken(base):
            return mark(base)
        if stem:
            alt = safe_filename(f"{base} - {stem}")
            if not taken(alt):
                return mark(alt)
        n = 2
        while taken(f"{base} ({n})"):
            n += 1
        return mark(f"{base} ({n})")

    def _target_dir(self, dwg_path: str) -> str:
        if not self.keep_subdir:
            return self.out
        rel = os.path.relpath(os.path.dirname(dwg_path), self.src)
        return self.out if rel in (".", "") else os.path.join(self.out, rel)

    # -------------------------------------------------- 主流程
    def run(self) -> dict:
        if not os.path.isdir(self.src):
            raise RuntimeError(f"源文件夹不存在：{self.src}")
        os.makedirs(self.out, exist_ok=True)

        if self.single_file:
            dwg_files = [self.single_file]
        else:
            dwg_files = find_dwg_files(self.src, self.recursive, exclude_dirs=[self.out])
        if not dwg_files:
            self._log(f"在 {self.src} 里没有找到任何 .dwg 文件")
            self._emit("done", total=0, ok=0, fail=0, skip=0, results=[])
            return {"total": 0, "ok": 0, "fail": 0, "skip": 0, "results": []}

        self._log(f"找到 {len(dwg_files)} 个 DWG，输出目录：{self.out}")
        self._emit("scan_done", dwg_count=len(dwg_files))

        plotter = AcadPlotter(pc3=self.pc3, include_model=self.include_model,
                              frames_mode=self.frames_mode, log=self._log)
        plotter.connect()

        counters = {"ok": 0, "fail": 0, "skip": 0}
        preview_count = 0
        try:
            for file_no, dwg in enumerate(dwg_files, 1):
                if self.cancel():
                    self._log("用户已停止转换")
                    break

                self._emit("progress", current=file_no, total=len(dwg_files),
                           dwg=os.path.basename(dwg))
                self._log(f"[{file_no}/{len(dwg_files)}] {os.path.basename(dwg)}")

                folder = self._target_dir(dwg)
                stem = os.path.splitext(os.path.basename(dwg))[0]

                def namer(targets, folder=folder, stem=stem):
                    plan = []
                    for n, target in enumerate(targets, 1):
                        base = build_filename(self.name_tpl, self.prefix,
                                              target["name"], stem, n)
                        base = self._unique_name(base, folder, stem)
                        plan.append((target, os.path.join(folder, base + ".pdf")))
                    return plan

                if not self.dry_run:
                    try:
                        results = plotter.export_drawing(
                            dwg, namer, skip_existing=self.skip_existing)
                    except Exception as exc:             # noqa: BLE001
                        msg = _humanize_com_error(exc)
                        self._log(f"  × 处理失败：{msg}，正在重新连接 AutoCAD 后重试…")
                        try:
                            plotter.connect()
                            results = plotter.export_drawing(
                                dwg, namer, skip_existing=self.skip_existing)
                        except Exception as exc2:        # noqa: BLE001
                            results = [LayoutResult(dwg, "", "", "fail",
                                                    _humanize_com_error(exc2))]
                else:
                    results = plotter.walk_layouts(dwg, namer)

                if not results:
                    self._log("  没有可导出的图纸（该文件只有模型空间，或没有布局）")

                for r in results:
                    if r.status != "preview":
                        counters[r.status] = counters.get(r.status, 0) + 1
                        self.results.append(r)
                    fname = os.path.basename(r.pdf) or "-"
                    if r.status == "ok":
                        self._log(f"  √ {fname}  ({r.layout})")
                    elif r.status == "preview":
                        preview_count += 1
                        self.results.append(r)
                        self._log(f"  预览 → {fname}   （图纸：{r.layout}）")
                    elif r.status == "skip":
                        self._log(f"  - 跳过 {fname}")
                    else:
                        self._log(f"  × {r.layout or os.path.basename(r.dwg)}：{r.message}")
                    self._emit("result", result=r.to_dict())
        finally:
            plotter.close(keep_acad=self.keep_acad)

        summary = {"total": len(self.results), **counters, "preview": preview_count}
        self._log("")
        if self.dry_run:
            self._log(f"预览结束：实际会生成 {preview_count} 个 PDF（本次未出图）")
        else:
            self._log(f"完成：成功 {counters['ok']} 个，跳过 {counters['skip']} 个，"
                      f"失败 {counters['fail']} 个")
        self._emit("done", results=[r.to_dict() for r in self.results], **summary)
        return {**summary, "results": [r.to_dict() for r in self.results]}


# ---------------------------------------------------------------- 命令行

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="dwg2pdf",
        description="把 DWG 里的每张图纸（布局）批量导出为单个 PDF，文件名为图纸名",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("-s", "--src", required=True, help="源文件夹（存放 DWG）")
    parser.add_argument("-o", "--out", required=True, help="输出 PDF 的文件夹")
    parser.add_argument("-p", "--prefix", default="", help="文件名前缀，如 A-（可选）")
    parser.add_argument("-t", "--tpl", default="{prefix}{layout}",
                        help="命名模板，默认 {prefix}{layout}\n"
                             "可用占位符：{prefix} {layout} {dwg} {n}")
    parser.add_argument("--include-model", action="store_true",
                        help="同时导出模型空间（默认只导出图纸布局）")
    parser.add_argument("--frames", dest="frames_mode", action="store_const", const="frames",
                        help="强制按模型空间图框拆分：每张图框（A0-A4）输出一个 PDF")
    parser.add_argument("--layouts", dest="frames_mode", action="store_const", const="layouts",
                        help="只按图纸布局输出，不做图框拆分")
    parser.set_defaults(frames_mode="auto")
    parser.add_argument("--no-recursive", action="store_true", help="只处理第一层，不含子文件夹")
    parser.add_argument("--keep-subdir", action="store_true", help="在输出目录中保持源文件夹结构")
    parser.add_argument("--skip-existing", action="store_true", help="跳过已存在的 PDF（断点续转）")
    parser.add_argument("--pc3", default=None, help=f"强制使用指定打印设备，默认 {DWG_TO_PDF}")
    parser.add_argument("--keep-acad", action="store_true",
                        help="转换完成后保留 AutoCAD 进程（下次启动更快）")
    parser.add_argument("--dry-run", action="store_true", help="只列出会生成哪些 PDF，不出图")
    parser.add_argument("--check", action="store_true", help="只检查 AutoCAD 与 DWG 数量")
    parser.add_argument("--json", action="store_true", help="把结果以 JSON 打印到标准输出")
    parser.add_argument("--log", default=None, help="同时把日志写入指定文件")
    args = parser.parse_args(argv)

    log_fp = open(args.log, "w", encoding="utf-8") if args.log else None

    def on_event(ev):
        if ev.get("type") == "log":
            line = ev["message"]
            print(line, flush=True)
            if log_fp:
                log_fp.write(line + "\n")
                log_fp.flush()

    if args.check:
        plotter = AcadPlotter(pc3=args.pc3, include_model=args.include_model,
                              log=print)
        try:
            plotter.connect()
        except Exception as exc:                        # noqa: BLE001
            print(f"× {exc}")
            return 3
        files = find_dwg_files(args.src, not args.no_recursive) if os.path.isdir(args.src) else []
        print(f"AutoCAD 版本：{plotter.version}")
        print(f"DWG 文件数：{len(files)}")
        for f in files[:20]:
            print("  " + os.path.basename(f))
        if len(files) > 20:
            print(f"  …共 {len(files)} 个")
        plotter.close()
        return 0

    conv = Converter(
        src=args.src, out=args.out, prefix=args.prefix, name_tpl=args.tpl,
        include_model=args.include_model, recursive=not args.no_recursive,
        keep_subdir=args.keep_subdir, skip_existing=args.skip_existing,
        pc3=args.pc3, on_event=on_event, dry_run=args.dry_run,
        keep_acad=args.keep_acad, frames_mode=args.frames_mode,
    )
    try:
        summary = conv.run()
    except Exception as exc:                            # noqa: BLE001
        print(f"× 运行失败：{exc}")
        traceback.print_exc()
        return 3
    finally:
        if log_fp:
            log_fp.close()

    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary.get("fail", 0) == 0 else 2


if __name__ == "__main__":
    sys.exit(main())

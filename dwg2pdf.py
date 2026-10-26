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
import math
import os
import re
import shutil
import sys
import tempfile
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


# ---------------------------------------------------------------- 页序（LAYOUT / BOM）

PAGE_RANK_LAYOUT = 0          # 布置图：默认排第 1 页
PAGE_RANK_BOM = 1             # 物料清单（BOM）：默认排第 2 页
PAGE_RANK_OTHER = 2           # 其它图纸：保持原顺序


def _natural_key(name: str):
    """
    按名字里的数字做「自然排序」：「布局2」排在「布局10」前面，
    而不是按字符串比大小（那样 布局10 < 布局2）。
    只作为取不到 TabOrder 时的兜底。
    """
    parts = re.split(r"(\d+)", str(name or ""))
    return [(0, int(p), "") if p.isdigit() else (1, 0, p) for p in parts if p != ""]


def _page_rank(*sources: str) -> int:
    """
    判断这张图纸该排第几页：LAYOUT 图 → 0，BOM 图 → 1，其它 → 2。
    依次看每个来源（图纸名、页面文字），谁先能确定就用谁：
    - 名字/文字里同时出现 LAYOUT 和 BOM（例如布置图的标题栏里印了 “BOM” 字段）时
      分不清，判为「其它」，宁可不动也不能把图序搞乱。
    """
    for src in sources:
        raw = src or ""
        key = re.sub(r"[^A-Z0-9]", "", raw.upper())
        if not key and not raw:
            continue
        is_layout = "LAYOU" in key or "布置图" in raw    # LAYOUT / LAYOU / 平面布置图
        is_bom = ("BOM" in key or "BILLOFMATERIAL" in key
                  or "物料清单" in raw or "设备清单" in raw)
        if is_layout and not is_bom:
            return PAGE_RANK_LAYOUT
        if is_bom and not is_layout:
            return PAGE_RANK_BOM
        if is_layout and is_bom:
            return PAGE_RANK_OTHER           # 两者都有，说不清，别动
    return PAGE_RANK_OTHER


def _pdf_head_text(path: str, limit: int = 6000) -> str:
    """
    读一页 PDF 的文字（只取前 limit 个字符），用来判断这页画的是哪类图。
    读不到（没装 pypdf、扫描件、文字被转成曲线）就返回空串，调用方按「其它」处理。
    """
    try:
        from pypdf import PdfReader
    except ImportError:
        return ""
    # AutoCAD 出的 PDF 常有重复的 /PageMode 等键，pypdf 会刷一堆告警，
    # 对结果没有影响，压掉免得刷满日志
    try:
        import logging
        logging.getLogger("pypdf").setLevel(logging.ERROR)
    except Exception:                         # noqa: BLE001
        pass
    try:
        reader = PdfReader(path)
        if not reader.pages:
            return ""
        return (reader.pages[0].extract_text() or "")[:limit]
    except Exception:                         # noqa: BLE001
        return ""


def _merge_pdfs(paths: list[str], out_path: str) -> int:
    """
    把多个单页 PDF 按给定顺序合并成一个多页 PDF（每页原样保留尺寸）。
    优先用 pypdf（纯 Python，打包后体积小）；没装则退回 PyMuPDF。
    返回页数。
    """
    paths = [p for p in paths if p and os.path.isfile(p)]
    if not paths:
        raise RuntimeError("没有可合并的 PDF")
    tmp = out_path + TEMP_EXT

    # AutoCAD 出的 PDF 里常见重复的 /PageMode 之类的键，pypdf 会刷一堆告警，
    # 这些对结果没有影响，压掉，免得把日志刷得看不清
    try:
        import logging
        logging.getLogger("pypdf").setLevel(logging.ERROR)
    except Exception:
        pass

    try:
        from pypdf import PdfWriter
    except ImportError:
        PdfWriter = None                      # type: ignore[assignment]

    if PdfWriter is not None:
        writer = PdfWriter()
        try:
            for p in paths:
                writer.append(p)
            pages = len(writer.pages)
            with open(tmp, "wb") as f:
                writer.write(f)
        finally:
            try:
                writer.close()
            except Exception:
                pass
    else:
        # 兜底用 PyMuPDF。这里用 importlib 动态导入：
        # 写成静态 import 会被 PyInstaller 当成依赖，白白把 exe 撑大 90 MB。
        import importlib
        _pdf = None
        for mod in ("pymupdf", "fitz"):
            try:
                _pdf = importlib.import_module(mod)
                break
            except ImportError:
                continue
        if _pdf is None:
            raise RuntimeError("缺少 PDF 合并库（pypdf 或 PyMuPDF），请先 "
                               "pip install pypdf")
        doc = _pdf.open()
        try:
            for p in paths:
                with _pdf.open(p) as src:
                    doc.insert_pdf(src)
            pages = doc.page_count
            doc.save(tmp)
        finally:
            doc.close()

    if not os.path.isfile(tmp) or os.path.getsize(tmp) == 0:
        raise RuntimeError("合并后没有生成文件")
    os.replace(tmp, out_path)
    return pages


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
        """
        取出所有要出图的布局名，**按布局空间的选项卡序号（TabOrder）排序**：
        「布局1」第 1 张、「布局2」第 2 张……

        注意：AutoCAD 的 Layouts 集合是按**名称字母序**枚举的（布局1、布局10、布局11…
        布局2、布局3…），并不是选项卡顺序；只有 Layout.TabOrder 才是图纸标签上的真实
        位置。不排序就会让 布局10 排到 布局2 前面。
        """
        picked = []                                  # (tab_order, 枚举序号, 名称)
        skipped = 0
        count = doc.Layouts.Count
        for i in range(count):
            try:
                layout = doc.Layouts.Item(i)
                name = str(layout.Name)
            except Exception:
                skipped += 1                         # 个别布局取不到，别把整份文件搞挂
                continue
            if not self.include_model and name.strip().lower() == MODEL_LAYOUT:
                continue
            try:
                tab = int(layout.TabOrder)
            except Exception:
                tab = -1                             # 取不到就先按枚举顺序
            picked.append((tab, i, name))

        if skipped:
            self.log(f"  ! 有 {skipped} 个布局读不出来，已跳过（其余照常出图）")

        # 取不到 TabOrder（个别版本会返回同一个值）时，退回按名字里的数字排序
        if len(picked) > 1 and len({t for t, _, _ in picked}) <= 1:
            picked.sort(key=lambda t: (_natural_key(t[2]), t[1]))
        else:
            picked.sort(key=lambda t: (t[0], t[1]))

        names = [n for _, _, n in picked]

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

    def _pick_media(self, layout, paper: str, paper_mm, want_landscape: bool,
                    allow_bleed: bool = False):
        """
        按图纸的实际毫米尺寸挑纸张：
        1) 有完全一致的纸张就直接用（A0-A4 都能命中）；
        2) 没有（例如 A1 加长 1051x594）就找短边一致的纸张，按缩放输出；
        3) 实在没有就按图幅名前缀找.
        allow_bleed=True 时优先用「满幅无边距」（full_bleed）纸：图纸画到纸边时，
        普通纸的可打印边距（实测 DWG To PDF 是 5/17mm）会把边上的内容裁掉。
        返回 (纸张名, 是否横版, 是否尺寸完全一致)。
        """
        try:
            names = [str(n) for n in layout.GetCanonicalMediaNames()]
        except Exception:
            return "", None, False
        pool = [n for n in names if "FULL_BLEED" not in n.upper()] or names
        if allow_bleed:
            bleed = [n for n in names if "FULL_BLEED" in n.upper()]
            if bleed:
                pool = bleed
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

    # -------------------------------------------------- 布局页面设置修复

    def _paper_size_mm(self, layout):
        """当前打印设备下这张纸的物理尺寸（毫米）"""
        try:
            w, h = layout.GetPaperSize()
            w, h = float(w), float(h)
            if w > 10 and h > 10:
                return (w, h)
        except Exception:
            pass
        return None

    def _pspace_parts(self, layout, paper):
        """
        图纸空间里的内容分成两拨：视口框、其它对象（图框线、文字、表格…）。
        比纸大好几倍的对象直接忽略——那多半是跑出纸外的残留图形，
        拿它当窗口就会把整张图缩成一个小角。
        返回 (视口框列表, 其它对象的总框)
        """
        pw, ph = paper
        limit = pw * ph * 4.0
        vps, other = [], None
        total = com_get(lambda: layout.Block.Count, 0)
        for i in range(total):
            def read(i=i):
                e = layout.Block.Item(i)
                mn, mx = e.GetBoundingBox()
                is_vp = False
                try:
                    is_vp = str(e.ObjectName) == "AcDbViewport"
                except Exception:
                    pass
                return (min(mn[0], mx[0]), min(mn[1], mx[1]),
                        max(mn[0], mx[0]), max(mn[1], mx[1]), is_vp)
            r = com_get(read, None)
            if not r:
                continue
            b = r[:4]
            if (b[2] - b[0]) * (b[3] - b[1]) > limit:
                continue
            if r[4]:
                vps.append(b)
            else:
                other = b if other is None else (min(other[0], b[0]), min(other[1], b[1]),
                                                 max(other[2], b[2]), max(other[3], b[3]))
        return vps, other

    @staticmethod
    def _union(boxes):
        box = None
        for b in boxes:
            if not b:
                continue
            box = b if box is None else (min(box[0], b[0]), min(box[1], b[1]),
                                         max(box[2], b[2]), max(box[3], b[3]))
        return box

    def _pspace_box(self, layout, paper):
        """图纸空间里「真正画在纸上」的内容范围（毫米）——退回用，含所有视口"""
        vps, other = self._pspace_parts(layout, paper)
        return self._union(vps + [other])

    def _sheet_box(self, layout, paper):
        """
        取「图框」范围。布局里常常不止一个视口：
        - 一个把整个模型都框进来的「总览」视口，面积是纸的好几倍；
        - 一个局部放大用的小视口；
        - 以及真正要打印的图框视口（长宽比≈纸张、面积≈纸张）。
        拿并集当窗口，整张图就会被总览视口压成一小块，所以要先挑出图框：

        1. 候选 = 长宽比与纸张吻合（±8%）、面积不小于纸张 5% 的视口；
        2. 只有 1 个候选 → 就是它；
        3. 多个候选 → 只留「面积在纸张 0.5~2 倍」的（一页真有多张图平铺时，
           这几张都会落在这个区间，取它们的并集，绝不裁切）；
        4. 一个都不落在区间里（图框偏小、总览偏大）→ 老实取全部候选的并集。
        """
        vps, other = self._pspace_parts(layout, paper)
        if len(vps) >= 2:
            pw, ph = paper
            paper_area = pw * ph
            want = max(pw, ph) / min(pw, ph)
            if paper_area > 0:
                cands = []
                for b in vps:
                    w, h = b[2] - b[0], b[3] - b[1]
                    if w <= 5 or h <= 5:
                        continue
                    ar = max(w, h) / min(w, h)
                    if abs(ar - want) <= 0.08 * want and w * h >= 0.05 * paper_area:
                        cands.append(b)
                if len(cands) == 1:
                    return self._union([cands[0], other])
                if len(cands) >= 2:
                    band = [b for b in cands
                            if 0.5 * paper_area <= (b[2] - b[0]) * (b[3] - b[1]) <= 2.0 * paper_area]
                    return self._union((band or cands) + [other])
        return self._union(vps + [other])

    def _restore_media(self, layout, orig_media: str, orig_paper) -> bool:
        """
        换完打印设备后把纸张还原回去，返回「是否精确还原」。
        1) 先按原纸名试（两个设备都有同名纸时最准）；
        2) 名字对不上（打印机驱动的 "A4" 在 DWG To PDF 里没有）就按物理尺寸找同规格。
        """
        if orig_media:
            try:
                layout.CanonicalMediaName = orig_media
            except Exception:
                pass
            try:
                now = str(layout.CanonicalMediaName or "")
            except Exception:
                now = ""
            if now and now.strip().upper() == orig_media.strip().upper():
                return True
        if not orig_paper:
            return False
        want_land = orig_paper[0] >= orig_paper[1]
        media, _land, exact = self._pick_media(layout, "", orig_paper, want_land)
        if media:
            try:
                layout.CanonicalMediaName = media
            except Exception:
                pass
            if not exact:
                self.log(f"    · 没找到 {round(orig_paper[0])}x{round(orig_paper[1])} mm 的纸，"
                         f"改用 {media}")
            return exact
        return False

    @staticmethod
    def _fits_in(cw: float, ch: float, paper) -> bool:
        """
        内容尺寸能不能在 1:1 下装进这张纸。
        按「长边比长边、短边比短边」比，不考虑纸张当前的横竖方向——
        页面设置里的纸常常是竖版 210x297，而图纸内容是横的，直接比宽高会误判成装不下。
        """
        pw, ph = paper
        return (max(cw, ch) <= max(pw, ph) + 0.5
                and min(cw, ch) <= min(pw, ph) + 0.5)

    def _fit_layout_to_page(self, layout) -> None:
        """
        页面设置没法照搬时（原打印设备是打印机驱动，纸名在 DWG To PDF 里不存在，
        会被换成 Letter 一类的默认纸），按图纸空间里真实的内容重新框定窗口：
        纸张方向跟着图走、窗口取内容、居中，避免「图只占一个角」或「图被纸边裁掉」。

        两种情形分开处理，这是关键：
        - 内容**装得下**这张纸 → 按 1:1 出图，比例一律不改（工程图纸比例是硬要求）；
        - 内容**比纸还大** → 用「布满图纸」把整张纸缩进去。有些图的图纸空间内容
          （图框 + 视口）本来就比页面设置里的纸大一圈、或者根本没落在纸张矩形里，
          这时如果还按 1:1 或按纸张矩形裁窗口，出来的就是空白页或被切掉一半的图。
        """
        try:
            if str(layout.Name).strip().lower() == MODEL_LAYOUT:
                return                       # 模型空间没有「纸内纸外」之分，不改
        except Exception:
            pass
        paper = self._paper_size_mm(layout)
        if not paper:
            return
        box = self._sheet_box(layout, paper)
        if not box:
            return
        cw, ch = box[2] - box[0], box[3] - box[1]
        if cw <= 5 or ch <= 5:
            return

        pw0, ph0 = paper
        fits_first = self._fits_in(cw, ch, paper)
        fill_first = (cw * ch / (pw0 * ph0)) if pw0 * ph0 else 0.0
        # 纸张方向跟着图走：图是横的就用横版纸，别再靠「旋转 90°」硬塞，
        # 那样图会被纸边裁掉一半（实测就是这么出的问题）。
        # 用满幅无边距纸（full_bleed）的两种情况：
        # - 内容按 1:1 就铺满整张纸（≥75%）：普通纸 5/17 mm 的可打印边距会把纸边那圈内容裁掉；
        # - 需要按图框缩放（内容比纸大、或图框只占纸的一小块）：既然要缩，就缩到纸边，
        #   否则普通纸的边距会让整页四周留出一大圈白边（实测只剩 81%，看着就是「没铺满」）。
        edge_to_edge = (fits_first and fill_first >= 0.75) or (not fits_first) or fill_first < 0.80
        media, mland, _exact = self._pick_media(layout, "", paper, cw >= ch,
                                                allow_bleed=edge_to_edge)
        if edge_to_edge and not _exact:
            # 满幅纸里没有这个尺寸（非标纸幅、英制纸等）：退回普通纸，别硬凑一张别的尺寸
            m2, l2, e2 = self._pick_media(layout, "", paper, cw >= ch, allow_bleed=False)
            if e2:
                media, mland, _exact = m2, l2, e2
        if media:
            try:
                layout.CanonicalMediaName = media
            except Exception:
                pass
            paper = self._paper_size_mm(layout) or paper
        if mland is not None:
            try:
                layout.PlotRotation = 0
            except Exception:
                pass
        pw, ph = paper

        fits = self._fits_in(cw, ch, paper)
        # 图框占纸张的比例。正常图纸的图框≈纸张（实测 92%~100%）；
        # 有的图图框比纸小一大截（例如 122x88 mm 的图框配 A4 纸，只占 17%），
        # 那是页面设置和图纸对不上，1:1 出来就是「图缩在页面中间、四周一圈白边」。
        # 这种就不再坚持 1:1，直接把图框铺满纸张。
        fill = (cw * ch / (pw * ph)) if pw * ph else 0.0
        keep_scale = fits and fill >= 0.80

        # 窗口四周统一留 1mm：出图时图框线不会正好压在可打印区边界上（否则线会被裁掉一半）
        pad = 1.0
        if not keep_scale and not fits:
            # 内容比纸大时这点余量没意义，反而让缩放的基数变模糊，干脆不留
            pad = 0.0

        try:
            layout.SetWindowToPlot(self._double_array([box[0] - pad, box[1] - pad]),
                                   self._double_array([box[2] + pad, box[3] + pad]))
            layout.PlotType = 4                            # 4 = 窗口
            layout.CenterPlot = True
            if keep_scale:
                layout.UseStandardScale = True
                layout.StandardScale = 16                  # 1:1
                if not getattr(self, "_fit_noted", False):
                    self._fit_noted = True
                    self.log("    · 原打印设备不是 PDF 设备，页面设置已按图纸内容重新框定并居中"
                             "（比例不变）")
            else:
                layout.UseStandardScale = True
                layout.StandardScale = 0                   # 0 = 布满图纸
                if fits:
                    self.log(f"    · 图框只有 {round(cw)}x{round(ch)} mm，仅占 "
                             f"{round(pw)}x{round(ph)} mm 纸张的 {fill * 100:.0f}%（页面设置"
                             f"与图纸不匹配），已按图框铺满本页")
                else:
                    self.log(f"    · 图纸内容 {round(cw)}x{round(ch)} mm 比纸张 "
                             f"{round(pw)}x{round(ph)} mm 大，已按「布满图纸」缩放到本页"
                             f"（整张图完整、不裁切）")
        except Exception as exc:                           # noqa: BLE001
            self.log(f"    · 按图纸内容重设窗口失败，仍按原页面设置出图：{exc}")

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
        # 纸张尺寸（毫米）必须趁还没换设备先取：换完设备就是这个设备的默认纸了
        orig_paper = self._paper_size_mm(layout)

        def switch_to_pdf_device():
            try:
                layout.ConfigName = self.pc3
            except Exception:
                pass
            # 原来的纸张（如 ISO A0）在 DWG To PDF 里一般也有同名规格，试着还原；
            # 名字对不上（例如打印机驱动里的 "A4"）就按物理尺寸去找同一规格。
            return self._restore_media(layout, orig_media, orig_paper)

        # 候选方案：先用图纸原设备（显式传名），不行再用 DWG To PDF.pc3
        plans: list[str] = []
        native_ok = (not self.force_pc3) and orig_cfg.lower().endswith(".pc3")
        if native_ok:
            plans.append(orig_cfg)
        if orig_cfg.strip().lower() != self.pc3.strip().lower():
            plans.append(self.pc3)

        temp_path = pdf_path + TEMP_EXT          # .part：无关联扩展名，不会被自动打开
        last_err = ""

        def plot_once(cfg: str) -> bool:
            _remove_quietly(temp_path)
            plot.PlotToFile(temp_path, cfg)     # ← 传设备名，且写临时扩展名
            return self._wait_file(temp_path, timeout=20)

        for cfg in plans:
            fitted = False
            if cfg.strip().lower() != orig_cfg.strip().lower():
                # 换了打印设备：先把纸张按原尺寸还原；换设备后原来的页面设置
                # （尤其是「旋转」和可打印边距）含义会变，所以再按图纸内容
                # 重新框一次窗口并居中，避免「图缩在一个角、其余大片空白」。
                switch_to_pdf_device()
                self._fit_layout_to_page(layout)
                fitted = True
            for _try in range(2):
                try:
                    ok = plot_once(cfg)
                except Exception as exc:               # noqa: BLE001
                    last_err = str(exc)
                    break
                if not ok:
                    last_err = "AutoCAD 未生成文件"
                    break
                # 出成了一张空白页：说明页面设置里的纸和图纸实际内容对不上
                # （内容跑到纸张矩形外面去了）。这时按图纸内容重新框定再出一次。
                if not fitted and self._pdf_is_blank(temp_path):
                    self.log("    · 这一页是空白（页面设置与图纸内容不匹配），"
                             "按图纸内容重新框定后重出")
                    switch_to_pdf_device()
                    self._fit_layout_to_page(layout)
                    fitted = True
                    continue
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
                        break

        _remove_quietly(temp_path)
        raise RuntimeError(
            f"出图未生成文件（已尝试 {len(plans)} 种打印设备："
            + "、".join(plans) + (f"；最后错误：{last_err}" if last_err else "") + "）"
        )

    @staticmethod
    def _pdf_is_blank(path: str) -> bool:
        """
        判断刚出的 PDF 是不是「整页空白」：只看内容流是否为空。
        读不了（没装 pypdf、文件还没写完）就当作不空白，绝不影响正常出图。
        """
        try:
            from pypdf import PdfReader
        except Exception:
            return False
        try:
            reader = PdfReader(path)
            if not reader.pages:
                return False
            page = reader.pages[0]
            contents = page.get_contents()
            if contents is None:
                return True
            data = contents.get_data() or b""
            return len(data.strip()) < 40
        except Exception:
            return False

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
        self._fit_noted = False

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
                       skip_existing: bool = False,
                       on_sheet: Optional[Callable[[int, int, str], None]] = None) -> list[LayoutResult]:
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

            for idx, (target, pdf_path) in enumerate(plan, 1):
                name = target["name"]
                is_sheet = target.get("x0") is not None
                if on_sheet is not None:
                    try:
                        on_sheet(idx, len(plan), name)
                    except Exception:
                        pass
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
                 frames_mode: str = "auto", one_pdf: bool = True,
                 order_pages: bool = True):
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
        self.one_pdf = bool(one_pdf)
        # 合并模式下默认把 LAYOUT 图排第 1 页、BOM 图排第 2 页
        self.order_pages = bool(order_pages)
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

    def _merged_path(self, folder: str, stem: str) -> str:
        """「一个 DWG 一个 PDF」时的成品路径：用 DWG 文件名（可加前缀）"""
        base = safe_filename(f"{self.prefix}{stem}")
        return os.path.join(folder, self._unique_name(base, folder, stem) + ".pdf")

    def _order_for_merge(self, items: list[LayoutResult]):
        """
        合并前按「LAYOUT 图排第 1 页、BOM 图排第 2 页」重排。
        判断依据依次是：图纸名（布局名/图框标题）→ 这一页 PDF 里的文字。
        同类图纸出现不止一张、或根本判断不出来时保持原顺序，绝不乱动。
        返回 (新列表, 说明文字)。
        """
        if not self.order_pages or len(items) < 2:
            return items, ""

        ranks = []
        for r in items:
            rank = _page_rank(r.layout)
            if rank == PAGE_RANK_OTHER and r.pdf:
                rank = _page_rank(_pdf_head_text(r.pdf))
            ranks.append(rank)

        lay = [i for i, k in enumerate(ranks) if k == PAGE_RANK_LAYOUT]
        bom = [i for i, k in enumerate(ranks) if k == PAGE_RANK_BOM]
        if len(lay) > 1 or len(bom) > 1:
            self._log(f"  · 识别到 {len(lay)} 张 LAYOUT 图、{len(bom)} 张 BOM 图，"
                      f"无法唯一确定页序，保持图纸原顺序")
            return items, ""
        if not lay and not bom:
            return items, ""

        if lay and bom:
            a, b = lay[0], bom[0]
            rest = [r for i, r in enumerate(items) if i not in (a, b)]
            new = [items[a], items[b]] + rest
            note = "LAYOUT 图排第 1 页、BOM 图排第 2 页"
        elif lay:
            a = lay[0]
            new = [items[a]] + [r for i, r in enumerate(items) if i != a]
            note = "LAYOUT 图排第 1 页"
        else:
            b = bom[0]
            rest = [r for i, r in enumerate(items) if i != b]
            if not rest:
                return items, ""
            new = rest[:1] + [items[b]] + rest[1:]
            note = "BOM 图排第 2 页"

        if new == items:
            return items, ""
        return new, note

    def _export_merged(self, plotter, dwg: str, folder: str,
                       stem: str) -> list[LayoutResult]:
        """
        一个 DWG → 一个多页 PDF：先把每张图纸出到临时目录，再按图纸顺序合并。
        合并失败就退回「每张图纸一个文件」，结果不会丢。
        """
        final = self._merged_path(folder, stem)
        if self.skip_existing and os.path.isfile(final) and os.path.getsize(final) > 0:
            return [LayoutResult(dwg, "", final, "skip", "PDF 已存在，已跳过")]

        tmpdir = tempfile.mkdtemp(prefix="dwg2pdf_")
        try:
            def namer(targets):
                return [(t, os.path.join(tmpdir, "%04d.pdf" % n))
                        for n, t in enumerate(targets, 1)]

            def on_sheet(idx, total, name):
                self._log(f"     第 {idx}/{total} 张：{name}")

            try:
                raw = plotter.export_drawing(dwg, namer, on_sheet=on_sheet)
            except Exception as exc:                   # noqa: BLE001
                self._log(f"  × 处理失败：{_humanize_com_error(exc)}，"
                          f"正在重新连接 AutoCAD 后重试…")
                plotter.connect()
                raw = plotter.export_drawing(dwg, namer, on_sheet=on_sheet)

            ok = [r for r in raw if r.status == "ok" and r.pdf and os.path.isfile(r.pdf)]
            bad = [r for r in raw if r not in ok]
            if not ok:
                return raw or [LayoutResult(dwg, "", "", "fail",
                                            "没有可导出的图纸（该文件没有布局）")]

            os.makedirs(folder, exist_ok=True)
            # 默认把 LAYOUT 图放第 1 页、BOM 图放第 2 页（--no-page-order 可关掉）
            ordered, note = self._order_for_merge(ok)
            if note:
                self._log(f"  · 页序已调整：{note}")
            try:
                pages = _merge_pdfs([r.pdf for r in ordered], final)
            except Exception as exc:                   # noqa: BLE001
                self._log(f"  ！合并成一个 PDF 失败：{exc}；已改为每张图纸一个文件")
                fallback: list[LayoutResult] = []
                for r in ok:
                    base = safe_filename(f"{self.prefix}{stem} - {r.layout}")
                    dst = os.path.join(folder, self._unique_name(base, folder, stem) + ".pdf")
                    try:
                        shutil.move(r.pdf, dst)
                    except OSError as exc2:
                        fallback.append(LayoutResult(dwg, r.layout, r.pdf, "fail",
                                                     f"文件无法移动：{exc2}"))
                        continue
                    fallback.append(LayoutResult(dwg, r.layout, dst, "ok",
                                                 "合并失败，已按单张输出"))
                return fallback + bad

            size_mb = os.path.getsize(final) / 1024 / 1024
            msg = f"{pages} 页 · {size_mb:.1f} MB"
            if bad:
                msg += f"，另有 {len(bad)} 张未出图"
            return [LayoutResult(dwg, f"共 {pages} 张图纸", final, "ok", msg)] + bad
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def _preview_merged(self, plotter, dwg: str, folder: str, stem: str) -> list[LayoutResult]:
        """试运行：只算出会合并成几页，不出图"""
        final = self._merged_path(folder, stem)
        got = plotter.walk_layouts(dwg, lambda targets: [(t, "") for t in targets])
        if not got:
            return []
        return [LayoutResult(dwg, f"共 {len(got)} 张图纸", final, "preview",
                             f"合并为 1 个 PDF，共 {len(got)} 页")]

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

                if self.dry_run:
                    results = (self._preview_merged(plotter, dwg, folder, stem)
                               if self.one_pdf
                               else plotter.walk_layouts(dwg, namer))
                elif self.one_pdf:
                    try:
                        results = self._export_merged(plotter, dwg, folder, stem)
                    except Exception as exc:             # noqa: BLE001
                        results = [LayoutResult(dwg, "", "", "fail",
                                                _humanize_com_error(exc))]
                else:
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

                if not results:
                    self._log("  没有可导出的图纸（该文件只有模型空间，或没有布局）")

                for r in results:
                    if r.status != "preview":
                        counters[r.status] = counters.get(r.status, 0) + 1
                        self.results.append(r)
                    fname = os.path.basename(r.pdf) or "-"
                    if r.status == "ok":
                        if self.one_pdf and r.message:
                            self._log(f"  √ {fname}  ({r.message})")
                        else:
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
    parser.add_argument("--one-pdf", dest="one_pdf", action="store_true",
                        help="一个 DWG 只输出一个多页 PDF（每页一张图纸），默认就是这种方式")
    parser.add_argument("--split", dest="one_pdf", action="store_false",
                        help="每张图纸各输出一个 PDF（旧方式）")
    parser.set_defaults(one_pdf=True)
    parser.add_argument("--no-page-order", dest="order_pages", action="store_false",
                        help="合并时不要自动调整页序（默认会把 LAYOUT 图排第 1 页、BOM 图排第 2 页）")
    parser.set_defaults(order_pages=True)
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
        one_pdf=args.one_pdf, order_pages=args.order_pages,
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

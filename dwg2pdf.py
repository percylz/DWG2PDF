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

# 双语文案：tr(中文, English)；语言由界面设置或启动参数决定，见 i18n.py
from i18n import is_en, t as tr

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
# PDF 内容流里的算子：出图判空用（见 AcadPlotter._pdf_page_is_empty）
TEXT_SHOW = re.compile(rb"(?:\bTj\b|\bTJ\b)")          # 显示文字
PATH_BUILD = re.compile(rb"(?:\d\s+[mlc]\b|\bre\b)")   # 画线/曲线/矩形

SHEET_CODE = re.compile(r"^[A-Za-z]{1,4}[-_ ]?\d{1,4}$")   # 图号，如 FD-09 / A-101
# 长图号：KUL01-BLKA1-SC-E-70001-Z 这种「好几段、用横线连」的编号（东南亚项目常用）
LONG_CODE = re.compile(r"^[A-Za-z0-9]{1,10}(?:[-_][A-Za-z0-9]{1,10}){3,}$")
# 说明性文字不是图名
NOTE_PREFIX = re.compile(r"^(note|notes|remark|remarks|说明|备注|注)[\s:：]", re.I)
MTEXT_CODES = re.compile(r"\\[A-Za-z][^;\\]*;?|[{}]")       # MText 里的格式控制符


# ---------------------------------------------------------------- 数据结构

class EmptySheet(Exception):
    """这个布局出图后整页是空的（DWG 里的空布局），不该写进 PDF"""


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
        raise RuntimeError(tr("没有可合并的 PDF", "No PDF pages to merge"))
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
            raise RuntimeError(tr("缺少 PDF 合并库（pypdf 或 PyMuPDF），请先 "
                                 "pip install pypdf",
                                 "PDF merge library missing (pypdf or PyMuPDF). "
                                 "Run: pip install pypdf"))
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
        raise RuntimeError(tr("合并后没有生成文件", "The merged PDF was not created"))
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
        s = tr("未命名图纸", "Untitled")
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
                 paper_size: str = "", log: Optional[Callable[[str], None]] = None):
        self.pc3 = pc3 or DWG_TO_PDF
        self.include_model = include_model
        self.force_pc3 = force_pc3
        self.frames_mode = frames_mode      # auto / frames / layouts
        # 出图纸张：''/auto = 跟着图纸走；A0~A4 = 不小于图纸的那张（只放大、不缩小）
        self.paper_size = (paper_size or "").strip().upper()
        self._paper_noted: set[str] = set()
        self._insert_cache = None      # (选择集, 块名→序号, 总数)
        self._text_pos_cache = None    # (选择集, [(x, y, 序号)])
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
            raise RuntimeError(tr(
                "缺少 pywin32 组件，无法调用 AutoCAD。\n"
                "请在命令行执行：pip install pywin32",
                "pywin32 is missing, so AutoCAD cannot be automated.\n"
                "Run: pip install pywin32"
            ))

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
                raise RuntimeError(tr(
                    "无法启动 AutoCAD。请确认已正确安装并已授权（未过试用期/已登录）。\n"
                    f"最后一条错误：{last_err}",
                    "Cannot start AutoCAD. Make sure it is installed, licensed and able "
                    "to start (trial not expired / signed in).\n"
                    f"Last error: {last_err}"
                ))

        try:
            self.version = str(self.app.Version)
        except Exception:
            self.version = tr("未知", "unknown")

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

        self.log(tr(f"已连接 AutoCAD {self.version}"
                    f"（{'复用已打开的实例' if not self._owned else '新启动实例'}）",
                    f"Connected to AutoCAD {self.version} "
                    f"({'reusing the running instance' if not self._owned else 'new instance started'})"))

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
                    self.log(tr(f"  AutoCAD 已就绪（等了 {time.time() - t0:.0f} 秒）",
                                f"  AutoCAD is ready (waited {time.time() - t0:.0f} s)"))
                return
            except Exception:
                pass
            if time.time() - t0 > timeout:
                self.log(tr("  ⚠ AutoCAD 长时间没有响应，仍继续尝试出图",
                            "  ⚠ AutoCAD has been unresponsive for a while; still trying to plot"))
                return
            if not waiting:
                waiting = True
                self.log(tr("  AutoCAD 还在启动，等它就绪…",
                            "  AutoCAD is still starting up, waiting for it…"))
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
            self.log(tr(f"  ! 有 {skipped} 个布局读不出来，已跳过（其余照常出图）",
                        f"  ! {skipped} layout(s) could not be read and were skipped "
                        f"(the others are plotted as usual)"))

        # 取不到 TabOrder（个别版本会返回同一个值）时，退回按名字里的数字排序
        if len(picked) > 1 and len({t for t, _, _ in picked}) <= 1:
            picked.sort(key=lambda t: (_natural_key(t[2]), t[1]))
        else:
            picked.sort(key=lambda t: (t[0], t[1]))

        names = [n for _, _, n in picked]

        # 只有模型空间、没有任何图纸布局的图（简单图常见）：按一张图纸输出，
        # 否则这种文件会「什么都没生成」，容易让人以为程序没工作。
        if not names and not self.include_model:
            self.log(tr("  （该文件没有图纸布局，按模型空间输出一张 Model.pdf）",
                        "  (no paper-space layout in this file — model space is plotted as one Model.pdf)"))
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

    def _insert_index(self, doc):
        """
        一次遍历，把所有块参照按「块名 → 序号列表」分好组，供识图复用。

        单个 `ss.Item(i).Name` 是一次跨进程 COM 调用（实测 2.8 ms）：一份 20 MB 的
        电气图有 2395 个块参照，逐个读名字就是 6~7 秒；图框识别和整图纸块识别都要
        这份名单，所以只读一次、两个函数共用。

        本来想用过滤选择集按块名一次选出来，但 AutoCAD COM 的 Select 过滤器在
        Mode=5 下会被静默忽略（传了 filter 仍返回全部 2395 个），只能这样折中。
        """
        if self._insert_cache is not None:
            return self._insert_cache
        try:
            ss = self._ss(doc, "d2p_ins_all")
            ss.Select(5, 0, 0, self._short_array([0]), self._variant_array(["INSERT"]))
            total = com_get(lambda: ss.Count, 0)
        except Exception:
            self._insert_cache = (None, {}, 0)
            return self._insert_cache

        by_name: dict[str, list] = {}
        for i in range(total):
            if total > 800 and i and i % 800 == 0:
                self.log(tr(f"    · 已读取 {i}/{total} 个块参照的名字…",
                            f"    · read names of {i}/{total} block references…"))
            nm = com_get(lambda i=i: str(ss.Item(i).Name).strip(), "")
            if nm:
                by_name.setdefault(nm, []).append(i)
        self._insert_cache = (ss, by_name, total)
        return self._insert_cache

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
        ss, by_name, total = self._insert_index(doc)
        if ss is None or not by_name:
            return frames
        # 按「块名」遍历而不是按「块参照」遍历：2395 个块参照只有 42 个不同名字，
        # 名字不像图框的整组跳过，连 InsertionPoint 都不用读。
        for name, idxs in by_name.items():
            # 先看名字像不像图框，再决定要不要去翻块定义。
            # 建筑图里常有搬来的匿名块（A$Ca5f0b8dd 之类，块内上千个实体），
            # 逐个翻一遍会白白卡上几分钟。
            if not self._name_may_be_frame(name):
                continue
            rect = self._block_frame_rect(doc, name)
            if not rect:
                continue
            for idx in idxs:
                f = com_get(lambda idx=idx: self._frame_from_insert(ss, idx, name, rect), None)
                if f:
                    frames.append(f)
        return frames

    def _frame_from_insert(self, ss, idx: int, name: str, rect) -> Optional[dict]:
        """一个块参照 + 它的块定义矩形 → 图框字典（模型坐标，已算缩放和旋转）"""
        import math
        e = ss.Item(idx)
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
                self.log(tr(f"    · 已检查 {i}/{total} 条多段线…",
                            f"    · checked {i}/{total} polylines…"))
            f = com_get(lambda i=i: read_one(i), None)
            if f:
                frames.append(f)
        return frames

    def _sheets_from_repeated_blocks(self, doc, frame: dict) -> list[dict]:
        """
        大外框里并排画了 N 张图时，图与图之间常常没有分割线，但每张图都有**同一个图签块**
        （同名、同尺寸、按行或按列排开）。用这些重复块把外框拆成 N 张图纸。

        认的条件（都满足才算）：
        - 同名块出现 ≥2 次，尺寸一致（毫米取整后相同）；
        - 每个块的短边 ≥100 mm、长边 ≤1300 mm，长宽比 0.5~2.0（像个图幅，不像图例表格）；
        - 块都在外框里，且这些块加起来占外框面积 ≥35%（否则只是图框里的小图标）；
        - 取「出现次数最多」的那一组。
        """
        scale = frame.get("scale") or 1
        try:
            ss = self._ss(doc, "d2p_ins2")
            ss.Select(5, 0, 0, self._short_array([0]), self._variant_array(["INSERT"]))
            total = com_get(lambda: ss.Count, 0)
        except Exception:
            return []
        if not total:
            return []

        fx0, fy0, fx1, fy1 = frame["x0"], frame["y0"], frame["x1"], frame["y1"]
        groups: dict[tuple, list[tuple]] = {}

        def read_one(idx: int):
            e = ss.Item(idx)
            mn, mx = e.GetBoundingBox()
            return (str(e.Name), min(mn[0], mx[0]), min(mn[1], mx[1]),
                    max(mn[0], mx[0]), max(mn[1], mx[1]))

        for i in range(total):
            r = com_get(lambda i=i: read_one(i), None)
            if not r:
                continue
            name, x0, y0, x1, y1 = r
            w, h = x1 - x0, y1 - y0
            if w <= 0 or h <= 0:
                continue
            # 允许压线（图签块常常正好落在外框上），但不能明显跑到框外
            if x0 < fx0 - w * 0.2 or y0 < fy0 - h * 0.2 \
                    or x1 > fx1 + w * 0.2 or y1 > fy1 + h * 0.2:
                continue
            groups.setdefault((name, round(w / scale), round(h / scale)),
                              []).append((x0, y0, x1, y1))

        best = None
        frame_area = max(1.0, (fx1 - fx0) * (fy1 - fy0))
        for key, rects in groups.items():
            if len(rects) < 2:
                continue
            _name, wmm, hmm = key
            if min(wmm, hmm) < 100 or max(wmm, hmm) > 1300:
                continue
            if not (0.5 <= wmm / max(hmm, 1) <= 2.0):
                continue
            cover = sum((b[2] - b[0]) * (b[3] - b[1]) for b in rects) / frame_area
            if cover < 0.35:
                continue
            if best is None or len(rects) > len(best[1]):
                best = (key, rects)
        if not best:
            return []
        (name, wmm, hmm), rects = best
        self.log(tr(f"  （外框里识别到 {len(rects)} 个同款图签块「{name}」{wmm}x{hmm} mm，按它拆成 "
                    f"{len(rects)} 张图）",
                    f"  (found {len(rects)} identical title blocks \"{name}\" {wmm}x{hmm} mm "
                    f"inside the outer frame — splitting into {len(rects)} sheets)"))
        return [{"paper": self._paper_label(wmm, hmm), "paper_mm": (wmm, hmm),
                 "scale": scale, "x0": b[0], "x1": b[2], "y0": b[1], "y1": b[3]}
                for b in rects]

    @staticmethod
    def _is_pdf_device(cfg: str) -> bool:
        """
        图纸原本的打印设备是不是「真能出 PDF 的那个 .pc3」。

        实测（KUL01-BLKA1-DR-E-10200）：原设备是同事自定义的 `M_PDF.pc3`，
        名字里带 pdf、却是本机没有的绘图仪配置——调 PlotToFile 既不报错、也不
        生成文件，只能等满超时才换成 DWG To PDF，每页白等 20 秒。
        所以只认 AutoCAD 自带的两个 PDF 设备，其余一律直接用 DWG To PDF.pc3。
        """
        s = (cfg or "").strip().lower()
        if not s.endswith(".pc3"):
            return False
        return ("dwg to pdf" in s) or ("autocad pdf" in s)

    @staticmethod
    def _frame_area(f) -> float:
        return (f["x1"] - f["x0"]) * (f["y1"] - f["y0"])

    def _sheets_from_title_blocks(self, doc) -> list[dict]:
        """
        连一个图框都认不出来时，最后再试一招：找「同名、同尺寸、成排摆放」的
        **整张图纸模板块**——把「图框 + 标题栏 + 全部内容」做成一个块，复制 N 份并排画。

        实测 KUL01-BLKA1-SC-E-70001~3（MCB BOARD DETAILS）就是这么画的：4 个「图签1」
        块 267.5x189.2 mm 严丝合缝地并排，既没有外框、没有分割线，也没有 A0-A4 命名的
        图框块，旧逻辑只能把整个模型空间当一张图，4 张图纸全挤在一页里。

        认的条件（全部满足）：
        - 同名块 ≥2 个，尺寸一致；
        - 这些块互不重叠，加起来占它们并集的 ≥85%（每个块本身就是一张纸；
          「每张纸右下角放一个小图签」那种只有几个百分点，不会被误认）；
        - 长宽比 1.15~1.7（像纸，A 系列是 1.414）、换算成毫米后短边 120~1300 mm。
        """
        ss, by_name, total = self._insert_index(doc)
        if ss is None or total < 2:
            return []

        std_longs = [max(v) for v in PAPER_SIZES.values()]

        def pick_scale(w_units, h_units):
            """挑一个比例，让这个块看起来像一张纸：返回 (scale, wmm, hmm) 或 None"""
            best = None
            for scale in KNOWN_SCALES:
                wmm, hmm = w_units / scale, h_units / scale
                lo, hi = min(wmm, hmm), max(wmm, hmm)
                if not (120 <= lo <= 1300 and 150 <= hi <= 1500):
                    continue
                # 纸的长宽比：A 系列 1.414、美制 Letter 1.29。放宽到 1.15~1.7，
                # 太方或太长的块（图例表格、说明条）一律不当图纸
                if not (1.15 <= hi / max(lo, 1e-6) <= 1.7):
                    continue
                gap = min(abs(hi - v) for v in std_longs)
                if best is None or gap < best[0]:
                    best = (gap, scale, wmm, hmm)
            if not best:
                return None
            return best[1], best[2], best[3]

        def bbox_of(idx: int):
            e = ss.Item(idx)
            mn, mx = e.GetBoundingBox()
            return (min(mn[0], mx[0]), min(mn[1], mx[1]),
                    max(mn[0], mx[0]), max(mn[1], mx[1]))

        # 先按块名分组、每组只抽 2 个实例算包围盒：尺寸对不上、或换算成毫米后
        # 长宽比不像纸的组直接扔掉。否则 2000 多个设备符号（每个包围盒 2.9 ms）
        # 光算包围盒就要 7 秒，而我们只关心那几个「整张图纸」块。
        groups: dict[tuple, list[tuple]] = {}
        for name, idxs in by_name.items():
            if len(idxs) < 2:
                continue
            sample = []
            for idx in idxs[:2]:
                b = com_get(lambda idx=idx: bbox_of(idx), None)
                if b:
                    sample.append(b)
            if len(sample) < 2:
                continue
            w0, h0 = sample[0][2] - sample[0][0], sample[0][3] - sample[0][1]
            w1, h1 = sample[1][2] - sample[1][0], sample[1][3] - sample[1][1]
            if w0 <= 0 or h0 <= 0:
                continue
            # 同一块名的尺寸必须一致（差 ≤2%），否则不是「同一张模板复制出来的」
            if abs(w1 - w0) > w0 * 0.02 or abs(h1 - h0) > h0 * 0.02:
                continue
            if not pick_scale(w0, h0):
                continue
            for idx in idxs:
                b = com_get(lambda idx=idx: bbox_of(idx), None)
                if not b:
                    continue
                w, h = b[2] - b[0], b[3] - b[1]
                if w <= 0 or h <= 0:
                    continue
                # 同一块名的尺寸一般完全一致，取整到「块单位的 0.1」归堆
                # 连 idx 一起存：后面要按块参照算「内容签名」，判断是不是重复图框
                groups.setdefault((name, round(w, 1), round(h, 1)),
                                  []).append((b, idx))

        # 顺序说明：pick_scale 已经挑过，这里不再重复筛；后续按覆盖面积挑最佳组

        best_group = None
        for key, items in groups.items():
            if len(items) < 2:
                continue
            rects = [it[0] for it in items]
            name, w_units, h_units = key
            scaled = pick_scale(w_units, h_units)
            if not scaled:
                continue
            _scale, wmm, hmm = scaled
            # 互不重叠：任两块的重叠面积不超过其中较小那块的一半
            ok, sum_area = True, 0.0
            for a in rects:
                sum_area += (a[2] - a[0]) * (a[3] - a[1])
            ux0 = min(r[0] for r in rects)
            uy0 = min(r[1] for r in rects)
            ux1 = max(r[2] for r in rects)
            uy1 = max(r[3] for r in rects)
            union_area = max(1.0, (ux1 - ux0) * (uy1 - uy0))
            if sum_area / union_area < 0.85:
                ok = False            # 只在每张纸的角落有个小图签，不是整张纸的模板
            if ok:
                for a in rects:
                    for b in rects:
                        if a is b:
                            continue
                        ix = min(a[2], b[2]) - max(a[0], b[0])
                        iy = min(a[3], b[3]) - max(a[1], b[1])
                        if ix <= 0 or iy <= 0:
                            continue
                        small = min((a[2] - a[0]) * (a[3] - a[1]),
                                    (b[2] - b[0]) * (b[3] - b[1]))
                        if ix * iy > 0.5 * small:
                            ok = False
                            break
                    if not ok:
                        break
            if not ok:
                continue
            if best_group is None or len(rects) > len(best_group[1]):
                best_group = ((name, wmm, hmm, _scale), rects, [it[1] for it in items])

        if not best_group:
            return []
        (name, wmm, hmm, scale), rects, idxs = best_group
        self.log(tr(f"  （没找到图框块，但识别到 {len(rects)} 个同款整图纸块「{name}」"
                    f"{round(wmm)}x{round(hmm)} mm，按它拆成 {len(rects)} 张图）",
                    f"  (no frame block found, but {len(rects)} identical full-sheet blocks "
                    f"\"{name}\" {round(wmm)}x{round(hmm)} mm were found — splitting into "
                    f"{len(rects)} sheets)"))
        self._warn_duplicate_blocks(ss, idxs, name)
        return [{"paper": self._paper_label(wmm, hmm), "paper_mm": (wmm, hmm),
                 "scale": scale, "x0": b[0], "x1": b[2], "y0": b[1], "y1": b[3]}
                for b in rects]

    def _block_signature(self, ss, idx: int):
        """
        整图纸块的「内容签名」= 块定义名 + 该块参照的属性值。

        同一个块定义的所有实例，几何和块内文字必然一模一样；差异只能来自属性
        （图号、图名、日期这类）。所以「块名相同 + 属性值全相同」⇒ 两张图完全一样。

        返回 None 表示属性读不到（CAD 忙/块已损坏），这种不参与比对，免得误报。
        """
        def read():
            e = ss.Item(idx)
            try:
                nm = e.EffectiveName
            except Exception:                       # noqa: BLE001
                nm = e.Name
            try:
                attrs = e.GetAttributes()
            except Exception:                       # noqa: BLE001
                return None
            vals = []
            for a in attrs or ():
                try:
                    vals.append((str(a[0]), str(a[1])))
                except Exception:                   # noqa: BLE001
                    continue
            return (str(nm), tuple(sorted(vals)))

        return com_get(read, None)

    def _warn_duplicate_blocks(self, ss, idxs, name: str) -> None:
        """
        拆页后核对：同款整图纸块里，如果有两个连属性值都一样，就提示「重复图框」。

        实测（泰国数据中心第二批 3-IP GW- FWU1-01 等 3 个文件）：布局里并排贴了
        两个一模一样的图框，拆页后第 17、18 页完全相同——这是图纸的问题，
        不是程序出的错，但审核前应该看得见，所以这里记一条日志。只提示，不改行为。
        """
        seen: dict[tuple, list[int]] = {}
        for pos, idx in enumerate(idxs):
            sig = self._block_signature(ss, idx)
            if sig is None:
                continue
            seen.setdefault(sig, []).append(pos + 1)
        for sig, pages in seen.items():
            if len(pages) < 2:
                continue
            where = (", ".join(f"sheet {p}" for p in pages) if is_en()
                     else "、".join(f"第 {p} 张" for p in pages))
            if is_en():
                extra = f", all {len(sig[1])} attributes match too" if sig[1] else ""
            else:
                extra = f"，{len(sig[1])} 项属性也一致" if sig[1] else ""
            self.log(tr(f"  ⚠ 疑似重复图框：{where}图内容完全相同"
                        f"（块「{name}」{extra}），请核对图纸",
                        f"  ⚠ duplicate title block: {where} are identical"
                        f" (block \"{name}\"{extra}) — please check the drawing"))

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

    def _text_positions(self, doc):
        """
        只读所有文字的插入点（不读内容），并缓存。

        实测一份 2.6 MB 的电气图有 2317 个文字，每次 `.Item(i)` + `InsertionPoint`
        + `TextString` + `Height` 约 1.3 ms，全读要 11.7 秒；而标题栏只占图框
        右下角一小块、几十个字。所以先只读位置（约 6 秒，且整份图只做一次），
        按范围筛出来再读内容——文字越多省得越多。
        """
        if getattr(self, "_text_pos_cache", None) is not None:
            return self._text_pos_cache
        ss = None
        pos: list[tuple] = []
        try:
            ss = self._ss(doc, "d2p_alltxt")
            ss.Select(5, 0, 0, self._short_array([0]), self._variant_array(["TEXT,MTEXT"]))
            total = com_get(lambda: ss.Count, 0)
            for i in range(total):
                p = com_get(lambda i=i: ss.Item(i).InsertionPoint, None)
                if p is not None:
                    pos.append((float(p[0]), float(p[1]), i))
        except Exception:
            ss, pos = None, []
        self._text_pos_cache = (ss, pos)
        return self._text_pos_cache

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
        items = []
        ss, pos = self._text_positions(doc)
        if ss is not None and pos:
            for x, y, idx in pos:
                if not (rx0 <= x <= rx1 and ry0 <= y <= ry1):
                    continue
                t = com_get(lambda idx=idx: self._one_text(ss, idx), _MISS)
                if t is _MISS or not t:
                    continue
                items.append(t)
        else:
            # 位置索引没取到（AutoCAD 忙），退回全量读
            items = [t for t in self._all_texts(doc)
                     if rx0 <= t[0] <= rx1 and ry0 <= t[1] <= ry1]
        if not items:
            return "", ""

        def is_code(s: str) -> bool:
            return bool((SHEET_CODE.match(s) and not s.isdigit()) or LONG_CODE.match(s))

        codes = [t for t in items if is_code(t[3])]
        texts = [t for t in items if not is_code(t[3]) and len(t[3]) >= 2
                 and not NOTE_PREFIX.match(t[3])]
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
        source = tr("块", "blocks")
        if not frames:
            frames = self._frames_from_polylines(doc)
            source = tr("矩形", "rectangles")
        if not frames:
            # 一个图框都认不出来：试试「整张图纸模板块」（MCB BOARD DETAILS 是这种画法，
            # 4 张图并排、没有外框也没有分割线，靠重复出现的「图签1」块拆页）
            frames = self._sheets_from_title_blocks(doc)
            if frames:
                source = tr("整图纸块", "full-sheet blocks")
        if not frames:
            return []
        frames = self._dedupe_frames(frames)
        # 一个又长又扁的大框，很可能不是「一张加长图纸」，而是好几张图并排画在
        # 一个外框里（外框只是装饰，图之间没有分割线）。这时用重复出现的图签块拆开。
        if len(frames) == 1:
            f = frames[0]
            fw, fh = f["x1"] - f["x0"], f["y1"] - f["y0"]
            ar = max(fw, fh) / max(min(fw, fh), 1e-6)
            if ar >= 1.8 or ar <= 0.55:
                sub = self._sheets_from_repeated_blocks(doc, f)
                if len(sub) >= 2:
                    frames = sub
                    source = tr("重复图签块", "repeated title blocks")
        frames.sort(key=lambda f: (-round(f["y1"] / 100.0), f["x0"]))
        for i, f in enumerate(frames, 1):
            f["index"] = i
            f["title"], f["code"] = self._read_title_block(doc, f)
            if not f["title"]:
                f["title"] = f["code"] or tr(f"图纸{i:02d}", f"Sheet {i:02d}")
            f["name"] = f["title"]
        self.log(tr(f"  （按{source}识别出 {len(frames)} 张图纸："
                    + "、".join(f"{f['paper']}" for f in frames) + "）",
                    f"  (found {len(frames)} sheet(s) by {source}: "
                    + ", ".join(f"{f['paper']}" for f in frames) + ")"))
        return frames

    @staticmethod
    def _media_landscape(name: str):
        """从纸张名里读出方向：'ISO_A1_(841.00_x_594.00_MM)' → 横版"""
        up = (name or "").strip().upper()
        # 部分 PDF 打印机的自定义纸用 P_/L_ 前缀标方向，P_(1191_x_843) 实际是竖版
        if up.startswith("P_("):
            return False
        if up.startswith("L_("):
            return True
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

    def _apply_paper_choice(self, paper_mm):
        """
        用户在界面上选了纸张（A0~A4）时，用它替掉图纸自身的纸幅。

        规则是「只放大、不缩小」：纸张比图纸大就按用户选的出（图会等比放大，
        字更大、打印出来更清楚）；图纸本身比选定的纸还大，就保持图纸原来的尺寸——
        缩小出图只会让字更小，跟「提高清晰度」正好相反。
        返回 (纸幅 mm, 纸名标签)；没选或不用改时纸名标签为空。
        """
        choice = getattr(self, "paper_size", "") or ""
        if choice in ("", "AUTO"):
            return paper_mm, ""
        want = PAPER_SIZES.get(choice)
        if not want:
            return paper_mm, ""
        if not paper_mm:
            return want, choice
        sl, ss = max(paper_mm), min(paper_mm)
        wl, ws = max(want), min(want)
        if wl + 1 >= sl and ws + 1 >= ss:
            return want, choice
        return paper_mm, ""

    def _note_paper_choice(self, tag: str, sheet_mm, used_mm) -> None:
        """同一张纸只在日志里说一次，别每页刷一句"""
        if not tag or tag in self._paper_noted:
            return
        self._paper_noted.add(tag)
        self.log(tr(f"    · 按界面选定的 {tag} 出图（{round(max(used_mm))}x{round(min(used_mm))} mm，"
                    f"图纸 {round(max(sheet_mm))}x{round(min(sheet_mm))} mm 等比放大）",
                    f"    · plotting on {tag} as selected in the UI "
                    f"({round(max(used_mm))}x{round(min(used_mm))} mm; the "
                    f"{round(max(sheet_mm))}x{round(min(sheet_mm))} mm sheet is scaled up "
                    f"proportionally)"))

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
            # 尺寸对不上任何标准纸（非标图幅、按图签块算出来的小图框等）：
            # 挑一张「装得下」的最小标准纸，宁可大一圈也不要把图塞进一张小纸里
            containers = []
            for n in pool:
                mw, mh = self._media_size(n)
                ml, ms = max(mw, mh), min(mw, mh)
                if ml + 1 >= t_long and ms + 1 >= t_short:
                    # 面积相同的横竖两版（A4 297x210 / 210x297）优先挑方向一致的
                    containers.append((ml * ms, self._media_landscape(n) != want_landscape, n))
            if containers:
                containers.sort()
                pick = containers[0][2]
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
        """
        当前打印设备下这张纸的**实际出图尺寸**（毫米）。
        GetPaperSize() 只给纸张的标称尺寸，不看「旋转」：ISO_A0 永远是 841x1189。
        但 PlotRotation = 90/270 时出图是横的（PDF 页面 1189x841）。
        不换算就会把横图当成竖纸来比，1:1 出图左右直接裁掉一半。
        """
        try:
            w, h = layout.GetPaperSize()
            w, h = float(w), float(h)
            if w > 10 and h > 10:
                try:
                    rot = int(layout.PlotRotation or 0)
                except Exception:
                    rot = 0
                return (h, w) if rot in (1, 3) else (w, h)
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
                # 标准图幅（A0~A4）的图框一律留着：页面设置里的纸选错成 Letter 时，
                # 真正的图框会比「4 倍纸面积」还大，误删就会把图框和标题栏切掉。
                w, h = b[2] - b[0], b[3] - b[1]
                long_, short_ = max(w, h), min(w, h)
                if not any(abs(long_ - pw2) <= 6 and abs(short_ - ph2) <= 6
                           for pw2, ph2 in PAPER_SIZES.values()):
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
            # 图纸空间里的图形（图框线 + 标题栏 + 标注）本身差不多正好是一张纸：
            # 它才是图框。有的图把视口画得比纸还宽（实测 10200-C-0：视口 1546mm 宽、
            # 纸只有 1191mm），照视口出图就被压成半页空白。
            if other and paper_area > 0:
                ow, oh = other[2] - other[0], other[3] - other[1]
                if ow > 5 and oh > 5:
                    oar = max(ow, oh) / min(ow, oh)
                    oarea = ow * oh / paper_area
                    if abs(oar - want) <= 0.08 * want and 0.8 <= oarea <= 1.35:
                        return other
            # 一个都不像图框（总览视口又大又扁、局部视口又小）：挑「面积最接近纸张」的
            # 那个视口——它通常就是真正画在纸上的视图；再和图纸空间里的图框线取并集。
            if vps:
                def _closeness(b):
                    area = max((b[2] - b[0]) * (b[3] - b[1]), 1e-6)
                    return abs(math.log(area / paper_area))
                best = min(vps, key=_closeness)
                return self._union([best, other])
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
                self.log(tr(f"    · 没找到 {round(orig_paper[0])}x{round(orig_paper[1])} mm 的纸，"
                            f"改用 {media}",
                            f"    · no {round(orig_paper[0])}x{round(orig_paper[1])} mm paper "
                            f"available, using {media}"))
            return exact
        return False

    def _rotation_for(self, media, cw: float, ch: float) -> int:
        """
        纸和图的横竖不一致时，出图要旋转 90°（0 不转）。
        DWG To PDF 的 ISO/ANSI 纸名只有一种写法（A0 就是 841x1189 竖版），
        横图要 1:1 出进去只能靠旋转——旋转后 PDF 页面本身也变成横的
        （实测 PlotRotation=1/3 → 页面 1189x841）。
        """
        size = self._media_size(media) if media else None
        if not size:
            return 0
        media_land = size[0] >= size[1]
        want_land = cw >= ch
        return 1 if media_land != want_land else 0

    def _sheet_size_after(self, media, rot: int):
        """选完纸、定好旋转之后，PDF 页面实际是多少毫米"""
        size = self._media_size(media) if media else None
        if not size:
            return None
        w, h = size
        return (h, w) if rot in (1, 3) else (w, h)

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
        # 纸本身选错了也得先换纸：原打印设备的纸名在 DWG To PDF 里没有时，AutoCAD 会
        # 退回默认的 Letter，A0 的图被硬塞进 Letter 只会糊成一团（实测就是这么出的）。
        # 判据：图比纸大，而且纸的短边不到图短边的 7 成。
        target = paper
        if (not fits_first) and min(pw0, ph0) < 0.7 * min(cw, ch):
            target = (cw, ch)
        # 界面里指定了纸张（A0~A4）时，图装得下就按它出——纸大一圈，图就等比放大一圈，
        # 字更大、打印更清楚。图纸本身比选定的纸大时 _apply_paper_choice 不会改，
        # 免得把 A0 的图缩到 A1 反而更糊。
        meta_mm, meta_tag = self._apply_paper_choice((cw, ch))
        if meta_tag:
            target = meta_mm
        # 纸张方向跟着图走：图是横的就用横版纸，别再靠「旋转 90°」硬塞，
        # 那样图会被纸边裁掉一半（实测就是这么出的问题）。
        # 用满幅无边距纸（full_bleed）的两种情况：
        # - 内容按 1:1 就铺满整张纸（≥75%）：普通纸 5/17 mm 的可打印边距会把纸边那圈内容裁掉；
        # - 需要按图框缩放（内容比纸大、或图框只占纸的一小块）：既然要缩，就缩到纸边，
        #   否则普通纸的边距会让整页四周留出一大圈白边（实测只剩 81%，看着就是「没铺满」）。
        edge_to_edge = (fits_first and fill_first >= 0.75) or (not fits_first) or fill_first < 0.80
        media, mland, _exact = self._pick_media(layout, "", target, cw >= ch,
                                                allow_bleed=edge_to_edge)
        if edge_to_edge and not _exact:
            # 满幅纸里没有这个尺寸（非标纸幅、英制纸等）：退回普通纸，别硬凑一张别的尺寸
            m2, l2, e2 = self._pick_media(layout, "", target, cw >= ch, allow_bleed=False)
            if e2:
                media, mland, _exact = m2, l2, e2
        if media:
            try:
                layout.CanonicalMediaName = media
            except Exception:
                pass
        # 纸张方向必须跟图一致：DWG To PDF 的 A 系列纸名只有竖版（A0 就是 841x1189），
        # 想把横图（1189x841）1:1 出进去，只能把出图旋转 90°（页面本身会变成 1189x841）。
        rot = self._rotation_for(media, cw, ch)
        pw, ph = self._sheet_size_after(media, rot) or paper

        fits = self._fits_in(cw, ch, (pw, ph))
        # 图框占纸张的比例。正常图纸的图框≈纸张（实测 92%~100%）；
        # 有的图图框比纸小一大截（例如 122x88 mm 的图框配 A4 纸，只占 17%），
        # 那是页面设置和图纸对不上，1:1 出来就是「图缩在页面中间、四周一圈白边」。
        # 这种就不再坚持 1:1，直接把图框铺满纸张。
        fill = (cw * ch / (pw * ph)) if pw * ph else 0.0
        keep_scale = fits and fill >= 0.80
        if meta_tag:
            self._note_paper_choice(meta_tag, (cw, ch), (pw, ph))

        # 窗口四周统一留 1mm：出图时图框线不会正好压在可打印区边界上（否则线会被裁掉一半）
        pad = 1.0
        if not keep_scale and not fits:
            # 内容比纸大时这点余量没意义，反而让缩放的基数变模糊，干脆不留
            pad = 0.0
        if keep_scale and (cw + 2 * pad > pw + 0.5 or ch + 2 * pad > ph + 0.5):
            # 图框本来就顶到纸边（差不到 1mm）时再加余量，出图会被纸边裁掉一圈
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
                    self.log(tr("    · 原打印设备不是 PDF 设备，页面设置已按图纸内容重新框定并居中"
                                "（比例不变）",
                                "    · the original plot device is not a PDF device; the page "
                                "setup was re-fitted and centred on the drawing content "
                                "(scale unchanged)"))
            else:
                layout.UseStandardScale = True
                layout.StandardScale = 0                   # 0 = 布满图纸
                if fits:
                    self.log(tr(f"    · 图框只有 {round(cw)}x{round(ch)} mm，仅占 "
                                f"{round(pw)}x{round(ph)} mm 纸张的 {fill * 100:.0f}%（页面设置"
                                f"与图纸不匹配），已按图框铺满本页",
                                f"    · the frame covers only {round(cw)}x{round(ch)} mm, i.e. "
                                f"{fill * 100:.0f}% of the {round(pw)}x{round(ph)} mm paper "
                                f"(page setup does not match the drawing); the frame was "
                                f"fitted to fill the page"))
                else:
                    self.log(tr(f"    · 图纸内容 {round(cw)}x{round(ch)} mm 比纸张 "
                                f"{round(pw)}x{round(ph)} mm 大，已按「布满图纸」缩放到本页"
                                f"（整张图完整、不裁切）",
                                f"    · the drawing is {round(cw)}x{round(ch)} mm, larger than "
                                f"the {round(pw)}x{round(ph)} mm paper; it was scaled to "
                                f"\"fit to paper\" (whole sheet, nothing cropped)"))
            # 旋转放到最后设：先设旋转再设窗口，窗口坐标会被当成「已旋转的坐标系」来解释，
            # 出图结果是横躺并被裁掉（实测）。顺序反了就是这个下场。
            layout.PlotRotation = rot
        except Exception as exc:                           # noqa: BLE001
            self.log(tr(f"    · 按图纸内容重设窗口失败，仍按原页面设置出图：{exc}",
                        f"    · could not re-fit the plot window to the drawing content; "
                        f"plotting with the original page setup: {exc}"))

    def _plot_sheet(self, doc, sheet: dict, pdf_path: str) -> None:
        """把模型空间里的一张图（按图框窗口 + 对应图幅）打印成一个 PDF"""
        model = None
        for i in range(doc.Layouts.Count):
            if str(doc.Layouts.Item(i).Name).strip().lower() == MODEL_LAYOUT:
                model = doc.Layouts.Item(i)
                break
        if model is None:
            raise RuntimeError(tr("找不到模型空间布局", "Model-space layout not found"))

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
                self.log(tr(f"    · 图纸未设打印样式表，改用 {MONO_CTB} 输出（否则白字会看不见）",
                            f"    · the drawing has no plot style table; using {MONO_CTB} "
                            f"(otherwise white text would be invisible)"))
        except Exception:
            pass
        w, h = sheet["x1"] - sheet["x0"], sheet["y1"] - sheet["y0"]
        want_landscape = w >= h
        # 用户指定了纸张（界面上的「纸张大小」）就用它，只放大不缩小
        meta_mm, meta_tag = self._apply_paper_choice(sheet.get("paper_mm"))
        # 模型空间的图一律「布满图纸」出图，所以配满幅无边距纸：普通纸 5/17 mm 的
        # 可打印边距会把纸边那圈内容裁掉，也让整页四周多出一圈白边。
        media, media_landscape, media_exact = self._pick_media(
            model, sheet["paper"], meta_mm, want_landscape, allow_bleed=True)
        if not media:
            media, media_landscape, media_exact = self._pick_media(
                model, sheet["paper"], meta_mm, want_landscape)
        if media:
            try:
                model.CanonicalMediaName = media
            except Exception:
                pass
            if meta_tag and sheet.get("paper_mm"):
                self._note_paper_choice(meta_tag, sheet["paper_mm"], meta_mm)
            elif not media_exact:
                self.log(tr(f"    · 图幅 {sheet['paper']} 没有对应纸张，按 {media} 缩放输出",
                            f"    · no paper matches sheet size {sheet['paper']}; "
                            f"plotting scaled onto {media}"))


        # 注意顺序：必须先设置窗口范围，再切到「按窗口出图」。
        # 反过来时，如果图纸原有的窗口范围是空的（很多图 PlotType=0 显示），
        # AutoCAD 会直接报「输入无效」。
        # 窗口四周留 0.4 mm（换算成图形单位）：满幅纸上图框线正好压在纸边，
        # 不留这点余量，图框最外那条线会被裁掉一半。
        pad = 0.4 * float(sheet.get("scale") or 1)
        model.SetWindowToPlot(self._double_array([sheet["x0"] - pad, sheet["y0"] - pad]),
                              self._double_array([sheet["x1"] + pad, sheet["y1"] + pad]))
        model.PlotType = 4                                  # 4 = 窗口
        model.CenterPlot = True
        model.UseStandardScale = True
        model.StandardScale = 0                             # 0 = 布满图纸
        # 注意：模型空间的窗口是「图形坐标」，出图旋转会把图形本身转 90°（实测出来的
        # 图是躺着的、还会被纸边裁掉）。所以这里靠「挑一张方向一致的纸」解决，
        # 不能靠 PlotRotation。真没有同方向的纸（A0 只有竖版）就宁可留白也不要转。
        try:
            model.PlotRotation = 0
        except Exception:
            pass

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
            if self._wait_file(temp_path, timeout=8):
                if self._pdf_page_is_empty(temp_path):
                    _remove_quietly(temp_path)
                    raise EmptySheet(tr(f"图框 {sheet.get('paper', '')} 里没有内容（整页空白）",
                                        f"nothing inside the {sheet.get('paper', '')} sheet "
                                        f"frame (blank page)"))
                try:
                    os.replace(temp_path, pdf_path)
                    return
                except OSError as exc:
                    _remove_quietly(pdf_path)
                    try:
                        os.replace(temp_path, pdf_path)
                        return
                    except OSError:
                        last_err = tr(f"PDF 改名失败：{exc}",
                                      f"Failed to rename the PDF: {exc}")
                        continue
            last_err = tr("AutoCAD 未生成文件", "AutoCAD produced no file")
            time.sleep(1.0)
        _remove_quietly(temp_path)
        raise RuntimeError(tr(f"出图未生成文件（图幅 {sheet['paper']}"
                              + (f"；最后错误：{last_err}" if last_err else "") + "）",
                              f"No file was produced (sheet size {sheet['paper']}"
                              + (f"; last error: {last_err}" if last_err else "") + ")"))

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
        native_ok = (not self.force_pc3) and self._is_pdf_device(orig_cfg)
        if native_ok:
            plans.append(orig_cfg)
        if orig_cfg.strip().lower() != self.pc3.strip().lower():
            plans.append(self.pc3)

        temp_path = pdf_path + TEMP_EXT          # .part：无关联扩展名，不会被自动打开
        last_err = ""

        def plot_once(cfg: str) -> bool:
            _remove_quietly(temp_path)
            plot.PlotToFile(temp_path, cfg)     # ← 传设备名，且写临时扩展名
            return self._wait_file(temp_path, timeout=2.5)

        # 页面设置里的出图窗口比纸还大：照它出图一定在纸上留一大片空白，
        # 而且这个窗口本身就没有意义，直接按图纸内容重新框定再出（比出完再判断快得多）。
        window_bad = self._window_overflows_paper(layout, orig_paper)

        for cfg in plans:
            fitted = False
            if (cfg.strip().lower() != orig_cfg.strip().lower()) or window_bad:
                # 换了打印设备：先把纸张按原尺寸还原；换设备后原来的页面设置
                # （尤其是「旋转」和可打印边距）含义会变，所以再按图纸内容
                # 重新框一次窗口并居中，避免「图缩在一个角、其余大片空白」。
                switch_to_pdf_device()
                self._fit_layout_to_page(layout)
                fitted = True
                window_bad = False
            for _try in range(2):
                try:
                    ok = plot_once(cfg)
                except Exception as exc:               # noqa: BLE001
                    last_err = str(exc)
                    break
                if not ok:
                    last_err = tr("AutoCAD 未生成文件", "AutoCAD produced no file")
                    break
                stats = self._pdf_page_stats(temp_path)
                # 出成了一张空白页：先按图纸内容重新框定再试一次（页面设置里的纸
                # 跟图纸对不上时，内容会跑到纸张矩形外面，出图就是白的）。
                if stats is not None and not fitted and self._stats_are_empty(stats):
                    self.log(tr("    · 这一页是空白（页面设置与图纸内容不匹配），"
                                "按图纸内容重新框定后重出",
                                "    · the page came out blank (page setup does not match the "
                                "drawing); re-fitting to the drawing content and plotting again"))
                    switch_to_pdf_device()
                    self._fit_layout_to_page(layout)
                    fitted = True
                    continue
                # 重框过还是空白 → 这个布局本身就没内容（DWG 里留的空布局，
                # AutoCAD 里手动打印同样是白纸），不该占 PDF 的一页。
                if self._stats_are_empty(stats):
                    _remove_quietly(temp_path)
                    raise EmptySheet(tr(f"{layout_name} 是空布局（整页没有图形和文字）",
                                        f"{layout_name} is an empty layout "
                                        f"(no graphics and no text on the page)"))
                if self._stats_are_sparse(stats):
                    self.log(tr(f"    · 布局 {layout_name} 几乎没有图形，只出到图框和文字，"
                                f"请核对（多半是视口没内容、图层被关掉或外部参照丢了）",
                                f"    · layout {layout_name} has almost no graphics — only the "
                                f"frame and text were plotted, please check it (usually an empty "
                                f"viewport, layers switched off, or a missing xref)"))
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
                        last_err = tr(f"PDF 改名失败：{exc}",
                                      f"Failed to rename the PDF: {exc}")
                        break

        _remove_quietly(temp_path)
        raise RuntimeError(tr(
            f"出图未生成文件（已尝试 {len(plans)} 种打印设备："
            + "、".join(plans) + (f"；最后错误：{last_err}" if last_err else "") + "）",
            f"No file was produced (tried {len(plans)} plot device(s): "
            + ", ".join(plans) + (f"; last error: {last_err}" if last_err else "") + ")"
        ))

    @staticmethod
    def _window_overflows_paper(layout, paper) -> bool:
        """
        页面设置里的「窗口」比纸张还大 → 照它出图必然在纸上空出一大片
        （实测 10200-C-0：窗口 1690x841 mm、纸 1191x843 mm，图只占左半边 49%）。
        图纸空间里的窗口坐标就是毫米，可以直接跟纸幅比。
        """
        if not paper:
            return False
        try:
            if int(layout.PlotType) != 4:
                return False
            win = layout.GetWindowToPlot()
            ll, ur = win[0], win[1]
            ww = abs(float(ur[0]) - float(ll[0]))
            wh = abs(float(ur[1]) - float(ll[1]))
        except Exception:
            return False
        if ww <= 1 or wh <= 1:
            return False
        # 窗口的任一边都不该超过纸张的长边（正放/横放都算上，取宽松的那条）
        limit = max(paper) * 1.08
        return ww > limit or wh > limit

    @staticmethod
    def _pdf_page_is_empty(path: str) -> bool:
        """
        判断刚出的这一页是不是「整页空白」——一个图形一个字都没有。

        实测（泰国 DEC 那批 DWG）页面内容流分层很清楚：
          · 真图纸：文字算子 60~800 个、路径算子 6000~160000 个、流 170 KB~2.5 MB
          · 空布局：文字算子 0 个、路径算子 10 个（只剩纸边那一个矩形）、流 203 字节
        判据是「没有文字 + 路径极少 + 流很短」三条同时成立，跟真图纸之间隔着一个
        数量级以上的余量，不会把稀疏但正常的图纸误判成空。

        读不了（没装 pypdf、文件还没写完）一律当作不空——宁可多出一页空白，
        也不能因为解析失败把真图纸丢掉。
        """
        stats = AcadPlotter._pdf_page_stats(path)
        if stats is None:
            return False
        return AcadPlotter._stats_are_empty(stats)

    @staticmethod
    def _pdf_page_stats(path: str):
        """
        读一页 PDF，返回 (文字算子数, 路径算子数, 内容流字节数)。
        读不了返回 None——调用方一律按「不空、正常」处理。
        """
        try:
            import logging
            from pypdf import PdfReader
            # AutoCAD 出的 PDF 里偶尔有重复键，pypdf 每读一页就打一条
            # "Multiple definitions in dictionary" 到控制台，纯噪音，关掉
            logging.getLogger("pypdf").setLevel(logging.ERROR)
        except Exception:
            return None
        try:
            reader = PdfReader(path)
            if not reader.pages:
                return None
            contents = reader.pages[0].get_contents()
            data = contents.get_data() if contents is not None else b""
        except Exception:
            return None
        if not data:
            return (0, 0, 0)
        return (len(TEXT_SHOW.findall(data)), len(PATH_BUILD.findall(data)), len(data))

    @staticmethod
    def _stats_are_empty(stats) -> bool:
        """整页空白：一个字都没有，只剩纸边那一两条线"""
        if not stats:
            return False
        n_text, n_path, n_bytes = stats
        if n_bytes >= 120000:              # 这么大的内容流不可能是空的
            return False
        if n_text > 0:                     # 有文字就一定不是空页
            return False
        return n_path < 1000

    @staticmethod
    def _stats_are_sparse(stats) -> bool:
        """
        有字、但几乎没有线条：典型是「只出到图框和标题栏，图本身没出来」
        （视口里没内容、图层被关掉、参照文件丢失）。
        实测真图纸最少也有 5776 个路径算子、106 KB 内容流，这里的门槛留了 2~4 倍余量。
        """
        if not stats:
            return False
        n_text, n_path, n_bytes = stats
        return n_text > 0 and n_path < 2500 and n_bytes < 90000

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
        self._insert_cache = None      # 块参照名字索引（每个 DWG 只建一次）
        self._text_pos_cache = None    # 文字插入点索引
        self._fonts_repaired = False
        self._font_cache = None
        self._fit_noted = False
        self._paper_noted = set()

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
                            self.log(tr(f"  × 图框识别出错：{_humanize_com_error(exc)}",
                                        f"  × sheet detection failed: {_humanize_com_error(exc)}"))
                            break
                        wait = 2 * (attempt + 1)
                        self.log(tr(f"  · 识别图框时 AutoCAD 正忙，{wait} 秒后重试…",
                                    f"  · AutoCAD is busy while detecting sheets; "
                                    f"retrying in {wait} s…"))
                        time.sleep(wait)
                if sheets:
                    return sheets
                if self.frames_mode == "frames":
                    raise RuntimeError(tr("已指定按图框拆分，但没识别到图框"
                                          "（模型空间里需要有 A0-A4 图框块或标准图幅的矩形边框）",
                                          "Frame splitting was requested but no frame was found "
                                          "(model space needs A0-A4 frame blocks or rectangles of "
                                          "standard sheet size)"))
                self.log(tr("  ! 没识别到图框，改为按图纸布局输出"
                            "（这张图如果多张图纸都排在模型空间里，就会只出一张）",
                            "  ! No frame detected; falling back to plotting one page per layout "
                            "(if several sheets are laid out in model space, only one page "
                            "will come out)"))
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
            self.log(tr(f"    · 图纸有 {fixed} 个文字样式用了本机没有的字体，"
                        f"已临时替换（否则这些字不会出图）",
                        f"    · {fixed} text style(s) use fonts missing on this machine; "
                        f"temporarily substituted (otherwise that text would not plot)"))
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
            return [LayoutResult(win_path, "", "", "fail",
                                 tr(f"打开图纸失败：{msg}", f"Failed to open the drawing: {msg}"))]

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
                                                tr("PDF 已存在，已跳过",
                                                   "PDF already exists — skipped")))
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
                except EmptySheet:
                    results.append(LayoutResult(win_path, name, "", "empty",
                                                tr("空布局，已跳过",
                                                   "Empty layout — skipped")))
                    continue
                except Exception as exc:          # noqa: BLE001
                    results.append(LayoutResult(win_path, name, pdf_path, "fail",
                                                tr(f"出图失败：{_humanize_com_error(exc)}",
                                                   f"Plot failed: {_humanize_com_error(exc)}")))
                    continue

                if self._wait_file(pdf_path):
                    size_kb = os.path.getsize(pdf_path) / 1024
                    extra = f"{target['paper']} · " if is_sheet and target.get("paper") else ""
                    results.append(LayoutResult(win_path, name, pdf_path, "ok",
                                                f"{extra}{size_kb:.0f} KB"))
                else:
                    results.append(LayoutResult(win_path, name, pdf_path, "fail",
                                                tr("没有生成 PDF 文件（可能是图纸为空或打印设备不可用）",
                                                   "No PDF was created (the sheet may be empty or "
                                                   "the plot device unavailable)")))
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
            self.log(tr(f"  × 打开失败：{_humanize_com_error(exc)}",
                        f"  × open failed: {_humanize_com_error(exc)}"))
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
                             tr(f"图幅 {target['paper']}", f"sheet {target['paper']}")
                             if target.get("paper") else tr("预览", "preview"))
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
                self.log(tr("AutoCAD 将保留在后台，下次转换启动更快",
                            "AutoCAD is kept running in the background — the next run starts faster"))
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
        return tr("该 DWG 已在 AutoCAD 中打开，请先关闭后重试",
                  "This DWG is already open in AutoCAD — close it and try again")
    if "拒绝访问" in text or "access is denied" in low:
        return tr("文件被占用或没有读取权限",
                  "The file is locked or cannot be read (permission denied)")
    if "找不到" in text or "not found" in low or "无法打开" in text:
        return tr("文件不存在、已损坏，或由更高版本 CAD 保存（需用同版本或更高版本打开）",
                  "The file is missing, damaged, or saved by a newer CAD version "
                  "(open it with the same or a newer version)")
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
                 order_pages: bool = True, paper_size: str = ""):
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
        # 合并模式下是否把 LAYOUT 图排第 1 页、BOM 图排第 2 页（界面默认不勾）
        self.order_pages = bool(order_pages)
        # 出图纸张：'' / auto = 跟图纸走；A0~A4 = 不小于图纸的那张
        self.paper_size = (paper_size or "").strip().upper()
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

        names = [str(getattr(r, "layout", "") or "") for r in items]
        if len(set(names)) < len(names):
            # 图纸名有重复：多半是标题没读出来、几张图都取了同一段说明文字，
            # 这时只能靠页面文字猜，猜错了会把页序搞乱，所以宁可不调。
            self._log(tr("  · 图纸标题有重复，无法可靠判断页序，保持图纸原顺序",
                         "  · Sheet titles repeat, so page order cannot be decided reliably — "
                         "keeping the original order"))
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
            self._log(tr(f"  · 识别到 {len(lay)} 张 LAYOUT 图、{len(bom)} 张 BOM 图，"
                         f"无法唯一确定页序，保持图纸原顺序",
                         f"  · found {len(lay)} LAYOUT sheet(s) and {len(bom)} BOM sheet(s); "
                         f"page order is ambiguous — keeping the original order"))
            return items, ""
        if not lay and not bom:
            return items, ""

        if lay and bom:
            a, b = lay[0], bom[0]
            rest = [r for i, r in enumerate(items) if i not in (a, b)]
            new = [items[a], items[b]] + rest
            note = tr("LAYOUT 图排第 1 页、BOM 图排第 2 页",
                      "LAYOUT sheet first, BOM sheet second")
        elif lay:
            a = lay[0]
            new = [items[a]] + [r for i, r in enumerate(items) if i != a]
            note = tr("LAYOUT 图排第 1 页", "LAYOUT sheet first")
        else:
            b = bom[0]
            rest = [r for i, r in enumerate(items) if i != b]
            if not rest:
                return items, ""
            new = rest[:1] + [items[b]] + rest[1:]
            note = tr("BOM 图排第 2 页", "BOM sheet second")

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
            return [LayoutResult(dwg, "", final, "skip",
                                 tr("PDF 已存在，已跳过", "PDF already exists — skipped"))]

        tmpdir = tempfile.mkdtemp(prefix="dwg2pdf_")
        try:
            def namer(targets):
                return [(t, os.path.join(tmpdir, "%04d.pdf" % n))
                        for n, t in enumerate(targets, 1)]

            def on_sheet(idx, total, name):
                self._log(tr(f"     第 {idx}/{total} 张：{name}",
                             f"     sheet {idx}/{total}: {name}"))

            try:
                raw = plotter.export_drawing(dwg, namer, on_sheet=on_sheet)
            except Exception as exc:                   # noqa: BLE001
                self._log(tr(f"  × 处理失败：{_humanize_com_error(exc)}，"
                             f"正在重新连接 AutoCAD 后重试…",
                             f"  × failed: {_humanize_com_error(exc)}; "
                             f"reconnecting to AutoCAD and retrying…"))
                plotter.connect()
                raw = plotter.export_drawing(dwg, namer, on_sheet=on_sheet)

            ok = [r for r in raw if r.status == "ok" and r.pdf and os.path.isfile(r.pdf)]
            empty = [r for r in raw if r.status == "empty"]
            bad = [r for r in raw if r not in ok and r not in empty]
            if not ok:
                if empty:
                    return empty
                return raw or [LayoutResult(dwg, "", "", "fail",
                                            tr("没有可导出的图纸（该文件没有布局）",
                                               "Nothing to export (this file has no layout)"))]

            os.makedirs(folder, exist_ok=True)
            # 默认把 LAYOUT 图放第 1 页、BOM 图放第 2 页（--no-page-order 可关掉）
            ordered, note = self._order_for_merge(ok)
            if note:
                self._log(tr(f"  · 页序已调整：{note}", f"  · page order adjusted: {note}"))
            try:
                pages = _merge_pdfs([r.pdf for r in ordered], final)
            except Exception as exc:                   # noqa: BLE001
                self._log(tr(f"  ！合并成一个 PDF 失败：{exc}；已改为每张图纸一个文件",
                             f"  ! merging into one PDF failed: {exc}; "
                             f"falling back to one PDF per sheet"))
                fallback: list[LayoutResult] = []
                for r in ok:
                    base = safe_filename(f"{self.prefix}{stem} - {r.layout}")
                    dst = os.path.join(folder, self._unique_name(base, folder, stem) + ".pdf")
                    try:
                        shutil.move(r.pdf, dst)
                    except OSError as exc2:
                        fallback.append(LayoutResult(dwg, r.layout, r.pdf, "fail",
                                                     tr(f"文件无法移动：{exc2}",
                                                        f"Cannot move the file: {exc2}")))
                        continue
                    fallback.append(LayoutResult(dwg, r.layout, dst, "ok",
                                                 tr("合并失败，已按单张输出",
                                                    "Merge failed — written as a single-sheet PDF")))
                return fallback + bad

            size_mb = os.path.getsize(final) / 1024 / 1024
            msg = tr(f"{pages} 页 · {size_mb:.1f} MB",
                     f"{pages} pages · {size_mb:.1f} MB")
            if empty:
                msg += tr(f"，已跳过 {len(empty)} 个空布局",
                          f", {len(empty)} empty layout(s) skipped")
            if bad:
                msg += tr(f"，另有 {len(bad)} 张未出图",
                          f", {len(bad)} sheet(s) not plotted")
            return (empty + [LayoutResult(dwg, tr(f"共 {pages} 张图纸", f"{pages} sheets"),
                                          final, "ok", msg)]
                    + bad)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def _preview_merged(self, plotter, dwg: str, folder: str, stem: str) -> list[LayoutResult]:
        """试运行：只算出会合并成几页，不出图"""
        final = self._merged_path(folder, stem)
        got = plotter.walk_layouts(dwg, lambda targets: [(t, "") for t in targets])
        if not got:
            return []
        return [LayoutResult(dwg, tr(f"共 {len(got)} 张图纸", f"{len(got)} sheets"),
                             final, "preview",
                             tr(f"合并为 1 个 PDF，共 {len(got)} 页",
                                f"merged into 1 PDF, {len(got)} pages"))]

    # -------------------------------------------------- 主流程
    def run(self) -> dict:
        if not os.path.isdir(self.src):
            raise RuntimeError(tr(f"源文件夹不存在：{self.src}",
                                  f"Source folder does not exist: {self.src}"))
        os.makedirs(self.out, exist_ok=True)

        if self.single_file:
            dwg_files = [self.single_file]
        else:
            dwg_files = find_dwg_files(self.src, self.recursive, exclude_dirs=[self.out])
        if not dwg_files:
            self._log(tr(f"在 {self.src} 里没有找到任何 .dwg 文件",
                         f"No .dwg file found in {self.src}"))
            self._emit("done", total=0, ok=0, fail=0, skip=0, results=[])
            return {"total": 0, "ok": 0, "fail": 0, "skip": 0, "results": []}

        self._log(tr(f"找到 {len(dwg_files)} 个 DWG，输出目录：{self.out}",
                     f"Found {len(dwg_files)} DWG file(s); output folder: {self.out}"))
        self._emit("scan_done", dwg_count=len(dwg_files))

        plotter = AcadPlotter(pc3=self.pc3, include_model=self.include_model,
                              frames_mode=self.frames_mode,
                              paper_size=self.paper_size, log=self._log)
        plotter.connect()

        counters = {"ok": 0, "fail": 0, "skip": 0, "empty": 0}
        preview_count = 0
        try:
            for file_no, dwg in enumerate(dwg_files, 1):
                if self.cancel():
                    self._log(tr("用户已停止转换", "Stopped by the user"))
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
                        self._log(tr(f"  × 处理失败：{msg}，正在重新连接 AutoCAD 后重试…",
                                     f"  × failed: {msg}; reconnecting to AutoCAD and retrying…"))
                        try:
                            plotter.connect()
                            results = plotter.export_drawing(
                                dwg, namer, skip_existing=self.skip_existing)
                        except Exception as exc2:        # noqa: BLE001
                            results = [LayoutResult(dwg, "", "", "fail",
                                                    _humanize_com_error(exc2))]

                if not results:
                    self._log(tr("  没有可导出的图纸（该文件只有模型空间，或没有布局）",
                                 "  Nothing to export (this file has only model space, or no layout)"))

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
                        self._log(tr(f"  预览 → {fname}   （图纸：{r.layout}）",
                                     f"  preview → {fname}   (sheet: {r.layout})"))
                    elif r.status == "skip":
                        self._log(tr(f"  - 跳过 {fname}", f"  - skipped {fname}"))
                    elif r.status == "empty":
                        # 空布局：DWG 里留着、但页面上一个图形一个字都没有，
                        # 出成 PDF 就是一张白纸，占着页数反而碍事，直接不出页。
                        self._log(tr(f"  - 空布局不出页：{r.layout}",
                                     f"  - empty layout, no page: {r.layout}"))
                    else:
                        self._log(f"  × {r.layout or os.path.basename(r.dwg)}：{r.message}")
                    self._emit("result", result=r.to_dict())
        finally:
            plotter.close(keep_acad=self.keep_acad)

        summary = {"total": len(self.results), **counters, "preview": preview_count}
        self._log("")
        if self.dry_run:
            self._log(tr(f"预览结束：实际会生成 {preview_count} 个 PDF（本次未出图）",
                         f"Preview finished: {preview_count} PDF(s) would be created "
                         f"(nothing was plotted this time)"))
        else:
            tail = (tr(f"，空布局 {counters['empty']} 个已跳过",
                       f", {counters['empty']} empty layout(s) skipped")
                    if counters.get("empty") else "")
            self._log(tr(f"完成：成功 {counters['ok']} 个，跳过 {counters['skip']} 个，"
                         f"失败 {counters['fail']} 个{tail}",
                         f"Done: {counters['ok']} succeeded, {counters['skip']} skipped, "
                         f"{counters['fail']} failed{tail}"))
        self._emit("done", results=[r.to_dict() for r in self.results], **summary)
        return {**summary, "results": [r.to_dict() for r in self.results]}


# ---------------------------------------------------------------- 命令行

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="dwg2pdf",
        description=tr("把 DWG 里的每张图纸（布局）批量导出为单个 PDF，文件名为图纸名",
                       "Export every sheet (layout) of a DWG to PDF(s), named after the sheet"),
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("-s", "--src", required=True,
                        help=tr("源文件夹（存放 DWG）", "Source folder holding the DWGs"))
    parser.add_argument("-o", "--out", required=True,
                        help=tr("输出 PDF 的文件夹", "Folder for the output PDFs"))
    parser.add_argument("-p", "--prefix", default="",
                        help=tr("文件名前缀，如 A-（可选）", "Filename prefix, e.g. A- (optional)"))
    parser.add_argument("-t", "--tpl", default="{prefix}{layout}",
                        help=tr("命名模板，默认 {prefix}{layout}\n"
                                "可用占位符：{prefix} {layout} {dwg} {n}",
                                "Name template, default {prefix}{layout}\n"
                                "Placeholders: {prefix} {layout} {dwg} {n}"))
    parser.add_argument("--include-model", action="store_true",
                        help=tr("同时导出模型空间（默认只导出图纸布局）",
                                "Also export model space (default: layouts only)"))
    parser.add_argument("--frames", dest="frames_mode", action="store_const", const="frames",
                        help=tr("强制按模型空间图框拆分：每张图框（A0-A4）输出一个 PDF",
                                "Force splitting by model-space frames: one PDF per A0-A4 frame"))
    parser.add_argument("--layouts", dest="frames_mode", action="store_const", const="layouts",
                        help=tr("只按图纸布局输出，不做图框拆分",
                                "Plot layouts only, no frame splitting"))
    parser.set_defaults(frames_mode="auto")
    parser.add_argument("--no-recursive", action="store_true",
                        help=tr("只处理第一层，不含子文件夹",
                                "Only the top folder, skip sub-folders"))
    parser.add_argument("--keep-subdir", action="store_true",
                        help=tr("在输出目录中保持源文件夹结构",
                                "Mirror the source folder structure in the output folder"))
    parser.add_argument("--skip-existing", action="store_true",
                        help=tr("跳过已存在的 PDF（断点续转）",
                                "Skip PDFs that already exist (resume where you left off)"))
    parser.add_argument("--pc3", default=None,
                        help=tr(f"强制使用指定打印设备，默认 {DWG_TO_PDF}",
                                f"Force a specific plot device, default {DWG_TO_PDF}"))
    parser.add_argument("--one-pdf", dest="one_pdf", action="store_true",
                        help=tr("一个 DWG 只输出一个多页 PDF（每页一张图纸），默认就是这种方式",
                                "One multi-page PDF per DWG (one sheet per page) — the default"))
    parser.add_argument("--split", dest="one_pdf", action="store_false",
                        help=tr("每张图纸各输出一个 PDF（旧方式）",
                                "One PDF per sheet (legacy behaviour)"))
    parser.set_defaults(one_pdf=True)
    parser.add_argument("--page-order", dest="order_pages", action="store_true",
                        help=tr("合并时把 LAYOUT 图排第 1 页、BOM 图排第 2 页（默认不排，保持图纸顺序）",
                                "Put the LAYOUT sheet first and the BOM sheet second when merging "
                                "(off by default: keep the sheet order)"))
    parser.add_argument("--no-page-order", dest="order_pages", action="store_false")
    parser.set_defaults(order_pages=False)
    parser.add_argument("--paper", default="A1", choices=["auto", "A0", "A1", "A2", "A3", "A4"],
                        help=tr("出图纸张，默认 A1：图纸比它小就等比放大（字更清楚），"
                                "比它大就保持图纸原尺寸；选 auto 则完全跟着图纸走",
                                "Paper size, default A1: sheets smaller than it are scaled up "
                                "(clearer text), larger sheets keep their own size; auto = follow "
                                "each sheet"))
    parser.add_argument("--keep-acad", action="store_true",
                        help=tr("转换完成后保留 AutoCAD 进程（下次启动更快）",
                                "Keep AutoCAD running afterwards (faster next start)"))
    parser.add_argument("--dry-run", action="store_true",
                        help=tr("只列出会生成哪些 PDF，不出图",
                                "Only list the PDFs that would be created, do not plot"))
    parser.add_argument("--check", action="store_true",
                        help=tr("只检查 AutoCAD 与 DWG 数量",
                                "Only check AutoCAD and count the DWGs"))
    parser.add_argument("--json", action="store_true",
                        help=tr("把结果以 JSON 打印到标准输出",
                                "Print the result as JSON to stdout"))
    parser.add_argument("--log", default=None,
                        help=tr("同时把日志写入指定文件", "Also write the log to this file"))
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
                              paper_size=getattr(args, "paper", ""), log=print)
        try:
            plotter.connect()
        except Exception as exc:                        # noqa: BLE001
            print(f"× {exc}")
            return 3
        files = find_dwg_files(args.src, not args.no_recursive) if os.path.isdir(args.src) else []
        print(tr(f"AutoCAD 版本：{plotter.version}", f"AutoCAD version: {plotter.version}"))
        print(tr(f"DWG 文件数：{len(files)}", f"DWG files: {len(files)}"))
        for f in files[:20]:
            print("  " + os.path.basename(f))
        if len(files) > 20:
            print(tr(f"  …共 {len(files)} 个", f"  …{len(files)} in total"))
        plotter.close()
        return 0

    conv = Converter(
        src=args.src, out=args.out, prefix=args.prefix, name_tpl=args.tpl,
        include_model=args.include_model, recursive=not args.no_recursive,
        keep_subdir=args.keep_subdir, skip_existing=args.skip_existing,
        pc3=args.pc3, on_event=on_event, dry_run=args.dry_run,
        keep_acad=args.keep_acad, frames_mode=args.frames_mode,
        one_pdf=args.one_pdf, order_pages=args.order_pages,
        paper_size=args.paper,
    )
    try:
        summary = conv.run()
    except Exception as exc:                            # noqa: BLE001
        print(tr(f"× 运行失败：{exc}", f"× run failed: {exc}"))
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

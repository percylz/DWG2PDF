#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
i18n.py —— 中英双语（简体中文 / English）

只做三件事，保持极简，避免把界面的复杂度带进来：

1. `system_lang()`：看这台机器 Windows 的界面语言是中文还是别的，
   中文（简/繁/港/新加坡）→ "zh"，其余 → "en"；
2. `set_lang()/get_lang()`：当前使用的语言，'zh' 或 'en'；
3. `t(zh, en)`：取词。两种语言的文案就写在调用处，一眼能对照着改，
   不用维护一套 key→文案的字典（那种做法在 3000 行的出图内核里最容易漏改）。

注意：语言是进程级全局状态。改语言只影响「之后」产生的文字，
已经输出到页面上的日志不会回溯翻译——这是刻意为之，历史日志保留当时的样子。
"""

from __future__ import annotations

import ctypes
import os

SUPPORTED = ("zh", "en")

_LANG = "zh"


def system_lang() -> str:
    """按运行环境猜语言：这台机器看起来是中文环境就用中文，否则英文。

    判断不看单一开关，而是几个信号「有一个是中文就算中文」，因为实测同一台机器
    这几项会互相矛盾（用户界面语言是 en-US、系统的却是 zh-CN、代码页 936）：

    1. 用户界面语言 / 系统界面语言的主语言号 == 0x04（zh-CN / zh-TW / zh-HK / zh-SG）；
    2. 区域名称（GetUserDefaultLocaleName）以 zh 开头；
    3. 非 Unicode 程序用的代码页是 936（GBK）/ 54936（GB18030）/ 950（Big5）。

    这样「中文用户装的英文版 Windows」仍是中文，而外国同事的机器没有任何中文信号 → 英文。
    """
    forced = (os.environ.get("DWG2PDF_LANG") or "").strip().lower()
    if forced in SUPPORTED:                      # 方便自动化测试 / 特殊部署
        return forced
    if _windows_is_chinese():
        return "zh"
    if _windows_signals():
        return "en"
    try:                                          # 非 Windows 或 API 不可用时的兜底
        import locale
        loc = ""
        try:
            loc = locale.getlocale()[0] or ""
        except Exception:
            loc = ""
        return "zh" if loc.lower().startswith("zh") else "en"
    except Exception:
        return "zh"


def _windows_signals() -> bool:
    """Windows 语言 API 能取到值就返回 True（表示判断有效）"""
    return bool(_win_lang_ids() or _win_locale_name() or _win_acp())


def _win_lang_ids() -> list[int]:
    out = []
    try:
        k = ctypes.windll.kernel32
        for fn in ("GetUserDefaultUILanguage", "GetSystemDefaultUILanguage"):
            try:
                out.append(int(getattr(k, fn)()) & 0x3FF)
            except Exception:
                continue
    except Exception:
        return []
    return out


def _win_locale_name() -> str:
    try:
        buf = ctypes.create_unicode_buffer(85)
        if ctypes.windll.kernel32.GetUserDefaultLocaleName(buf, 85):
            return buf.value or ""
    except Exception:
        pass
    return ""


def _win_acp() -> int:
    try:
        return int(ctypes.windll.kernel32.GetACP())
    except Exception:
        return 0


def _windows_is_chinese() -> bool:
    if 0x04 in _win_lang_ids():                  # 0x04 = LANG_CHINESE
        return True
    if _win_locale_name().lower().startswith("zh"):
        return True
    return _win_acp() in (936, 54936, 950)


def normalize(lang: str) -> str:
    """把各种写法收敛成 'zh' / 'en'；空串返回空（表示「跟随系统」）"""
    s = (lang or "").strip().lower()
    if not s:
        return ""
    if s.startswith("zh") or s in ("cn", "chinese", "中文", "chs", "cht"):
        return "zh"
    if s.startswith("en") or s in ("english", "英文"):
        return "en"
    return ""


def resolve(saved: str) -> str:
    """最终该用哪种语言：用户存过的优先，没存过就看系统"""
    return normalize(saved) or system_lang()


def set_lang(lang: str) -> str:
    """切换语言，返回切换后的值。传空 / 'auto' 表示跟随系统。"""
    global _LANG
    picked = normalize(lang) or system_lang()
    _LANG = picked
    return _LANG


def get_lang() -> str:
    return _LANG


def is_en() -> bool:
    return _LANG == "en"


def t(zh: str, en: str) -> str:
    """取词：英文环境返回 en，否则返回 zh"""
    return en if _LANG == "en" else zh

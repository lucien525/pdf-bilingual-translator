# -*- coding: utf-8 -*-
"""字体扫描与自动选择。FONTS_MAP 在 import 时生成。"""

import os
import re
from functools import lru_cache

from core import config

_FONT_EXT = (".otf", ".ttf", ".ttc", ".otc")
_FONT_SCAN_MAX_DEPTH = 8


def scan_fonts():
    fonts = {}
    for root_dir in config.FONT_ROOTS:
        if not root_dir or not os.path.isdir(root_dir):
            continue
        root_dir = os.path.abspath(root_dir)
        for root, dirs, files in os.walk(root_dir):
            depth = root[len(root_dir):].count(os.sep)
            if depth > _FONT_SCAN_MAX_DEPTH:
                dirs[:] = []
                continue
            for f in files:
                if not f.lower().endswith(_FONT_EXT):
                    continue
                full = os.path.join(root, f)
                try:
                    rel = os.path.relpath(full, root_dir)
                except Exception:
                    rel = f
                name = os.path.splitext(rel)[0].replace("\\", "/")
                if name not in fonts:
                    fonts[name] = full
    return fonts


FONTS_MAP = scan_fonts()


def _default_font_display():
    if config.FONT_PATH and os.path.exists(config.FONT_PATH):
        ap = os.path.abspath(config.FONT_PATH)
        for name, p in FONTS_MAP.items():
            if os.path.abspath(p) == ap:
                return name
    if FONTS_MAP:
        return sorted(FONTS_MAP.keys())[0]
    return "(内置宋体)"


_LANG_FONT_PATTERNS = {
    "ar": ["Naskh", "Amiri", "Arabic", "NotoSansArabic", "NotoNaskhArabic"],
    "ko": ["NotoSansKR", "NotoSerifKR", "KR", "Korean", "Hangul"],
    "ja": ["NotoSansJP", "NotoSerifJP", "JP", "Japanese"],
    "ru": ["DejaVu", "NotoSans", "Arial", "Liberation"],
}


@lru_cache(maxsize=2048)
def _pattern_hit(name_lower, pat_lower):
    return re.search(r'(?:^|[^a-z])' + re.escape(pat_lower) + r'(?:[^a-z]|$)',
                     name_lower) is not None


def _auto_pick_font(target_lang, user_font_path):
    if user_font_path and os.path.exists(user_font_path):
        return user_font_path
    patterns = _LANG_FONT_PATTERNS.get(target_lang, [])
    if patterns and FONTS_MAP:
        for name in sorted(FONTS_MAP.keys()):
            low = name.lower()
            for pat in patterns:
                if _pattern_hit(low, pat.lower()):
                    return FONTS_MAP[name]
    return config.FONT_PATH or ""

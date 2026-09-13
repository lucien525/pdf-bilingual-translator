# -*- coding: utf-8 -*-
"""
PDF / Word / PPT 翻译器
"""

import sys
sys.dont_write_bytecode = True

import os
import json
import time
import re
import glob
import base64
import hashlib
import shutil
import signal
import atexit
import threading
import uuid
import io
import socket
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Optional
from html import escape
from functools import lru_cache   # ★ 优化：正则缓存

import pymupdf as fitz
try:
    fitz.TOOLS.mupdf_display_errors(False)
except Exception:
    pass

from PIL import Image
from openai import OpenAI
import gradio as gr
from dotenv import load_dotenv

# ★ 修复：Pillow 9.4+ 弃用 Image.LANCZOS，用兼容写法
_RESAMPLE = getattr(Image, "Resampling", Image).LANCZOS

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    from pptx import Presentation
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
    except Exception:
        MSO_SHAPE_TYPE = None
    HAS_OFFICE = True
except ImportError:
    HAS_OFFICE = False

try:
    import notes_builder as NB
    HAS_NOTES = True
except ImportError:
    NB = None
    HAS_NOTES = False

# ================= .env =================
load_dotenv()
DEFAULT_API_KEY = os.getenv("DEEPSEEK_API_KEY", "").strip()
DEFAULT_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat").strip()
if DEFAULT_MODEL not in ("deepseek-chat", "deepseek-reasoner"):
    DEFAULT_MODEL = "deepseek-chat"

# ================= 字体扫描 =================
_HERE = os.path.dirname(os.path.abspath(__file__))

FONT_ROOTS = [
    r"D:\file\translate\word_type",
    os.path.join(_HERE, "fonts"),
]

_DEFAULT_FONT_CANDIDATES = [
    os.getenv("TRANSLATE_FONT_PATH", "").strip(),
    r"D:\file\translate\word_type\09_SourceHanSerifSC\OTF\SimplifiedChinese\SourceHanSerifSC-Regular.otf",
    os.path.join(_HERE, "fonts", "SourceHanSerifSC-Regular.otf"),
]
FONT_PATH = next((p for p in _DEFAULT_FONT_CANDIDATES
                  if p and os.path.exists(p)), "")

_FONT_EXT = (".otf", ".ttf", ".ttc", ".otc")

# ★ 优化：字体扫描深度从 5 提升到 8
_FONT_SCAN_MAX_DEPTH = 8


def scan_fonts():
    fonts = {}
    for root_dir in FONT_ROOTS:
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
    if FONT_PATH and os.path.exists(FONT_PATH):
        ap = os.path.abspath(FONT_PATH)
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


# ★ 优化：正则结果 lru_cache 缓存，避免每次切换语言都重新编译
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
    return FONT_PATH or ""


# ================= 字号 =================
FONT_SIZE_CHOICES = [
    ("很小 · 8pt", 8.0),
    ("小 · 9pt", 9.0),
    ("中 · 10pt", 10.0),
    ("标准 · 11pt（默认）", 11.0),
    ("大 · 12pt", 12.0),
    ("很大 · 14pt", 14.0),
    ("特大 · 16pt", 16.0),
]
DEFAULT_FONT_SIZE = 11.0

# ================= 双语 PDF 质量预设 =================
PDF_QUALITY_PRESETS = {
    "省流 · 小体积（~1/3）": {
        "zoom": 1.0, "jpeg": 60, "garbage": 3,
        "hint": "文字仍清晰，适合自己看 · 300 页 ≈ 60 MB",
    },
    "标准 · 推荐": {
        "zoom": 1.4, "jpeg": 72, "garbage": 3,
        "hint": "清晰度与体积均衡（默认）· 300 页 ≈ 180 MB",
    },
    "清晰": {
        "zoom": 1.8, "jpeg": 85, "garbage": 2,
        "hint": "文字边缘锐利，适合放平板 · 300 页 ≈ 330 MB",
    },
    "印刷级 · 大体积": {
        "zoom": 2.5, "jpeg": 92, "garbage": 1,
        "hint": "接近无损，可打印 · 300 页 ≈ 720 MB",
    },
}
DEFAULT_PDF_QUALITY = "标准 · 推荐"

# ================= 其他配置 =================
RESULT_ROOT = "trans_result"

RENDER_ZOOM = 2.0
BILINGUAL_ZOOM = 1.4
BILINGUAL_JPEG_QUALITY = 72

PREVIEW_PAGES = 5
PREVIEW_MAX_WIDTH = 1400
PREVIEW_JPEG_QUALITY = 88
PREVIEW_PARAS = 5

CHECKPOINT_EVERY = 10
PAGE_CN_THRESHOLD = 20
PAGE_CN_RATIO = 0.2
VALID_RATIO_THRESHOLD = 0.5

OFFICE_BATCH_SIZE = 20
STATE_SAVE_INTERVAL = 1.5
SIZE_REFRESH_INTERVAL = 5.0

LANG_NAMES = {
    "zh-CN": "简体中文", "zh-TW": "繁体中文",
    "en": "英语", "ja": "日语", "ko": "韩语",
    "fr": "法语", "de": "德语", "es": "西班牙语",
    "pt": "葡萄牙语", "ru": "俄语", "ar": "阿拉伯语", "it": "意大利语",
}

NON_LATIN_SCRIPT_LANGS = ("zh-CN", "zh-TW", "ja", "ko", "ru", "ar")
CJK_LIKE_LANGS = NON_LATIN_SCRIPT_LANGS
RTL_LANGS = ("ar",)

_MARK_RE = re.compile(r'\[\[B(\d+)\]\]')

_IO_LOCK = threading.Lock()
_CREATE_LOCK = threading.Lock()
_LAST_PREVIEW_SIG = {}
_PREVIEW_HTML_CACHE = {}
_DATAURL_CACHE = OrderedDict()
_DATAURL_CACHE_MAX = 200
_SELECTED_STOP_VALUE = None

_STATE_SAVE_TS = {}
# ★ 修复：改为 RLock，避免 Ctrl+C 与正常 save 冲突时死锁
_STATE_SAVE_LOCK = threading.RLock()

# ★ 修复：弹窗只弹一次
_MODAL_SHOWN = set()


# ============================================================
# 提示词
# ============================================================

def build_system_prompt(target_lang, reader_profile="", want_terms=True):
    lang = LANG_NAMES.get(target_lang, "简体中文")
    base = (
        f"你是一位资深文学翻译家，精通多国语言，译笔力求神似而非字对字。"
        f"请把用户发来的内容翻译成【{lang}】，并遵循："
        f"1) 译文必须符合{lang}母语者的阅读习惯和审美，流畅、有文采；"
        f"2) 对话要自然生动，符合人物身份；3) 修辞、隐喻、双关尽量找到{lang}对应表达，"
        f"实在无法对应则意译并保留神韵；4) 不遗漏任何内容，不总结，不输出任何解释。\n\n"
        f"【格式要求】用户会给出一页原文的多个段落，每段以 [[B0]] [[B1]] [[B2]] ... 标记开头。"
        f"你必须严格保留所有标记、保持顺序，标记后紧跟该段译文。"
        f"除标记和译文外，不要输出任何其他文字、不加解释、不用代码块。"
    )
    if HAS_NOTES and NB is not None:
        base += NB.build_terms_instruction(reader_profile, want_terms)
    return base


def build_simple_system(target_lang):
    lang = LANG_NAMES.get(target_lang, "简体中文")
    return (
        f"你是一位资深文学翻译家，精通多国语言，译笔力求神似而非字对字。"
        f"请把用户发来的内容翻译成流畅、有文采的【{lang}】，符合{lang}母语者的阅读习惯。"
        f"保留段落结构，不遗漏内容，不总结，不输出解释，直接给出译文。"
    )


def cache_prefix(target_lang):
    return "" if target_lang == "zh-CN" else f"{target_lang}_"


# ============================================================
# 任务管理
# ============================================================

@dataclass
class TaskState:
    task_id: str
    kind: str
    src_path: str
    src_name: str
    out_dir: str
    work_dir: str
    target_lang: str = "zh-CN"
    created_at: float = field(default_factory=time.time)
    status: str = "queued"
    current: int = 0
    total: int = 0
    label: str = ""
    log: list = field(default_factory=list)
    output_files: list = field(default_factory=list)
    preview_images: list = field(default_factory=list)
    preview_html: str = ""
    error: str = ""
    estimated_size: int = 0
    current_size: int = 0
    last_index: int = 0
    pdf_quality: str = DEFAULT_PDF_QUALITY
    make_bilingual: bool = True
    _preview_scanned: bool = False
    _last_size_ts: float = 0.0
    stop_event: threading.Event = field(default_factory=threading.Event)
    thread: Optional[threading.Thread] = None
    _lock: threading.RLock = field(
        default_factory=threading.RLock, repr=False, compare=False
    )

    def to_dict(self):
        with self._lock:
            log_snap = list(self.log[-60:])
            files_snap = list(self.output_files or [])
            return {
                "task_id": self.task_id, "kind": self.kind,
                "src_name": self.src_name, "out_dir": self.out_dir,
                "target_lang": self.target_lang, "created_at": self.created_at,
                "status": self.status if self.status != "stopping" else "paused",
                "current": self.current, "total": self.total,
                "label": self.label, "log": log_snap,
                "output_files": files_snap, "error": self.error,
                "estimated_size": self.estimated_size,
                "current_size": self.current_size,
                "last_index": self.last_index,
                "pdf_quality": self.pdf_quality,
                "make_bilingual": self.make_bilingual,
            }

    def log_msg(self, m):
        with self._lock:
            self.log.append(m)
            if len(self.log) > 200:
                del self.log[:-200]


class TaskManager:
    def __init__(self):
        self._tasks = {}
        self._lock = threading.Lock()

    def create(self, kind, src_path, src_name, out_dir, work_dir,
               target_lang="zh-CN"):
        tid = uuid.uuid4().hex[:8]
        t = TaskState(task_id=tid, kind=kind, src_path=src_path,
                      src_name=src_name, out_dir=out_dir, work_dir=work_dir,
                      target_lang=target_lang)
        with self._lock:
            self._tasks[tid] = t
        return t

    def get(self, tid):
        with self._lock:
            return self._tasks.get(tid)

    def all_sorted(self):
        with self._lock:
            return sorted(list(self._tasks.values()),
                          key=lambda t: -t.created_at)

    def running(self):
        with self._lock:
            return [t for t in self._tasks.values()
                    if t.status in ("queued", "running", "stopping")]

    def find_active_by_out_dir(self, out_dir):
        target = os.path.normcase(os.path.abspath(out_dir))
        with self._lock:
            for t in self._tasks.values():
                if os.path.normcase(os.path.abspath(t.out_dir)) == target and \
                   t.status in ("queued", "running", "stopping"):
                    return t
        return None

    def load_from_disk(self, state_dict):
        tid = state_dict.get("task_id")
        if not tid:
            return
        with self._lock:
            if tid in self._tasks:
                return
        t = TaskState(
            task_id=tid, kind=state_dict.get("kind", "pdf"),
            src_path="", src_name=state_dict.get("src_name", ""),
            out_dir=state_dict.get("out_dir", ""),
            work_dir=os.path.join(state_dict.get("out_dir", ""), "_work"),
            target_lang=state_dict.get("target_lang", "zh-CN"),
            created_at=state_dict.get("created_at", time.time()),
            status="paused",
            current=state_dict.get("current", 0),
            total=state_dict.get("total", 0),
            label=state_dict.get("label", ""),
            log=state_dict.get("log", []),
            output_files=state_dict.get("output_files", []) or [],
            error=state_dict.get("error", ""),
            estimated_size=state_dict.get("estimated_size", 0),
            current_size=state_dict.get("current_size", 0),
            last_index=state_dict.get("last_index", 0),
            pdf_quality=state_dict.get("pdf_quality", DEFAULT_PDF_QUALITY),
            make_bilingual=state_dict.get("make_bilingual", True),
        )
        with self._lock:
            self._tasks[tid] = t


MANAGER = TaskManager()
SELECTED_TASK_ID = None


def save_state(task, force=False):
    if task is None:
        return
    now = time.monotonic()
    if not force:
        with _STATE_SAVE_LOCK:
            last = _STATE_SAVE_TS.get(task.task_id, 0.0)
            if now - last < STATE_SAVE_INTERVAL:
                return
            _STATE_SAVE_TS[task.task_id] = now
    else:
        with _STATE_SAVE_LOCK:
            _STATE_SAVE_TS[task.task_id] = now

    try:
        path = os.path.join(task.work_dir, "state.json")
        os.makedirs(task.work_dir, exist_ok=True)
        data = task.to_dict()
        with _IO_LOCK:
            tmp = path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(tmp, path)
    except Exception:
        pass


def _add_output(task, *paths_):
    if task is None:
        return
    with task._lock:
        for p in paths_:
            if p and p not in task.output_files:
                task.output_files.append(p)


def scan_all_states():
    root = os.path.abspath(RESULT_ROOT)
    if not os.path.isdir(root):
        return
    for name in os.listdir(root):
        p = os.path.join(root, name, "_work", "state.json")
        if os.path.isfile(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    MANAGER.load_from_disk(json.load(f))
            except Exception:
                pass


def cleanup_orphan_files():
    root = os.path.abspath(RESULT_ROOT)
    if not os.path.isdir(root):
        return 0
    removed = 0
    for name in os.listdir(root):
        d = os.path.join(root, name)
        if not os.path.isdir(d):
            continue
        for f in os.listdir(d):
            fp = os.path.join(d, f)
            try:
                if time.time() - os.path.getmtime(fp) < 86400:
                    continue
            except Exception:
                continue
            if f.endswith(".tmp.pdf") or f.endswith(".bak.pdf"):
                try:
                    os.remove(fp)
                    removed += 1
                    print(f"   [clean] {name}/{f}")
                except Exception:
                    pass
    return removed


# ============================================================
# 路径 / 工具
# ============================================================

def safe_dirname(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name or "").strip()
    name = name.replace("..", "_")
    name = name.rstrip(". ")
    return name[:80] or "untitled"


def prepare_paths(src_path, out_subdir=None):
    base = os.path.basename(src_path)
    book, ext = os.path.splitext(base)
    book = safe_dirname(book)
    if not out_subdir:
        out_subdir = book
    else:
        out_subdir = safe_dirname(out_subdir)
    out_dir = os.path.abspath(os.path.join(RESULT_ROOT, out_subdir))
    work = os.path.join(out_dir, "_work")
    input_dir = os.path.join(work, "input")
    notes_dir = os.path.join(out_dir, f"_{book}")
    os.makedirs(work, exist_ok=True)
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(notes_dir, exist_ok=True)
    return {
        "out_dir": out_dir, "work": work, "input_dir": input_dir,
        "notes_dir": notes_dir,
        "output_pdf":    os.path.join(out_dir, "translated.pdf"),
        "bilingual_pdf": os.path.join(out_dir, "bilingual.pdf"),
        "trans_with_notes_pdf": os.path.join(out_dir, "translated_with_notes.pdf"),
        "cache_file":    os.path.join(work, "translate_cache.json"),
        "output_html":   os.path.join(work, "bilingual.html"),
        "img_dir":       os.path.join(work, "bilingual_pages"),
        "preview_dir":   os.path.join(work, "preview"),
        "progress_file": os.path.join(work, "progress.json"),
        "office_cn":     os.path.join(out_dir, f"{book}_cn{ext.lower()}"),
        "office_bi":     os.path.join(out_dir, f"{book}_bilingual{ext.lower()}"),
    }


def h(s):
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:20]


def load_json_file(path, default):
    if path and os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def save_json_file(path, data):
    if not path:
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with _IO_LOCK:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        os.replace(tmp, path)


def load_progress_file(path):
    d = load_json_file(path, {})
    return set(int(x) for x in d.get("done_pages", []))


def save_progress_file(path, done_pages):
    save_json_file(path, {"done_pages": sorted(int(x) for x in done_pages)})


def clear_dir_preview(folder):
    for pat in ("*.png", "*.jpg", "*.jpeg"):
        try:
            for f in glob.glob(os.path.join(folder, pat)):
                try:
                    os.remove(f)
                except Exception:
                    pass
        except Exception:
            pass


# ============================================================
# 体积预估 / 显示
# ============================================================

def fmt_size(n):
    try:
        n = float(n)
    except Exception:
        return "—"
    if n <= 0:
        return "—"
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} PB"


def _sum_dir_size(folder, skip_subdir="_work"):
    total = 0
    if not folder or not os.path.isdir(folder):
        return total
    for root, dirs, files in os.walk(folder):
        if skip_subdir and skip_subdir in dirs:
            dirs.remove(skip_subdir)
        for f in files:
            try:
                total += os.path.getsize(os.path.join(root, f))
            except Exception:
                pass
    return total


def _task_disk_size(task):
    total = 0
    for f in list(task.output_files or []):
        try:
            if f and os.path.exists(f):
                total += os.path.getsize(f)
        except Exception:
            pass
    if total > 0:
        return total
    return _sum_dir_size(task.out_dir)


def estimate_output_size(kind, src_path, total_pages,
                         target_lang="zh-CN", want_terms=True,
                         pdf_quality=None, make_bilingual=True):
    try:
        if kind == "pdf":
            if target_lang in ("zh-CN", "zh-TW", "ja", "ko"):
                per_page = 26 * 1024
            else:
                per_page = 36 * 1024
            main = int(total_pages * per_page)

            if not make_bilingual:
                bilingual = 0
            else:
                base = 220 * 1024
                preset = PDF_QUALITY_PRESETS.get(
                    pdf_quality or DEFAULT_PDF_QUALITY,
                    PDF_QUALITY_PRESETS[DEFAULT_PDF_QUALITY]
                )
                zoom = preset.get("zoom", 1.4)
                jpeg = preset.get("jpeg", 72)
                scale = (zoom / 1.4) ** 2 * (jpeg / 72)
                bilingual = int(total_pages * base * scale)

            notes = 200 * 1024 if want_terms else 0
            return {
                "main": main,
                "bilingual": bilingual,
                "notes": notes,
                "total": main + bilingual + notes,
            }
        elif kind in ("docx", "pptx"):
            try:
                src_size = os.path.getsize(src_path)
            except Exception:
                src_size = 300 * 1024
            out = int(src_size * 1.2)
            return {
                "main": out,
                "bilingual": out,
                "notes": 0,
                "total": out * 2,
            }
    except Exception:
        pass
    return {"main": 0, "bilingual": 0, "notes": 0, "total": 0}


def refresh_task_size(task, force=False):
    if task is None:
        return
    now = time.time()
    if not force and (now - task._last_size_ts) < SIZE_REFRESH_INTERVAL:
        return
    task._last_size_ts = now
    try:
        task.current_size = _task_disk_size(task)
    except Exception:
        pass


def safe_save_pdf(doc, out_path, retries=5, garbage=3):
    tmp_path = out_path + ".tmp.pdf"
    try:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    except Exception:
        pass

    doc.save(tmp_path, deflate=True, garbage=garbage)

    last_err = None
    for attempt in range(retries):
        try:
            os.replace(tmp_path, out_path)
            return out_path
        except (PermissionError, OSError) as e:
            last_err = e
            time.sleep(0.8 * (attempt + 1))

    fallback = out_path[:-4] + ".bak.pdf"
    try:
        if os.path.exists(fallback):
            try:
                os.remove(fallback)
            except Exception:
                pass
        os.replace(tmp_path, fallback)
        return fallback
    except Exception:
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        raise last_err if last_err else RuntimeError("无法保存 PDF")


def _pix_to_pil(pix):
    try:
        n = pix.n
        w, h_ = pix.width, pix.height
        if n == 3:
            return Image.frombytes("RGB", (w, h_), pix.samples)
        if n == 4:
            return Image.frombytes("RGBA", (w, h_), pix.samples).convert("RGB")
        if n == 1:
            return Image.frombytes("L", (w, h_), pix.samples).convert("RGB")
    except Exception:
        pass
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def img_to_base64_dataurl(path):
    try:
        st = os.stat(path)
        key = (path, st.st_mtime_ns, st.st_size)
    except Exception as e:
        print(f"[preview] stat failed: {path} ({e})")
        return None

    cached = _DATAURL_CACHE.get(key)
    if cached is not None:
        _DATAURL_CACHE.move_to_end(key)
        return cached

    try:
        with open(path, "rb") as f:
            data = f.read()
        if not data:
            print(f"[preview] empty file: {path}")
            return None
        b64 = base64.b64encode(data).decode("ascii")
        low = path.lower()
        mime = "image/png" if low.endswith(".png") else "image/jpeg"
        url = f"data:{mime};base64,{b64}"
        _DATAURL_CACHE[key] = url
        while len(_DATAURL_CACHE) > _DATAURL_CACHE_MAX:
            _DATAURL_CACHE.popitem(last=False)
        return url
    except Exception as e:
        print(f"[preview] base64 failed: {path} ({e})")
        return None


def preview_signature(preview_imgs):
    parts = []
    for p in preview_imgs or []:
        try:
            parts.append(f"{os.path.basename(p)}:{os.path.getmtime(p):.0f}")
        except Exception:
            parts.append(os.path.basename(p) if p else "")
    return tuple(parts)


def build_preview_html(preview_imgs, preview_html=None):
    if preview_html:
        return preview_html

    if not preview_imgs:
        return '''
        <div style="padding:60px 24px;color:#8b8578;text-align:center;font-size:13.5px;
                    background:#fdfcf9;border:1.5px dashed #ebe5d8;border-radius:16px;
                    line-height:1.9">
            <div style="font-size:36px;margin-bottom:12px;opacity:.4">📄</div>
            <div style="color:#5a5a5a;font-weight:500">暂无预览</div>
            <div style="font-size:12px;color:#a9a49a;margin-top:6px">
                翻译启动后会显示前 5 页的左右对照
            </div>
        </div>
        '''

    imgs_html = []
    failed = []
    for p in preview_imgs:
        data_url = img_to_base64_dataurl(p)
        if not data_url:
            failed.append(os.path.basename(p))
            continue
        imgs_html.append(
            f'<img src="{data_url}" '
            f'style="width:100%;display:block;margin:0 0 16px 0;'
            f'box-shadow:0 4px 16px rgba(15,61,62,.12);border-radius:8px;">'
        )

    if not imgs_html:
        failed_str = "、".join(failed[:6]) if failed else "(未知)"
        return f'''
        <div style="padding:40px 24px;color:#a33;text-align:center;font-size:13.5px;
                    background:#fdf4f4;border:1px dashed #e0c0c0;border-radius:16px">
            <div style="font-size:32px;margin-bottom:10px">⚠️</div>
            <div>预览图片加载失败</div>
            <div style="font-size:12px;color:#b88;margin-top:8px">
                失败文件：{escape(failed_str)}
            </div>
        </div>
        '''

    extra = ""
    if failed:
        extra = (
            f'<div style="color:#c33;font-size:11px;text-align:center;padding:6px 0">'
            f'（{len(failed)} 张未能加载：{escape("、".join(failed[:3]))}）</div>'
        )

    return f'''
    <div style="background:#0f1f1f;border-radius:16px;padding:18px;
                max-height:820px;overflow-y:auto;scroll-behavior:smooth;
                box-shadow:inset 0 0 40px rgba(0,0,0,.3)">
      <div style="color:#c9a961;font-size:12px;text-align:center;
                  padding:6px 0 14px 0;letter-spacing:1px;
                  font-weight:500;text-transform:uppercase">
        · 左右对照 · {len(imgs_html)} 页 ·
      </div>
      {''.join(imgs_html)}
      {extra}
      <div style="color:#6b6b6b;font-size:11px;text-align:center;padding:6px 0">
        — 仅显示前 {len(imgs_html)} 页 —
      </div>
    </div>
    '''


def _build_docx_preview_html(pairs):
    if not pairs:
        return ""
    rows = []
    for idx, (src, dst) in enumerate(pairs):
        rows.append(f'''
        <div style="background:#fff;border-radius:12px;padding:20px 22px;margin-bottom:12px;
                    box-shadow:0 2px 12px rgba(15,61,62,.08);
                    border-left:3px solid #c9a961">
          <div style="font-size:11px;color:#c9a961;font-weight:700;
                      letter-spacing:1px;margin-bottom:8px;text-transform:uppercase">
            第 {idx+1} 段 · 原文
          </div>
          <div style="font-size:14px;color:#5a5a5a;line-height:1.8;margin-bottom:16px">{escape(str(src or ""))}</div>
          <div style="font-size:11px;color:#c9a961;font-weight:700;
                      letter-spacing:1px;margin-bottom:8px;text-transform:uppercase">
            第 {idx+1} 段 · 译文
          </div>
          <div style="font-size:14.5px;color:#0f3d3e;line-height:1.9">{escape(str(dst or ""))}</div>
        </div>''')
    return (
        '<div style="background:#f5f3ee;border-radius:16px;padding:16px;'
        'max-height:820px;overflow-y:auto;scroll-behavior:smooth">'
        '<div style="color:#8b8578;font-size:12px;text-align:center;'
        'padding:6px 0 14px 0;letter-spacing:.5px">'
        f'文本对照预览（前 {len(pairs)} 段）</div>'
        + "".join(rows) +
        '<div style="color:#a9a49a;font-size:11px;text-align:center;padding:6px 0">'
        '— 完整结果请下载 Word 查看 —</div></div>'
    )


def _build_pptx_preview_html(pairs):
    if not pairs:
        return ""
    rows = []
    for idx, (page_no, src, dst) in enumerate(pairs):
        rows.append(f'''
        <div style="background:#fff;border-radius:12px;padding:20px 22px;margin-bottom:12px;
                    box-shadow:0 2px 12px rgba(15,61,62,.08);
                    border-left:3px solid #c9a961">
          <div style="font-size:11px;color:#c9a961;font-weight:700;
                      letter-spacing:1px;margin-bottom:8px;text-transform:uppercase">
            第 {page_no} 张 · 原文
          </div>
          <div style="font-size:14px;color:#5a5a5a;line-height:1.8;margin-bottom:16px">{escape(str(src or ""))}</div>
          <div style="font-size:11px;color:#c9a961;font-weight:700;
                      letter-spacing:1px;margin-bottom:8px;text-transform:uppercase">
            第 {page_no} 张 · 译文
          </div>
          <div style="font-size:14.5px;color:#0f3d3e;line-height:1.9">{escape(str(dst or ""))}</div>
        </div>''')
    return (
        '<div style="background:#f5f3ee;border-radius:16px;padding:16px;'
        'max-height:820px;overflow-y:auto;scroll-behavior:smooth">'
        '<div style="color:#8b8578;font-size:12px;text-align:center;'
        'padding:6px 0 14px 0;letter-spacing:.5px">'
        f'文本对照预览（前 {len(pairs)} 段）</div>'
        + "".join(rows) +
        '<div style="color:#a9a49a;font-size:11px;text-align:center;padding:6px 0">'
        '— 完整结果请下载 PPT 查看 —</div></div>'
    )


def make_progress_html(done, total, label="",
                       current_size=0, estimated_size=0, warning=""):
    if total <= 0:
        total = 1
    done = max(0, min(done, total))
    pct = int(done * 100 / total)

    size_line = ""
    if estimated_size > 0 or current_size > 0:
        cur_s = fmt_size(current_size) if current_size else "0 B"
        est_s = fmt_size(estimated_size) if estimated_size else "—"
        size_line = f'''
        <div style="display:flex;justify-content:space-between;
                    margin-top:10px;font-size:11.5px;color:#8b8578;
                    font-family:'SF Mono',Consolas,monospace">
          <span>💾 已落盘 {cur_s}</span>
          <span>预估总产出 ≈ {est_s}</span>
        </div>
        '''

    warning_html = ""
    if warning:
        warning_html = f'''
        <div style="margin-top:12px;padding:9px 13px;
                    background:#fdf1e0;border-left:3px solid #d99a2b;
                    border-radius:6px;font-size:12.5px;color:#8b5a1f;
                    line-height:1.65">
          ⚠️ {escape(warning)}
        </div>
        '''

    return f'''
    <div style="padding:6px 2px">
      <div style="display:flex;justify-content:space-between;align-items:baseline;
                  font-size:13px;color:#5a5a5a;margin-bottom:10px;gap:10px">
        <span style="font-weight:500;color:#0f3d3e;
                     overflow:hidden;text-overflow:ellipsis;white-space:nowrap">
          {escape(label or "等待开始")}
        </span>
        <span style="font-family:'SF Mono',Consolas,monospace;font-size:12.5px;
                     color:#0f3d3e;font-weight:600;flex-shrink:0">
          {done} / {total}
        </span>
      </div>
      <div style="height:8px;background:#ebe5d8;border-radius:999px;overflow:hidden;
                  position:relative">
        <div style="height:100%;width:{pct}%;
                    background:linear-gradient(90deg,#0f3d3e,#1f5b5c 45%,#c9a961);
                    border-radius:999px;
                    transition:width .45s cubic-bezier(.4,0,.2,1);
                    box-shadow:0 0 12px rgba(201,169,97,.45);
                    position:relative"></div>
      </div>
      <div style="text-align:right;font-size:11px;color:#a9a49a;
                  font-family:'SF Mono',Consolas,monospace;margin-top:4px">
        {pct}%
      </div>
      {size_line}
      {warning_html}
    </div>
    '''


def build_done_modal_html(task):
    if task is None or task.status != "done":
        return ""

    files_html = ""
    for f in list(task.output_files or []):
        if f and os.path.exists(f):
            name = os.path.basename(f)
            try:
                sz = os.path.getsize(f)
                if sz < 1024 * 100:
                    size_str = f"{sz / 1024:.1f} KB"
                else:
                    size_str = f"{sz / 1024 / 1024:.2f} MB"
            except Exception:
                size_str = ""
            low = name.lower()
            if low.endswith(".pdf"):
                icon = "📕"
            elif low.endswith(".docx"):
                icon = "📘"
            elif low.endswith(".pptx"):
                icon = "📊"
            elif low.endswith(".csv"):
                icon = "📊"
            elif low.endswith(".md"):
                icon = "📝"
            elif low.endswith(".html"):
                icon = "🌐"
            else:
                icon = "📄"
            files_html += f'''
            <div style="display:flex;align-items:center;gap:12px;
                        padding:11px 14px;background:#f9f7f1;
                        border-radius:10px;margin-bottom:8px">
              <span style="font-size:18px;flex-shrink:0">{icon}</span>
              <span style="flex:1;font-size:13px;color:#1f2937;
                           overflow:hidden;text-overflow:ellipsis;
                           white-space:nowrap">{escape(name)}</span>
              <span style="font-size:11.5px;color:#8b8578;
                           font-family:'SF Mono',Consolas,monospace;
                           flex-shrink:0">{size_str}</span>
            </div>'''

    if not files_html:
        files_html = (
            '<div style="padding:14px;color:#8b8578;font-size:13px;'
            'text-align:center">（暂无输出文件）</div>'
        )

    lang_label = LANG_NAMES.get(getattr(task, "target_lang", "zh-CN"), "简体中文")
    actual_size = task.current_size or 0
    est_size = task.estimated_size or 0

    size_pill = ""
    if actual_size > 0:
        if est_size > 0:
            diff_pct = (actual_size - est_size) / est_size * 100
            if abs(diff_pct) < 15:
                color = "#0f3d3e"
                bg = "#e8f0ef"
            else:
                color = "#8b5a1f"
                bg = "#fdf1e0"
            size_pill = (
                f'<span style="font-size:12px;color:{color};background:{bg};'
                f'padding:5px 14px;border-radius:999px;font-weight:500">'
                f'📦 {fmt_size(actual_size)}</span>'
            )
        else:
            size_pill = (
                f'<span style="font-size:12px;color:#0f3d3e;background:#e8f0ef;'
                f'padding:5px 14px;border-radius:999px;font-weight:500">'
                f'📦 {fmt_size(actual_size)}</span>'
            )

    return f'''
    <div id="done_modal_overlay"
         onclick="if(event.target===this){{this.style.display='none'}}"
         style="position:fixed;inset:0;background:rgba(10,20,20,.55);
                backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);
                z-index:99999;display:flex;align-items:center;
                justify-content:center;padding:24px;
                animation:doneFadeIn .25s ease">
      <style>
        @keyframes doneFadeIn {{
          from {{opacity:0}}
          to {{opacity:1}}
        }}
        @keyframes donePopIn {{
          from {{opacity:0;transform:scale(.92) translateY(10px)}}
          to {{opacity:1;transform:scale(1) translateY(0)}}
        }}
        #done_modal_card {{
          animation:donePopIn .35s cubic-bezier(.2,1.2,.4,1);
        }}
        #done_modal_close_btn:hover {{
          transform:translateY(-1px);
          box-shadow:0 6px 20px rgba(15,61,62,.4) !important;
        }}
        #done_modal_close_btn:active {{
          transform:translateY(0);
        }}
      </style>
      <div id="done_modal_card"
           style="background:#fff;max-width:540px;width:100%;
                  border-radius:24px;padding:36px 32px 28px;
                  box-shadow:0 24px 72px rgba(0,0,0,.4);
                  position:relative;max-height:90vh;overflow-y:auto">
        <div style="text-align:center;margin-bottom:24px">
          <div style="width:76px;height:76px;border-radius:50%;
                      background:linear-gradient(135deg,#0f3d3e,#1f5b5c);
                      display:inline-flex;align-items:center;justify-content:center;
                      font-size:38px;margin-bottom:16px;
                      box-shadow:0 8px 28px rgba(15,61,62,.35);
                      position:relative">
            🎉
            <div style="position:absolute;inset:-4px;border-radius:50%;
                        border:2px solid #c9a961;opacity:.4"></div>
          </div>
          <div style="font-size:25px;font-weight:700;color:#0f3d3e;
                      font-family:'Noto Serif SC',Georgia,serif;
                      margin-bottom:8px;letter-spacing:.5px">
            翻译完成
          </div>
          <div style="font-size:13.5px;color:#8b8578;
                      overflow:hidden;text-overflow:ellipsis;
                      white-space:nowrap;padding:0 20px">
            {escape(task.src_name)}
          </div>
        </div>

        <div style="display:flex;gap:8px;margin-bottom:22px;
                    justify-content:center;flex-wrap:wrap">
          <span style="font-size:12px;color:#0f3d3e;background:#e8f0ef;
                       padding:5px 14px;border-radius:999px;font-weight:500">
            🌐 {lang_label}
          </span>
          <span style="font-size:12px;color:#0f3d3e;background:#e8f0ef;
                       padding:5px 14px;border-radius:999px;font-weight:500">
            ✅ {task.total} 页
          </span>
          <span style="font-size:12px;color:#0f3d3e;background:#e8f0ef;
                       padding:5px 14px;border-radius:999px;font-weight:500">
            💾 {len(task.output_files)} 个文件
          </span>
          {size_pill}
        </div>

        <div style="max-height:240px;overflow-y:auto;margin-bottom:18px;
                    padding-right:2px">
          {files_html}
        </div>

        <div style="font-size:11.5px;color:#8b8578;background:#f9f7f1;
                    padding:10px 14px;border-radius:10px;margin-bottom:20px;
                    word-break:break-all;line-height:1.6;
                    font-family:'SF Mono',Consolas,monospace">
          📁 {escape(task.out_dir)}
        </div>

        <button id="done_modal_close_btn"
                onclick="document.getElementById('done_modal_overlay').style.display='none';event.stopPropagation();"
                style="width:100%;padding:15px;
                       background:linear-gradient(135deg,#0f3d3e,#1f5b5c);
                       color:#fff;border:none;border-radius:12px;font-size:15px;
                       font-weight:600;cursor:pointer;
                       box-shadow:0 4px 14px rgba(15,61,62,.3);
                       transition:transform .12s,box-shadow .18s;
                       font-family:inherit;letter-spacing:.5px">
          知道了
        </button>
      </div>
    </div>
    '''


def parse_marked(text, _n=None):
    result = {}
    if not text:
        return result
    if HAS_NOTES and NB is not None:
        idx = text.rfind(NB.TERM_MARK)
        if idx >= 0:
            text = text[:idx]
    matches = list(_MARK_RE.finditer(text))
    for i, m in enumerate(matches):
        idx = int(m.group(1))
        if idx in result:
            continue
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        result[idx] = text[start:end].strip()
    return result


# ============================================================
# 内容校验
# ============================================================

def _has_target_chars(text, target_lang):
    if not text:
        return False
    if target_lang in ("zh-CN", "zh-TW"):
        return any('\u4e00' <= c <= '\u9fff' for c in text)
    elif target_lang == "ja":
        return any(
            ('\u3040' <= c <= '\u30ff') or ('\u4e00' <= c <= '\u9fff')
            for c in text
        )
    elif target_lang == "ko":
        return any('\uac00' <= c <= '\ud7a3' for c in text)
    elif target_lang == "ru":
        return any('\u0400' <= c <= '\u04ff' for c in text)
    elif target_lang == "ar":
        return any('\u0600' <= c <= '\u06ff' for c in text)
    return True


def page_is_translated(page, target_lang="zh-CN", src_page=None):
    if target_lang not in NON_LATIN_SCRIPT_LANGS:
        return True
    try:
        text = page.get_text()
    except Exception:
        return False

    # ★ 修复：空页判断更温和——原文也空则视为已处理
    if not text or not text.strip():
        if src_page is not None:
            try:
                if not src_page.get_text().strip():
                    return True
            except Exception:
                pass
        return False

    if src_page is None:
        return True

    try:
        src_text = src_page.get_text()
        if src_text.strip() == text.strip():
            return False
    except Exception:
        pass

    total_nonspace = sum(1 for c in text if not c.isspace())
    # ★ 修复：短页（标题页/扉页）不要误判为未翻译
    if total_nonspace < 5:
        if src_page is not None:
            try:
                src_ns = sum(
                    1 for c in src_page.get_text() if not c.isspace())
            except Exception:
                src_ns = 0
            if src_ns < 5:
                return True
        return False

    if target_lang in ("zh-CN", "zh-TW"):
        rng = ('\u4e00', '\u9fff')
    elif target_lang == "ja":
        cnt_kana = 0
        for c in text:
            if '\u3040' <= c <= '\u30ff':
                cnt_kana += 1
        if cnt_kana >= max(3, PAGE_CN_THRESHOLD // 4):
            return True
        rng = ('\u4e00', '\u9fff')
    elif target_lang == "ko":
        rng = ('\uac00', '\ud7a3')
    elif target_lang == "ru":
        rng = ('\u0400', '\u04ff')
    elif target_lang == "ar":
        rng = ('\u0600', '\u06ff')
    else:
        return True

    lo, hi = rng
    cnt = 0
    for c in text:
        if lo <= c <= hi:
            cnt += 1

    need_chars = max(3, min(PAGE_CN_THRESHOLD, total_nonspace // 3))
    if cnt < need_chars:
        return False
    return (cnt / total_nonspace) >= PAGE_CN_RATIO


def translations_look_valid(translations, blocks, target_lang="zh-CN"):
    if not translations or not blocks:
        return False
    valid = 0
    for i, b in enumerate(blocks):
        v = (translations.get(i) or "").strip()
        src = (b[4] or "").strip()
        if not v:
            continue
        if v.lower() == src.lower():
            continue
        if target_lang in NON_LATIN_SCRIPT_LANGS:
            if not _has_target_chars(v, target_lang) and len(v) > 5:
                continue
        valid += 1
    need = max(1, int(len(blocks) * VALID_RATIO_THRESHOLD))
    return valid >= need


# ============================================================
# API
# ============================================================

def _extract_status(err):
    for attr in ("status_code", "code"):
        v = getattr(err, attr, None)
        if isinstance(v, int):
            return v
    resp = getattr(err, "response", None)
    if resp is not None:
        v = getattr(resp, "status_code", None)
        if isinstance(v, int):
            return v
    return None


def _is_fatal_api_error(msg):
    if not msg:
        return False
    m = str(msg)
    low = m.lower()
    return ("余额" in m
            or "API Key 无效" in m
            or "无效或已过期" in m
            or "insufficient" in low
            or "unauthorized" in low)


def _interruptible_sleep(seconds, stop_event):
    if stop_event is None:
        time.sleep(seconds)
        return False
    end = time.time() + seconds
    while True:
        remaining = end - time.time()
        if remaining <= 0:
            return False
        if stop_event.is_set():
            return True
        time.sleep(min(0.25, remaining))


def _api_call(client, model, system, user_content, retries=4, stop_event=None):
    last_err = ""
    for attempt in range(retries):
        if stop_event is not None and stop_event.is_set():
            raise RuntimeError("用户已停止")
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user_content},
                ],
                temperature=0.3,
            )
            return resp.choices[0].message.content or ""
        except Exception as e:
            last_err = str(e)
            status = _extract_status(e)
            low = last_err.lower()
            if status == 402 or "insufficient" in low or "余额" in last_err:
                raise RuntimeError("账户余额不足，请去 DeepSeek 平台充值后再继续")
            if status == 401:
                raise RuntimeError("API Key 无效或已过期，请检查后重试")
            if status == 429:
                _interruptible_sleep(10 * (attempt + 1), stop_event)
                continue
            _interruptible_sleep(6 * (attempt + 1), stop_event)

    raise RuntimeError(f"API 连续失败：{last_err}")


def call_api(client, model, text, target_lang="zh-CN", retries=4,
             stop_event=None, reader_profile="", want_terms=True):
    system = build_system_prompt(target_lang, reader_profile, want_terms)
    return _api_call(client, model, system, text, retries, stop_event=stop_event)


def call_simple_api(client, model, text, target_lang="zh-CN", retries=4,
                    stop_event=None):
    system = build_simple_system(target_lang)
    return _api_call(client, model, system, text, retries, stop_event=stop_event)


# ============================================================
# PDF 字体与排版
# ============================================================

def _resolve_font(font_path):
    if font_path and os.path.exists(font_path):
        return font_path, "user_font"
    return None, "china-s"


def _pdf_font_setup(page, font_path=None):
    fp, name = _resolve_font(font_path)
    if fp:
        try:
            page.insert_font(fontname=name, fontfile=fp, set_simple=False)
            return name
        except Exception:
            pass
    return "china-s"


def estimate_lines(width, text, fs):
    total = 0
    for para in text.split("\n"):
        if not para:
            total += 1
            continue
        w = 0
        lines = 1
        for ch in para:
            cw = fs if ord(ch) > 0x2E80 else fs * 0.55
            if w + cw > width:
                lines += 1
                w = cw
            else:
                w += cw
        total += lines
    return total


def find_fontsize(rect, text, max_size=11.0):
    fs = float(max_size)
    min_fs = 4.5
    while fs >= min_fs:
        if estimate_lines(rect.width, text, fs) * fs * 1.35 <= rect.height + 2:
            return fs
        fs -= 0.5
    return min_fs


def _clean_block_text(t):
    return re.sub(r'[ \t]*\n[ \t]*', ' ', t or "").strip()


def apply_translations(page, blocks, translations,
                       font_path=None, max_font_size=11.0,
                       target_lang="zh-CN"):
    items = []
    for i, b in enumerate(blocks):
        text = (translations.get(i) or "").strip()
        if not text:
            continue
        try:
            rect = fitz.Rect(b[0], b[1], b[2], b[3])
        except Exception:
            continue
        if rect.is_empty or rect.width <= 0 or rect.height <= 0:
            continue
        items.append((rect, text))

    if not items:
        return 0, 0

    for rect, _ in items:
        try:
            page.add_redact_annot(rect)
        except Exception:
            pass

    try:
        page.apply_redactions(
            images=fitz.PDF_REDACT_IMAGE_NONE,
            graphics=fitz.PDF_REDACT_LINE_ART_NONE,
            text=fitz.PDF_REDACT_TEXT_REMOVE,
        )
    except Exception:
        try:
            page.apply_redactions()
        except Exception:
            pass

    fontname = _pdf_font_setup(page, font_path)
    align = 2 if target_lang in RTL_LANGS else 0
    written = 0
    failed = 0

    for rect, text in items:
        fs = find_fontsize(rect, text, max_font_size)

        def _try(r, size):
            try:
                return page.insert_textbox(
                    r, text, fontname=fontname, fontsize=size,
                    color=(0, 0, 0), align=align, overlay=True,
                )
            except Exception:
                return -1

        rc = _try(rect, fs)
        while rc < 0 and fs > 4.5:
            fs -= 0.5
            rc = _try(rect, fs)

        if rc < 0:
            ext = fitz.Rect(
                rect.x0, rect.y0,
                rect.x1,
                max(rect.y1, page.rect.y1 - 20),
            )
            rc = _try(ext, max(fs, 5.0))

        if rc < 0:
            try:
                page.insert_text(
                    fitz.Point(rect.x0, rect.y1 - 2),
                    text, fontname=fontname, fontsize=max(fs, 5.0),
                    color=(0, 0, 0), overlay=True,
                )
            except Exception:
                failed += 1
                continue

        written += len(text)

    return written, failed


def translate_page(client, model, blocks, cache, cache_file,
                   target_lang="zh-CN", stop_event=None,
                   reader_profile="", want_terms=True):
    pref = cache_prefix(target_lang)
    cleaned = [_clean_block_text(b[4]) for b in blocks]
    marked = "\n\n".join(f"[[B{i}]] {t}" for i, t in enumerate(cleaned))

    terms_suffix = "|T" if (want_terms and HAS_NOTES) else ""
    key = f"pg_{pref}" + h(marked + terms_suffix)

    if key in cache:
        try:
            entry = cache[key]
            if isinstance(entry, dict) and "paragraphs" in entry:
                cached = {int(k): v for k, v in entry["paragraphs"].items()}
                cached_terms = entry.get("terms", [])
            else:
                cached = {int(k): v for k, v in entry.items()}
                cached_terms = []
            cached_complete = all(
                (cached.get(i) or "").strip() for i in range(len(blocks))
            )
            if cached_complete and translations_look_valid(
                    cached, blocks, target_lang):
                return cached, cached_terms
        except Exception:
            pass

    raw = call_api(client, model, marked, target_lang,
                   stop_event=stop_event,
                   reader_profile=reader_profile,
                   want_terms=(want_terms and HAS_NOTES))

    if HAS_NOTES and NB is not None:
        body, terms_raw = NB.split_translation_and_terms(raw)
    else:
        body, terms_raw = raw, ""

    parsed = parse_marked(body, len(blocks))
    terms = NB.parse_terms_block(terms_raw) \
        if (HAS_NOTES and NB is not None and want_terms) else []

    missing = [i for i in range(len(blocks))
               if i not in parsed or not parsed[i].strip()]

    if missing and not (stop_event is not None and stop_event.is_set()):
        try:
            missing_texts = [cleaned[i] for i in missing]
            batch_marked = "\n\n".join(
                f"[[B{j}]] {t}" for j, t in enumerate(missing_texts)
            )
            r = call_api(client, model, batch_marked, target_lang,
                         stop_event=stop_event,
                         reader_profile=reader_profile, want_terms=False)
            if HAS_NOTES and NB is not None:
                body2, _ = NB.split_translation_and_terms(r)
            else:
                body2 = r
            sub = parse_marked(body2, len(missing))
            for j, i in enumerate(missing):
                v = (sub.get(j) or "").strip()
                if v:
                    parsed[i] = v
        except RuntimeError as e:
            if _is_fatal_api_error(str(e)):
                raise
        except Exception:
            pass

        still_missing = [i for i in missing
                         if not (parsed.get(i) or "").strip()]
        for i in still_missing:
            if stop_event is not None and stop_event.is_set():
                break
            bk = f"bk_{pref}" + h(cleaned[i])
            if bk in cache and isinstance(cache[bk], str) and cache[bk].strip():
                parsed[i] = cache[bk]
                continue
            tr = ""
            try:
                r = call_api(client, model, f"[[B0]] {cleaned[i]}",
                             target_lang, stop_event=stop_event,
                             reader_profile=reader_profile, want_terms=False)
                if HAS_NOTES and NB is not None:
                    body2, _ = NB.split_translation_and_terms(r)
                else:
                    body2 = r
                sub = parse_marked(body2, 1)
                tr = (sub.get(0) or "").strip()
            except RuntimeError as e:
                if _is_fatal_api_error(str(e)):
                    raise
                tr = ""
            except Exception:
                tr = ""
            parsed[i] = tr
            if tr:
                cache[bk] = tr
                save_json_file(cache_file, cache)

    parsed_complete = all(
        (parsed.get(i) or "").strip() for i in range(len(blocks))
    )
    if parsed_complete and translations_look_valid(parsed, blocks, target_lang):
        cache[key] = {
            "paragraphs": {str(k): v for k, v in parsed.items()},
            "terms": terms,
        }
        save_json_file(cache_file, cache)

    return parsed, terms


def render_preview_only(trans_path, paths, task, n_preview=PREVIEW_PAGES):
    clear_dir_preview(paths["preview_dir"])
    os.makedirs(paths["preview_dir"], exist_ok=True)

    if not os.path.exists(trans_path):
        return []

    try:
        with open(trans_path, "rb") as f:
            data = f.read()
        trans = fitz.open(stream=data, filetype="pdf")
    except Exception as e:
        print(f"[preview] open trans failed: {e}")
        return []

    try:
        orig = fitz.open(task.src_path)
    except Exception as e:
        print(f"[preview] open src failed: {e}")
        trans.close()
        return []

    n = min(n_preview, len(orig), len(trans))
    preview_imgs = []
    zoom_mat = fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM)
    for i in range(n):
        if task.stop_event.is_set():
            break
        try:
            o_pix = orig[i].get_pixmap(matrix=zoom_mat)
            t_pix = trans[i].get_pixmap(matrix=zoom_mat)
            o = _pix_to_pil(o_pix)
            t = _pix_to_pil(t_pix)
            hh = max(o.height, t.height)
            if o.height != hh:
                o = o.resize((int(o.width * hh / o.height), hh), _RESAMPLE)
            if t.height != hh:
                t = t.resize((int(t.width * hh / t.height), hh), _RESAMPLE)
            gap = 8
            canvas = Image.new("RGB", (o.width + gap + t.width, hh), (40, 40, 40))
            canvas.paste(o, (0, 0))
            canvas.paste(t, (o.width + gap, 0))
            if canvas.width > PREVIEW_MAX_WIDTH:
                ratio = PREVIEW_MAX_WIDTH / canvas.width
                canvas = canvas.resize(
                    (PREVIEW_MAX_WIDTH, int(canvas.height * ratio)),
                    _RESAMPLE,
                )
            p = os.path.join(paths["preview_dir"], f"compare_{i:04d}.jpg")
            canvas.save(p, "JPEG", quality=PREVIEW_JPEG_QUALITY, optimize=True)
            preview_imgs.append(p)
        except Exception as e:
            print(f"[preview] page {i} render failed: {e}")
            continue
    orig.close()
    trans.close()
    return preview_imgs


def _collect_existing_previews(task):
    if task is None:
        return []
    work = getattr(task, "work_dir", "") or ""
    if not work:
        return []
    d = os.path.join(work, "preview")
    if not os.path.isdir(d):
        return []
    try:
        files = sorted(glob.glob(os.path.join(d, "compare_*.jpg")))
    except Exception:
        return []
    return [f for f in files
            if os.path.exists(f) and os.path.getsize(f) > 0]


def make_bilingual_pdf(trans_path, paths, task,
                       n_pages=None, done_pages=None,
                       zoom=None, jpeg_quality=None, garbage=3):
    if not os.path.exists(trans_path):
        task.log_msg(f"⚠️ 双语 PDF 源不存在：{trans_path}")
        return
    try:
        with open(trans_path, "rb") as f:
            data = f.read()
        trans = fitz.open(stream=data, filetype="pdf")
    except Exception as e:
        task.log_msg(f"⚠️ 双语 PDF 打开源失败：{e}")
        return

    try:
        orig = fitz.open(task.src_path)
        total = min(len(orig), len(trans))

        if done_pages:
            pages = sorted(p for p in done_pages if 1 <= p <= total)
        else:
            cap = total if n_pages is None else min(n_pages, total)
            pages = list(range(1, cap + 1))

        if not pages:
            task.log_msg("⚠️ 双语 PDF 无有效页")
            orig.close()
            trans.close()
            return

        out_doc = fitz.open()
        done_count = 0
        n_total = len(pages)

        zoom_val = float(zoom) if zoom else BILINGUAL_ZOOM
        jpeg_val = int(jpeg_quality) if jpeg_quality else BILINGUAL_JPEG_QUALITY
        zoom_mat = fitz.Matrix(zoom_val, zoom_val)

        for idx, pno in enumerate(pages):
            if task.stop_event.is_set():
                task.log_msg(
                    f"⏸ 双语 PDF 生成时收到停止信号，已生成 {done_count} 页")
                break
            i = pno - 1
            try:
                o_pix = orig[i].get_pixmap(matrix=zoom_mat)
                t_pix = trans[i].get_pixmap(matrix=zoom_mat)
                o = _pix_to_pil(o_pix)
                t = _pix_to_pil(t_pix)

                hh = max(o.height, t.height)
                if o.height != hh:
                    o = o.resize((int(o.width * hh / o.height), hh),
                                 _RESAMPLE)
                if t.height != hh:
                    t = t.resize((int(t.width * hh / t.height), hh),
                                 _RESAMPLE)

                gap = 10
                canvas = Image.new(
                    "RGB", (o.width + gap + t.width, hh), (255, 255, 255))
                canvas.paste(o, (0, 0))
                canvas.paste(t, (o.width + gap, 0))

                buf = io.BytesIO()
                canvas.save(buf, "JPEG",
                            quality=jpeg_val, optimize=False)
                img_bytes = buf.getvalue()
                buf.close()

                w, ph = canvas.size
                page = out_doc.new_page(width=w, height=ph)
                page.insert_image(fitz.Rect(0, 0, w, ph), stream=img_bytes)

                canvas.close()
                o.close()
                t.close()
                done_count += 1

                if done_count % 10 == 0 or done_count == n_total:
                    task.log_msg(f"📐 双语 PDF {done_count}/{n_total} 页")
            except Exception as e:
                task.log_msg(f"⚠️ 双语 PDF 第 {pno} 页失败：{e}")
                continue

        orig.close()
        trans.close()

        if done_count > 0:
            try:
                saved = safe_save_pdf(out_doc, paths["bilingual_pdf"],
                                      garbage=garbage)
                if saved != paths["bilingual_pdf"]:
                    task.log_msg(
                        f"⚠️ 双语 PDF 保存到备用路径：{os.path.basename(saved)}"
                    )
                    paths["bilingual_pdf"] = saved
                task.log_msg(
                    f"✅ 双语 PDF 生成完成：{done_count} 页 → "
                    f"{os.path.basename(paths['bilingual_pdf'])}"
                )
            except Exception as e:
                task.log_msg(f"⚠️ 双语 PDF 保存失败：{e}")
        else:
            task.log_msg("⚠️ 双语 PDF 生成 0 页")
        out_doc.close()
    except Exception as e:
        import traceback
        task.log_msg(f"⚠️ 双语 PDF 生成异常：{e}")
        task.log_msg(traceback.format_exc()[:500])


# ============================================================
# 术语页 PDF
# ============================================================

def build_notes_pdf(global_terms, book_title, reader_profile="",
                    total_pages=0, lang_label="", font_path=None):
    items = list(global_terms.get("terms", {}).values())
    if not items:
        return None
    items.sort(key=lambda x: (x.get("first_page", 999999), x.get("term", "")))

    A4_W, A4_H = 595, 842
    MARGIN = 44
    BOTTOM = A4_H - MARGIN

    doc = fitz.open()
    state = {"page": None, "fn": None, "y": MARGIN}

    def new_page():
        p = doc.new_page(width=A4_W, height=A4_H)
        fn = _pdf_font_setup(p, font_path)
        state["page"] = p
        state["fn"] = fn
        state["y"] = MARGIN

    new_page()

    def write(text, size, color=(0, 0, 0), gap=4, indent=0):
        page = state["page"]
        fn = state["fn"]
        y = state["y"]

        rect = fitz.Rect(MARGIN + indent, y, A4_W - MARGIN, BOTTOM)
        rc = -1
        try:
            rc = page.insert_textbox(
                rect, text, fontname=fn, fontsize=size,
                color=color, align=0, overlay=True,
            )
        except Exception:
            rc = -1

        if rc < 0:
            new_page()
            page = state["page"]
            fn = state["fn"]
            y = MARGIN
            rect = fitz.Rect(MARGIN + indent, y, A4_W - MARGIN, BOTTOM)
            try:
                rc = page.insert_textbox(
                    rect, text, fontname=fn, fontsize=size,
                    color=color, align=0, overlay=True,
                )
            except Exception:
                rc = -1

        if rc < 0:
            try:
                page.insert_text(
                    fitz.Point(MARGIN + indent, y + size),
                    text, fontname=fn, fontsize=size,
                    color=color, overlay=True,
                )
            except Exception:
                pass
            state["y"] = y + size + gap
            return

        used = rect.height - rc
        state["y"] = y + used + gap

    write(f"《{book_title}》阅读笔记", 20, color=(0.08, 0.08, 0.08), gap=14)

    meta = f"术语 {len(items)} 个"
    if total_pages:
        meta += f"    |    全书 {total_pages} 页"
    if lang_label:
        meta += f"    |    译文 {lang_label}"
    write(meta, 9.5, color=(0.42, 0.42, 0.42), gap=4)

    if reader_profile:
        write(f"读者背景：{reader_profile}", 9.5,
              color=(0.42, 0.42, 0.42), gap=4)

    write("注：本笔记由 AI 在翻译过程中同步生成，仅供参考。"
          "关键概念请务必核对原书定义。",
          9, color=(0.65, 0.25, 0.25), gap=18)

    by_chapter = {}
    for t in items:
        ch = t.get("chapter") or "术语总表（未识别章节）"
        by_chapter.setdefault(ch, []).append(t)

    for ch, ch_items in by_chapter.items():
        write(f"§ {ch}", 14, color=(0.42, 0.36, 0.2), gap=12)

        for t in ch_items:
            term = t.get("term", "")
            trans = t.get("translation", "")
            note = t.get("note", "")
            first = t.get("first_page", "?")
            cnt = t.get("count", 1)

            write(f"· {term}  —  {trans}", 11,
                  color=(0.08, 0.08, 0.08), gap=2)
            write(f"    首次出现 p.{first} · 全书出现 {cnt} 次", 8.5,
                  color=(0.55, 0.55, 0.55), gap=4)

            if note:
                write(f"    {note}", 10, color=(0.28, 0.28, 0.28), gap=6)
            else:
                state["y"] += 2

    state["y"] += 10
    write("—— 术语表结束 · 正文自此开始 ——", 11,
          color=(0.55, 0.42, 0.2), gap=8)

    return doc


def insert_notes_into_pdf(pdf_path, global_terms, book_title,
                          reader_profile="", total_pages=0, lang_label="",
                          font_path=None, position="front"):
    if not pdf_path or not os.path.exists(pdf_path):
        return False
    if not global_terms or not global_terms.get("terms"):
        return False

    notes_doc = build_notes_pdf(
        global_terms, book_title,
        reader_profile=reader_profile,
        total_pages=total_pages,
        lang_label=lang_label,
        font_path=font_path,
    )
    if notes_doc is None or len(notes_doc) == 0:
        if notes_doc:
            notes_doc.close()
        return False

    notes_pages = len(notes_doc)
    main_doc = None
    try:
        with open(pdf_path, "rb") as f:
            data = f.read()
        main_doc = fitz.open(stream=data, filetype="pdf")

        orig_toc = []
        try:
            orig_toc = main_doc.get_toc(simple=True) or []
        except Exception:
            pass

        # ★ 修复：显式指定 from_page / to_page，避免默认 -1 语义歧义
        if position == "front":
            main_doc.insert_pdf(notes_doc,
                                from_page=0,
                                to_page=len(notes_doc) - 1,
                                start_at=0)
        else:
            main_doc.insert_pdf(notes_doc)

        try:
            new_toc = []
            if position == "front":
                new_toc.append([1, "📖 术语表", 1])
                for item in orig_toc:
                    if not isinstance(item, (list, tuple)) or len(item) < 3:
                        continue
                    level, title, page = item[0], item[1], item[2]
                    try:
                        page = int(page) + notes_pages
                    except Exception:
                        continue
                    if page < 1:
                        continue
                    new_toc.append([level, title, page])
            else:
                for item in orig_toc:
                    if isinstance(item, (list, tuple)) and len(item) >= 3:
                        new_toc.append(list(item))
                new_toc.append([1, "📖 术语表",
                                max(1, len(main_doc) - notes_pages + 1)])

            if new_toc:
                main_doc.set_toc(new_toc)
        except Exception:
            pass

        tmp = pdf_path + ".tmp.pdf"
        main_doc.save(tmp, deflate=True, garbage=3)

        last_err = None
        for attempt in range(5):
            try:
                os.replace(tmp, pdf_path)
                return True
            except (PermissionError, OSError) as e:
                last_err = e
                time.sleep(0.8 * (attempt + 1))
        try:
            os.remove(tmp)
        except Exception:
            pass
        print(f"插入术语页失败（无法替换文件）：{last_err}")
        return False
    except Exception as e:
        print(f"插入术语页失败：{e}")
        return False
    finally:
        try:
            if main_doc is not None:
                main_doc.close()
        except Exception:
            pass
        try:
            notes_doc.close()
        except Exception:
            pass


# ============================================================
# Office
# ============================================================

def _docx_replace_para_text(para, new_text):
    """
    替换段落文本。

    ★ 说明：多 run 场景下，用第一个 run 的字符格式承载全部译文，
    其余 run 清空。这是 python-docx 下最稳妥的策略——按比例拆
    分到各 run 反而会因为译文长度变化导致格式错乱。
    """
    if para.runs:
        para.runs[0].text = new_text
        for r in para.runs[1:]:
            r.text = ""
    else:
        para.add_run(new_text)


def _docx_insert_after(para, text):
    new_p = OxmlElement("w:p")
    para._p.addnext(new_p)
    new_para = Paragraph(new_p, para._parent)
    if text:
        new_para.add_run(text)
    try:
        new_para.style = para.style
    except Exception:
        pass
    try:
        pPr = new_para._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F5F1E6")
        pPr.append(shd)
    except Exception:
        pass
    return new_para


def _pptx_set_para_text(para, new_text):
    """★ 修复：空 run 时用 add_run，避免 para.text 在某些版本上抛异常。"""
    if para.runs:
        para.runs[0].text = new_text
        for r in para.runs[1:]:
            r.text = ""
    else:
        try:
            run = para.add_run()
            run.text = new_text
        except Exception:
            try:
                para.text = new_text
            except Exception:
                pass


def _iter_pptx_shapes(shapes):
    for shape in shapes:
        if MSO_SHAPE_TYPE is not None and \
                getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            try:
                yield from _iter_pptx_shapes(shape.shapes)
                continue
            except Exception:
                pass
        yield shape


def translate_batch_office(client, model, texts, cache, cache_file,
                           target_lang="zh-CN", stop_event=None, task=None):
    pref = cache_prefix(target_lang)
    results = [None] * len(texts)
    failed = set()
    any_success = False

    def ckey(t):
        return f"t_{pref}" + h(t)

    pending = []
    for i, t in enumerate(texts):
        k = ckey(t)
        if k in cache and isinstance(cache[k], str) and cache[k].strip():
            results[i] = cache[k]
            any_success = True
        else:
            pending.append((i, t))

    for start in range(0, len(pending), OFFICE_BATCH_SIZE):
        if stop_event is not None and stop_event.is_set():
            break
        chunk = pending[start:start + OFFICE_BATCH_SIZE]
        marked = "\n\n".join(f"[[B{j}]] {t}" for j, (_, t) in enumerate(chunk))
        try:
            system = build_simple_system(target_lang)
            raw = _api_call(client, model, system, marked, stop_event=stop_event)
            parsed = parse_marked(raw, len(chunk))
            for j, (gi, t) in enumerate(chunk):
                tr = (parsed.get(j) or "").strip()
                if tr:
                    results[gi] = tr
                    cache[ckey(t)] = tr
                    any_success = True
                else:
                    failed.add(gi)
        except RuntimeError as e:
            if _is_fatal_api_error(str(e)):
                raise
            if task:
                task.log_msg(f"⚠️ 批量翻译失败（{e}），改为逐条重试")
            for gi, t in chunk:
                if stop_event is not None and stop_event.is_set():
                    break
                try:
                    tr = call_simple_api(client, model, t, target_lang,
                                         stop_event=stop_event)
                    results[gi] = tr
                    cache[ckey(t)] = tr
                    any_success = True
                except RuntimeError as e2:
                    if _is_fatal_api_error(str(e2)):
                        raise
                    failed.add(gi)
                    if task:
                        task.log_msg(f"⚠️ 段落 #{gi} 翻译失败：{e2}")
                    continue

        save_json_file(cache_file, cache)

    if pending and not any_success and not (
        stop_event is not None and stop_event.is_set()
    ):
        raise RuntimeError("所有段落翻译均失败，请检查 API Key / 余额 / 网络")

    fail_count = 0
    for i in range(len(results)):
        if results[i] is None:
            results[i] = texts[i]
            fail_count += 1

    return results, fail_count


# ============================================================
# PDF Worker
# ============================================================

def pdf_worker(task, paths, real_key, model, trial, target_lang="zh-CN",
               reader_profile="", want_terms=True,
               font_path="", max_font_size=DEFAULT_FONT_SIZE):
    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_json_file(paths["cache_file"], {})
    done_pages = load_progress_file(paths["progress_file"])

    preset = PDF_QUALITY_PRESETS.get(
        task.pdf_quality, PDF_QUALITY_PRESETS[DEFAULT_PDF_QUALITY]
    )
    if task.make_bilingual:
        task.log_msg(
            f"🎨 双语 PDF 质量：{task.pdf_quality}"
            f"（zoom={preset['zoom']}，JPEG={preset['jpeg']}）"
        )
    else:
        task.log_msg("🎨 已关闭双语 PDF 生成（仅输出译文）")

    enable_terms = bool(want_terms and HAS_NOTES and NB is not None)
    terms_file = os.path.join(task.work_dir, "terms.json")
    global_terms = NB.load_terms(terms_file) if enable_terms else {"terms": {}}

    chapters = []
    if enable_terms:
        try:
            with fitz.open(task.src_path) as _d:
                chapters = NB.detect_chapters(_d)
            if chapters:
                task.log_msg(f"📚 识别到 {len(chapters)} 个章节，术语将按章归类")
            else:
                task.log_msg("📚 未检测到 PDF 目录，术语不按章节归类")
        except Exception as e:
            task.log_msg(f"⚠️ 章节识别失败：{e}")

    if font_path and os.path.exists(font_path):
        task.log_msg(f"🔤 正文字体：{os.path.basename(font_path)}")
    else:
        task.log_msg("🔤 正文字体：PyMuPDF 内置宋体（china-s）")
    task.log_msg(f"📏 正文字号上限：{max_font_size}pt")
    if target_lang in RTL_LANGS:
        task.log_msg(f"↔️ 目标语言为 RTL，使用右对齐")

    try:
        with fitz.open(task.src_path) as _tmp_doc:
            total_src_pages = len(_tmp_doc)
    except Exception:
        total_src_pages = 0
    est = estimate_output_size(
        "pdf", task.src_path, total_src_pages,
        target_lang=target_lang, want_terms=enable_terms,
        pdf_quality=task.pdf_quality,
        make_bilingual=task.make_bilingual,
    )
    task.estimated_size = est.get("total", 0)
    if task.estimated_size:
        task.log_msg(
            f"📦 预估总产出 ≈ {fmt_size(task.estimated_size)}"
            f"（译文 ≈ {fmt_size(est.get('main', 0))} · "
            f"双语 ≈ {fmt_size(est.get('bilingual', 0))} · "
            f"笔记 ≈ {fmt_size(est.get('notes', 0))}）"
        )
    save_state(task, force=True)

    pdf_exists = os.path.exists(paths["output_pdf"])
    if done_pages and pdf_exists:
        try:
            with open(paths["output_pdf"], "rb") as f:
                data = f.read()
            doc = fitz.open(stream=data, filetype="pdf")
            task.log_msg(f"📂 从已翻译 PDF 续传（进度记录说已翻 {len(done_pages)} 页）")
        except Exception as e:
            task.log_msg(f"⚠️ 打开旧译文失败，从头开始：{e}")
            doc = fitz.open(task.src_path)
            done_pages = set()
            save_progress_file(paths["progress_file"], done_pages)
    else:
        doc = fitz.open(task.src_path)
        if done_pages and not pdf_exists:
            task.log_msg("⚠️ 检测到进度记录但缺少译文 PDF，从头开始")
            done_pages = set()
            save_progress_file(paths["progress_file"], done_pages)

    if pdf_exists and target_lang in NON_LATIN_SCRIPT_LANGS:
        try:
            orig_check = fitz.open(task.src_path)
            actually_done = set()
            checked = 0
            for i in range(min(len(doc), len(orig_check))):
                if task.stop_event.is_set():
                    break
                checked += 1
                if page_is_translated(doc[i], target_lang, src_page=orig_check[i]):
                    actually_done.add(i + 1)
            orig_check.close()

            reported = len(done_pages)
            actual = len(actually_done)
            task.log_msg(f"🔎 内容校验：译文共 {checked} 页，其中已翻译的 {actual} 页")

            if reported != actual:
                missing = sorted(set(range(1, checked + 1)) - actually_done)
                task.log_msg(f"⚠️ 进度记录说已翻 {reported} 页，但实际只有 {actual} 页")
                if missing:
                    task.log_msg(f"   待补翻页码前 20 个：{missing[:20]}"
                                 f"{'...' if len(missing) > 20 else ''}")
                    task.log_msg(f"   共 {len(missing)} 页需要补翻")
                else:
                    task.log_msg(f"   进度多余 {reported - actual} 页，已修正")
            else:
                task.log_msg(f"✅ 进度与实际内容一致")

            done_pages = actually_done
            save_progress_file(paths["progress_file"], done_pages)
        except Exception as e:
            task.log_msg(f"⚠️ 内容校验失败（不影响继续）：{e}")

    total = len(doc)
    limit = min(5, total) if trial else total

    task.status = "running"
    task.total = limit
    task.label = f"PDF · 目标前 {limit} 页"
    task.current = sum(1 for p in done_pages if p <= limit)
    task.log_msg(f"✅ PDF 共 {total} 页，本次目标 {limit} 页，已翻 {task.current} 页")
    save_state(task, force=True)

    if pdf_exists:
        try:
            task.preview_images = render_preview_only(paths["output_pdf"], paths, task)
            _add_output(task, paths["output_pdf"])
            refresh_task_size(task, force=True)
            save_state(task, force=True)
        except Exception as e:
            task.log_msg(f"⚠️ 初始预览生成失败：{e}")

    orig_doc_for_check = None
    try:
        orig_doc_for_check = fitz.open(task.src_path)
    except Exception as e:
        task.log_msg(f"⚠️ 无法打开原 PDF 用于校验/回滚：{e}（本次运行跳过回滚保护）")

    error_msg = None
    all_pages_ok = True

    for pno in range(total):
        if task.stop_event.is_set():
            task.log_msg("⏸ 检测到停止信号，结束当前循环")
            all_pages_ok = False
            break
        page_num = pno + 1
        if page_num > limit and page_num not in done_pages:
            break
        if page_num in done_pages:
            continue

        page = doc[pno]
        blocks = [b for b in page.get_text("blocks") if b[6] == 0 and b[4].strip()]

        if not blocks:
            done_pages.add(page_num)
            save_progress_file(paths["progress_file"], done_pages)
            if page_num <= limit:
                task.current += 1
            task.label = f"已跳过插图页 {page_num}"
            save_state(task)
            continue

        try:
            trans, page_terms = translate_page(
                client, model, blocks, cache,
                paths["cache_file"], target_lang,
                stop_event=task.stop_event,
                reader_profile=reader_profile,
                want_terms=enable_terms,
            )
            trans_complete = all(
                (trans.get(i) or "").strip() for i in range(len(blocks))
            )
            if not trans_complete or not translations_look_valid(
                    trans, blocks, target_lang):
                task.log_msg(f"⚠️ 第 {page_num} 页翻译结果不完整，跳过，稍后重试")
                all_pages_ok = False
                continue

            written, failed = apply_translations(
                page, blocks, trans,
                font_path=font_path,
                max_font_size=max_font_size,
                target_lang=target_lang,
            )
            if failed:
                task.log_msg(
                    f"⚠️ 第 {page_num} 页有 {failed}/{len(blocks)} 段写入失败"
                    f"（该页将进行字符校验，可能回滚重试）"
                )
                all_pages_ok = False

            if target_lang in NON_LATIN_SCRIPT_LANGS:
                src_check_page = None
                if (orig_doc_for_check is not None
                        and pno < len(orig_doc_for_check)):
                    src_check_page = orig_doc_for_check[pno]

                if not page_is_translated(page, target_lang,
                                          src_page=src_check_page):
                    try:
                        after_text = page.get_text() or ""
                    except Exception:
                        after_text = ""
                    total_ns = sum(1 for c in after_text if not c.isspace())
                    if target_lang in ("zh-CN", "zh-TW"):
                        cnt_target = sum(
                            1 for c in after_text if '\u4e00' <= c <= '\u9fff')
                    elif target_lang == "ja":
                        cnt_target = sum(
                            1 for c in after_text
                            if '\u3040' <= c <= '\u30ff')
                    elif target_lang == "ko":
                        cnt_target = sum(
                            1 for c in after_text if '\uac00' <= c <= '\ud7a3')
                    elif target_lang == "ru":
                        cnt_target = sum(
                            1 for c in after_text if '\u0400' <= c <= '\u04ff')
                    elif target_lang == "ar":
                        cnt_target = sum(
                            1 for c in after_text if '\u0600' <= c <= '\u06ff')
                    else:
                        cnt_target = 0
                    ratio = (cnt_target / total_ns) if total_ns else 0.0
                    n_translated = sum(
                        1 for v in trans.values() if (v or "").strip())
                    task.log_msg(
                        f"⚠️ 第 {page_num} 页写入后未检出目标语言字符 "
                        f"(块={len(blocks)}, 译段={n_translated}, "
                        f"目标字符={cnt_target}, 非空白={total_ns}, "
                        f"占比={ratio:.1%}, "
                        f"阈值需≥{PAGE_CN_THRESHOLD}且≥{PAGE_CN_RATIO:.0%})，"
                        f"回滚该页"
                    )
                    if (orig_doc_for_check is not None
                            and pno < len(orig_doc_for_check)):
                        try:
                            doc.insert_pdf(
                                orig_doc_for_check,
                                from_page=pno, to_page=pno, start_at=pno,
                            )
                            doc.delete_page(pno + 1)
                            page = doc[pno]
                        except Exception as _e:
                            error_msg = (
                                f"回滚第 {page_num} 页失败，"
                                f"PDF 结构可能已损坏：{_e}"
                            )
                            task.log_msg(f"❌ {error_msg}")
                            all_pages_ok = False
                            break
                    all_pages_ok = False
                    continue

            if enable_terms and page_terms:
                ch_title = NB.guess_chapter_for_page(page_num, chapters)
                NB.merge_page_terms(global_terms, page_terms, page_num, ch_title)
                NB.save_terms(terms_file, global_terms)
                task.log_msg(
                    f"📝 第 {page_num} 页抽取术语 {len(page_terms)} 个"
                    f"（全局 {len(global_terms.get('terms', {}))} 个）"
                )
        except RuntimeError as e:
            error_msg = str(e)
            break
        except Exception as e:
            error_msg = f"未知错误：{e}"
            break

        done_pages.add(page_num)
        save_progress_file(paths["progress_file"], done_pages)
        if page_num <= limit:
            task.current += 1
        task.label = f"已翻 {task.current}/{limit} 页"
        task.log_msg(f"✅ 第 {page_num} 页完成")
        save_state(task, force=(page_num % 5 == 0))

        should_save = (page_num <= PREVIEW_PAGES) or (page_num % CHECKPOINT_EVERY == 0)
        if should_save:
            try:
                saved = safe_save_pdf(doc, paths["output_pdf"], garbage=1)
                if saved != paths["output_pdf"]:
                    task.log_msg(
                        f"⚠️ 原文件被占用，已保存到备用路径："
                        f"{os.path.basename(saved)}"
                    )
                    paths["output_pdf"] = saved
                task.log_msg(f"💾 已落盘（前 {page_num} 页）")
                if page_num <= PREVIEW_PAGES:
                    task.preview_images = render_preview_only(
                        paths["output_pdf"], paths, task)
                _add_output(task, paths["output_pdf"])
                refresh_task_size(task, force=True)
                save_state(task, force=True)
            except Exception as e:
                task.log_msg(f"⚠️ 落盘失败：{e}")

    save_failed = False
    try:
        saved = safe_save_pdf(doc, paths["output_pdf"], garbage=3)
        if saved != paths["output_pdf"]:
            task.log_msg(
                f"⚠️ 原文件被占用，已保存到备用路径："
                f"{os.path.basename(saved)}"
            )
            paths["output_pdf"] = saved
    except Exception as e:
        save_failed = True
        error_msg = error_msg or f"保存译文 PDF 失败：{e}"
    finally:
        try:
            doc.close()
        except Exception:
            pass
        try:
            if orig_doc_for_check is not None:
                orig_doc_for_check.close()
        except Exception:
            pass

    if save_failed or not os.path.exists(paths["output_pdf"]):
        task.status = "error"
        task.error = error_msg or "译文 PDF 保存失败，无法继续生成双语 / 笔记"
        task.label = "出错（保存失败）"
        task.log_msg(f"❌ {task.error}")
        refresh_task_size(task, force=True)
        save_state(task, force=True)
        return

    target_pages = set(range(1, limit + 1))
    done_in_target = target_pages & done_pages
    completed = (done_in_target == target_pages) and all_pages_ok and not error_msg

    if not done_pages and not pdf_exists:
        if task.stop_event.is_set():
            task.status = "paused"
            task.error = ""
            task.label = "已暂停（未翻译任何页）"
            task.log_msg("🛑 已暂停：尚未开始翻译")
        elif error_msg:
            task.status = "error"
            task.error = error_msg
            task.log_msg(f"❌ {error_msg}")
        else:
            task.status = "error"
            task.error = "本次没有任何页面成功翻译"
            task.log_msg("⚠️ " + task.error)
        save_state(task, force=True)
        return

    try:
        dp = {p for p in done_pages if p <= limit}
        if not dp:
            task.log_msg("ℹ️ 无已翻译页，跳过双语 PDF 生成")
        elif not task.make_bilingual:
            task.log_msg("ℹ️ 用户设置不生成双语 PDF，跳过")
            _add_output(task, paths["output_pdf"])
        else:
            task.log_msg(
                f"🖼 生成左右对照双语 PDF（{len(dp)} 页，"
                f"质量「{task.pdf_quality}」，请稍候）……")
            make_bilingual_pdf(
                paths["output_pdf"], paths, task,
                n_pages=None, done_pages=dp,
                zoom=preset["zoom"],
                jpeg_quality=preset["jpeg"],
                garbage=preset["garbage"],
            )
            if os.path.exists(paths["bilingual_pdf"]):
                _add_output(task, paths["output_pdf"], paths["bilingual_pdf"])
            else:
                _add_output(task, paths["output_pdf"])
    except Exception as e:
        task.log_msg(f"⚠️ 收尾生成双语 PDF 失败：{e}")

    if enable_terms and global_terms.get("terms"):
        try:
            book_title = os.path.splitext(task.src_name)[0]
            lang_label = LANG_NAMES.get(target_lang, target_lang)

            notes_dir = paths["notes_dir"]
            os.makedirs(notes_dir, exist_ok=True)

            notes_md   = os.path.join(notes_dir, "notes.md")
            notes_html = os.path.join(notes_dir, "notes.html")
            terms_csv  = os.path.join(notes_dir, "terms.csv")

            md = NB.build_notes_markdown(
                book_title, global_terms,
                reader_profile=reader_profile,
                total_pages=limit, lang_label=lang_label,
            )
            with open(notes_md, "w", encoding="utf-8") as f:
                f.write(md)

            try:
                html = NB.build_notes_html(
                    book_title, global_terms,
                    reader_profile=reader_profile,
                    total_pages=limit, lang_label=lang_label,
                    pdf_rel_path="../translated.pdf",
                )
            except TypeError:
                html = NB.build_notes_html(
                    book_title, global_terms,
                    reader_profile=reader_profile,
                    total_pages=limit, lang_label=lang_label,
                )
            with open(notes_html, "w", encoding="utf-8") as f:
                f.write(html)

            NB.build_terms_csv(global_terms, terms_csv)

            _add_output(task, notes_md, notes_html, terms_csv)

            task.log_msg(
                f"📓 阅读笔记已生成（{len(global_terms['terms'])} 个术语）"
                f" → {os.path.basename(notes_dir)}/"
            )
        except Exception as e:
            task.log_msg(f"⚠️ 生成阅读笔记失败：{e}")

    if enable_terms and global_terms.get("terms"):
        try:
            book_title = os.path.splitext(task.src_name)[0]
            lang_label = LANG_NAMES.get(target_lang, target_lang)

            if task.make_bilingual and os.path.exists(paths["bilingual_pdf"]):
                if insert_notes_into_pdf(
                    paths["bilingual_pdf"], global_terms, book_title,
                    reader_profile, limit, lang_label,
                    font_path=font_path, position="front",
                ):
                    task.log_msg("📎 术语表已插入 bilingual.pdf 开头（带书签）")

            try:
                with open(paths["output_pdf"], "rb") as f:
                    base_data = f.read()
                base_doc = fitz.open(stream=base_data, filetype="pdf")
                trans_with_notes = paths["trans_with_notes_pdf"]
                base_doc.save(trans_with_notes, deflate=True, garbage=3)
                base_doc.close()

                if insert_notes_into_pdf(
                    trans_with_notes, global_terms, book_title,
                    reader_profile, limit, lang_label,
                    font_path=font_path, position="front",
                ):
                    task.log_msg(
                        "📎 术语表已插入 translated_with_notes.pdf 开头（带书签）")
                    _add_output(task, trans_with_notes)
            except Exception as e:
                task.log_msg(f"⚠️ 生成 translated_with_notes.pdf 失败：{e}")
        except Exception as e:
            task.log_msg(f"⚠️ 插入术语页失败：{e}")

    try:
        task.preview_images = render_preview_only(paths["output_pdf"], paths, task)
    except Exception:
        pass

    if completed:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 全部完成！共 {len(done_pages)} 页，最大页码 {max(done_pages)}")
    elif task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停（半成品），共翻 {len(done_pages)} 页"
        if done_pages:
            task.log_msg(f"🛑 已暂停。下次点开始会从第 {max(done_pages) + 1} 页继续")
        else:
            task.log_msg("🛑 已暂停。")
    elif error_msg:
        task.status = "error"
        task.error = error_msg
        task.label = "出错（半成品）"
        task.log_msg(f"❌ {error_msg}")
    else:
        task.status = "paused"
        task.label = "半成品（部分页未成功）"
    refresh_task_size(task, force=True)
    save_state(task, force=True)


def docx_worker(task, paths, real_key, model, target_lang="zh-CN"):
    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_json_file(paths["cache_file"], {})

    est = estimate_output_size("docx", task.src_path, 0,
                               target_lang=target_lang, want_terms=False)
    task.estimated_size = est.get("total", 0)
    if task.estimated_size:
        task.log_msg(f"📦 预估总产出 ≈ {fmt_size(task.estimated_size)}")

    try:
        wdoc = Document(task.src_path)
    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        task.status = "error"
        task.error = f"打开 Word 失败：{e}"
        task.log_msg(f"❌ 打开 Word 失败：{e}")
        if "Package not found" in err_detail or "not a zip" in err_detail.lower():
            task.log_msg("   ⚠️ 这通常意味着 .docx 其实是从 .doc 改后缀来的。")
        elif "PermissionError" in err_detail:
            task.log_msg("   ⚠️ 文件被其他程序占用。")
        save_state(task, force=True)
        return

    para_targets = [p for p in wdoc.paragraphs if p.text.strip()]
    cell_targets = []
    seen = set()
    for table in wdoc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                for p in cell.paragraphs:
                    if p.text.strip():
                        cell_targets.append(p)

    all_targets = para_targets + cell_targets
    total_targets = len(all_targets)

    task.status = "running"
    task.total = total_targets
    task.current = 0
    task.label = f"Word · 共 {total_targets} 段"
    task.log_msg(f"📘 Word 已打开，共 {total_targets} 段，将批量翻译")
    save_state(task, force=True)

    all_texts = [p.text for p in all_targets]

    try:
        translated, fail_count = translate_batch_office(
            client, model, all_texts, cache, paths["cache_file"],
            target_lang, stop_event=task.stop_event, task=task,
        )
    except RuntimeError as e:
        task.status = "error"
        task.error = str(e)
        task.log_msg(f"❌ {e}")
        save_state(task, force=True)
        return

    preview_pairs = []
    for i, (para, tr) in enumerate(zip(all_targets, translated)):
        if task.stop_event.is_set():
            break
        _docx_replace_para_text(para, tr)
        if len(preview_pairs) < PREVIEW_PARAS:
            preview_pairs.append((all_texts[i], tr))
            task.preview_html = _build_docx_preview_html(preview_pairs)
        task.current = i + 1
        task.label = f"Word 应用 {i+1}/{total_targets}"
        if (i + 1) % 5 == 0:
            save_state(task)

    try:
        wdoc.save(paths["office_cn"])
        task.log_msg(f"✅ 译文版已保存：{paths['office_cn']}")
    except Exception as e:
        task.status = "error"
        task.error = f"保存 Word 失败：{e}"
        task.log_msg("❌ " + task.error)
        save_state(task, force=True)
        return

    bi_ok = False
    if not task.stop_event.is_set():
        try:
            doc_bi = Document(task.src_path)
            para_bi = [p for p in doc_bi.paragraphs if p.text.strip()]
            cell_bi = []
            seen_bi = set()
            for table in doc_bi.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell._tc in seen_bi:
                            continue
                        seen_bi.add(cell._tc)
                        for p in cell.paragraphs:
                            if p.text.strip():
                                cell_bi.append(p)
            all_bi = para_bi + cell_bi
            for i, para in enumerate(all_bi):
                if task.stop_event.is_set():
                    break
                src = para.text
                key = f"t_{cache_prefix(target_lang)}" + h(src)
                tr = cache.get(key) or src
                _docx_insert_after(para, tr)
            doc_bi.save(paths["office_bi"])
            task.log_msg(f"✅ 双语版已保存：{paths['office_bi']}")
            bi_ok = True
        except Exception as e:
            task.log_msg(f"⚠️ 保存双语版失败：{e}")

    _add_output(task, paths["office_cn"])
    if bi_ok:
        _add_output(task, paths["office_bi"])

    if task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停 {task.current}/{total_targets}"
        task.log_msg(f"🛑 已暂停（半成品），共处理 {task.current}/{total_targets} 段")
    elif fail_count > 0:
        task.status = "paused"
        task.error = f"{fail_count} 段翻译失败，未替换为译文"
        task.label = f"半成品（{fail_count} 段失败）"
        task.log_msg(f"⚠️ {task.error}（可再次点击开始以重试）")
    else:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 Word 翻译完成！共 {total_targets} 段")
    refresh_task_size(task, force=True)
    save_state(task, force=True)


def pptx_worker(task, paths, real_key, model, target_lang="zh-CN"):
    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_json_file(paths["cache_file"], {})

    est = estimate_output_size("pptx", task.src_path, 0,
                               target_lang=target_lang, want_terms=False)
    task.estimated_size = est.get("total", 0)
    if task.estimated_size:
        task.log_msg(f"📦 预估总产出 ≈ {fmt_size(task.estimated_size)}")

    try:
        prs = Presentation(task.src_path)
    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        task.status = "error"
        task.error = f"打开 PPT 失败：{e}"
        task.log_msg(f"❌ 打开 PPT 失败：{e}")
        if "Package not found" in err_detail or "not a zip" in err_detail.lower():
            task.log_msg("   ⚠️ 这通常意味着 .pptx 其实是从 .ppt 改后缀来的。")
        save_state(task, force=True)
        return

    targets = []
    texts = []
    for si, slide in enumerate(prs.slides):
        for shape in _iter_pptx_shapes(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            for para in shape.text_frame.paragraphs:
                text = "".join(r.text for r in para.runs)
                if text.strip():
                    targets.append((si, para))
                    texts.append(text)

    total_targets = len(targets)
    task.status = "running"
    task.total = total_targets
    task.current = 0
    task.label = f"PPT · 共 {total_targets} 段"
    task.log_msg(f"📊 PPT 已打开，共 {total_targets} 段")
    save_state(task, force=True)

    try:
        translated, fail_count = translate_batch_office(
            client, model, texts, cache, paths["cache_file"],
            target_lang, stop_event=task.stop_event, task=task,
        )
    except RuntimeError as e:
        task.status = "error"
        task.error = str(e)
        task.log_msg(f"❌ {e}")
        save_state(task, force=True)
        return

    preview_pairs = []
    for i, ((si, para), tr) in enumerate(zip(targets, translated)):
        if task.stop_event.is_set():
            break
        _pptx_set_para_text(para, tr)
        if len(preview_pairs) < PREVIEW_PARAS:
            preview_pairs.append((si + 1, texts[i], tr))
            task.preview_html = _build_pptx_preview_html(preview_pairs)
        task.current = i + 1
        task.label = f"PPT {i+1}/{total_targets}（第 {si+1} 张）"
        if (i + 1) % 5 == 0:
            save_state(task)

    try:
        prs.save(paths["office_cn"])
        _add_output(task, paths["office_cn"])
    except Exception as e:
        task.status = "error"
        task.error = f"保存 PPT 失败：{e}"
        task.log_msg("❌ " + task.error)
        save_state(task, force=True)
        return

    if task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停 {task.current}/{total_targets}"
        task.log_msg(f"🛑 已暂停（半成品），共处理 {task.current}/{total_targets} 段")
    elif fail_count > 0:
        task.status = "paused"
        task.error = f"{fail_count} 段翻译失败，未替换为译文"
        task.label = f"半成品（{fail_count} 段失败）"
        task.log_msg(f"⚠️ {task.error}（可再次点击开始以重试）")
    else:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 PPT 翻译完成！共 {total_targets} 段")
    refresh_task_size(task, force=True)
    save_state(task, force=True)


# ============================================================
# 任务控制
# ============================================================

def _gradio_upload_path(doc_file):
    if doc_file is None:
        return None, ""
    path = getattr(doc_file, "name", None)
    orig = getattr(doc_file, "orig_name", None)
    if path is None and isinstance(doc_file, str):
        path = doc_file
    if not path:
        return None, ""
    disp = orig or os.path.basename(path)
    return path, disp


def start_task(kind, upload_path, real_key, model, trial, target_lang="zh-CN",
               reader_profile="", want_terms=True,
               font_path="", max_font_size=DEFAULT_FONT_SIZE,
               pdf_quality=DEFAULT_PDF_QUALITY, make_bilingual=True,
               display_name=None, out_subdir=None, warnings=None):
    global SELECTED_TASK_ID
    raw_name = display_name or os.path.basename(upload_path)
    src_name = safe_dirname(os.path.basename(raw_name))
    if not os.path.splitext(src_name)[1]:
        src_name = src_name + os.path.splitext(upload_path)[1]

    paths = prepare_paths(src_name, out_subdir=out_subdir)
    src_copy = os.path.join(paths["input_dir"], src_name)
    try:
        shutil.copy2(upload_path, src_copy)
    except Exception as e:
        raise RuntimeError(f"复制上传文件失败：{e}")

    task = MANAGER.create(kind, src_copy, src_name, paths["out_dir"],
                          paths["work"], target_lang)
    SELECTED_TASK_ID = task.task_id

    task.pdf_quality = pdf_quality or DEFAULT_PDF_QUALITY
    task.make_bilingual = bool(make_bilingual)

    lang_label = LANG_NAMES.get(target_lang, target_lang)
    task.log_msg(f"🆔 任务 {task.task_id} 已创建（{kind}，目标语言：{lang_label}）")
    task.log_msg(f"📁 结果目录：{paths['out_dir']}")
    for w in (warnings or []):
        task.log_msg(w)

    def runner():
        try:
            if kind == "pdf":
                pdf_worker(task, paths, real_key, model, trial, target_lang,
                           reader_profile=reader_profile,
                           want_terms=want_terms,
                           font_path=font_path,
                           max_font_size=max_font_size)
            elif kind == "docx":
                docx_worker(task, paths, real_key, model, target_lang)
            elif kind == "pptx":
                pptx_worker(task, paths, real_key, model, target_lang)
        except Exception as e:
            task.status = "error"
            task.error = str(e)
            task.log_msg(f"❌ 未捕获错误：{e}")
            save_state(task, force=True)

    task.thread = threading.Thread(target=runner, daemon=True)
    task.thread.start()
    return task


def on_start(api_key, model, doc_file, mode, trial, target_lang,
             reader_profile, want_terms,
             font_choice, font_size,
             pdf_quality, make_bilingual_cb,
             stop_dd_value=None):
    global SELECTED_TASK_ID

    keep_upload = gr.update()

    real_key = (api_key or "").strip() or DEFAULT_API_KEY
    if not real_key or doc_file is None:
        return on_refresh_fast(stop_dd_value) + ([], keep_upload)

    model = (model or DEFAULT_MODEL).strip()
    if model not in ("deepseek-chat", "deepseek-reasoner"):
        model = "deepseek-chat"

    if target_lang not in LANG_NAMES:
        target_lang = "zh-CN"

    user_font_path = FONTS_MAP.get(font_choice, "") if font_choice else ""
    sel_font_path = _auto_pick_font(target_lang, user_font_path)

    try:
        sel_font_size = float(font_size)
    except Exception:
        sel_font_size = DEFAULT_FONT_SIZE
    if sel_font_size < 6 or sel_font_size > 24:
        sel_font_size = DEFAULT_FONT_SIZE

    if pdf_quality not in PDF_QUALITY_PRESETS:
        pdf_quality = DEFAULT_PDF_QUALITY

    upload_path, display_name = _gradio_upload_path(doc_file)
    if not upload_path:
        return on_refresh_fast(stop_dd_value) + ([], keep_upload)

    display_name = safe_dirname(os.path.basename(display_name or upload_path))
    _ext = os.path.splitext(upload_path)[1].lower()
    if not display_name.lower().endswith(_ext):
        display_name = safe_dirname(os.path.splitext(display_name)[0]) + _ext

    name_lower = display_name.lower()

    if name_lower.endswith(".pdf"):
        kind = "pdf"
    elif name_lower.endswith(".docx"):
        kind = "docx"
    elif name_lower.endswith(".pptx"):
        kind = "pptx"
    else:
        print(f"[on_start] 不支持的文件类型：{display_name}")
        return on_refresh_fast(stop_dd_value) + ([], keep_upload)

    if kind in ("docx", "pptx") and not HAS_OFFICE:
        print(f"[on_start] 未安装 python-docx / python-pptx，无法处理 {display_name}")
        return on_refresh_fast(stop_dd_value) + ([], keep_upload)

    warnings = []

    # ★ 优化：校验文档类型选择与文件后缀是否一致
    _mode_ext = {
        "📕 PDF 书籍": ".pdf",
        "📘 Word 文档": ".docx",
        "📊 PPT 演示": ".pptx",
    }
    if mode in _mode_ext and not name_lower.endswith(_mode_ext[mode]):
        warnings.append(
            f"ℹ️ 文档类型选择与文件后缀不一致，已按后缀（{_ext}）处理"
        )

    if kind == "pdf" and (not sel_font_path or not os.path.exists(sel_font_path)):
        warnings.append("⚠️ 未选中有效字体，将使用 PyMuPDF 内置宋体（china-s）")
    if (kind == "pdf" and target_lang in ("ar", "ko")
            and (not sel_font_path
                 or os.path.basename(sel_font_path).lower()
                    .find("sourcehanserifsc") >= 0)):
        warnings.append(
            f"⚠️ 目标语言为 {LANG_NAMES[target_lang]}，"
            f"当前字体可能不含所需字符，建议选择支持该语言的字体"
        )
    if trial and kind in ("docx", "pptx"):
        warnings.append("ℹ️ 试翻模式仅对 PDF 生效，Word / PPT 将完整翻译")

    src_name = display_name
    base = os.path.splitext(src_name)[0]
    book = safe_dirname(base)
    short_hash = h(display_name)[:6]
    out_dir_name = f"{book}_{short_hash}"
    target_out_dir = os.path.abspath(os.path.join(RESULT_ROOT, out_dir_name))

    with _CREATE_LOCK:
        existing = MANAGER.find_active_by_out_dir(target_out_dir)
        if existing is not None:
            SELECTED_TASK_ID = existing.task_id
            existing.log_msg(f"⚠️ 重复创建请求已拒绝，任务 #{existing.task_id} 已在运行")
            save_state(existing, force=True)
            return on_refresh_fast(stop_dd_value) + ([], keep_upload)

        try:
            start_task(kind, upload_path, real_key, model, trial, target_lang,
                       reader_profile=reader_profile,
                       want_terms=want_terms,
                       font_path=sel_font_path,
                       max_font_size=sel_font_size,
                       pdf_quality=pdf_quality,
                       make_bilingual=make_bilingual_cb,
                       display_name=src_name,
                       out_subdir=out_dir_name,
                       warnings=warnings)
        except Exception as e:
            print(f"[on_start] 创建任务失败：{e}")

    return on_refresh_fast(stop_dd_value) + ([], keep_upload)


def on_stop_all(stop_dd_value=None):
    running = MANAGER.running()
    if not running:
        current = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
        if current:
            current.log_msg("⚠️ 没有正在运行的任务")
            save_state(current, force=True)
        return on_refresh_fast(stop_dd_value)

    for t in running:
        t.stop_event.set()
        t.status = "stopping"
        t.label = f"⏸ 停止信号已发出…（{t.current}/{t.total}）"
        t.log_msg("⏸ 收到停止信号 —— 当前一步完成后暂停")
        save_state(t, force=True)

    return on_refresh_fast(stop_dd_value)


def on_stop_selected(label, stop_dd_value=None):
    global SELECTED_TASK_ID
    if not label:
        return on_refresh_fast(stop_dd_value)

    m = re.match(r'#([0-9a-f]{8})', label)
    if not m:
        return on_refresh_fast(stop_dd_value)

    tid = m.group(1)
    t = MANAGER.get(tid)
    if t is None:
        return on_refresh_fast(stop_dd_value)

    if t.status in ("queued", "running"):
        t.stop_event.set()
        t.status = "stopping"
        t.label = f"⏸ 停止信号已发出…（{t.current}/{t.total}）"
        t.log_msg("⏸ 收到单独停止信号")
        save_state(t, force=True)
        SELECTED_TASK_ID = tid
    elif t.status == "stopping":
        t.log_msg("⏸ 该任务已在停止中")
        save_state(t, force=True)

    return on_refresh_fast(stop_dd_value)


def on_load_preview():
    current = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
    if current is None:
        tasks = MANAGER.all_sorted()
        current = tasks[0] if tasks else None
    if current is None:
        return build_preview_html([]), []

    imgs = [p for p in (current.preview_images or [])
            if p and os.path.exists(p)]
    # ★ 优化：内存为空时从磁盘恢复（每次点击都重新扫，但只在真的空时才扫）
    if not imgs:
        imgs = _collect_existing_previews(current)
        if imgs:
            current.preview_images = imgs
    ph = getattr(current, "preview_html", "") or ""
    files = [os.path.abspath(f) for f in (current.output_files or [])
             if f and os.path.exists(f)]
    return build_preview_html(imgs, ph), files


def build_task_list_html(tasks):
    if not tasks:
        return ('<div style="padding:18px;color:#a9a49a;font-size:13px;'
                'text-align:center">暂无任务</div>')

    status_color = {"queued": "#8b8578", "running": "#0f3d3e",
                    "stopping": "#b8860b", "paused": "#b8860b",
                    "done": "#0f3d3e", "error": "#c0392b"}
    status_bg = {"queued": "#f0ebe0", "running": "#e8f0ef",
                 "stopping": "#fdf1e0", "paused": "#fdf1e0",
                 "done": "#e8f0ef", "error": "#fbeaea"}
    status_text = {"queued": "排队中", "running": "运行中", "stopping": "停止中…",
                   "paused": "已暂停", "done": "已完成", "error": "出错"}
    kind_icon = {"pdf": "📕", "docx": "📘", "pptx": "📊"}

    rows = []
    for t in tasks[:8]:
        col = status_color.get(t.status, "#666")
        bg = status_bg.get(t.status, "#f0ebe0")
        stx = status_text.get(t.status, t.status)
        kc = kind_icon.get(t.kind, "📄")
        pct = int(t.current * 100 / t.total) if t.total else 0
        lang_short = LANG_NAMES.get(getattr(t, "target_lang", "zh-CN"), "")
        size_hint = ""
        if t.current_size or t.estimated_size:
            size_hint = (
                f'<span style="font-size:11px;color:#8b8578;'
                f'font-family:\'SF Mono\',Consolas,monospace">'
                f'{fmt_size(t.current_size) if t.current_size else "0 B"}'
                f' / {fmt_size(t.estimated_size) if t.estimated_size else "—"}'
                f'</span>'
            )
        rows.append(f'''
        <div style="padding:11px 14px;border-bottom:1px solid #f2ede0;
                    display:flex;align-items:center;gap:10px;font-size:13px">
          <span style="font-size:15px;flex-shrink:0">{kc}</span>
          <span style="font-family:'SF Mono',Consolas,monospace;
                       color:#8b8578;font-size:11.5px;flex-shrink:0">#{t.task_id}</span>
          <span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;
                       white-space:nowrap;color:#1a1a1a"
                title="{escape(t.src_name)}">{escape(t.src_name)}</span>
          {size_hint}
          <span style="font-size:11.5px;color:#8b8578;background:#f5f1e6;
                       padding:2px 8px;border-radius:999px;flex-shrink:0">{lang_short}</span>
          <span style="color:{col};background:{bg};font-weight:600;
                       font-size:11.5px;padding:2px 10px;border-radius:999px;
                       flex-shrink:0">{stx}</span>
          <span style="color:#8b8578;font-family:'SF Mono',Consolas,monospace;
                       font-size:11.5px;flex-shrink:0">{t.current}/{t.total} ({pct}%)</span>
        </div>''')

    rows_html = "".join(rows)

    # ★ 优化：超过 8 个任务时提示还有多少未显示
    if len(tasks) > 8:
        rows_html += (
            f'<div style="padding:10px 14px;text-align:center;'
            f'color:#a9a49a;font-size:12px;background:#fdfcf8">'
            f'… 还有 {len(tasks) - 8} 个任务未显示</div>'
        )

    return (f'<div style="background:#fff;border:1px solid #ebe5d8;'
            f'border-radius:14px;overflow:hidden">{rows_html}</div>')


def on_refresh_fast(stop_dd_value=None):
    global _SELECTED_STOP_VALUE
    tasks = MANAGER.all_sorted()
    task_list_html = build_task_list_html(tasks)

    current = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
    if current is None and tasks:
        current = tasks[0]

    if current is None:
        progress_html = make_progress_html(0, 1, "等待开始")
        log_text = ""
        previews = []
        preview_html = ""
    else:
        try:
            refresh_task_size(current)
        except Exception:
            pass
        warning = ""
        if current.status == "paused" and current.error:
            warning = current.error
        progress_html = make_progress_html(
            current.current, current.total, current.label,
            current_size=current.current_size,
            estimated_size=current.estimated_size,
            warning=warning,
        )
        log_text = "\n".join(current.log[-40:])
        previews = [p for p in (current.preview_images or [])
                    if p and os.path.exists(p)]
        if not previews and not getattr(current, "_preview_scanned", False):
            current._preview_scanned = True
            previews = _collect_existing_previews(current)
            if previews:
                current.preview_images = previews
        preview_html = getattr(current, "preview_html", "") or ""

    if current is not None:
        if preview_html:
            sig = ("html", hashlib.md5(preview_html.encode("utf-8")).hexdigest())
        else:
            sig = preview_signature(previews)

        if _LAST_PREVIEW_SIG.get(current.task_id) != sig:
            _LAST_PREVIEW_SIG[current.task_id] = sig
            html = build_preview_html(previews, preview_html)
            _PREVIEW_HTML_CACHE[current.task_id] = html
            gallery_value = html
        else:
            gallery_value = _PREVIEW_HTML_CACHE.get(current.task_id)
            if gallery_value is None:
                gallery_value = build_preview_html([])
    else:
        gallery_value = build_preview_html([])

    active = [t for t in tasks if t.status in ("queued", "running", "stopping")]
    choices = []
    for t in active:
        short = t.src_name if len(t.src_name) <= 50 else t.src_name[:47] + "..."
        choices.append(f"#{t.task_id}  {short}  [{t.status}]")

    preserve = stop_dd_value if stop_dd_value in choices else _SELECTED_STOP_VALUE
    if preserve not in choices:
        preserve = None
    _SELECTED_STOP_VALUE = preserve
    dd_update = gr.update(choices=choices, value=preserve)

    # ★ 修复：用全部任务 id 作为 alive 集合。
    #   之前用 tasks[:50] 会导致超过 50 个任务时，早完成的 done 任务
    #   被清出 _MODAL_SHOWN → 弹窗反复触发。
    alive_ids = {t.task_id for t in tasks}
    for k in list(_LAST_PREVIEW_SIG.keys()):
        if k not in alive_ids:
            _LAST_PREVIEW_SIG.pop(k, None)
    for k in list(_PREVIEW_HTML_CACHE.keys()):
        if k not in alive_ids:
            _PREVIEW_HTML_CACHE.pop(k, None)
    for k in list(_STATE_SAVE_TS.keys()):
        if k not in alive_ids:
            _STATE_SAVE_TS.pop(k, None)
    for k in list(_MODAL_SHOWN):
        if k not in alive_ids:
            _MODAL_SHOWN.discard(k)

    # ★ 修复：弹窗只弹一次；默认返回 gr.update() 表示"不改"，
    #        避免下一次 tick 用空字符串把已弹出的弹窗清掉。
    modal_html = gr.update()
    for t in tasks:
        if t.status == "done" and t.task_id not in _MODAL_SHOWN:
            _MODAL_SHOWN.add(t.task_id)
            try:
                refresh_task_size(t, force=True)
            except Exception:
                pass
            modal_html = build_done_modal_html(t)
            break

    return task_list_html, progress_html, log_text, dd_update, gallery_value, modal_html


def open_result_folder():
    current = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
    if current is None:
        tasks = MANAGER.all_sorted()
        current = tasks[0] if tasks else None
    if current is None:
        return "⚠️ 还没有任务"
    folder = os.path.abspath(current.out_dir)
    if not os.path.exists(folder):
        return f"⚠️ 结果文件夹不存在：{folder}"
    try:
        if sys.platform.startswith("win"):
            os.startfile(folder)  # type: ignore
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", folder])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", folder])
        return f"✅ 已打开：{folder}"
    except Exception as e:
        return f"❌ 打开失败：{e}\n路径：{folder}"


# ============================================================
# 优雅退出（Ctrl+C 秒退）
# ============================================================

_sigint_count = [0]


def _shutdown():
    """
    保存所有运行中任务的状态。

    ★ 修复：直接写盘，绕过 _STATE_SAVE_LOCK / _IO_LOCK。
    信号处理函数跑在主线程，如果此时 worker 线程正持有锁，
    用锁会阻塞 Ctrl+C，用户感觉"卡住"。反正紧接着就 os._exit(0)，
    锁没有意义，直接原子写即可。

    ★ 修复：tmp 文件名用 .exit.tmp，避免和 worker 线程的
    state.json.tmp 撞名（可能导致读到截断的 state.json）。
    """
    try:
        for t in MANAGER.running():
            try:
                t.stop_event.set()
                t.log_msg("🛑 程序退出，正在停止…")

                path = os.path.join(t.work_dir, "state.json")
                os.makedirs(t.work_dir, exist_ok=True)
                data = t.to_dict()

                tmp = path + ".exit.tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                os.replace(tmp, path)
            except Exception:
                pass
    except Exception:
        pass


atexit.register(_shutdown)


def _sigint(signum, frame):
    _sigint_count[0] += 1

    if _sigint_count[0] >= 2:
        print("\n⏹ 强制退出", flush=True)
        os._exit(0)

    print("\n⏸ 收到 Ctrl+C，正在保存任务状态……", flush=True)
    _shutdown()
    print("✅ 状态已保存，退出", flush=True)
    os._exit(0)


try:
    signal.signal(signal.SIGINT, _sigint)
except Exception:
    pass


# ============================================================
# UI
# ============================================================
with gr.Blocks(
    title="PDF / Word / PPT 翻译器",
    theme=gr.themes.Base(
        primary_hue=gr.themes.colors.gray,
        neutral_hue=gr.themes.colors.gray,
        font=[gr.themes.GoogleFont("Noto Sans SC"), "system-ui", "sans-serif"],
    ),
    css="""
    body, .gradio-container {
        background: #faf8f2 !important;
        color: #1a1a1a !important;
        font-size: 15px !important;
    }
    .gradio-container {
        max-width: 1380px !important;
        margin: 0 auto !important;
        padding: 14px 28px 44px !important;
    }
    .card-head {
        background: linear-gradient(135deg, #f7f3e8 0%, #f0ebdc 100%);
        border: 1px solid #e2dccb;
        border-radius: 18px;
        padding: 28px 32px 24px;
        margin-bottom: 18px;
        position: relative; overflow: hidden;
        box-shadow: 0 3px 14px rgba(15,61,62,.06);
        text-align: center;
    }
    .card-head::before {
        content: ""; position: absolute; top: 0; left: 0; right: 0;
        height: 3px;
        background: linear-gradient(90deg, #0f3d3e, #c9a961, #0f3d3e);
    }
    .card-head h1 {
        font-size: 27px; font-weight: 700; color: #0f3d3e; margin: 0;
        font-family: "Noto Serif SC", Georgia, serif;
        letter-spacing: .6px;
    }
    .card-head .sub {
        color: #5a5a5a; font-size: 13.5px; margin-top: 12px;
        line-height: 1.9;
        display: flex; justify-content: center; flex-wrap: wrap;
        gap: 4px 0;
    }
    .card-head .sub .dot { color: #c9a961; margin: 0 10px; }

    .section-title {
        font-size: 14.5px; font-weight: 600; color: #0f3d3e;
        padding: 0 0 12px 0;
        border-bottom: 1px solid #ebe5d8;
        margin-bottom: 16px;
        font-family: "Noto Serif SC", Georgia, serif;
        letter-spacing: .4px;
        display: flex; align-items: center; gap: 8px;
    }

    #main_row {
        gap: 18px !important;
        align-items: stretch !important;
    }
    #main_row > .gr-column,
    #main_row > div {
        flex: 1 1 0 !important;
        min-width: 0 !important;
    }
    #setup_col, #status_col {
        background: #ffffff !important;
        border: 1px solid #ebe5d8 !important;
        border-radius: 18px !important;
        padding: 22px 24px !important;
        box-shadow: 0 3px 14px rgba(15,61,62,.05);
        display: flex !important;
        flex-direction: column !important;
        min-height: 620px;
    }
    #status_col { background: #fdfcf8 !important; }

    @media (max-width: 960px) {
        #main_row > .gr-column,
        #main_row > div { min-width: 100% !important; }
        #setup_col, #status_col { min-height: auto; }
    }

    #log_row {
        margin-top: 18px !important;
        background: #ffffff !important;
        border: 1px solid #ebe5d8 !important;
        border-radius: 18px !important;
        padding: 22px 24px !important;
        box-shadow: 0 3px 14px rgba(15,61,62,.05);
    }

    #preview_row {
        margin-top: 18px !important;
        background: #ffffff !important;
        border: 1px solid #ebe5d8 !important;
        border-radius: 18px !important;
        padding: 22px 24px !important;
        box-shadow: 0 3px 14px rgba(15,61,62,.05);
    }

    label span, .gr-box > label > span {
        color: #1a1a1a !important;
        font-size: 13.5px !important;
        font-weight: 500 !important;
    }
    .gradio-container input:not([type="checkbox"]):not([type="radio"]),
    .gradio-container textarea,
    .gradio-container select {
        background: #ffffff !important; color: #111 !important;
        border: 1px solid #e2dccb !important; font-size: 14px !important;
        border-radius: 9px !important;
    }
    .gradio-container input:focus,
    .gradio-container textarea:focus {
        border-color: #c9a961 !important;
        box-shadow: 0 0 0 3px rgba(201,169,97,.15) !important;
    }
    .gradio-container .block {
        background: transparent !important;
        border: none !important;
    }
    #setup_col .block, #status_col .block,
    #log_row .block, #preview_row .block {
        border: none !important;
        box-shadow: none !important;
    }

    #file_mode { padding: 4px 0 !important; }
    #file_mode .wrap, #file_mode > div > div {
        display: flex !important; flex-direction: row !important;
        gap: 20px !important; flex-wrap: wrap !important;
    }
    #file_mode label {
        font-size: 14px !important; font-weight: 500 !important;
        cursor: pointer !important;
    }
    #file_mode input[type="radio"] {
        -webkit-appearance: radio !important; appearance: radio !important;
        width: 16px !important; height: 16px !important;
        min-width: 16px !important; max-width: 16px !important;
        accent-color: #0f3d3e !important; margin-right: 7px !important;
    }

    #trial_cb, #terms_cb, #bi_cb {
        background: transparent !important; border: none !important;
        padding: 4px 2px !important;
    }
    #trial_cb *, #terms_cb *, #bi_cb * { cursor: pointer !important; }
    #trial_cb input[type="checkbox"], #terms_cb input[type="checkbox"],
    #bi_cb input[type="checkbox"] {
        -webkit-appearance: checkbox !important; appearance: checkbox !important;
        width: 16px !important; height: 16px !important;
        min-width: 16px !important; max-width: 16px !important;
        accent-color: #0f3d3e !important; margin-right: 9px !important;
    }
    #trial_cb label, #terms_cb label, #bi_cb label {
        cursor: pointer !important; font-size: 13.5px !important;
    }

    #pdf_upload {
        min-height: 100px !important;
        max-height: 170px !important;
        overflow-y: auto !important;
        box-sizing: border-box !important;
        border: 1.5px dashed #e2dccb !important;
        border-radius: 12px !important;
        background: #fdfcf8 !important;
        transition: border-color .2s, background .2s;
    }
    #pdf_upload:hover {
        border-color: #c9a961 !important;
        background: #faf8f2 !important;
    }
    #pdf_upload > div { padding: 10px 14px !important; }
    #pdf_upload button {
        padding: 7px 18px !important; font-size: 13px !important;
        min-height: 34px !important; cursor: pointer !important;
        background: #f5f1e6 !important; border: 1px solid #e2dccb !important;
        color: #0f3d3e !important; border-radius: 8px !important;
        font-weight: 500 !important;
    }
    #pdf_upload button:hover { background: #ebe5d5 !important; }
    #pdf_upload .file {
        padding: 7px 11px !important; margin: 5px 0 !important;
        font-size: 13px !important; background: #f5f1e6 !important;
        border-radius: 8px !important;
        border: 1px solid #ebe5d8 !important;
    }

    #action_grid { margin-top: 4px; }
    #action_grid .gr-row,
    #action_grid .row {
        display: grid !important;
        grid-template-columns: 1fr 1fr !important;
        gap: 10px !important;
        margin: 0 !important;
    }
    #action_grid .gr-row > *,
    #action_grid .row > * {
        min-width: 0 !important;
        width: 100% !important;
    }
    #action_grid button {
        width: 100% !important;
        height: 46px !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
        letter-spacing: .3px;
        transition: transform .08s, box-shadow .15s, background .15s;
    }
    #action_grid button:hover { transform: translateY(-1px); }
    #action_grid .primary {
        background: linear-gradient(135deg, #0f3d3e, #1f5b5c) !important;
        color: #faf8f2 !important;
        border: none !important;
        box-shadow: 0 3px 12px rgba(15,61,62,.28);
    }
    #action_grid .primary:hover {
        background: linear-gradient(135deg, #0a2e2f, #0f3d3e) !important;
        box-shadow: 0 5px 16px rgba(15,61,62,.38);
    }
    #action_grid .secondary {
        background: #f5f1e6 !important; color: #0f3d3e !important;
        border: 1px solid #e2dccb !important;
    }
    #action_grid .secondary:hover {
        background: #ebe5d5 !important;
        border-color: #c9a961 !important;
    }

    #stop_one_row {
        gap: 10px !important; align-items: stretch !important;
    }
    #stop_dd { flex: 1 1 auto !important; min-width: 0 !important; }
    #stop_one_btn {
        flex: 0 0 auto !important; min-width: 128px !important;
        height: 44px !important; border-radius: 10px !important;
        background: linear-gradient(135deg, #8b3a3a, #6b2222) !important;
        color: #fff !important;
        border: none !important; font-size: 14px !important;
        font-weight: 600 !important;
        transition: transform .08s, background .15s;
    }
    #stop_one_btn:hover {
        background: linear-gradient(135deg, #6b2222, #4a1414) !important;
        transform: translateY(-1px);
    }

    #task_log textarea {
        background: #0f1f1f !important;
        border: 1px solid #1f3838 !important;
        font-family: "SF Mono", "Consolas", "Menlo", monospace !important;
        font-size: 12px !important;
        line-height: 1.7 !important;
        color: #c9a961 !important;
        padding: 14px 16px !important;
        resize: vertical !important;
        min-height: 260px !important;
        max-height: 420px !important;
        white-space: pre !important;
        overflow-y: auto !important;
        border-radius: 12px !important;
    }
    #task_log {
        background: transparent !important;
        border: none !important;
    }

    #preview_box::-webkit-scrollbar { width: 10px; }
    #preview_box::-webkit-scrollbar-track {
        background: #e8e1cc; border-radius: 5px;
    }
    #preview_box::-webkit-scrollbar-thumb {
        background: #c9a961; border-radius: 5px;
    }

    .gradio-container .accordion-header {
        background: #f5f1e6 !important; color: #0f3d3e !important;
        font-size: 14px !important; font-weight: 500 !important;
        border-radius: 12px !important;
        border: 1px solid #ebe5d8 !important;
    }
    .gradio-container h3 {
        color: #0f3d3e !important; font-size: 18px !important;
        margin-top: 20px !important;
    }
    .gradio-container .prose, .gradio-container .prose * {
        color: #1a1a1a !important;
    }

    .gradio-container .tab-nav {
        border-bottom: 2px solid #ebe5d8 !important;
        margin-bottom: 12px !important;
    }
    .gradio-container .tab-nav button {
        font-size: 14.5px !important; font-weight: 600 !important;
        color: #8b8578 !important;
        padding: 10px 22px !important;
        border-radius: 10px 10px 0 0 !important;
        transition: color .15s, background .15s;
    }
    .gradio-container .tab-nav button.selected {
        color: #0f3d3e !important;
        background: #f5f1e6 !important;
    }
    .gradio-container .tab-nav button:hover {
        color: #0f3d3e !important;
        background: #faf8f2 !important;
    }

    #open_hint {
        min-height: 0 !important;
        margin-top: 8px !important;
    }
    #open_hint p {
        font-size: 12.5px !important;
        color: #0f3d3e !important;
        margin: 4px 0 0 0 !important;
        padding: 7px 12px !important;
        background: #f5f1e6 !important;
        border-left: 3px solid #c9a961 !important;
        border-radius: 6px !important;
        word-break: break-all;
    }

    #quality_hint {
        min-height: 0 !important;
        margin: -10px 0 4px 0 !important;
    }
    #quality_hint p {
        font-size: 12px !important;
        color: #8b8578 !important;
        margin: 0 !important;
        padding: 0 4px !important;
        line-height: 1.5 !important;
    }
    """,
) as demo:

    modal_html = gr.HTML(value="", elem_id="modal_host")

    gr.HTML("""
    <div class="card-head">
      <h1>📖 PDF / Word / PPT 翻译器</h1>
      <div class="sub">
        保留原排版<span class="dot">·</span>多任务并行
        <span class="dot">·</span>断点可续<span class="dot">·</span>
        12 种语言<span class="dot">·</span>术语笔记
        <span class="dot">·</span>质量可选
      </div>
    </div>
    """)

    with gr.Accordion("💡 使用说明 / 输出文件说明（点击展开）", open=False):
        gr.HTML("""
        <div style="padding:14px 20px;font-size:13.5px;line-height:2;color:#444;
                    background:#fdfcf8;border:1px solid #ebe5d8;border-radius:12px">
          <div><b>🔑 密钥</b>　留空取 <code>.env</code> 中的默认值；填写 → 临时覆盖</div>
          <div><b>🌐 语言</b>　简体/繁体中文、英、日、韩、法、德、西、葡、俄、阿、意</div>
          <div><b>🔤 字体</b>　自动扫描 <code>D:\\file\\translate\\word_type</code> 目录下的字体（仅对 PDF 生效）</div>
          <div><b>📏 字号</b>　正文字号上限，实际会根据原文框自动缩小</div>
          <div><b>🎨 质量</b>　仅对 PDF 的 <code>bilingual.pdf</code> 生效；取消勾选可完全跳过双语 PDF（体积减少约 90%）</div>
          <div><b>📓 术语表</b>　勾选后，术语页会插在 PDF 正文前（带书签）；notes.html 支持鼠标悬停查词 + 点击定位到原文页</div>
          <div><b>📁 输出位置</b>　译文在 <code>trans_result/&lt;文件名&gt;/</code>，笔记在 <code>_&lt;文件名&gt;/</code></div>
          <div style="margin-top:10px;padding-top:10px;border-top:1px dashed #e2dccb">
            <b>📚 三种成品</b>
            <div style="margin-left:16px">
              · <code>translated.pdf</code>　纯译文<br>
              · <code>translated_with_notes.pdf</code>　术语页 + 译文<br>
              · <code>bilingual.pdf</code>　术语页 + 左右对照（可关闭）
            </div>
          </div>
        </div>
        """)

    # ================= 上部：左设置 / 右状态 =================
    with gr.Row(equal_height=False, elem_id="main_row"):

        with gr.Column(scale=1, min_width=440, elem_id="setup_col"):
            gr.HTML('<div class="section-title">⚙️ 翻译设置</div>')

            with gr.Row(equal_height=True):
                api_key = gr.Textbox(
                    label="🔑 DeepSeek API Key",
                    type="password",
                    placeholder="留空用 .env",
                    scale=3,
                )
                model = gr.Dropdown(
                    choices=["deepseek-chat", "deepseek-reasoner"],
                    value=DEFAULT_MODEL,
                    label="🧠 模型",
                    scale=2,
                    elem_id="model_dd",
                )

            target_lang = gr.Dropdown(
                choices=[(v, k) for k, v in LANG_NAMES.items()],
                value="zh-CN",
                label="🌐 目标语言",
                elem_id="lang_dd",
            )

            file_mode = gr.Radio(
                choices=["📕 PDF 书籍", "📘 Word 文档", "📊 PPT 演示"],
                value="📕 PDF 书籍",
                label="📂 文档类型（仅参考，实际按后缀自动判断）",
                elem_id="file_mode",
            )
            doc_file = gr.File(
                label="📄 上传文档（.pdf / .docx / .pptx）",
                file_types=[".pdf", ".docx", ".pptx"],
                elem_id="pdf_upload",
            )

            _font_choices = sorted(FONTS_MAP.keys())
            with gr.Row(equal_height=True):
                font_dd = gr.Dropdown(
                    choices=_font_choices if _font_choices
                            else ["(无可用字体，用内置宋体)"],
                    value=_default_font_display() if _font_choices
                          else "(无可用字体，用内置宋体)",
                    label="🔤 正文字体（仅 PDF）",
                    interactive=True,
                    elem_id="font_dd",
                    scale=3,
                )
                fontsize_dd = gr.Dropdown(
                    choices=[(name, val) for name, val in FONT_SIZE_CHOICES],
                    value=DEFAULT_FONT_SIZE,
                    label="📏 字号上限",
                    interactive=True,
                    elem_id="fontsize_dd",
                    scale=2,
                )

            with gr.Row(equal_height=True):
                pdf_quality_dd = gr.Dropdown(
                    choices=list(PDF_QUALITY_PRESETS.keys()),
                    value=DEFAULT_PDF_QUALITY,
                    label="🎨 双语 PDF 质量（仅 PDF）",
                    interactive=True,
                    elem_id="quality_dd",
                    scale=3,
                )
                make_bilingual_cb = gr.Checkbox(
                    label="生成左右对照双语 PDF",
                    value=True,
                    elem_id="bi_cb",
                    scale=2,
                )

            quality_hint = gr.Markdown(
                value=f"💡 {PDF_QUALITY_PRESETS[DEFAULT_PDF_QUALITY]['hint']}",
                elem_id="quality_hint",
            )

            with gr.Row(equal_height=True):
                want_terms_cb = gr.Checkbox(
                    label="📓 生成术语表与阅读笔记",
                    value=True,
                    elem_id="terms_cb",
                    scale=1,
                )
                trial = gr.Checkbox(
                    label="🧪 试翻（PDF 只翻前 5 页）",
                    value=False,
                    elem_id="trial_cb",
                    scale=1,
                )

            reader_profile = gr.Textbox(
                label="👤 读者背景（决定术语说明深浅，可留空）",
                placeholder="例如：有 Python 基础，但没接触过机器学习",
                lines=1,
            )

            with gr.Column(elem_id="action_grid"):
                with gr.Row(equal_height=True):
                    btn = gr.Button("▶ 开始翻译", variant="primary")
                    stop_btn = gr.Button("⏹ 全部停止", variant="secondary")
                    refresh_btn = gr.Button("🔄 刷新", variant="secondary")
                    open_btn = gr.Button("📁 打开目录", variant="secondary")

            open_hint = gr.Markdown(value="", elem_id="open_hint")

        with gr.Column(scale=1, min_width=440, elem_id="status_col"):
            gr.HTML('<div class="section-title">📋 任务列表</div>')
            task_list_html = gr.HTML(
                value='<div style="padding:18px;color:#a9a49a;font-size:13px;'
                      'text-align:center">暂无任务</div>'
            )

            gr.HTML('<div class="section-title" style="margin-top:18px">'
                    '🛑 单独停止某个任务</div>')
            with gr.Row(elem_id="stop_one_row"):
                stop_dd = gr.Dropdown(
                    label="选择要停止的任务（只列运行中）",
                    choices=[],
                    value=None,
                    interactive=True,
                    elem_id="stop_dd",
                )
                stop_one_btn = gr.Button("⏹ 停止选中", elem_id="stop_one_btn")

            gr.HTML('<div class="section-title" style="margin-top:18px">'
                    '📊 当前进度</div>')
            progress_bar = gr.HTML(value=make_progress_html(0, 1, "等待开始"))

    # ================= 中部：任务日志（全宽） =================
    with gr.Column(elem_id="log_row"):
        gr.HTML('<div class="section-title">📋 任务日志</div>')
        log = gr.Textbox(
            label="",
            lines=14,
            interactive=False,
            show_label=False,
            elem_id="task_log",
        )

    # ================= 下部：效果预览 / 下载文件（全宽 Tabs） =================
    with gr.Column(elem_id="preview_row"):
        with gr.Tabs():
            with gr.TabItem("👀 效果预览"):
                gallery = gr.HTML(
                    value=build_preview_html([]),
                    elem_id="preview_box",
                )
                load_btn = gr.Button(
                    "🔍 加载当前任务的预览图和下载文件",
                    variant="secondary",
                )
            with gr.TabItem("💾 下载文件"):
                out_files = gr.File(
                    label="", file_count="multiple",
                    interactive=False, show_label=False,
                )

    fast_outputs = [task_list_html, progress_bar, log, stop_dd, gallery, modal_html]
    full_outputs = [task_list_html, progress_bar, log, stop_dd,
                    gallery, modal_html, out_files, doc_file]

    btn.click(
        on_start,
        [api_key, model, doc_file, file_mode, trial, target_lang,
         reader_profile, want_terms_cb, font_dd, fontsize_dd,
         pdf_quality_dd, make_bilingual_cb, stop_dd],
        full_outputs,
        concurrency_limit=3,
        concurrency_id="start",
    )
    stop_btn.click(on_stop_all, [stop_dd], fast_outputs,
                   concurrency_limit=None, concurrency_id="stop")
    stop_one_btn.click(on_stop_selected, [stop_dd], fast_outputs,
                       concurrency_limit=None, concurrency_id="stop_one")
    refresh_btn.click(on_refresh_fast, [stop_dd], fast_outputs,
                      concurrency_limit=None, concurrency_id="manual")
    open_btn.click(open_result_folder, None, [open_hint])
    load_btn.click(on_load_preview, None, [gallery, out_files])

    pdf_quality_dd.change(
        lambda name: f"💡 {PDF_QUALITY_PRESETS.get(name, {}).get('hint', '')}",
        [pdf_quality_dd], [quality_hint],
    )

    try:
        timer = gr.Timer(3.0)
        timer.tick(on_refresh_fast, [stop_dd], fast_outputs,
                   concurrency_limit=None, concurrency_id="tick")
    except Exception:
        pass

    demo.load(on_refresh_fast, [stop_dd], fast_outputs,
              concurrency_limit=None, concurrency_id="load")


# ============================================================
# main
# ============================================================

def _find_free_port(start=7860, end=7899):   # ★ 优化：扩展端口搜索范围
    for p in range(start, end + 1):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                s.bind(("127.0.0.1", p))
                return p
        except OSError:
            continue
    return start


if __name__ == "__main__":
    PORT = _find_free_port(7860, 7899)
    URL = "http://127.0.0.1:" + str(PORT)

    try:
        with open("访问网址.txt", "w", encoding="utf-8") as f:
            f.write("PDF / Word / PPT 翻译器\n\n")
            f.write("浏览器访问网址：\n")
            f.write(URL + "\n\n")
            f.write("（这个文件由程序自动生成，改动此文件无效）\n")
            f.write("（如果打不开，说明程序已停止，请双击 重启翻译器.bat）\n")
    except Exception:
        pass

    print()
    print("=" * 64)
    print("  PDF / Word / PPT 翻译器 已启动")
    print()
    print("  在浏览器输入以下网址进入：")
    print()
    print("      " + URL)
    print()
    print("  别关这个终端窗口，否则程序停止")
    print("  网址也保存在 访问网址.txt 中，随时可查")
    print("=" * 64)
    print()

    print(f"   结果目录：{os.path.abspath(RESULT_ROOT)}")
    if DEFAULT_API_KEY:
        print(f"   API Key：已从 .env 读取（{DEFAULT_API_KEY[:6]}...）")
    else:
        print("   API Key：未配置，需在网页填写")

    print(f"   字体扫描目录：")
    for _root in FONT_ROOTS:
        exists = "v" if os.path.isdir(_root) else "x"
        print(f"      [{exists}] {_root}")
    print(f"   已找到字体：{len(FONTS_MAP)} 个")
    if FONTS_MAP:
        _default = _default_font_display()
        print(f"   默认字体：{_default}")

    print(f"   双语 PDF 质量预设：{len(PDF_QUALITY_PRESETS)} 档")
    for _name, _cfg in PDF_QUALITY_PRESETS.items():
        print(f"      · {_name}（zoom={_cfg['zoom']}, JPEG={_cfg['jpeg']}）")

    if not HAS_OFFICE:
        print("   [!] 未装 python-docx / python-pptx，Word / PPT 不可用")
    if not HAS_NOTES:
        print("   [!] 未找到 notes_builder.py，术语表功能不可用")
    else:
        print("   术语表模块：notes_builder.py 已加载")
    print()

    n_removed = cleanup_orphan_files()
    if n_removed:
        print(f"   已清理 {n_removed} 个冗余/临时文件")
    scan_all_states()
    n = len(MANAGER.all_sorted())
    if n:
        print(f"   已恢复 {n} 个历史任务")
    print()

    demo.queue(default_concurrency_limit=None)
    demo.launch(
        server_name="127.0.0.1",
        server_port=PORT,
        inbrowser=True,
        show_error=False,
        quiet=True,
        allowed_paths=[os.path.abspath(RESULT_ROOT)],
    )
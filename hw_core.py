# -*- coding: utf-8 -*-
"""作业解题器 · 基础设施：环境变量 / 常量 / 工具 / JSON 节流 / 字体扫描。"""

import sys
sys.dont_write_bytecode = True

import os
import json
import time
import re
import glob
import hashlib
import threading

import pymupdf as fitz
try:
    fitz.TOOLS.mupdf_display_errors(False)
except Exception:
    pass

from dotenv import load_dotenv

# ================= 可选依赖：Word / PPT =================
# ★ 这些导入用于探测依赖是否安装（HAS_OFFICE），名字由 hw_tasks.py 自行导入
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
    Document = qn = OxmlElement = Paragraph = Presentation = None
    MSO_SHAPE_TYPE = None
    HAS_OFFICE = False

# 标记为已使用（仅探测，实际使用见 hw_tasks.py）
_ = (Document, qn, OxmlElement, Paragraph, Presentation, MSO_SHAPE_TYPE)

# ================= .env =================
load_dotenv()
DEFAULT_API_KEY = os.getenv("DEEPSEEK_API_KEY", "").strip()
DEFAULT_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat").strip()
if DEFAULT_MODEL not in ("deepseek-chat", "deepseek-reasoner"):
    DEFAULT_MODEL = "deepseek-chat"
# ★ 修复：base_url / 超时进 .env，不再硬编码
DEEPSEEK_BASE_URL = os.getenv(
    "DEEPSEEK_BASE_URL", "https://api.deepseek.com"
).rstrip("/")
try:
    API_TIMEOUT = float(os.getenv("DEEPSEEK_TIMEOUT", "180"))
except Exception:
    API_TIMEOUT = 180.0

# ================= 配置 =================
RESULT_ROOT = "homework_result"
RENDER_ZOOM = 2.0
PREVIEW_PAGES = 5
PREVIEW_PARAS = 5
PREVIEW_MAX_WIDTH = 1400
PREVIEW_JPEG_QUALITY = 88
CHECKPOINT_EVERY = 10

# ★ 修复：Word/PPT 批量解题（对齐双语版的批量翻译）
OFFICE_BATCH_SIZE = 20

# ★ 修复：试解页数可选（对齐双语版 3/5/10/20）
TRIAL_PAGE_CHOICES = [
    ("3 页", 3),
    ("5 页（默认）", 5),
    ("10 页", 10),
    ("20 页", 20),
]

# ★ 优化：公式渲染 dpi（提高清晰度；内存、速度会有一定增加）
_RENDER_DPI = 300

# ★ 优化：日志时间戳格式
_LOG_TS_FMT = "%H:%M:%S"

# ★ 优化：state 保存节流 / 大小刷新节流
STATE_SAVE_INTERVAL = 1.5
SIZE_REFRESH_INTERVAL = 5.0

# ★ 修复：缓存 JSON 写节流（避免每段/每页整文件重写放大 IO）
_JSON_DEBOUNCE_SEC = 3.0
_JSON_PENDING = {}
_JSON_LAST_WRITE = {}

# ★ 优化：答案位置可选
ANS_POS_CHOICES = [
    ("块内右下（默认）", "inside"),
    ("块正下方", "below"),
    ("块右侧留白（不够时退回下方）", "right"),
]

# ★ 优化：学科微调 Prompt
SUBJECT_CHOICES = [
    ("通用（默认）", "general"),
    ("数学", "math"),
    ("英语", "english"),
    ("物理", "physics"),
    ("化学", "chemistry"),
    ("语文", "chinese"),
]

# ★ 优化：任务列表筛选
TASK_FILTER_CHOICES = [
    ("全部", "all"),
    ("运行中", "running"),
    ("已完成", "done"),
    ("出错", "error"),
]

_IO_LOCK = threading.Lock()

# ================= 字体自动扫描 =================
_HERE = os.path.dirname(os.path.abspath(__file__))

FONT_ROOTS = [
    r"D:\file\translate\word_type",
    os.path.join(_HERE, "fonts"),
]

_FONT_EXT = (".otf", ".ttf", ".ttc", ".otc")
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


def _pick_cn_font_path():
    """
    ★ 优化：优先读 HOMEWORK_FONT_PATH，其次 TRANSLATE_FONT_PATH，
    避免和翻译器共用一个变量。
    """
    env_path = (os.getenv("HOMEWORK_FONT_PATH", "").strip()
                or os.getenv("TRANSLATE_FONT_PATH", "").strip())
    if env_path and os.path.exists(env_path):
        return env_path
    known = r"D:\file\translate\word_type\09_SourceHanSerifSC\OTF\SimplifiedChinese\SourceHanSerifSC-Regular.otf"
    if os.path.exists(known):
        return known
    if FONTS_MAP:
        for key in sorted(FONTS_MAP.keys()):
            low = key.lower()
            if "sourcehanserifsc" in low and "regular" in low:
                return FONTS_MAP[key]
        for key in sorted(FONTS_MAP.keys()):
            low = key.lower()
            if ("sourcehan" in low or "notosans" in low) and "sc" in low:
                return FONTS_MAP[key]
        return FONTS_MAP[sorted(FONTS_MAP.keys())[0]]
    return ""


CN_FONT_PATH = _pick_cn_font_path()


# ============================================================
# 工具
# ============================================================
def h(s):
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:20]


def safe_dirname(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name or "").strip()
    name = name.replace("..", "_")
    name = name.rstrip(". ")
    return name[:80] or "untitled"


def load_json_file(path, default):
    if path and os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[warn] 读取 JSON 失败 {path}: {e}")
            return default
    return default


def _write_json_file(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with _IO_LOCK:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        os.replace(tmp, path)


def save_json_file(path, data, force=False):
    """★ 修复：节流写 JSON——同一路径 3 秒内只记脏引用，
    超时或 force=True 才真正落盘，避免每页/每段整文件重写放大 IO。"""
    if not path:
        return
    now = time.monotonic()
    with _IO_LOCK:
        last = _JSON_LAST_WRITE.get(path)
        if not force and last is not None and (now - last) < _JSON_DEBOUNCE_SEC:
            _JSON_PENDING[path] = data
            return
        _JSON_LAST_WRITE[path] = now
    _write_json_file(path, data)


def flush_all_json():
    """把节流中未落盘的 JSON 全部写盘（worker 结束 / 程序退出时调用）。"""
    with _IO_LOCK:
        pending = list(_JSON_PENDING.items())
        _JSON_PENDING.clear()
    for path, data in pending:
        try:
            _write_json_file(path, data)
        except Exception as e:
            print(f"[warn] flush JSON 失败 {path}: {e}")


def load_progress_file(path):
    d = load_json_file(path, {})
    return set(int(x) for x in d.get("done_pages", []))


def save_progress_file(path, done_pages):
    save_json_file(path, {"done_pages": sorted(int(x) for x in done_pages)})


def clear_dir_preview(folder):
    for pat in ("*.png", "*.jpg", "*.jpeg"):
        for f in glob.glob(os.path.join(folder, pat)):
            try:
                os.remove(f)
            except Exception:
                pass


def safe_save_pdf(doc, out_path, retries=5):
    tmp_path = out_path + ".tmp.pdf"
    try:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    except Exception:
        pass
    doc.save(tmp_path, deflate=True)
    last_err = None
    for attempt in range(retries):
        try:
            os.replace(tmp_path, out_path)
            return out_path
        except (PermissionError, OSError) as e:
            last_err = e
            time.sleep(0.8 * (attempt + 1))
    # ★ 修复：备用名统一为 .bak.pdf（对齐双语版，可被孤儿清理识别）
    if os.path.basename(out_path).endswith(".bak.pdf"):
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        raise last_err if last_err else RuntimeError("无法保存 PDF")

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


def prepare_paths(src_path):
    base = os.path.basename(src_path)
    book, ext = os.path.splitext(base)
    book = safe_dirname(book)
    short_hash = h(base)[:6]
    out_dir = os.path.abspath(
        os.path.join(RESULT_ROOT, f"{book}_{short_hash}")
    )
    work = os.path.join(out_dir, "_work")
    input_dir = os.path.join(work, "input")
    os.makedirs(work, exist_ok=True)
    os.makedirs(input_dir, exist_ok=True)
    return {
        "out_dir": out_dir, "work": work, "input_dir": input_dir,
        "solved_pdf":   os.path.join(out_dir, "solved.pdf"),
        "analysis_md":  os.path.join(out_dir, "analysis.md"),
        "cache_file":   os.path.join(work, "solve_cache.json"),
        "preview_dir":  os.path.join(work, "preview"),
        "progress_file": os.path.join(work, "progress.json"),
        "office_out":   os.path.join(out_dir, f"{book}_solved{ext.lower()}"),
    }

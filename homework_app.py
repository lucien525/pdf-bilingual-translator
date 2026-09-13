# -*- coding: utf-8 -*-
"""
PDF / Word / PPT 作业解题器
- 答案直接写在题目块内（红色）
- 题目下方给出详细解析（蓝色，浅底）
- 数学公式用 matplotlib mathtext 渲染成 PNG 贴入
- 断点续传 / 试解 5 页 / 每页落盘
- UI 与 bilingual_app.py 一致（含完成模态弹窗）
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
import socket
import threading
import uuid
import io
from dataclasses import dataclass, field
from typing import Optional
from html import escape

import numpy as np
import pymupdf as fitz
try:
    fitz.TOOLS.mupdf_display_errors(False)
except Exception:
    pass

from PIL import Image
from openai import OpenAI
import gradio as gr
from dotenv import load_dotenv

import matplotlib
matplotlib.use("Agg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib import font_manager as mfont
matplotlib.rcParams["mathtext.fontset"] = "cm"
matplotlib.rcParams["axes.unicode_minus"] = False

_RESAMPLE = getattr(Image, "Resampling", Image).LANCZOS

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    from pptx import Presentation
    HAS_OFFICE = True
except ImportError:
    HAS_OFFICE = False

# ================= .env =================
load_dotenv()
DEFAULT_API_KEY = os.getenv("DEEPSEEK_API_KEY", "").strip()
DEFAULT_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat").strip()
if DEFAULT_MODEL not in ("deepseek-chat", "deepseek-reasoner"):
    DEFAULT_MODEL = "deepseek-chat"

# ================= 配置 =================
RESULT_ROOT = "homework_result"
RENDER_ZOOM = 2.0
PREVIEW_PAGES = 5
PREVIEW_PARAS = 5
PREVIEW_MAX_WIDTH = 1400
PREVIEW_JPEG_QUALITY = 88
CHECKPOINT_EVERY = 10

ANS_COLOR = "#c0392b"
SOL_COLOR = "#1a3a8a"
SOL_BG = (0.94, 0.96, 1.0)
ANS_BG = None

_IO_LOCK = threading.Lock()
_CREATE_LOCK = threading.Lock()
_MPL_LOCK = threading.Lock()
_LAST_PREVIEW_SIG = {}
_PREVIEW_HTML_CACHE = {}
_MODAL_SHOWN = set()

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
    env_path = os.getenv("TRANSLATE_FONT_PATH", "").strip()
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

CN_FONT = None
if CN_FONT_PATH and os.path.exists(CN_FONT_PATH):
    try:
        mfont.fontManager.addfont(CN_FONT_PATH)
        CN_FONT = mfont.FontProperties(fname=CN_FONT_PATH)
        print(f"[font] 已加载中文字体：{CN_FONT_PATH}")
    except Exception as e:
        print(f"[warn] 中文字体注册失败：{e}")
else:
    print("[warn] 未找到中文字体，中文可能显示为方框")
    print(f"       扫描目录：{FONT_ROOTS}")


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
    fallback = out_path[:-4] + f"_new_{int(time.time())}.pdf"
    try:
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
        "cache_file":   os.path.join(work, "solve_cache.json"),
        "preview_dir":  os.path.join(work, "preview"),
        "progress_file": os.path.join(work, "progress.json"),
        "office_out":   os.path.join(out_dir, f"{book}_solved{ext.lower()}"),
    }


# ============================================================
# matplotlib 渲染
# ============================================================
def render_rich_text_png(text, width_pt, fontsize=11, dpi=200, color="#c0392b"):
    text = (text or "").strip()
    if not text:
        return None, 0.0, 0.0

    width_inch = max(0.5, float(width_pt) / 72.0)

    with _MPL_LOCK:
        fig = Figure(figsize=(width_inch, 40), dpi=dpi)
        fig.patch.set_alpha(0.0)
        canvas = FigureCanvasAgg(fig)
        ax = fig.add_axes([0.0, 0.0, 1.0, 1.0])
        ax.patch.set_alpha(0.0)
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")

        kw = dict(
            fontsize=fontsize,
            color=color,
            va="top", ha="left",
            wrap=True,
            linespacing=1.55,
        )
        if CN_FONT is not None:
            kw["fontproperties"] = CN_FONT

        t = ax.text(0.0, 1.0, text, **kw)

        try:
            canvas.draw()
            renderer = canvas.get_renderer()
            bbox = t.get_window_extent(renderer=renderer)
        except Exception:
            return None, 0.0, 0.0

        buf = np.asarray(canvas.buffer_rgba())
        H, W = buf.shape[:2]

        pad = 3
        x0 = max(0, int(bbox.x0) - pad)
        x1 = min(W, int(bbox.x1) + pad)
        y_top = max(0, H - int(bbox.y1) - pad)
        y_bot = min(H, H - int(bbox.y0) + pad)

        if x1 <= x0 or y_bot <= y_top:
            return None, 0.0, 0.0

        crop = np.ascontiguousarray(buf[y_top:y_bot, x0:x1])
        img = Image.fromarray(crop, mode="RGBA")

        out = io.BytesIO()
        img.save(out, "PNG")

        w_pt = (x1 - x0) / dpi * 72.0
        h_pt = (y_bot - y_top) / dpi * 72.0
        return out.getvalue(), w_pt, h_pt


# ============================================================
# Prompt
# ============================================================
def build_solve_system():
    return (
        "你是一位经验丰富的中学/大学老师，擅长解答各科作业"
        "（数学、英语、物理、化学、生物、语文等）。\n"
        "用户会给你一页作业的多个段落，每段以 [[B0]] [[B1]] [[B2]] ... 开头。\n\n"
        "【任务】对每一段判断：\n"
        "A. 若是题目（选择题、填空题、解答题、完形填空、翻译、"
        "阅读理解小题、口语问答、语法练习等）→ 给出答案和解析\n"
        "B. 若不是题目（标题、页眉、页脚、说明文字、图片说明、题号等）"
        "→ 直接跳过，不输出这一段\n\n"
        "【输出格式】只输出 A 类段落，每段格式：\n"
        "[[B0]]\n<ANS>答案</ANS>\n<SOL>解析</SOL>\n"
        "[[B3]]\n<ANS>答案</ANS>\n<SOL>解析</SOL>\n"
        "（段号与输入一致，只输出有答案的段号，跳过的段不出现）\n\n"
        "【答案规范】\n"
        "- 选择题：给选项字母 + 选项内容，如「B. 因为...」\n"
        "- 填空题：给填空内容，多个空用「；」分隔\n"
        "- 解答题：给最终结果（数值 / 表达式）\n"
        "- 英语翻译：给译文\n"
        "- 英语口语问答：给一个自然回答，如「I'm fine, thank you.」\n"
        "- 阅读理解：给选项字母\n\n"
        "【解析规范】\n"
        "- 2~6 句，讲清思路和步骤\n"
        "- 选择题说明为什么选它、为什么排除其他\n"
        "- 数学题给出关键中间步骤\n\n"
        "【严格约束】\n"
        "- 不要输出空的 <ANS></ANS> 或 <SOL></SOL>\n"
        "- 判断不出是不是题目的，视为 A 类，给一个答案\n"
        "- 只输出 <ANS> 没解析也行，反之亦然\n"
        "- 不要 Markdown 标记，不要代码块\n"
        "- 不要前言、后记、总结\n\n"
        "【数学公式】\n"
        "- 行内用 $...$，独立用 $$...$$\n"
        "- 只能用 mathtext 支持的：上标 ^ 下标 _ \\frac{}{} \\sqrt{} "
        "\\sum \\int \\prod \\lim \\alpha \\beta \\pi \\sigma \\omega "
        "\\cdot \\times \\div \\pm \\leq \\geq \\neq \\approx \\infty "
        "\\rightarrow \\sin \\cos \\tan \\log \\ln\n"
        "- 不要在 $...$ 里放中文\n"
    )


# ============================================================
# API
# ============================================================
def _api_call(client, model, system, user_content, retries=4):
    last_err = ""
    for attempt in range(retries):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user_content},
                ],
                temperature=0.2,
            )
            return resp.choices[0].message.content or ""
        except Exception as e:
            last_err = str(e)
            if "402" in last_err or "insufficient" in last_err.lower() or "余额" in last_err:
                raise RuntimeError("账户余额不足，请去 DeepSeek 平台充值后再继续")
            if "401" in last_err or "Unauthorized" in last_err or "invalid" in last_err.lower():
                raise RuntimeError("API Key 无效或已过期，请检查后重试")
            time.sleep(5 * (attempt + 1))
    raise RuntimeError(f"API 连续失败：{last_err}")


def call_solve_api(client, model, text, retries=4):
    return _api_call(client, model, build_solve_system(), text, retries)


# ============================================================
# 解析模型返回
# ============================================================
def _clean_tags(s):
    s = re.sub(r'</?ANS>', '', s or "", flags=re.I)
    s = re.sub(r'</?SOL>', '', s, flags=re.I)
    return s.strip()


def parse_solution_response(text, n_blocks):
    """解析 [[Bn]] <ANS>..</ANS> <SOL>..</SOL>"""
    result = {}
    if not text:
        return result

    parts = re.split(r'\[\[B(\d+)\]\]', text)
    for i in range(1, len(parts) - 1, 2):
        try:
            idx = int(parts[i])
        except Exception:
            continue
        body = parts[i + 1]
        ans_m = re.search(r'<ANS>(.*?)</ANS>', body, re.S | re.I)
        sol_m = re.search(r'<SOL>(.*?)</SOL>', body, re.S | re.I)
        ans = _clean_tags(ans_m.group(1) if ans_m else "")
        sol = _clean_tags(sol_m.group(1) if sol_m else "")

        # 兜底：模型只给了裸文本
        if not ans and not sol:
            stripped = _clean_tags(body)
            stripped = re.sub(r'\[\[B\d+\]\]', '', stripped).strip()
            if stripped:
                sol = stripped

        # 关键：只有真的非空才记录
        if ans or sol:
            result[idx] = {"ans": ans, "sol": sol}
    return result


def solution_is_valid(res):
    if not res:
        return False
    return bool((res.get("ans") or "").strip() or (res.get("sol") or "").strip())


# ============================================================
# PDF：解题 + 排版
# ============================================================
def solve_page(client, model, blocks, cache, cache_file):
    marked = "\n\n".join(f"[[B{i}]] {b[4].strip()}" for i, b in enumerate(blocks))
    key = "solve_" + h(marked)
    if key in cache:
        try:
            cached = {int(k): v for k, v in cache[key].items()}
            if all(isinstance(v, dict) for v in cached.values()):
                return cached
        except Exception:
            pass

    raw = call_solve_api(client, model, marked)
    parsed = parse_solution_response(raw, len(blocks))

    cache[key] = {str(k): v for k, v in parsed.items()}
    save_json_file(cache_file, cache)
    return parsed


def apply_solution_to_page(page, blocks, results):
    page_rect = page.rect

    for i, b in enumerate(blocks):
        res = results.get(i)
        if not res:
            continue
        ans = (res.get("ans") or "").strip()
        sol = (res.get("sol") or "").strip()
        if not ans and not sol:
            continue

        x0, y0, x1, y1 = float(b[0]), float(b[1]), float(b[2]), float(b[3])
        if x1 - x0 < 8 or y1 - y0 < 6:
            continue
        rect = fitz.Rect(x0, y0, x1, y1)

        # 答案
        if ans:
            png, w_pt, h_pt = render_rich_text_png(
                ans, max(20.0, rect.width - 6),
                fontsize=11, color=ANS_COLOR,
            )
            if png and w_pt > 0 and h_pt > 0:
                max_w = max(10.0, rect.width - 4)
                max_h = max(8.0, rect.height * 0.6)
                sc = min(1.0, max_w / w_pt, max_h / h_pt)
                w2 = max(6.0, w_pt * sc)
                h2 = max(6.0, h_pt * sc)
                cx = (rect.x0 + rect.x1) / 2.0
                by1 = rect.y1 - 2
                by0 = by1 - h2
                if by0 < rect.y0:
                    by0 = rect.y0
                    by1 = by0 + h2
                dst = fitz.Rect(cx - w2 / 2.0, by0, cx + w2 / 2.0, by1)
                if ANS_BG is not None:
                    try:
                        page.draw_rect(dst, color=None, fill=ANS_BG,
                                       fill_opacity=0.9, overlay=True)
                    except Exception:
                        pass
                try:
                    page.insert_image(dst, stream=png, overlay=True)
                except Exception:
                    pass

        # 解析
        if sol:
            sol_x0 = rect.x0
            sol_x1 = min(page_rect.width - 20.0,
                         max(rect.x1, rect.x0 + 260.0))
            sol_w = sol_x1 - sol_x0
            if sol_w < 60.0:
                sol_w = max(60.0, rect.width)
                sol_x1 = sol_x0 + sol_w

            png, w_pt, h_pt = render_rich_text_png(
                sol, sol_w, fontsize=9, color=SOL_COLOR,
            )
            if png and w_pt > 0 and h_pt > 0:
                sc = min(1.0, sol_w / w_pt)
                w2 = max(20.0, w_pt * sc)
                h2 = max(10.0, h_pt * sc)

                sy0 = rect.y1 + 2
                sy1 = sy0 + h2

                if sy1 > page_rect.height - 8:
                    sy1 = page_rect.height - 8
                    sy0 = max(rect.y1 + 2, sy1 - h2)
                    sy1 = sy0 + h2

                dst = fitz.Rect(sol_x0, sy0, sol_x0 + w2, sy1)

                try:
                    page.draw_rect(dst, color=None, fill=SOL_BG,
                                   fill_opacity=0.88, overlay=True)
                except Exception:
                    pass
                try:
                    page.insert_image(dst, stream=png, overlay=True)
                except Exception:
                    pass


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
    current_size: int = 0
    estimated_size: int = 0
    _last_size_ts: float = 0.0
    stop_event: threading.Event = field(default_factory=threading.Event)
    thread: Optional[threading.Thread] = None

    def to_dict(self):
        return {
            "task_id": self.task_id,
            "kind": self.kind,
            "src_name": self.src_name,
            "out_dir": self.out_dir,
            "created_at": self.created_at,
            "status": self.status if self.status != "stopping" else "paused",
            "current": self.current,
            "total": self.total,
            "label": self.label,
            "log": self.log[-60:],
            "output_files": self.output_files,
            "error": self.error,
            "current_size": self.current_size,
            "estimated_size": self.estimated_size,
        }

    def log_msg(self, m):
        self.log.append(m)
        if len(self.log) > 200:
            self.log = self.log[-200:]


class TaskManager:
    def __init__(self):
        self._tasks = {}
        self._lock = threading.Lock()

    def create(self, kind, src_path, src_name, out_dir, work_dir):
        tid = uuid.uuid4().hex[:8]
        t = TaskState(task_id=tid, kind=kind, src_path=src_path,
                      src_name=src_name, out_dir=out_dir, work_dir=work_dir)
        with self._lock:
            self._tasks[tid] = t
        return t

    def get(self, tid):
        return self._tasks.get(tid)

    def all_sorted(self):
        with self._lock:
            return sorted(self._tasks.values(), key=lambda t: -t.created_at)

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

    def load_from_disk(self, d):
        tid = d.get("task_id")
        if not tid:
            return
        if tid in self._tasks:
            return
        t = TaskState(
            task_id=tid,
            kind=d.get("kind", "pdf"),
            src_path="",
            src_name=d.get("src_name", ""),
            out_dir=d.get("out_dir", ""),
            work_dir=os.path.join(d.get("out_dir", ""), "_work"),
            created_at=d.get("created_at", time.time()),
            status="paused",
            current=d.get("current", 0),
            total=d.get("total", 0),
            label=d.get("label", ""),
            log=d.get("log", []),
            output_files=d.get("output_files", []) or [],
            error=d.get("error", ""),
            current_size=d.get("current_size", 0),
            estimated_size=d.get("estimated_size", 0),
        )
        with self._lock:
            self._tasks[tid] = t


MANAGER = TaskManager()
SELECTED_TASK_ID = None


def save_state(task):
    if task is None:
        return
    try:
        path = os.path.join(task.work_dir, "state.json")
        os.makedirs(task.work_dir, exist_ok=True)
        with _IO_LOCK:
            tmp = path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(task.to_dict(), f, ensure_ascii=False, indent=2)
            os.replace(tmp, path)
    except Exception:
        pass


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


def _task_disk_size(task):
    total = 0
    for f in list(task.output_files or []):
        try:
            if f and os.path.exists(f):
                total += os.path.getsize(f)
        except Exception:
            pass
    return total


def refresh_task_size(task, force=False):
    if task is None:
        return
    now = time.time()
    if not force and (now - task._last_size_ts) < 5.0:
        return
    task._last_size_ts = now
    try:
        task.current_size = _task_disk_size(task)
    except Exception:
        pass


# ============================================================
# 预览
# ============================================================
def render_preview_only(solved_path, paths, task, n_preview=PREVIEW_PAGES):
    clear_dir_preview(paths["preview_dir"])
    os.makedirs(paths["preview_dir"], exist_ok=True)
    if not os.path.exists(solved_path):
        return []

    try:
        with open(solved_path, "rb") as f:
            data = f.read()
        solved = fitz.open(stream=data, filetype="pdf")
    except Exception:
        return []

    orig = fitz.open(task.src_path)
    n = min(n_preview, len(orig), len(solved))
    preview_imgs = []
    for i in range(n):
        if task.stop_event.is_set():
            break
        try:
            o_pix = orig[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM))
            s_pix = solved[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM))
            o = Image.open(io.BytesIO(o_pix.tobytes("png"))).convert("RGB")
            s = Image.open(io.BytesIO(s_pix.tobytes("png"))).convert("RGB")
            hh = max(o.height, s.height)
            if o.height != hh:
                o = o.resize((int(o.width * hh / o.height), hh), _RESAMPLE)
            if s.height != hh:
                s = s.resize((int(s.width * hh / s.height), hh), _RESAMPLE)
            gap = 8
            canvas = Image.new("RGB", (o.width + gap + s.width, hh), (40, 40, 40))
            canvas.paste(o, (0, 0))
            canvas.paste(s, (o.width + gap, 0))
            if canvas.width > PREVIEW_MAX_WIDTH:
                ratio = PREVIEW_MAX_WIDTH / canvas.width
                canvas = canvas.resize(
                    (PREVIEW_MAX_WIDTH, int(canvas.height * ratio)),
                    _RESAMPLE,
                )
            p = os.path.join(paths["preview_dir"], f"compare_{i:04d}.jpg")
            canvas.save(p, "JPEG", quality=PREVIEW_JPEG_QUALITY, optimize=True)
            preview_imgs.append(p)
        except Exception:
            continue
    orig.close()
    solved.close()
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


def img_to_base64_dataurl(path):
    try:
        with open(path, "rb") as f:
            data = f.read()
        b64 = base64.b64encode(data).decode("ascii")
        low = path.lower()
        mime = "image/png" if low.endswith(".png") else "image/jpeg"
        return f"data:{mime};base64,{b64}"
    except Exception:
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
                PDF 解题显示前 5 页左右对照；Word / PPT 显示前 5 段文本对照
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
        · 左右对照（左：原题 · 右：答案 + 解析） · {len(imgs_html)} 页 ·
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
            第 {idx+1} 段 · 解答
          </div>
          <div style="font-size:14.5px;color:#0f3d3e;line-height:1.9;white-space:pre-wrap">{escape(str(dst or ""))}</div>
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
            第 {page_no} 张 · 解答
          </div>
          <div style="font-size:14.5px;color:#0f3d3e;line-height:1.9;white-space:pre-wrap">{escape(str(dst or ""))}</div>
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


# ============================================================
# 进度
# ============================================================
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
                    box-shadow:0 0 12px rgba(201,169,97,.45)"></div>
      </div>
      <div style="text-align:right;font-size:11px;color:#a9a49a;
                  font-family:'SF Mono',Consolas,monospace;margin-top:4px">
        {pct}%
      </div>
      {size_line}
      {warning_html}
    </div>
    '''


# ============================================================
# 完成模态弹窗
# ============================================================
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

    actual_size = task.current_size or 0
    size_pill = ""
    if actual_size > 0:
        size_pill = (
            f'<span style="font-size:12px;color:#0f3d3e;background:#e8f0ef;'
            f'padding:5px 14px;border-radius:999px;font-weight:500">'
            f'📦 {fmt_size(actual_size)}</span>'
        )

    unit = "页" if getattr(task, "kind", "pdf") == "pdf" else "段"

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
            解题完成
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
            ✅ {task.total} {unit}
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


# ============================================================
# PDF worker
# ============================================================
def pdf_worker(task, paths, real_key, model, trial):
    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_json_file(paths["cache_file"], {})
    done_pages = load_progress_file(paths["progress_file"])

    pdf_exists = os.path.exists(paths["solved_pdf"])
    if done_pages and pdf_exists:
        try:
            with open(paths["solved_pdf"], "rb") as f:
                data = f.read()
            doc = fitz.open(stream=data, filetype="pdf")
            task.log_msg(f"📂 从已解题 PDF 续传（进度记录：{len(done_pages)} 页）")
        except Exception as e:
            task.log_msg(f"⚠️ 打开旧结果失败，从头开始：{e}")
            doc = fitz.open(task.src_path)
            done_pages = set()
            save_progress_file(paths["progress_file"], done_pages)
    else:
        doc = fitz.open(task.src_path)
        if done_pages and not pdf_exists:
            task.log_msg("⚠️ 有进度记录但缺结果 PDF，从头开始")
            done_pages = set()
            save_progress_file(paths["progress_file"], done_pages)

    total = len(doc)
    limit = min(5, total) if trial else total

    task.status = "running"
    task.total = limit
    task.label = f"PDF · 目标前 {limit} 页"
    task.current = len([p for p in done_pages if p <= limit])
    task.log_msg(f"✅ PDF 共 {total} 页，本次目标 {limit} 页，已完成 {task.current} 页")
    save_state(task)

    if pdf_exists:
        try:
            task.preview_images = render_preview_only(paths["solved_pdf"], paths, task)
            task.output_files = [paths["solved_pdf"]]
            refresh_task_size(task, force=True)
            save_state(task)
        except Exception as e:
            task.log_msg(f"⚠️ 初始预览失败：{e}")

    error_msg = None
    completed = False

    for pno in range(total):
        if task.stop_event.is_set():
            task.log_msg("⏸ 停止信号，结束当前循环")
            break
        page_num = pno + 1
        if page_num > limit and page_num not in done_pages:
            break
        if page_num in done_pages:
            continue

        page = doc[pno]
        blocks = [b for b in page.get_text("blocks")
                  if b[6] == 0 and b[4].strip()]

        if not blocks:
            done_pages.add(page_num)
            save_progress_file(paths["progress_file"], done_pages)
            task.current = len([p for p in done_pages if p <= limit])
            task.label = f"已跳过空白/图片页 {page_num}"
            save_state(task)
            continue

        try:
            results = solve_page(client, model, blocks, cache, paths["cache_file"])
            apply_solution_to_page(page, blocks, results)
        except RuntimeError as e:
            error_msg = str(e)
            break
        except Exception as e:
            error_msg = f"未知错误：{e}"
            break

        done_pages.add(page_num)
        save_progress_file(paths["progress_file"], done_pages)
        task.current = len([p for p in done_pages if p <= limit])
        task.label = f"已解 {task.current}/{limit} 页"
        task.log_msg(f"✅ 第 {page_num} 页完成")
        save_state(task)

        try:
            safe_save_pdf(doc, paths["solved_pdf"])
            task.log_msg(f"💾 已落盘（前 {page_num} 页）")
            if page_num <= PREVIEW_PAGES:
                task.preview_images = render_preview_only(
                    paths["solved_pdf"], paths, task)
            task.output_files = [paths["solved_pdf"]]
            refresh_task_size(task, force=True)
            save_state(task)
        except Exception as e:
            task.log_msg(f"⚠️ 落盘失败：{e}")
    else:
        completed = True

    try:
        safe_save_pdf(doc, paths["solved_pdf"])
    except Exception as e:
        error_msg = error_msg or f"保存结果 PDF 失败：{e}"
    finally:
        try:
            doc.close()
        except Exception:
            pass

    if not done_pages:
        task.status = "error"
        task.error = "本次没有任何页面成功处理"
        task.log_msg("⚠️ " + task.error)
        save_state(task)
        return

    try:
        task.preview_images = render_preview_only(paths["solved_pdf"], paths, task)
    except Exception:
        pass
    task.output_files = [paths["solved_pdf"]]

    if completed:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 全部完成！共 {len(done_pages)} 页")
    elif task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停（半成品），共 {len(done_pages)} 页"
        task.log_msg("🛑 已暂停，下次从未完成的页继续")
    elif error_msg:
        task.status = "error"
        task.error = error_msg
        task.label = "出错（半成品）"
        task.log_msg(f"❌ {error_msg}")
    else:
        task.status = "paused"
        task.label = "半成品"
    refresh_task_size(task, force=True)
    save_state(task)


# ============================================================
# Word / PPT worker
# ============================================================
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
        shd.set(qn("w:fill"), "EEF2FF")
        pPr.append(shd)
    except Exception:
        pass
    return new_para


def docx_worker(task, paths, real_key, model):
    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_json_file(paths["cache_file"], {})
    try:
        wdoc = Document(task.src_path)
    except Exception as e:
        task.status = "error"
        task.error = f"打开 Word 失败：{e}"
        task.log_msg(f"❌ {task.error}")
        save_state(task)
        return

    targets = [p for p in wdoc.paragraphs if p.text.strip()]
    total = len(targets)
    task.status = "running"
    task.total = total
    task.current = 0
    task.label = f"Word · 共 {total} 段"
    task.log_msg(f"📘 Word 已打开，共 {total} 段")
    save_state(task)

    error_msg = None
    done = 0
    solved_count = 0
    preview_pairs = []

    for para in targets:
        if task.stop_event.is_set():
            break
        text = para.text.strip()
        key = "solve_w_" + h(text)
        try:
            if key in cache and isinstance(cache[key], dict):
                res = cache[key]
            else:
                raw = call_solve_api(client, model, f"[[B0]] {text}")
                sub = parse_solution_response(raw, 1)
                res = sub.get(0, {"ans": "", "sol": ""})
                cache[key] = res
                save_json_file(paths["cache_file"], cache)
        except RuntimeError as e:
            error_msg = str(e)
            break

        ans = (res.get("ans") or "").strip()
        sol = (res.get("sol") or "").strip()
        if ans or sol:
            ins = []
            if ans:
                ins.append(f"【答案】{ans}")
            if sol:
                ins.append(f"【解析】{sol}")
            rendered = "\n".join(ins)
            _docx_insert_after(para, rendered)
            solved_count += 1
            # ★ 收集预览
            if len(preview_pairs) < PREVIEW_PARAS:
                preview_pairs.append((text, rendered))
                task.preview_html = _build_docx_preview_html(preview_pairs)

        done += 1
        task.current = done
        task.label = f"Word {done}/{total}（已解 {solved_count}）"
        if done % 5 == 0:
            save_state(task)

    try:
        wdoc.save(paths["office_out"])
        task.output_files = [paths["office_out"]]
        task.log_msg(f"✅ 已保存：{paths['office_out']}")
    except Exception as e:
        error_msg = error_msg or f"保存 Word 失败：{e}"

    if task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停 {done}/{total}"
    elif error_msg:
        task.status = "error"
        task.error = error_msg
    else:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 Word 解题完成！共处理 {done} 段，其中 {solved_count} 段有答案")
    refresh_task_size(task, force=True)
    save_state(task)


def pptx_worker(task, paths, real_key, model):
    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_json_file(paths["cache_file"], {})
    try:
        prs = Presentation(task.src_path)
    except Exception as e:
        task.status = "error"
        task.error = f"打开 PPT 失败：{e}"
        task.log_msg(f"❌ {task.error}")
        save_state(task)
        return

    targets = []
    for si, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                txt = "".join(r.text for r in para.runs)
                if txt.strip():
                    targets.append((si, para))

    total = len(targets)
    task.status = "running"
    task.total = total
    task.current = 0
    task.label = f"PPT · 共 {total} 段"
    task.log_msg(f"📊 PPT 已打开，共 {total} 段")
    save_state(task)

    error_msg = None
    done = 0
    solved_count = 0
    preview_pairs = []

    for si, para in targets:
        if task.stop_event.is_set():
            break
        text = "".join(r.text for r in para.runs).strip()
        key = "solve_p_" + h(text)
        try:
            if key in cache and isinstance(cache[key], dict):
                res = cache[key]
            else:
                raw = call_solve_api(client, model, f"[[B0]] {text}")
                sub = parse_solution_response(raw, 1)
                res = sub.get(0, {"ans": "", "sol": ""})
                cache[key] = res
                save_json_file(paths["cache_file"], cache)
        except RuntimeError as e:
            error_msg = str(e)
            break

        ans = (res.get("ans") or "").strip()
        sol = (res.get("sol") or "").strip()
        if ans or sol:
            parts = []
            if ans:
                parts.append(f"【答案】{ans}")
            if sol:
                parts.append(f"【解析】{sol}")
            rendered = "\n".join(parts)
            if para.runs:
                para.runs[0].text = text + "\n" + rendered
                for r in para.runs[1:]:
                    r.text = ""
            else:
                para.add_run(text + "\n" + rendered)
            solved_count += 1
            if len(preview_pairs) < PREVIEW_PARAS:
                preview_pairs.append((si + 1, text, rendered))
                task.preview_html = _build_pptx_preview_html(preview_pairs)

        done += 1
        task.current = done
        task.label = f"PPT {done}/{total}（第 {si+1} 张，已解 {solved_count}）"
        if done % 5 == 0:
            save_state(task)

    try:
        prs.save(paths["office_out"])
        task.output_files = [paths["office_out"]]
        task.log_msg(f"✅ 已保存：{paths['office_out']}")
    except Exception as e:
        error_msg = error_msg or f"保存 PPT 失败：{e}"

    if task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停 {done}/{total}"
    elif error_msg:
        task.status = "error"
        task.error = error_msg
    else:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 PPT 解题完成！共处理 {done} 段，其中 {solved_count} 段有答案")
    refresh_task_size(task, force=True)
    save_state(task)


# ============================================================
# 任务控制
# ============================================================
def start_task(kind, upload_path, real_key, model, trial):
    global SELECTED_TASK_ID
    src_name = os.path.basename(upload_path)
    paths = prepare_paths(src_name)
    src_copy = os.path.join(paths["input_dir"], src_name)
    try:
        shutil.copy2(upload_path, src_copy)
    except Exception as e:
        raise RuntimeError(f"复制上传文件失败：{e}")

    task = MANAGER.create(kind, src_copy, src_name,
                          paths["out_dir"], paths["work"])
    SELECTED_TASK_ID = task.task_id
    task.log_msg(f"🆔 任务 {task.task_id} 已创建（{kind}）")
    task.log_msg(f"📁 结果目录：{paths['out_dir']}")

    def runner():
        try:
            if kind == "pdf":
                pdf_worker(task, paths, real_key, model, trial)
            elif kind == "docx":
                docx_worker(task, paths, real_key, model)
            elif kind == "pptx":
                pptx_worker(task, paths, real_key, model)
        except Exception as e:
            task.status = "error"
            task.error = str(e)
            task.log_msg(f"❌ 未捕获错误：{e}")
            save_state(task)

    task.thread = threading.Thread(target=runner, daemon=True)
    task.thread.start()
    return task


def on_start(api_key, model, doc_file, trial):
    global SELECTED_TASK_ID

    real_key = (api_key or "").strip() or DEFAULT_API_KEY
    if not real_key or not doc_file:
        return on_refresh_fast() + ([], None)

    model = (model or DEFAULT_MODEL).strip()
    if model not in ("deepseek-chat", "deepseek-reasoner"):
        model = "deepseek-chat"

    name_lower = (doc_file.name or "").lower()
    if name_lower.endswith(".pdf"):
        kind = "pdf"
    elif name_lower.endswith(".docx"):
        kind = "docx"
    elif name_lower.endswith(".pptx"):
        kind = "pptx"
    else:
        cur = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
        if cur:
            cur.log_msg(f"❌ 不支持的文件类型：{doc_file.name}")
            save_state(cur)
        return on_refresh_fast() + ([], None)

    if kind in ("docx", "pptx") and not HAS_OFFICE:
        cur = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
        if cur:
            cur.log_msg("❌ 未安装 python-docx / python-pptx")
            save_state(cur)
        return on_refresh_fast() + ([], None)

    src_name = os.path.basename(doc_file.name)
    base = os.path.splitext(src_name)[0]
    book = safe_dirname(base)
    short_hash = h(src_name)[:6]
    target_out_dir = os.path.abspath(
        os.path.join(RESULT_ROOT, f"{book}_{short_hash}")
    )

    with _CREATE_LOCK:
        existing = MANAGER.find_active_by_out_dir(target_out_dir)
        if existing is not None:
            SELECTED_TASK_ID = existing.task_id
            existing.log_msg(f"⚠️ 已存在运行中的任务（#{existing.task_id}）")
            save_state(existing)
            return on_refresh_fast() + ([], None)

        try:
            start_task(kind, doc_file.name, real_key, model, trial)
        except Exception:
            pass

    return on_refresh_fast() + ([], None)


def on_stop_all():
    running = MANAGER.running()
    if not running:
        current = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
        if current:
            current.log_msg("⚠️ 没有正在运行的任务")
            save_state(current)
        return on_refresh_fast()
    for t in running:
        t.stop_event.set()
        t.status = "stopping"
        t.label = f"⏸ 停止信号已发出…（已处理 {t.current}/{t.total}）"
        t.log_msg("⏸ 收到停止信号 —— 当前一步完成后暂停")
        save_state(t)
    return on_refresh_fast()


def on_stop_selected(label):
    global SELECTED_TASK_ID
    if not label:
        return on_refresh_fast()
    m = re.match(r'#([0-9a-f]{8})', label)
    if not m:
        return on_refresh_fast()
    tid = m.group(1)
    t = MANAGER.get(tid)
    if t is None:
        return on_refresh_fast()
    if t.status in ("queued", "running"):
        t.stop_event.set()
        t.status = "stopping"
        t.label = f"⏸ 停止信号已发出…（已处理 {t.current}/{t.total}）"
        t.log_msg("⏸ 收到单独停止信号")
        save_state(t)
        SELECTED_TASK_ID = tid
    return on_refresh_fast()


def on_load_preview():
    current = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
    if current is None:
        tasks = MANAGER.all_sorted()
        current = tasks[0] if tasks else None
    if current is None:
        return build_preview_html([]), []

    imgs = [p for p in (current.preview_images or [])
            if p and os.path.exists(p)]
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
    status_text = {"queued": "排队中", "running": "运行中",
                   "stopping": "停止中…", "paused": "已暂停",
                   "done": "已完成", "error": "出错"}
    kind_icon = {"pdf": "📕", "docx": "📘", "pptx": "📊"}

    rows = []
    for t in tasks[:8]:
        col = status_color.get(t.status, "#666")
        bg = status_bg.get(t.status, "#f0ebe0")
        stx = status_text.get(t.status, t.status)
        kc = kind_icon.get(t.kind, "📄")
        pct = int(t.current * 100 / t.total) if t.total else 0
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
          <span style="color:{col};background:{bg};font-weight:600;
                       font-size:11.5px;padding:2px 10px;border-radius:999px;
                       flex-shrink:0">{stx}</span>
          <span style="color:#8b8578;font-family:'SF Mono',Consolas,monospace;
                       font-size:11.5px;flex-shrink:0">{t.current}/{t.total} ({pct}%)</span>
        </div>''')

    rows_html = "".join(rows)
    if len(tasks) > 8:
        rows_html += (
            f'<div style="padding:10px 14px;text-align:center;'
            f'color:#a9a49a;font-size:12px;background:#fdfcf8">'
            f'… 还有 {len(tasks) - 8} 个任务未显示</div>'
        )

    return (f'<div style="background:#fff;border:1px solid #ebe5d8;'
            f'border-radius:14px;overflow:hidden">{rows_html}</div>')


def on_refresh_fast():
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
        if not previews:
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
    dd_update = gr.update(choices=choices)

    alive_ids = {t.task_id for t in tasks}
    for k in list(_LAST_PREVIEW_SIG.keys()):
        if k not in alive_ids:
            _LAST_PREVIEW_SIG.pop(k, None)
    for k in list(_PREVIEW_HTML_CACHE.keys()):
        if k not in alive_ids:
            _PREVIEW_HTML_CACHE.pop(k, None)
    for k in list(_MODAL_SHOWN):
        if k not in alive_ids:
            _MODAL_SHOWN.discard(k)

    modal_html = gr.update(value=None)
    for t in tasks:
        if t.status == "done" and t.task_id not in _MODAL_SHOWN:
            _MODAL_SHOWN.add(t.task_id)
            try:
                refresh_task_size(t, force=True)
            except Exception:
                pass
            modal_html = build_done_modal_html(t)
            break

    return (task_list_html, progress_html, log_text,
            dd_update, gallery_value, modal_html)


def open_result_folder():
    global SELECTED_TASK_ID
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
# UI
# ============================================================
with gr.Blocks(
    title="PDF / Word / PPT 作业解题器",
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
        min-height: 480px;
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

    #trial_cb {
        background: transparent !important; border: none !important;
        padding: 4px 2px !important;
    }
    #trial_cb * { cursor: pointer !important; }
    #trial_cb input[type="checkbox"] {
        -webkit-appearance: checkbox !important; appearance: checkbox !important;
        width: 16px !important; height: 16px !important;
        min-width: 16px !important; max-width: 16px !important;
        accent-color: #0f3d3e !important; margin-right: 9px !important;
    }
    #trial_cb label { cursor: pointer !important; font-size: 13.5px !important; }

    #pdf_upload {
        min-height: 110px !important;
        max-height: 180px !important;
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
    """,
) as demo:

    modal_html = gr.HTML(value="", elem_id="modal_host")

    gr.HTML("""
    <div class="card-head">
      <h1>📝 PDF / Word / PPT 作业解题器</h1>
      <div class="sub">
        答案写在题目处<span class="dot">·</span>
        题目下方给解析<span class="dot">·</span>
        数学公式 mathtext 渲染<span class="dot">·</span>
        断点可续<span class="dot">·</span>
        试解 5 页
      </div>
    </div>
    """)

    with gr.Accordion("💡 使用说明（点击展开）", open=False):
        gr.HTML("""
        <div style="padding:14px 20px;font-size:13.5px;line-height:2;color:#444;
                    background:#fdfcf8;border:1px solid #ebe5d8;border-radius:12px">
          <div><b>🎨 PDF 效果</b>　答案 <span style="color:#c0392b;font-weight:600">红色</span> 写在题目块底部；解析 <span style="color:#1a3a8a;font-weight:600">蓝色</span> + 浅蓝底贴在题目下方</div>
          <div><b>📄 Word / PPT 效果</b>　在每个段落后插入【答案】【解析】段落（浅蓝底）</div>
          <div><b>📐 公式</b>　PDF 模式用 matplotlib mathtext 渲染成透明图片贴入；Word / PPT 直接输出文本</div>
          <div><b>🔑 密钥</b>　留空取 <code>.env</code> 中的默认值；填写 → 临时覆盖</div>
          <div><b>📄 上传</b>　拖入 .pdf / .docx / .pptx 自动识别类型；<b>PDF 只处理含文字层的，扫描件无法识别</b></div>
          <div><b>🧪 试解</b>　勾选后 PDF 只处理前 5 页；满意后取消勾选再全量跑</div>
          <div><b>📁 输出</b>　结果在 <code>homework_result/&lt;文件名_hash&gt;/</code></div>
          <div><b>👀 预览</b>　PDF 显示前 5 页图片对照；Word / PPT 显示前 5 段文本对照</div>
          <div><b>🎉 完成提示</b>　所有页/段完成后会弹出居中卡片，列出结果文件与保存位置</div>
        </div>
        """)

    # ================= 上部：左设置 / 右状态 =================
    with gr.Row(equal_height=False, elem_id="main_row"):

        with gr.Column(scale=1, min_width=440, elem_id="setup_col"):
            gr.HTML('<div class="section-title">⚙️ 解题设置</div>')

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
                )

            doc_file = gr.File(
                label="📄 上传作业文档（.pdf / .docx / .pptx）",
                file_types=[".pdf", ".docx", ".pptx"],
                elem_id="pdf_upload",
            )

            trial = gr.Checkbox(
                label="🧪 试解模式：PDF 只处理前 5 页（对 Word / PPT 无效）",
                value=False,
                elem_id="trial_cb",
            )

            with gr.Column(elem_id="action_grid"):
                with gr.Row(equal_height=True):
                    btn = gr.Button("▶ 开始解题", variant="primary")
                    stop_btn = gr.Button("⏹ 全部停止", variant="secondary")
                    refresh_btn = gr.Button("🔄 刷新", variant="secondary")
                    open_btn = gr.Button("📁 打开目录", variant="secondary")

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
                stop_one_btn = gr.Button("⏹ 停止选中",
                                         elem_id="stop_one_btn")

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

    fast_outputs = [task_list_html, progress_bar, log,
                    stop_dd, gallery, modal_html]
    full_outputs = [task_list_html, progress_bar, log,
                    stop_dd, gallery, modal_html, out_files, doc_file]

    btn.click(
        on_start,
        [api_key, model, doc_file, trial],
        full_outputs,
        concurrency_limit=3,
        concurrency_id="start",
    )
    stop_btn.click(on_stop_all, None, fast_outputs,
                   concurrency_limit=None, concurrency_id="stop")
    stop_one_btn.click(on_stop_selected, [stop_dd], fast_outputs,
                       concurrency_limit=None, concurrency_id="stop_one")
    refresh_btn.click(on_refresh_fast, None, fast_outputs,
                      concurrency_limit=None, concurrency_id="manual")
    open_btn.click(open_result_folder, None, [log])
    load_btn.click(on_load_preview, None, [gallery, out_files])

    try:
        timer = gr.Timer(3.0)
        timer.tick(on_refresh_fast, None, fast_outputs,
                   concurrency_limit=None, concurrency_id="tick")
    except Exception:
        pass

    demo.load(on_refresh_fast, None, fast_outputs,
              concurrency_limit=None, concurrency_id="load")


# ============================================================
# main
# ============================================================
def _find_free_port(start=7870, end=7899):
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
    PORT = _find_free_port(7870, 7899)
    URL = "http://127.0.0.1:" + str(PORT)

    try:
        with open("访问网址_解题器.txt", "w", encoding="utf-8") as f:
            f.write("PDF / Word / PPT 作业解题器\n\n")
            f.write("浏览器访问网址：\n" + URL + "\n\n")
            f.write("（如果打不开，说明程序已停止，请双击 启动解题器.bat）\n")
    except Exception:
        pass

    print()
    print("=" * 64)
    print("  PDF / Word / PPT 作业解题器 已启动")
    print()
    print("  在浏览器输入以下网址进入：")
    print()
    print("      " + URL)
    print()
    print("  别关这个终端窗口，否则程序停止")
    print("=" * 64)
    print()

    print(f"   结果目录：{os.path.abspath(RESULT_ROOT)}")
    if DEFAULT_API_KEY:
        print(f"   API Key：已从 .env 读取（{DEFAULT_API_KEY[:6]}...）")
    else:
        print("   API Key：未配置，需在网页填写")
    print(f"   中文字体：{CN_FONT_PATH or '（未找到，中文可能显示为方框）'}")
    if not HAS_OFFICE:
        print("   ⚠️ 未装 python-docx / python-pptx，Word / PPT 不可用")
    print()

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
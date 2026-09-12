# -*- coding: utf-8 -*-
"""
PDF / Word / PPT 翻译器（网页版 · 多任务并行 + 单独停止 + 断点续传）
"""

import os
import sys
import json
import time
import re
import glob
import base64
import hashlib
import shutil
import threading
import uuid
import io
from dataclasses import dataclass, field
from typing import Optional
from html import escape

import pymupdf as fitz
try:
    fitz.TOOLS.mupdf_display_errors(False)
except Exception:
    pass

from PIL import Image
from openai import OpenAI
import gradio as gr
from dotenv import load_dotenv

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    from pptx import Presentation
    HAS_OFFICE = True
except ImportError:
    HAS_OFFICE = False

# ================= 读取 .env =================
load_dotenv()
DEFAULT_API_KEY = os.getenv("DEEPSEEK_API_KEY", "").strip()
DEFAULT_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat").strip()
if DEFAULT_MODEL not in ("deepseek-chat", "deepseek-reasoner"):
    DEFAULT_MODEL = "deepseek-chat"

# ================= 配置 =================
FONT_PATH = r"D:\file\translate\word_type\09_SourceHanSerifSC\OTF\SimplifiedChinese\SourceHanSerifSC-Regular.otf"
RESULT_ROOT = "result"
RENDER_ZOOM = 2.0
PREVIEW_PAGES = 5
PREVIEW_MAX_WIDTH = 1400
PREVIEW_JPEG_QUALITY = 88
PREVIEW_PARAS = 5

_IO_LOCK = threading.Lock()
_LAST_PREVIEW_SIG = {}
_PREVIEW_HTML_CACHE = {}

SYSTEM_PROMPT = (
    "你是一位资深文学翻译家，精通中英双语，译笔力求神似而非字对字。"
    "请遵循：1) 译文必须符合中文母语者的阅读习惯和审美，流畅、有文采；"
    "2) 对话要自然生动，符合人物身份；3) 修辞、隐喻、双关尽量找到中文对应表达，"
    "实在无法对应则意译并保留神韵；4) 不遗漏任何内容，不总结，不输出任何解释。\n\n"
    "【格式要求】用户会给出一页英文的多个段落，每段以 [[B0]] [[B1]] [[B2]] ... 标记开头。"
    "你必须严格保留所有标记、保持顺序，标记后紧跟该段译文。"
    "除标记和译文外，不要输出任何其他文字、不加解释、不用代码块。"
)
SIMPLE_SYSTEM = (
    "你是一位资深文学翻译家，精通中英双语，译笔力求神似而非字对字。"
    "请把用户发来的英文翻译成流畅、有文采的中文，符合中文母语者的阅读习惯。"
    "保留段落结构，不遗漏内容，不总结，不输出解释，直接给出译文。"
)


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
            return [t for t in self._tasks.values() if t.status in ("queued", "running", "stopping")]

    def find_active_by_out_dir(self, out_dir):
        target = os.path.abspath(out_dir)
        with self._lock:
            for t in self._tasks.values():
                if os.path.abspath(t.out_dir) == target and t.status in ("queued", "running", "stopping"):
                    return t
        return None

    def load_from_disk(self, state_dict):
        tid = state_dict["task_id"]
        if tid in self._tasks:
            return
        t = TaskState(
            task_id=tid,
            kind=state_dict.get("kind", "pdf"),
            src_path="",
            src_name=state_dict.get("src_name", ""),
            out_dir=state_dict.get("out_dir", ""),
            work_dir=os.path.join(state_dict.get("out_dir", ""), "_work"),
            created_at=state_dict.get("created_at", time.time()),
            status=state_dict.get("status", "paused"),
            current=state_dict.get("current", 0),
            total=state_dict.get("total", 0),
            label=state_dict.get("label", ""),
            log=state_dict.get("log", []),
            output_files=state_dict.get("output_files", []),
            error=state_dict.get("error", ""),
        )
        if t.status in ("queued", "running", "stopping"):
            t.status = "paused"
        with self._lock:
            self._tasks[tid] = t


MANAGER = TaskManager()
SELECTED_TASK_ID = None


def save_state(task):
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
            if f.startswith("translated_new_") and f.endswith(".pdf"):
                try:
                    os.remove(os.path.join(d, f))
                    removed += 1
                    print(f"   🧹 清理冗余文件：{name}/{f}")
                except Exception:
                    pass
            if f.endswith(".tmp.pdf"):
                try:
                    os.remove(os.path.join(d, f))
                    removed += 1
                    print(f"   🧹 清理临时文件：{name}/{f}")
                except Exception:
                    pass
    return removed


# ============================================================
# 路径 / 工具
# ============================================================

def safe_dirname(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip().rstrip(". ")
    return name[:80] or "untitled"


def prepare_paths(src_path):
    base = os.path.basename(src_path)
    book, ext = os.path.splitext(base)
    book = safe_dirname(book)
    out_dir = os.path.abspath(os.path.join(RESULT_ROOT, book))
    work = os.path.join(out_dir, "_work")
    input_dir = os.path.join(work, "input")
    os.makedirs(work, exist_ok=True)
    os.makedirs(input_dir, exist_ok=True)
    return {
        "out_dir": out_dir, "work": work, "input_dir": input_dir,
        "output_pdf":    os.path.join(out_dir, "translated.pdf"),
        "bilingual_pdf": os.path.join(out_dir, "bilingual.pdf"),
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
        except PermissionError as e:
            last_err = e
            time.sleep(0.8 * (attempt + 1))
        except OSError as e:
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


def img_to_base64_dataurl(path):
    try:
        with open(path, "rb") as f:
            data = f.read()
        b64 = base64.b64encode(data).decode("ascii")
        low = path.lower()
        if low.endswith(".jpg") or low.endswith(".jpeg"):
            mime = "image/jpeg"
        elif low.endswith(".png"):
            mime = "image/png"
        else:
            mime = "image/jpeg"
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
        <div style="padding:30px 20px;color:#888;text-align:center;font-size:13.5px;
                    background:#faf8f2;border:1px dashed #ddd5c0;border-radius:10px;
                    line-height:1.8">
            📄 暂无预览<br>
            <span style="font-size:12px;color:#aaa">翻译启动后会显示前 5 页的左右对照</span>
        </div>
        '''

    imgs_html = []
    for p in preview_imgs:
        data_url = img_to_base64_dataurl(p)
        if not data_url:
            continue
        imgs_html.append(
            f'<img src="{data_url}" '
            f'style="width:100%;display:block;margin:0 0 14px 0;'
            f'box-shadow:0 2px 10px rgba(0,0,0,.18);border-radius:4px;">'
        )

    if not imgs_html:
        return '''
        <div style="padding:30px 20px;color:#888;text-align:center;font-size:13.5px;
                    background:#faf8f2;border:1px dashed #ddd5c0;border-radius:10px">
            ⚠️ 预览图片加载失败
        </div>
        '''

    return f'''
    <div style="background:#2b2b2b;border-radius:10px;padding:14px;
                max-height:820px;overflow-y:auto;
                scroll-behavior:smooth">
      <div style="color:#aaa;font-size:12px;text-align:center;
                  padding:6px 0 12px 0;letter-spacing:.5px">
        左右对照 · 上下滚动阅读（{len(imgs_html)} 页）
      </div>
      {''.join(imgs_html)}
      <div style="color:#666;font-size:11px;text-align:center;padding:6px 0">
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
        <div style="background:#fff;border-radius:8px;padding:16px 18px;margin-bottom:12px;
                    box-shadow:0 2px 10px rgba(0,0,0,.15)">
          <div style="font-size:11px;color:#b09b63;font-weight:600;
                      letter-spacing:.5px;margin-bottom:6px">第 {idx+1} 段 · 原文</div>
          <div style="font-size:14px;color:#555;line-height:1.75;margin-bottom:14px">{escape(src)}</div>
          <div style="font-size:11px;color:#b09b63;font-weight:600;
                      letter-spacing:.5px;margin-bottom:6px">第 {idx+1} 段 · 译文</div>
          <div style="font-size:14.5px;color:#111;line-height:1.9">{escape(dst)}</div>
        </div>''')
    return (
        '<div style="background:#f5f1e6;border-radius:10px;padding:14px;'
        'max-height:820px;overflow-y:auto;scroll-behavior:smooth">'
        '<div style="color:#888;font-size:12px;text-align:center;'
        'padding:6px 0 12px 0;letter-spacing:.5px">'
        f'文本对照预览（前 {len(pairs)} 段）'
        '</div>'
        + "".join(rows) +
        '<div style="color:#aaa;font-size:11px;text-align:center;padding:6px 0">'
        '— 完整结果请下载 Word 查看 —</div>'
        '</div>'
    )


def _build_pptx_preview_html(pairs):
    if not pairs:
        return ""
    rows = []
    for idx, (page_no, src, dst) in enumerate(pairs):
        rows.append(f'''
        <div style="background:#fff;border-radius:8px;padding:16px 18px;margin-bottom:12px;
                    box-shadow:0 2px 10px rgba(0,0,0,.15)">
          <div style="font-size:11px;color:#b09b63;font-weight:600;
                      letter-spacing:.5px;margin-bottom:6px">第 {page_no} 张 · 原文</div>
          <div style="font-size:14px;color:#555;line-height:1.75;margin-bottom:14px">{escape(src)}</div>
          <div style="font-size:11px;color:#b09b63;font-weight:600;
                      letter-spacing:.5px;margin-bottom:6px">第 {page_no} 张 · 译文</div>
          <div style="font-size:14.5px;color:#111;line-height:1.9">{escape(dst)}</div>
        </div>''')
    return (
        '<div style="background:#f5f1e6;border-radius:10px;padding:14px;'
        'max-height:820px;overflow-y:auto;scroll-behavior:smooth">'
        '<div style="color:#888;font-size:12px;text-align:center;'
        'padding:6px 0 12px 0;letter-spacing:.5px">'
        f'文本对照预览（前 {len(pairs)} 段）'
        '</div>'
        + "".join(rows) +
        '<div style="color:#aaa;font-size:11px;text-align:center;padding:6px 0">'
        '— 完整结果请下载 PPT 查看 —</div>'
        '</div>'
    )


def make_progress_html(done, total, label=""):
    if total <= 0:
        total = 1
    done = max(0, min(done, total))
    pct = int(done * 100 / total)
    return f'''
    <div style="padding:10px 4px">
      <div style="display:flex;justify-content:space-between;font-size:13px;color:#444;margin-bottom:6px">
        <span>{label}</span>
        <span><b>{done}</b> / {total}（{pct}%）</span>
      </div>
      <div style="height:16px;background:#e5e7eb;border-radius:8px;overflow:hidden;box-shadow:inset 0 1px 3px rgba(0,0,0,.08)">
        <div style="height:100%;width:{pct}%;background:linear-gradient(90deg,#4f46e5,#7c3aed);transition:width .35s ease"></div>
      </div>
    </div>
    '''


def parse_marked(text, n):
    result = {}
    pat = re.compile(r'\[\[B(\d+)\]\]')
    matches = list(pat.finditer(text))
    for i, m in enumerate(matches):
        idx = int(m.group(1))
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        result[idx] = text[start:end].strip()
    return result


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
                temperature=0.3,
            )
            return resp.choices[0].message.content
        except Exception as e:
            last_err = str(e)
            if "402" in last_err or "insufficient" in last_err.lower() or "余额" in last_err:
                raise RuntimeError("账户余额不足，请去 DeepSeek 平台充值后再继续")
            if "401" in last_err or "Unauthorized" in last_err or "invalid" in last_err.lower():
                raise RuntimeError("API Key 无效或已过期，请检查后重试")
            time.sleep(6 * (attempt + 1))
    raise RuntimeError(f"API 连续失败：{last_err}")


def call_api(client, model, text, retries=4):
    return _api_call(client, model, SYSTEM_PROMPT, text, retries)


def call_simple_api(client, model, text, retries=4):
    return _api_call(client, model, SIMPLE_SYSTEM, text, retries)


# ============================================================
# PDF
# ============================================================

def translate_page(client, model, blocks, cache, cache_file):
    marked = "\n\n".join(f"[[B{i}]] {b[4].strip()}" for i, b in enumerate(blocks))
    key = "pg_" + h(marked)
    if key in cache:
        try:
            return {int(k): v for k, v in cache[key].items()}
        except Exception:
            pass

    raw = call_api(client, model, marked)
    parsed = parse_marked(raw, len(blocks))

    missing = [i for i in range(len(blocks)) if i not in parsed or not parsed[i].strip()]
    for i in missing:
        bk = "bk_" + h(blocks[i][4])
        if bk in cache and cache[bk].strip():
            parsed[i] = cache[bk]
            continue
        r = call_api(client, model, f"[[B0]] {blocks[i][4].strip()}")
        sub = parse_marked(r, 1)
        parsed[i] = sub.get(0, "").strip()
        cache[bk] = parsed[i]
        save_json_file(cache_file, cache)

    cache[key] = {str(k): v for k, v in parsed.items()}
    save_json_file(cache_file, cache)
    return parsed


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


def find_fontsize(rect, text):
    for fs in [11, 10.5, 10, 9.5, 9, 8.5, 8, 7.5, 7, 6.5, 6, 5.5, 5, 4.5, 4]:
        if estimate_lines(rect.width, text, fs) * fs * 1.35 <= rect.height + 2:
            return fs
    return 4


def apply_translations(page, blocks, translations):
    page.insert_font(fontname="cn", fontfile=FONT_PATH, set_simple=False)
    for i, b in enumerate(blocks):
        text = translations.get(i, "").strip()
        if not text:
            continue
        rect = fitz.Rect(b[0], b[1], b[2], b[3])
        if rect.is_empty or rect.width <= 0 or rect.height <= 0:
            continue
        page.draw_rect(rect, color=None, fill=(1, 1, 1), overlay=True)
        fs = find_fontsize(rect, text)
        rc = page.insert_textbox(
            rect, text, fontname="cn", fontsize=fs,
            color=(0, 0, 0), align=0, overlay=True,
        )
        while rc < 0 and fs > 4:
            fs -= 0.5
            page.draw_rect(rect, color=None, fill=(1, 1, 1), overlay=True)
            rc = page.insert_textbox(
                rect, text, fontname="cn", fontsize=fs,
                color=(0, 0, 0), align=0, overlay=True,
            )


def render_preview_only(trans_path, paths, task, n_preview=PREVIEW_PAGES):
    clear_dir_preview(paths["preview_dir"])
    os.makedirs(paths["preview_dir"], exist_ok=True)

    if not os.path.exists(trans_path):
        return []

    try:
        with open(trans_path, "rb") as f:
            data = f.read()
        trans = fitz.open(stream=data, filetype="pdf")
    except Exception:
        return []

    orig = fitz.open(task.src_path)
    n = min(n_preview, len(orig), len(trans))
    preview_imgs = []
    for i in range(n):
        if task.stop_event.is_set():
            break
        try:
            o_pix = orig[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM))
            t_pix = trans[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM))
            o = Image.open(io.BytesIO(o_pix.tobytes("png"))).convert("RGB")
            t = Image.open(io.BytesIO(t_pix.tobytes("png"))).convert("RGB")
            hh = max(o.height, t.height)
            if o.height != hh:
                o = o.resize((int(o.width * hh / o.height), hh), Image.LANCZOS)
            if t.height != hh:
                t = t.resize((int(t.width * hh / t.height), hh), Image.LANCZOS)
            gap = 8
            canvas = Image.new("RGB", (o.width + gap + t.width, hh), (40, 40, 40))
            canvas.paste(o, (0, 0))
            canvas.paste(t, (o.width + gap, 0))
            if canvas.width > PREVIEW_MAX_WIDTH:
                ratio = PREVIEW_MAX_WIDTH / canvas.width
                canvas = canvas.resize(
                    (PREVIEW_MAX_WIDTH, int(canvas.height * ratio)),
                    Image.LANCZOS,
                )
            p = os.path.join(paths["preview_dir"], f"compare_{i:04d}.jpg")
            canvas.save(p, "JPEG", quality=PREVIEW_JPEG_QUALITY, optimize=True)
            preview_imgs.append(p)
        except Exception:
            continue
    orig.close()
    trans.close()
    return preview_imgs


def make_bilingual_pdf_preview(trans_path, paths, task, n_preview=PREVIEW_PAGES):
    if not os.path.exists(trans_path):
        return
    try:
        with open(trans_path, "rb") as f:
            data = f.read()
        trans = fitz.open(stream=data, filetype="pdf")
    except Exception:
        return

    try:
        orig = fitz.open(task.src_path)
        n = min(n_preview, len(orig), len(trans))
        pages = []
        for i in range(n):
            if task.stop_event.is_set():
                break
            try:
                o_pix = orig[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM))
                t_pix = trans[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM))
                o = Image.open(io.BytesIO(o_pix.tobytes("png"))).convert("RGB")
                t = Image.open(io.BytesIO(t_pix.tobytes("png"))).convert("RGB")
                hh = max(o.height, t.height)
                if o.height != hh:
                    o = o.resize((int(o.width * hh / o.height), hh), Image.LANCZOS)
                if t.height != hh:
                    t = t.resize((int(t.width * hh / t.height), hh), Image.LANCZOS)
                gap = 10
                canvas = Image.new("RGB", (o.width + gap + t.width, hh), (255, 255, 255))
                canvas.paste(o, (0, 0))
                canvas.paste(t, (o.width + gap, 0))
                pages.append(canvas)
            except Exception:
                continue
        orig.close()
        trans.close()
        if pages:
            pages[0].save(paths["bilingual_pdf"], save_all=True,
                          append_images=pages[1:], resolution=150.0)
    except Exception:
        pass


# ============================================================
# Office
# ============================================================

def _docx_replace_para_text(para, new_text):
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


# ============================================================
# Workers
# ============================================================

def pdf_worker(task, paths, real_key, model, trial):
    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_json_file(paths["cache_file"], {})
    done_pages = load_progress_file(paths["progress_file"])

    pdf_exists = os.path.exists(paths["output_pdf"])
    if done_pages and pdf_exists:
        try:
            with open(paths["output_pdf"], "rb") as f:
                data = f.read()
            doc = fitz.open(stream=data, filetype="pdf")
            task.log_msg(f"📂 从已翻译 PDF 续传（已翻 {len(done_pages)} 页，本次不重做）")
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

    total = len(doc)
    limit = min(5, total) if trial else total

    task.status = "running"
    task.total = limit
    task.label = f"PDF · 目标前 {limit} 页"
    task.current = len([p for p in done_pages if p <= limit])
    task.log_msg(f"✅ PDF 共 {total} 页，本次目标前 {limit} 页")
    save_state(task)

    if pdf_exists:
        try:
            task.log_msg("🖼 生成前 5 页预览……")
            task.preview_images = render_preview_only(paths["output_pdf"], paths, task)
            make_bilingual_pdf_preview(paths["output_pdf"], paths, task)
            task.output_files = [paths["output_pdf"], paths["bilingual_pdf"]]
            save_state(task)
            task.log_msg(f"✅ 预览图就绪（{len(task.preview_images)} 张）")
            save_state(task)
        except Exception as e:
            task.log_msg(f"⚠️ 初始预览生成失败：{e}")

    newly = []
    completed = False
    error_msg = None

    for pno in range(total):
        if task.stop_event.is_set():
            task.log_msg("⏸ 检测到停止信号，结束当前循环")
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
            task.current = len([p for p in done_pages if p <= limit])
            task.label = f"已跳过插图页 {page_num}"
            save_state(task)
            continue

        try:
            trans = translate_page(client, model, blocks, cache, paths["cache_file"])
            apply_translations(page, blocks, trans)
        except RuntimeError as e:
            error_msg = str(e)
            break
        except Exception as e:
            error_msg = f"未知错误：{e}"
            break

        done_pages.add(page_num)
        save_progress_file(paths["progress_file"], done_pages)
        newly.append(page_num)
        task.current = len([p for p in done_pages if p <= limit])
        task.label = f"已翻 {task.current}/{limit} 页"
        task.log_msg(f"✅ 第 {page_num} 页完成（本次新增 {len(newly)} 页）")
        save_state(task)

        if page_num <= PREVIEW_PAGES:
            try:
                safe_save_pdf(doc, paths["output_pdf"])
                task.preview_images = render_preview_only(paths["output_pdf"], paths, task)
                make_bilingual_pdf_preview(paths["output_pdf"], paths, task)
                task.output_files = [paths["output_pdf"], paths["bilingual_pdf"]]
                save_state(task)
            except Exception:
                pass
    else:
        completed = True

    try:
        safe_save_pdf(doc, paths["output_pdf"])
    except Exception as e:
        error_msg = error_msg or f"保存译文 PDF 失败：{e}"
    finally:
        try:
            doc.close()
        except Exception:
            pass

    if not done_pages:
        task.status = "error"
        task.error = "本次没有任何页面成功翻译"
        task.log_msg("⚠️ " + task.error)
        save_state(task)
        return

    try:
        task.preview_images = render_preview_only(paths["output_pdf"], paths, task)
        make_bilingual_pdf_preview(paths["output_pdf"], paths, task)
        task.output_files = [paths["output_pdf"], paths["bilingual_pdf"]]
    except Exception:
        pass

    if completed:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 全部完成！共 {len(done_pages)} 页")
    elif task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停（半成品），共翻 {len(done_pages)} 页"
        task.log_msg(f"🛑 已暂停。下次点开始会从第 {max(done_pages) + 1} 页继续，不重复扣费")
    elif error_msg:
        task.status = "error"
        task.error = error_msg
        task.label = "出错（半成品）"
        task.log_msg(f"❌ {error_msg}")
    else:
        task.status = "paused"
        task.label = "半成品"
    save_state(task)


def docx_worker(task, paths, real_key, model):
    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_json_file(paths["cache_file"], {})

    try:
        wdoc = Document(task.src_path)
    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        task.status = "error"
        task.error = f"打开 Word 失败：{e}"
        task.log_msg(f"❌ 打开 Word 失败：{e}")
        task.log_msg(f"   错误类型：{type(e).__name__}")
        try:
            task.log_msg(f"   文件路径：{task.src_path}")
            task.log_msg(f"   文件大小：{os.path.getsize(task.src_path)} 字节")
        except Exception:
            pass
        if "Package not found" in err_detail or "not a zip" in err_detail.lower():
            task.log_msg("   ⚠️ 这通常意味着 .docx 其实是从 .doc 改后缀来的，")
            task.log_msg("      请用 Word/WPS 打开后『另存为 → Word 文档(*.docx)』再上传。")
        elif "PermissionError" in err_detail or "拒绝访问" in err_detail:
            task.log_msg("   ⚠️ 文件被其他程序占用，请关闭 Word/WPS 后重试。")
        save_state(task)
        return

    try:
        doc_bi = Document(task.src_path)
    except Exception as e:
        task.status = "error"
        task.error = f"打开 Word（双语副本）失败：{e}"
        task.log_msg(task.error)
        save_state(task)
        return

    para_targets = [p for p in wdoc.paragraphs if p.text.strip()]
    cell_targets = []
    seen = set()
    for table in wdoc.tables:
        for row in table.rows:
            for cell in row.cells:
                cid = id(cell._tc)
                if cid in seen:
                    continue
                seen.add(cid)
                for p in cell.paragraphs:
                    if p.text.strip():
                        cell_targets.append(p)

    total_targets = len(para_targets) + len(cell_targets)
    task.status = "running"
    task.total = total_targets
    task.current = 0
    task.label = f"Word · 共 {total_targets} 段"
    task.log_msg(f"📘 Word 已打开，共 {total_targets} 段")
    save_state(task)

    def translate_text_cached(text):
        key = "t_" + h(text)
        if key in cache:
            return cache[key]
        tr = call_simple_api(client, model, text)
        cache[key] = tr
        save_json_file(paths["cache_file"], cache)
        return tr

    done = 0
    error_msg = None
    preview_pairs = []

    for para in para_targets:
        if task.stop_event.is_set():
            task.log_msg("⏸ 检测到停止信号，结束中文版段落翻译")
            break
        src_text = para.text
        try:
            tr = translate_text_cached(src_text)
        except RuntimeError as e:
            error_msg = str(e)
            break

        if len(preview_pairs) < PREVIEW_PARAS:
            preview_pairs.append((src_text, tr))
            task.preview_html = _build_docx_preview_html(preview_pairs)
            save_state(task)

        _docx_replace_para_text(para, tr)
        done += 1
        task.current = done
        task.label = f"中文版 {done}/{total_targets}"
        if done % 3 == 0:
            save_state(task)

    if not error_msg and not task.stop_event.is_set():
        for para in cell_targets:
            if task.stop_event.is_set():
                break
            try:
                tr = translate_text_cached(para.text)
            except RuntimeError as e:
                error_msg = str(e)
                break
            _docx_replace_para_text(para, tr)
            done += 1
            task.current = done
            task.label = f"中文版 {done}/{total_targets}"
            if done % 3 == 0:
                save_state(task)

    try:
        wdoc.save(paths["office_cn"])
        task.log_msg(f"✅ 中文版已保存：{paths['office_cn']}")
    except Exception as e:
        task.status = "error"
        task.error = f"保存 Word 失败：{e}"
        task.log_msg("❌ " + task.error)
        save_state(task)
        return
    save_state(task)

    bi_done = 0
    if not error_msg and not task.stop_event.is_set():
        para_bi = [p for p in doc_bi.paragraphs if p.text.strip()]
        for para in para_bi:
            if task.stop_event.is_set():
                break
            text = para.text
            key = "t_" + h(text)
            try:
                tr = cache[key] if key in cache else call_simple_api(client, model, text)
            except RuntimeError:
                break
            _docx_insert_after(para, tr)
            bi_done += 1
            task.label = f"双语版 {bi_done}/{total_targets}"
            if bi_done % 3 == 0:
                save_state(task)

    try:
        doc_bi.save(paths["office_bi"])
        task.log_msg(f"✅ 双语版已保存：{paths['office_bi']}")
        task.output_files = [paths["office_cn"], paths["office_bi"]]
    except Exception as e:
        task.log_msg(f"⚠️ 保存双语版失败：{e}")
        task.output_files = [paths["office_cn"]]

    if task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停 {done}/{total_targets}"
        task.log_msg(f"🛑 已暂停（半成品），共处理 {done}/{total_targets} 段")
    elif error_msg:
        task.status = "error"
        task.error = error_msg
        task.label = "出错"
        task.log_msg(f"❌ {error_msg}")
    else:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 Word 翻译完成！共 {done} 段")
    save_state(task)


def pptx_worker(task, paths, real_key, model):
    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_json_file(paths["cache_file"], {})

    try:
        prs = Presentation(task.src_path)
    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        task.status = "error"
        task.error = f"打开 PPT 失败：{e}"
        task.log_msg(f"❌ 打开 PPT 失败：{e}")
        task.log_msg(f"   错误类型：{type(e).__name__}")
        try:
            task.log_msg(f"   文件路径：{task.src_path}")
            task.log_msg(f"   文件大小：{os.path.getsize(task.src_path)} 字节")
        except Exception:
            pass
        if "Package not found" in err_detail or "not a zip" in err_detail.lower():
            task.log_msg("   ⚠️ 这通常意味着 .pptx 其实是从 .ppt 改后缀来的，")
            task.log_msg("      请用 PowerPoint/WPS 打开后『另存为 → PowerPoint 演示文稿(*.pptx)』再上传。")
        elif "PermissionError" in err_detail or "拒绝访问" in err_detail:
            task.log_msg("   ⚠️ 文件被其他程序占用，请关闭 PowerPoint/WPS 后重试。")
        save_state(task)
        return

    targets = []
    for si, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                text = "".join(r.text for r in para.runs)
                if text.strip():
                    targets.append((si, para))

    total_targets = len(targets)
    task.status = "running"
    task.total = total_targets
    task.current = 0
    task.label = f"PPT · 共 {total_targets} 段"
    task.log_msg(f"📊 PPT 已打开，共 {total_targets} 段")
    save_state(task)

    done = 0
    error_msg = None
    preview_pairs = []

    for si, para in targets:
        if task.stop_event.is_set():
            task.log_msg("⏸ 检测到停止信号，结束")
            break
        text = "".join(r.text for r in para.runs)
        key = "t_" + h(text)
        try:
            tr = cache[key] if key in cache else call_simple_api(client, model, text)
            cache[key] = tr
            save_json_file(paths["cache_file"], cache)
        except RuntimeError as e:
            error_msg = str(e)
            break

        if len(preview_pairs) < PREVIEW_PARAS:
            preview_pairs.append((si + 1, text, tr))
            task.preview_html = _build_pptx_preview_html(preview_pairs)
            save_state(task)

        if para.runs:
            para.runs[0].text = tr
            for r in para.runs[1:]:
                r.text = ""
        done += 1
        task.current = done
        task.label = f"PPT {done}/{total_targets}（第 {si+1} 张）"
        if done % 3 == 0:
            save_state(task)

    try:
        prs.save(paths["office_cn"])
        task.output_files = [paths["office_cn"]]
    except Exception as e:
        task.status = "error"
        task.error = f"保存 PPT 失败：{e}"
        task.log_msg("❌ " + task.error)
        save_state(task)
        return

    if task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停 {done}/{total_targets}"
        task.log_msg(f"🛑 已暂停（半成品），共处理 {done}/{total_targets} 段")
    elif error_msg:
        task.status = "error"
        task.error = error_msg
        task.label = "出错"
        task.log_msg(f"❌ {error_msg}")
    else:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 PPT 翻译完成！共 {done} 段")
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

    task = MANAGER.create(kind, src_copy, src_name, paths["out_dir"], paths["work"])
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


def on_start(api_key, model, doc_file, mode, trial):
    """自动按文件后缀判断类型，忽略页面上的选择"""
    global SELECTED_TASK_ID

    real_key = (api_key or "").strip() or DEFAULT_API_KEY
    if not real_key or not doc_file:
        return on_refresh_fast() + ([], None)

    model = (model or DEFAULT_MODEL).strip()
    if model not in ("deepseek-chat", "deepseek-reasoner"):
        model = "deepseek-chat"

    # 根据后缀自动判断类型
    name_lower = (doc_file.name or "").lower()

    if name_lower.endswith(".pdf"):
        kind = "pdf"
    elif name_lower.endswith(".docx"):
        kind = "docx"
    elif name_lower.endswith(".pptx"):
        kind = "pptx"
    elif name_lower.endswith(".doc"):
        cur = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
        if cur:
            cur.log_msg("❌ 不支持老版 .doc 格式")
            cur.log_msg("   请用 Word/WPS 打开后『另存为 → Word 文档 (*.docx)』再上传。")
            save_state(cur)
        return on_refresh_fast() + ([], None)
    elif name_lower.endswith(".ppt"):
        cur = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
        if cur:
            cur.log_msg("❌ 不支持老版 .ppt 格式")
            cur.log_msg("   请用 PowerPoint/WPS 打开后『另存为 → PowerPoint 演示文稿 (*.pptx)』再上传。")
            save_state(cur)
        return on_refresh_fast() + ([], None)
    else:
        cur = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
        if cur:
            cur.log_msg(f"❌ 不支持的文件类型：{doc_file.name}")
            cur.log_msg("   只支持 .pdf / .docx / .pptx")
            save_state(cur)
        return on_refresh_fast() + ([], None)

    if kind in ("docx", "pptx") and not HAS_OFFICE:
        cur = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
        if cur:
            cur.log_msg("❌ 未安装 python-docx / python-pptx，无法处理 Word / PPT")
            cur.log_msg("   请运行：pip install python-docx python-pptx")
            save_state(cur)
        return on_refresh_fast() + ([], None)

    if kind == "pdf" and not os.path.exists(FONT_PATH):
        cur = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
        if cur:
            cur.log_msg(f"❌ 字体文件不存在：{FONT_PATH}")
            save_state(cur)
        return on_refresh_fast() + ([], None)

    src_name = os.path.basename(doc_file.name)
    base = os.path.splitext(src_name)[0]
    book = safe_dirname(base)
    target_out_dir = os.path.abspath(os.path.join(RESULT_ROOT, book))

    existing = MANAGER.find_active_by_out_dir(target_out_dir)
    if existing is not None:
        SELECTED_TASK_ID = existing.task_id
        existing.log_msg(f"⚠️ 收到重复创建请求，已拒绝。本任务已存在并正在运行（#{existing.task_id}）")
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
        t.label = f"⏸ 停止信号已发出，等待当前步骤完成…（已处理 {t.current}/{t.total}）"
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
        t.label = f"⏸ 停止信号已发出，等待当前步骤完成…（已处理 {t.current}/{t.total}）"
        t.log_msg("⏸ 收到单独停止信号 —— 当前一步完成后暂停")
        save_state(t)
        SELECTED_TASK_ID = tid
    elif t.status == "stopping":
        t.log_msg("⏸ 该任务已在停止中，无需重复操作")
        save_state(t)

    return on_refresh_fast()


def on_load_preview():
    current = MANAGER.get(SELECTED_TASK_ID) if SELECTED_TASK_ID else None
    if current is None:
        tasks = MANAGER.all_sorted()
        current = tasks[0] if tasks else None
    if current is None:
        return build_preview_html([]), []

    imgs = current.preview_images if current.preview_images else []
    ph = getattr(current, "preview_html", "") or ""
    files = [os.path.abspath(f) for f in current.output_files if f and os.path.exists(f)]
    return build_preview_html(imgs, ph), files


def build_task_list_html(tasks):
    if not tasks:
        return '<div style="padding:14px;color:#888;font-size:13px">暂无任务</div>'

    status_icon = {
        "queued": "⏳", "running": "▶️", "stopping": "⏸️",
        "paused": "🛑", "done": "✅", "error": "❌",
    }
    status_color = {
        "queued": "#888", "running": "#0a7d32", "stopping": "#b8860b",
        "paused": "#b8860b", "done": "#0a7d32", "error": "#c0392b",
    }
    status_text = {
        "queued": "排队中", "running": "运行中", "stopping": "停止中…",
        "paused": "已暂停", "done": "已完成", "error": "出错",
    }
    kind_icon = {"pdf": "📕", "docx": "📘", "pptx": "📊"}

    rows = []
    for t in tasks[:8]:
        ic = status_icon.get(t.status, "•")
        col = status_color.get(t.status, "#666")
        stx = status_text.get(t.status, t.status)
        kc = kind_icon.get(t.kind, "📄")
        pct = int(t.current * 100 / t.total) if t.total else 0
        rows.append(f'''
        <div style="padding:10px 14px;border-bottom:1px solid #eee;display:flex;align-items:center;gap:10px;font-size:13.5px">
          <span style="font-size:16px">{ic}</span>
          <span style="font-size:16px">{kc}</span>
          <span style="font-family:Consolas,monospace;color:#666">#{t.task_id}</span>
          <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">{t.src_name}</span>
          <span style="color:{col};font-weight:600">{stx}</span>
          <span style="color:#666">{t.current}/{t.total} ({pct}%)</span>
        </div>''')

    return f'<div style="background:#fff;border:1px solid #e6dfce;border-radius:10px;overflow:hidden">{"".join(rows)}</div>'


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
        progress_html = make_progress_html(current.current, current.total, current.label)
        log_text = "\n".join(current.log[-40:])
        previews = current.preview_images or []
        preview_html = getattr(current, "preview_html", "") or ""

    if current is not None:
        if preview_html:
            sig = ("html", len(preview_html), preview_html[:30], preview_html[-30:])
        else:
            sig = preview_signature(previews)

        if _LAST_PREVIEW_SIG.get(current.task_id) != sig:
            _LAST_PREVIEW_SIG[current.task_id] = sig
            html = build_preview_html(previews, preview_html)
            _PREVIEW_HTML_CACHE[current.task_id] = html
            gallery_value = html
        else:
            gallery_value = _PREVIEW_HTML_CACHE.get(
                current.task_id, build_preview_html([])
            )
    else:
        gallery_value = build_preview_html([])

    active = [t for t in tasks if t.status in ("queued", "running", "stopping")]
    choices = []
    for t in active:
        short = t.src_name if len(t.src_name) <= 50 else t.src_name[:47] + "..."
        choices.append(f"#{t.task_id}  {short}  [{t.status}]")

    dd_update = gr.update(choices=choices)

    return task_list_html, progress_html, log_text, dd_update, gallery_value


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
            os.startfile(folder)
        elif sys.platform == "darwin":
            import subprocess; subprocess.Popen(["open", folder])
        else:
            import subprocess; subprocess.Popen(["xdg-open", folder])
        return f"✅ 已打开：{folder}"
    except Exception as e:
        return f"❌ 打开失败：{e}\n路径：{folder}"


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
        max-width: 1320px !important;
        margin: 0 auto !important;
        padding: 8px 28px 40px !important;
    }

    .card-head {
        background: #f5f1e6; border: 1px solid #e2dccb;
        border-radius: 14px; padding: 28px 32px 24px;
        margin-bottom: 22px; position: relative; overflow: hidden;
    }
    .card-head::before {
        content: ""; position: absolute; top: 0; left: 0; right: 0;
        height: 3px; background: linear-gradient(90deg, #2b2b2b, #8a7a4f, #2b2b2b);
    }
    .card-head h1 {
        font-size: 27px; font-weight: 700; color: #111; margin: 0;
        font-family: "Noto Serif SC", Georgia, serif;
    }
    .card-head .sub { color: #555; font-size: 14.5px; margin-top: 12px; line-height: 1.95; }
    .card-head .sub .dot { color: #b09b63; margin: 0 10px; }
    .card-head .meta {
        margin-top: 16px; padding-top: 16px;
        border-top: 1px dashed #d8d0bc;
        font-size: 13.5px; color: #555; line-height: 2.15;
    }
    .card-head code {
        background: #ebe5d5; color: #333;
        padding: 2px 8px; border-radius: 4px;
        font-size: 13px; font-family: Consolas, Monaco, monospace;
    }
    .card-head .k { color: #2b2b2b; font-weight: 600; margin-right: 6px; }

    label span, .gr-box > label > span {
        color: #1a1a1a !important;
        font-size: 14.5px !important; font-weight: 500 !important;
    }
    .gradio-container input:not([type="checkbox"]):not([type="radio"]),
    .gradio-container textarea,
    .gradio-container select {
        background: #ffffff !important; color: #111 !important;
        border: 1px solid #d8d2c0 !important; font-size: 14.5px !important;
    }
    .gradio-container .block {
        background: #ffffff !important;
        border: 1px solid #e6dfce !important; border-radius: 10px !important;
    }

    #file_mode { padding: 14px 18px !important; }
    #file_mode .wrap, #file_mode > div > div {
        display: flex !important; flex-direction: row !important;
        gap: 24px !important; flex-wrap: wrap !important;
    }
    #file_mode label { font-size: 15px !important; font-weight: 500 !important; cursor: pointer !important; }
    #file_mode input[type="radio"] {
        -webkit-appearance: radio !important; appearance: radio !important;
        width: 18px !important; height: 18px !important;
        min-width: 18px !important; max-width: 18px !important;
        accent-color: #2b2b2b !important; margin-right: 8px !important;
    }

    #trial_cb { background: transparent !important; border: none !important; padding: 8px 12px !important; }
    #trial_cb * { cursor: pointer !important; }
    #trial_cb input[type="checkbox"] {
        -webkit-appearance: checkbox !important; appearance: checkbox !important;
        width: 18px !important; height: 18px !important;
        min-width: 18px !important; max-width: 18px !important;
        accent-color: #2b2b2b !important; margin-right: 10px !important;
    }
    #trial_cb label { cursor: pointer !important; font-size: 14.5px !important; }

    #model_dd { min-width: 260px !important; }
    #model_dd input, #model_dd .wrap-inner { min-width: 240px !important; }

    #pdf_upload {
        min-height: 130px !important;
        max-height: 200px !important;
        overflow-y: auto !important;
        overflow-x: hidden !important;
        box-sizing: border-box !important;
    }
    #pdf_upload > div { padding: 10px 14px !important; }
    #pdf_upload button {
        padding: 8px 20px !important; font-size: 14px !important;
        min-height: 40px !important; height: auto !important;
        cursor: pointer !important; pointer-events: auto !important;
    }
    #pdf_upload .file {
        padding: 8px 12px !important; margin: 6px 0 !important;
        font-size: 14px !important; background: #faf8f2 !important;
        border-radius: 6px !important; overflow: hidden !important;
        text-overflow: ellipsis !important; white-space: nowrap !important;
    }
    #pdf_upload .file * { font-size: 14px !important; color: #1a1a1a !important; line-height: 1.5 !important; }

    #action_grid .gr-row, #action_grid .row {
        display: flex !important; gap: 12px !important; margin-bottom: 12px !important;
    }
    #action_grid .gr-row > *, #action_grid .row > * {
        flex: 1 1 0 !important; min-width: 0 !important;
    }
    #action_grid button {
        width: 100% !important; height: 52px !important;
        font-size: 15px !important; font-weight: 600 !important; border-radius: 10px !important;
    }
    #action_grid .primary {
        background: #2b2b2b !important; color: #faf8f2 !important; border: none !important;
    }
    #action_grid .primary:hover { background: #000 !important; }
    #action_grid .secondary {
        background: #f0ebdc !important; color: #222 !important; border: 1px solid #ddd5c0 !important;
    }
    #action_grid .secondary:hover {
        background: #e8e1cc !important; border-color: #c9bfa3 !important;
    }

    #stop_one_row { gap: 12px !important; align-items: stretch !important; }
    #stop_dd { flex: 3 1 0 !important; min-width: 0 !important; }
    #stop_one_btn {
        flex: 1 1 0 !important; min-width: 180px !important;
        height: 52px !important; border-radius: 10px !important;
        background: #8b3a3a !important; color: #fff !important;
        border: none !important; font-size: 15px !important; font-weight: 600 !important;
    }
    #stop_one_btn:hover { background: #6b2222 !important; }

    #preview_box::-webkit-scrollbar { width: 10px; }
    #preview_box::-webkit-scrollbar-track { background: #e8e1cc; border-radius: 5px; }
    #preview_box::-webkit-scrollbar-thumb { background: #b09b63; border-radius: 5px; }
    #preview_box::-webkit-scrollbar-thumb:hover { background: #8a7a4f; }

    .gradio-container .accordion-header {
        background: #f5f1e6 !important; color: #222 !important; font-size: 14.5px !important;
    }
    .gradio-container h3 { color: #111 !important; font-size: 18px !important; margin-top: 20px !important; }
    .gradio-container .prose, .gradio-container .prose * { color: #1a1a1a !important; }
    """,
) as demo:

    gr.HTML("""
    <div class="card-head">
      <h1>📖 PDF / Word / PPT 翻译器</h1>
      <div class="sub">
        三种格式<span class="dot">·</span>保留原排版
        <span class="dot">·</span>多任务并行<span class="dot">·</span>断点可续
      </div>
      <div class="meta">
        <span class="k">🔑 密钥</span>留空取 <code>.env</code> 中的默认值，亦可临时填入覆盖
        <br>
        <span class="k">📁 成果</span>归于 <code>result/&lt;文件名&gt;/</code>
        <br>
        <span class="k">📄 上传</span>拖入文件即自动识别类型（.pdf / .docx / .pptx），无需手动选。<br>
        　　　　　 下方"文档类型"选项只作参考，不影响实际处理。
        <br>
        <span class="k">🔄 多任务</span>上传文件 → 点「创建新任务」→ 上传框自动清空，可以继续传下一个。
        <br>
        <span class="k">⏸ 停止</span>下方可<b>单独停止</b>某个任务，也可<b>一键停止全部</b>。
        <br>
        <span class="k">👀 预览</span>PDF 显示前 5 页左右对照；Word / PPT 显示前 5 段文本对照。
      </div>
    </div>
    """)

    with gr.Row():
        api_key = gr.Textbox(
            label="🔑 DeepSeek API Key（留空用 .env 默认）",
            type="password",
            placeholder="sk-...　留空 → 用 .env；填入 → 临时覆盖",
            scale=3,
        )
        model = gr.Dropdown(
            choices=["deepseek-chat", "deepseek-reasoner"],
            value=DEFAULT_MODEL,
            label="🧠 翻译模型",
            scale=2,
            elem_id="model_dd",
        )

    file_mode = gr.Radio(
        choices=["📕 PDF 书籍", "📘 Word 文档", "📊 PPT 演示"],
        value="📕 PDF 书籍",
        label="📂 文档类型（仅参考，实际按文件后缀自动判断）",
        elem_id="file_mode",
    )

    doc_file = gr.File(
        label="📄 上传文档以创建新任务（自动识别 .pdf / .docx / .pptx）",
        file_types=[".pdf", ".docx", ".pptx"],
        elem_id="pdf_upload",
    )
    trial = gr.Checkbox(
        label="🧪 试翻模式：PDF 只翻前 5 页（对 Word / PPT 无效）",
        value=True,
        elem_id="trial_cb",
    )

    with gr.Column(elem_id="action_grid"):
        with gr.Row(equal_height=True):
            btn = gr.Button("▶ 创建新任务并开始", variant="primary")
            stop_btn = gr.Button("⏹ 停止所有任务", variant="secondary")
        with gr.Row(equal_height=True):
            refresh_btn = gr.Button("🔄 手动刷新状态", variant="secondary")
            open_btn = gr.Button("📁 打开当前任务文件夹", variant="secondary")

    gr.Markdown("### 📋 任务列表（最近 8 条，刷新页面后仍保留）")
    task_list_html = gr.HTML(value='<div style="padding:14px;color:#888;font-size:13px">暂无任务</div>')

    gr.Markdown("### 🛑 单独停止某个任务")
    with gr.Row(elem_id="stop_one_row"):
        stop_dd = gr.Dropdown(
            label="选择要停止的任务（只列正在运行的任务）",
            choices=[],
            value=None,
            interactive=True,
            elem_id="stop_dd",
        )
        stop_one_btn = gr.Button("⏹ 停止选中任务", elem_id="stop_one_btn")

    gr.Markdown("### 📊 当前任务进度")
    progress_bar = gr.HTML(value=make_progress_html(0, 1, "等待开始"))

    with gr.Accordion("📋 当前任务日志（点击展开 / 收起）", open=False):
        log = gr.Textbox(label="", lines=14, interactive=False, show_label=False)

    gr.Markdown("### 👀 效果预览（PDF 图片对照 · Word/PPT 文本对照）")
    gallery = gr.HTML(
        value=build_preview_html([]),
        elem_id="preview_box",
    )

    gr.Markdown("### 💾 下载（可选）")
    out_files = gr.File(
        label="",
        file_count="multiple",
        interactive=True,
        show_label=False,
    )

    load_btn = gr.Button("🔍 加载当前任务的预览图和下载文件", variant="secondary")

    fast_outputs = [task_list_html, progress_bar, log, stop_dd, gallery]
    full_outputs = [task_list_html, progress_bar, log, stop_dd, gallery, out_files, doc_file]

    btn.click(
        on_start,
        [api_key, model, doc_file, file_mode, trial],
        full_outputs,
        concurrency_limit=None,
        concurrency_id="start",
    )
    stop_btn.click(
        on_stop_all,
        None,
        fast_outputs,
        concurrency_limit=None,
        concurrency_id="stop",
    )
    stop_one_btn.click(
        on_stop_selected,
        [stop_dd],
        fast_outputs,
        concurrency_limit=None,
        concurrency_id="stop_one",
    )
    refresh_btn.click(
        on_refresh_fast,
        None,
        fast_outputs,
        concurrency_limit=None,
        concurrency_id="manual",
    )
    open_btn.click(open_result_folder, None, [log])
    load_btn.click(on_load_preview, None, [gallery, out_files])

    try:
        timer = gr.Timer(3.0)
        timer.tick(
            on_refresh_fast,
            None,
            fast_outputs,
            concurrency_limit=None,
            concurrency_id="tick",
        )
    except Exception:
        pass

    demo.load(
        on_refresh_fast,
        None,
        fast_outputs,
        concurrency_limit=None,
        concurrency_id="load",
    )


if __name__ == "__main__":
    PORT = 7860
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
    if not HAS_OFFICE:
        print("   ⚠️ 未装 python-docx / python-pptx，Word / PPT 不可用")
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
        allowed_paths=[os.path.abspath(".")],
    )
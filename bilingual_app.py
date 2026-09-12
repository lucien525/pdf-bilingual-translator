# -*- coding: utf-8 -*-
"""
PDF / Word / PPT 翻译器（网页版 · 多任务并行 + 单独停止 + 断点续传）
"""

import os
import sys
import json
import time
import re
import hashlib
import shutil
import threading
import uuid
from dataclasses import dataclass, field
from typing import Optional

import pymupdf as fitz
# 关闭 PDF 内部损坏对象产生的警告刷屏（如 "cannot find object in xref"）
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

# 全局 IO 锁：防止多线程同时写同一个 JSON 文件导致损坏
_IO_LOCK = threading.Lock()

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


def render_page_pngs(src_path, trans_path, num_pages, img_dir):
    os.makedirs(img_dir, exist_ok=True)
    orig = fitz.open(src_path)
    trans = fitz.open(trans_path)
    n = min(num_pages, len(orig), len(trans))
    orig_paths, trans_paths = [], []
    for i in range(n):
        o_path = os.path.join(img_dir, f"orig_{i:04d}.png")
        t_path = os.path.join(img_dir, f"trans_{i:04d}.png")
        orig[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM)).save(o_path)
        trans[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM)).save(t_path)
        orig_paths.append(o_path)
        trans_paths.append(t_path)
    orig.close()
    trans.close()
    return orig_paths, trans_paths


def make_side_by_side(orig_png, trans_png, out_png):
    o = Image.open(orig_png).convert("RGB")
    t = Image.open(trans_png).convert("RGB")
    hh = max(o.height, t.height)
    if o.height != hh:
        o = o.resize((int(o.width * hh / o.height), hh), Image.LANCZOS)
    if t.height != hh:
        t = t.resize((int(t.width * hh / t.height), hh), Image.LANCZOS)
    gap = 8
    canvas = Image.new("RGB", (o.width + gap + t.width, hh), (40, 40, 40))
    canvas.paste(o, (0, 0))
    canvas.paste(t, (o.width + gap, 0))
    canvas.thumbnail((1800, 1400), Image.LANCZOS)
    canvas.save(out_png)
    return out_png


def make_bilingual_pdf(orig_paths, trans_paths, out_pdf):
    pages = []
    for o_path, t_path in zip(orig_paths, trans_paths):
        o = Image.open(o_path).convert("RGB")
        t = Image.open(t_path).convert("RGB")
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
    if pages:
        pages[0].save(out_pdf, save_all=True, append_images=pages[1:], resolution=120.0)
    return out_pdf


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

    doc = fitz.open(task.src_path)
    total = len(doc)
    limit = min(5, total) if trial else total

    task.status = "running"
    task.total = limit
    task.label = f"PDF · 目标前 {limit} 页"
    task.current = len([p for p in done_pages if p <= limit])
    task.log_msg(f"✅ PDF 共 {total} 页，本次目标前 {limit} 页")
    if task.current > 0:
        task.log_msg(f"📚 已处理 {task.current} 页，从第 {max(done_pages)+1} 页继续（跳过缓存）")
    save_state(task)

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

        page = doc[pno]
        blocks = [b for b in page.get_text("blocks") if b[6] == 0 and b[4].strip()]

        if page_num in done_pages:
            if blocks:
                try:
                    trans = translate_page(client, model, blocks, cache, paths["cache_file"])
                    apply_translations(page, blocks, trans)
                except Exception:
                    pass
            continue

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
    else:
        completed = True

    try:
        doc.save(paths["output_pdf"], deflate=True)
    finally:
        doc.close()

    if not done_pages:
        task.status = "error"
        task.error = "本次没有任何页面成功翻译"
        task.log_msg("⚠️ " + task.error)
        save_state(task)
        return

    render_up_to = max(done_pages)
    task.label = f"渲染前 {render_up_to} 页"
    save_state(task)

    orig_paths, trans_paths = render_page_pngs(task.src_path, paths["output_pdf"], render_up_to, paths["img_dir"])
    make_bilingual_pdf(orig_paths, trans_paths, paths["bilingual_pdf"])

    os.makedirs(paths["preview_dir"], exist_ok=True)
    preview_imgs = []
    for i in range(min(6, len(orig_paths))):
        try:
            p = os.path.join(paths["preview_dir"], f"compare_{i:04d}.png")
            make_side_by_side(orig_paths[i], trans_paths[i], p)
            preview_imgs.append(p)
        except Exception:
            pass

    task.preview_images = preview_imgs
    task.output_files = [paths["output_pdf"], paths["bilingual_pdf"]]

    if completed:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 全部完成！共 {len(done_pages)} 页")
    elif task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停（半成品），续自第 {render_up_to + 1} 页"
        task.log_msg(f"🛑 已暂停。下次从第 {render_up_to + 1} 页继续，不重复扣费")
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
        doc_bi = Document(task.src_path)
    except Exception as e:
        task.status = "error"
        task.error = f"打开 Word 失败：{e}"
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
    for para in para_targets:
        if task.stop_event.is_set():
            task.log_msg("⏸ 检测到停止信号，结束中文版段落翻译")
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
        task.status = "error"
        task.error = f"打开 PPT 失败：{e}"
        task.log_msg(task.error)
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
    global SELECTED_TASK_ID

    real_key = (api_key or "").strip() or DEFAULT_API_KEY
    if not real_key or not doc_file:
        t, p, l, dd = on_refresh_fast()
        return t, p, l, [], [], doc_file, dd

    model = (model or DEFAULT_MODEL).strip()
    if model not in ("deepseek-chat", "deepseek-reasoner"):
        model = "deepseek-chat"

    is_pdf = mode.startswith("📕")
    is_docx = mode.startswith("📘")
    is_pptx = mode.startswith("📊")

    if (is_docx or is_pptx) and not HAS_OFFICE:
        t, p, l, dd = on_refresh_fast()
        return t, p, l, [], [], doc_file, dd
    if is_pdf and not os.path.exists(FONT_PATH):
        t, p, l, dd = on_refresh_fast()
        return t, p, l, [], [], doc_file, dd

    src_name = os.path.basename(doc_file.name)
    base = os.path.splitext(src_name)[0]
    book = safe_dirname(base)
    target_out_dir = os.path.abspath(os.path.join(RESULT_ROOT, book))

    existing = MANAGER.find_active_by_out_dir(target_out_dir)
    if existing is not None:
        SELECTED_TASK_ID = existing.task_id
        existing.log_msg(f"⚠️ 收到重复创建请求，已拒绝。本任务已存在并正在运行（#{existing.task_id}）")
        existing.log_msg("   如需重新开始，请先停掉它")
        save_state(existing)
        t, p, l, dd = on_refresh_fast()
        return t, p, l, [], [], None, dd

    kind = "pdf" if is_pdf else ("docx" if is_docx else "pptx")
    try:
        start_task(kind, doc_file.name, real_key, model, trial)
    except Exception as e:
        print(f"启动任务失败：{e}")

    t, p, l, dd = on_refresh_fast()
    return t, p, l, [], [], None, dd


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
    """停止用户在下拉框里选中的那一个任务"""
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
        return [], []

    imgs = current.preview_images if current.preview_images else []
    files = [os.path.abspath(f) for f in current.output_files if f and os.path.exists(f)]
    return imgs, files


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
    else:
        progress_html = make_progress_html(current.current, current.total, current.label)
        log_text = "\n".join(current.log[-40:])

    # 只列活跃任务（queued / running / stopping）
    active = [t for t in tasks if t.status in ("queued", "running", "stopping")]
    choices = []
    for t in active:
        short = t.src_name if len(t.src_name) <= 50 else t.src_name[:47] + "..."
        choices.append(f"#{t.task_id}  {short}  [{t.status}]")

    dd_update = gr.update(choices=choices)

    return task_list_html, progress_html, log_text, dd_update


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

    /* 单独停止任务的下拉 + 按钮 */
    #stop_one_row { gap: 12px !important; align-items: stretch !important; }
    #stop_dd { flex: 3 1 0 !important; min-width: 0 !important; }
    #stop_one_btn {
        flex: 1 1 0 !important; min-width: 180px !important;
        height: 52px !important; border-radius: 10px !important;
        background: #8b3a3a !important; color: #fff !important;
        border: none !important; font-size: 15px !important; font-weight: 600 !important;
    }
    #stop_one_btn:hover { background: #6b2222 !important; }

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
        <span class="k">🔄 多任务</span>上传文件 → 点「创建新任务」→ 上传框自动清空，可以继续传下一个（不同书）。<br>
        　　　　　 <b>同一本书只允许一个任务运行</b>，重复点会被拒绝。
        <br>
        <span class="k">⏸ 停止</span>下方可<b>单独停止</b>某个任务，也可<b>一键停止全部</b>。
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
        label="📂 文档类型（先选类型，再上传文件）",
        elem_id="file_mode",
    )

    doc_file = gr.File(
        label="📄 上传文档以创建新任务（支持 .pdf / .docx / .pptx，刷新后不保留）",
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

    gr.Markdown("### 👀 效果预览（PDF 任务显示左右对照，最多前 6 张）")
    gallery = gr.Gallery(
        label="对照预览",
        columns=1,
        height=600,
        object_fit="contain",
        show_label=False,
    )

    gr.Markdown("### 💾 下载（可选）")
    out_files = gr.File(
        label="",
        file_count="multiple",
        interactive=True,
        show_label=False,
    )

    load_btn = gr.Button("🔍 加载当前任务的预览图和下载文件", variant="secondary")

    fast_outputs = [task_list_html, progress_bar, log, stop_dd]
    full_outputs = [task_list_html, progress_bar, log, gallery, out_files, doc_file, stop_dd]

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
    except Exception as e:
        print(f"⚠️ 未能启用自动刷新（Timer 不可用）：{e}")

    demo.load(
        on_refresh_fast,
        None,
        fast_outputs,
        concurrency_limit=None,
        concurrency_id="load",
    )


if __name__ == "__main__":
    if DEFAULT_API_KEY:
        print(f"✅ 已从 .env 读取默认 Key（{DEFAULT_API_KEY[:6]}...）")
    else:
        print("ℹ️  .env 中未找到 DEEPSEEK_API_KEY，需在网页里手动填写")
    print(f"ℹ️  默认模型：{DEFAULT_MODEL}")
    print(f"ℹ️  结果目录：{os.path.abspath(RESULT_ROOT)}")

    scan_all_states()
    n = len(MANAGER.all_sorted())
    if n:
        print(f"ℹ️  已恢复 {n} 个历史任务")

    if not HAS_OFFICE:
        print("⚠️  未检测到 python-docx / python-pptx，Word / PPT 模式不可用")
        print("    如需使用，请运行： pip install python-docx python-pptx")

    demo.queue(default_concurrency_limit=None)
    demo.launch(
        inbrowser=True,
        allowed_paths=[os.path.abspath(".")],
    )
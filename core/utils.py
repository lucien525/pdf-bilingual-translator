# -*- coding: utf-8 -*-
"""基础工具：路径 / 哈希 / JSON 读写节流 / 体积预估。"""

import os
import re
import json
import time
import glob
import hashlib
import threading

from core import config

# ★ 修复：缓存 JSON 写节流（避免每页/每段整文件重写放大 IO）
_IO_LOCK = threading.Lock()
_JSON_DEBOUNCE_SEC = 3.0
_JSON_PENDING = {}
_JSON_LAST_WRITE = {}


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
    out_dir = os.path.abspath(os.path.join(config.RESULT_ROOT, out_subdir))
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
                preset = config.PDF_QUALITY_PRESETS.get(
                    pdf_quality or config.DEFAULT_PDF_QUALITY,
                    config.PDF_QUALITY_PRESETS[config.DEFAULT_PDF_QUALITY]
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

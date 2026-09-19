# -*- coding: utf-8 -*-
"""作业解题器 · Gradio 界面：build_ui() + 所有回调 + UI 专属全局状态。
布局与翻译器（ui.py）保持一致：全宽折叠条 ①~⑤。"""

import os
import re
import sys
import json
import uuid
import shutil
import atexit
import signal
import hashlib
import threading
from collections import OrderedDict
from html import escape

import gradio as gr

import hw_core
from hw_tasks import (MANAGER, save_state, refresh_task_size, _STATE_SAVE_TS,
                      pdf_worker, docx_worker, pptx_worker)
from hw_solve import _SUBJECT_PROMPTS
from hw_preview import (build_preview_html, make_progress_html,
                        build_done_modal_html, preview_signature,
                        _collect_existing_previews)

# ================= UI 专属全局状态 =================
_TASK_FILTER = "all"
_CREATE_LOCK = threading.Lock()
_REFRESH_LOCK = threading.RLock()
_LAST_PREVIEW_SIG = {}
_PREVIEW_HTML_CACHE = {}
_MODAL_SHOWN = OrderedDict()
_MODAL_SHOWN_MAX = 500
_SELECTED_STOP_VALUE = None
SELECTED_TASK_ID = None


# ============================================================
# 任务控制
# ============================================================
def start_task(kind, upload_path, real_key, model, trial,
               subject="general", extra_prompt="",
               ans_position="inside", parse_only=False,
               trial_pages=5):
    global SELECTED_TASK_ID
    src_name = os.path.basename(upload_path)
    paths = hw_core.prepare_paths(src_name)
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
                pdf_worker(task, paths, real_key, model, trial,
                           subject=subject, extra_prompt=extra_prompt,
                           ans_position=ans_position, parse_only=parse_only,
                           trial_pages=trial_pages)
            elif kind == "docx":
                docx_worker(task, paths, real_key, model,
                            subject=subject, extra_prompt=extra_prompt)
            elif kind == "pptx":
                pptx_worker(task, paths, real_key, model,
                            subject=subject, extra_prompt=extra_prompt)
        except Exception as e:
            task.status = "error"
            task.error = str(e)
            task.log_msg(f"❌ 未捕获错误：{e}")
            save_state(task, force=True)
        finally:
            # ★ 修复：任务结束（含出错/停止）时把节流的缓存 JSON 落盘
            hw_core.flush_all_json()

    task.thread = threading.Thread(target=runner, daemon=True)
    task.thread.start()
    return task


def on_start(api_key, model, doc_file, mode, trial, trial_pages,
             subject, custom_prompt, ans_position, parse_only,
             stop_dd_value=None):
    global SELECTED_TASK_ID

    keep_upload = gr.update()

    real_key = (api_key or "").strip() or hw_core.DEFAULT_API_KEY
    if not real_key or not doc_file:
        # ★ 修复：给用户可见反馈
        gr.Warning("请先上传作业文档（并确认已填写 API Key）")
        return on_refresh_fast(stop_dd_value) + ([], keep_upload)

    model = (model or hw_core.DEFAULT_MODEL).strip()
    if model not in ("deepseek-chat", "deepseek-reasoner"):
        model = "deepseek-chat"

    if subject not in _SUBJECT_PROMPTS:
        subject = "general"

    # ★ 修复：试解页数可选（3/5/10/20）
    try:
        trial_pages = int(trial_pages)
    except Exception:
        trial_pages = 5
    if trial_pages not in (3, 5, 10, 20):
        trial_pages = 5

    name_lower = (doc_file.name or "").lower()
    _ext = os.path.splitext(doc_file.name)[1].lower()
    if name_lower.endswith(".pdf"):
        kind = "pdf"
    elif name_lower.endswith(".docx"):
        kind = "docx"
    elif name_lower.endswith(".pptx"):
        kind = "pptx"
    else:
        # ★ 修复：错误直接反馈给用户，不再写进旧任务的日志
        gr.Warning(f"不支持的文件类型：{doc_file.name}"
                   f"（仅支持 .pdf / .docx / .pptx）")
        return on_refresh_fast(stop_dd_value) + ([], keep_upload)

    _mode_ext = {
        "📕 PDF 书籍": ".pdf",
        "📘 Word 文档": ".docx",
        "📊 PPT 演示": ".pptx",
    }
    if mode in _mode_ext and not name_lower.endswith(_mode_ext[mode]):
        # ★ 对齐翻译器：文档类型选择与后缀不一致时提示
        gr.Warning(f"ℹ️ 文档类型选择与文件后缀不一致，已按后缀（{_ext}）处理")

    if kind in ("docx", "pptx") and not hw_core.HAS_OFFICE:
        gr.Warning("未安装 python-docx / python-pptx，无法处理 Word / PPT")
        return on_refresh_fast(stop_dd_value) + ([], keep_upload)

    src_name = os.path.basename(doc_file.name)
    base = os.path.splitext(src_name)[0]
    book = hw_core.safe_dirname(base)
    short_hash = hw_core.h(src_name)[:6]
    target_out_dir = os.path.abspath(
        os.path.join(hw_core.RESULT_ROOT, f"{book}_{short_hash}")
    )

    with _CREATE_LOCK:
        existing = MANAGER.find_active_by_out_dir(target_out_dir)
        if existing is not None:
            SELECTED_TASK_ID = existing.task_id
            existing.log_msg(f"⚠️ 已存在运行中的任务（#{existing.task_id}）")
            save_state(existing, force=True)
            return on_refresh_fast(stop_dd_value) + ([], keep_upload)

        try:
            start_task(kind, doc_file.name, real_key, model, trial,
                       subject=subject,
                       extra_prompt=custom_prompt or "",
                       ans_position=ans_position or "inside",
                       parse_only=bool(parse_only),
                       trial_pages=trial_pages)
        except Exception as e:
            print(f"[on_start] 创建任务失败：{e}")
            # ★ 修复：错误反馈到网页，而不是只进控制台
            gr.Warning(f"创建任务失败：{e}")

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
        t.label = f"⏸ 停止信号已发出…（已处理 {t.current}/{t.total}）"
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
        t.label = f"⏸ 停止信号已发出…（已处理 {t.current}/{t.total}）"
        t.log_msg("⏸ 收到单独停止信号")
        save_state(t, force=True)
        SELECTED_TASK_ID = tid
    elif t.status == "stopping":
        # ★ 修复：对齐双语版，重复点击停止也有反馈
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
    if not imgs:
        imgs = _collect_existing_previews(current)
        if imgs:
            current.preview_images = imgs
    ph = getattr(current, "preview_html", "") or ""
    files = [os.path.abspath(f) for f in (current.output_files or [])
             if f and os.path.exists(f)]
    return build_preview_html(imgs, ph), files


def on_filter_change(value):
    """★ 新增：任务列表筛选切换"""
    global _TASK_FILTER
    with _REFRESH_LOCK:
        _TASK_FILTER = value or "all"
    return on_refresh_fast()


def build_task_list_html(tasks):
    if not tasks:
        return ('<div style="padding:18px;color:#a9a49a;font-size:13px;'
                'text-align:center">暂无任务</div>')

    # ★ 新增：按 _TASK_FILTER 过滤
    total_before = len(tasks)
    if _TASK_FILTER == "running":
        tasks = [t for t in tasks
                 if t.status in ("queued", "running", "stopping")]
    elif _TASK_FILTER == "done":
        tasks = [t for t in tasks if t.status == "done"]
    elif _TASK_FILTER == "error":
        tasks = [t for t in tasks if t.status == "error"]

    if not tasks:
        return ('<div style="padding:18px;color:#a9a49a;font-size:13px;'
                'text-align:center">当前筛选下没有任务</div>')

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
        size_hint = ""
        if t.current_size or t.estimated_size:
            size_hint = (
                f'<span style="font-size:11px;color:#8b8578;'
                f'font-family:\'SF Mono\',Consolas,monospace">'
                f'{hw_core.fmt_size(t.current_size) if t.current_size else "0 B"}'
                f' / {hw_core.fmt_size(t.estimated_size) if t.estimated_size else "—"}'
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

    footer = ""
    if _TASK_FILTER != "all":
        footer = (
            f'<div style="padding:8px 14px;text-align:center;'
            f'color:#8b8578;font-size:11.5px;background:#fdfcf8;'
            f'border-top:1px solid #f2ede0">'
            f'筛选视图：{_TASK_FILTER}（共 {total_before} 个任务）</div>'
        )

    return (f'<div style="background:#fff;border:1px solid #ebe5d8;'
            f'border-radius:14px;overflow:hidden">{rows_html}{footer}</div>')


def on_refresh_fast(stop_dd_value=None):
    with _REFRESH_LOCK:
        return _on_refresh_fast_locked(stop_dd_value)


def _on_refresh_fast_locked(stop_dd_value=None):
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
        log_text = current.log_tail(40)
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
            sig = ("html",
                   hashlib.md5(preview_html.encode("utf-8")).hexdigest())
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
    # ★ 优化：删掉之前永不命中的 _MODAL_SHOWN 遍历清理（靠 _MODAL_SHOWN_MAX 兜底）

    modal_html = gr.update()
    for t in tasks:
        if t.status == "done" and t.task_id not in _MODAL_SHOWN:
            _MODAL_SHOWN[t.task_id] = True
            while len(_MODAL_SHOWN) > _MODAL_SHOWN_MAX:
                _MODAL_SHOWN.popitem(last=False)
            try:
                refresh_task_size(t, force=True)
            except Exception:
                pass
            modal_html = build_done_modal_html(t)
            break

    return (task_list_html, progress_html, log_text,
            dd_update, gallery_value, modal_html)


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
# 优雅退出
# ============================================================
_sigint_count = [0]


def _shutdown():
    try:
        for t in MANAGER.running():
            try:
                t.stop_event.set()

                path = os.path.join(t.work_dir, "state.json")
                os.makedirs(t.work_dir, exist_ok=True)

                data = {
                    "task_id": t.task_id,
                    "kind": t.kind,
                    "src_name": t.src_name,
                    "out_dir": t.out_dir,
                    "created_at": t.created_at,
                    "status": "paused",
                    "current": t.current,
                    "total": t.total,
                    "label": t.label,
                    "log": [x for x in t.log_tail(60).split("\n") if x],
                    "output_files": list(t.output_files or []),
                    "error": t.error,
                    "current_size": t.current_size,
                    "estimated_size": t.estimated_size,
                }
                tmp = path + f".exit.{uuid.uuid4().hex[:8]}.tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                os.replace(tmp, path)
            except Exception as e:
                print(f"[shutdown] 保存任务 {getattr(t, 'task_id', '?')} "
                      f"失败：{e}", flush=True)
    except Exception as e:
        print(f"[shutdown] 整体失败：{e}", flush=True)
    finally:
        # ★ 修复：退出前把节流的缓存 JSON 落盘
        hw_core.flush_all_json()


def _sigint(signum, frame):
    _sigint_count[0] += 1
    if _sigint_count[0] >= 2:
        print("\n⏹ 强制退出", flush=True)
        os._exit(0)
    print("\n⏸ 收到 Ctrl+C，正在保存任务状态……", flush=True)
    _shutdown()
    print("✅ 状态已保存，退出", flush=True)
    os._exit(0)


# ============================================================
# UI（与翻译器 ui.py 布局一致）
# ============================================================
UI_THEME = gr.themes.Base(
    primary_hue=gr.themes.colors.gray,
    neutral_hue=gr.themes.colors.gray,
    font=[gr.themes.GoogleFont("Noto Sans SC"), "system-ui", "sans-serif"],
)

# ★ 修复：Gradio 6 中 css / theme 必须传给 launch()，
# 放在 Blocks() 里会被忽略（旧版写法导致样式一直不生效）
UI_CSS = """
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

    /* 分组标题已并入各长条面板的标题栏 */

    /* 布局为全宽长条堆叠，无需左右两栏 */

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

    #trial_cb, #parse_only_cb {
        background: transparent !important; border: none !important;
        padding: 4px 2px !important;
    }
    #trial_cb *, #parse_only_cb * { cursor: pointer !important; }
    #trial_cb input[type="checkbox"], #parse_only_cb input[type="checkbox"] {
        -webkit-appearance: checkbox !important; appearance: checkbox !important;
        width: 16px !important; height: 16px !important;
        min-width: 16px !important; max-width: 16px !important;
        accent-color: #0f3d3e !important; margin-right: 9px !important;
    }
    #trial_cb label, #parse_only_cb label {
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
        grid-template-columns: 1fr 1fr 1fr !important;
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
        font-size: 13.5px !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
        letter-spacing: .3px;
        transition: transform .08s, box-shadow .15s, background .15s;
    }
    #action_grid button:hover { transform: translateY(-1px); }
    #action_grid .secondary {
        background: #f5f1e6 !important; color: #0f3d3e !important;
        border: 1px solid #e2dccb !important;
    }
    #action_grid .secondary:hover {
        background: #ebe5d5 !important;
        border-color: #c9a961 !important;
    }
    #start_btn {
        height: 54px !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        letter-spacing: 1.5px;
        background: linear-gradient(135deg, #0f3d3e, #1f5b5c) !important;
        color: #faf8f2 !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(15,61,62,.3);
    }
    #start_btn:hover {
        background: linear-gradient(135deg, #0a2e2f, #0f3d3e) !important;
        box-shadow: 0 6px 18px rgba(15,61,62,.4);
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

    #task_head_row {
        gap: 10px !important; align-items: center !important;
        justify-content: space-between !important;
    }
    #task_list {
        flex: 1 1 auto !important;
        min-height: 260px !important;
        max-height: 520px !important;
        overflow-y: auto !important;
        margin: 2px 0 0 0 !important;
    }
    #task_filter .wrap {
        display: flex !important; flex-wrap: wrap !important;
        gap: 6px !important; justify-content: flex-end !important;
    }
    #task_filter label {
        font-size: 12px !important; font-weight: 500 !important;
        color: #5a5a5a !important; background: #f5f1e6 !important;
        border: 1px solid #e2dccb !important;
        border-radius: 999px !important; padding: 4px 12px !important;
        cursor: pointer !important; margin: 0 !important;
        transition: background .15s, color .15s, border-color .15s;
    }
    #task_filter label.selected {
        background: #0f3d3e !important; color: #fff !important;
        border-color: #0f3d3e !important;
    }
    #task_filter input[type="radio"] {
        width: 14px !important; height: 14px !important;
        accent-color: #0f3d3e !important;
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

    #preview_box { min-height: 300px !important; }
    #preview_box::-webkit-scrollbar { width: 10px; }
    #preview_box::-webkit-scrollbar-track {
        background: #e8e1cc; border-radius: 5px;
    }
    #preview_box::-webkit-scrollbar-thumb {
        background: #c9a961; border-radius: 5px;
    }

    /* 全宽长条：每个功能分组 = 一条可折叠面板 */
    .sec-fold {
        border: 1px solid #ebe5d8 !important;
        border-radius: 14px !important;
        background: #ffffff !important;
        box-shadow: 0 3px 14px rgba(15,61,62,.05);
        margin-top: 12px !important;
        overflow: hidden !important;
    }
    .sec-fold > button.label-wrap {
        width: 100% !important;
        background: #f5f1e6 !important;
        border: none !important;
        padding: 14px 20px !important;
        font-size: 14.5px !important;
        font-weight: 600 !important;
        color: #0f3d3e !important;
        cursor: pointer !important;
        display: flex !important; align-items: center !important;
        gap: 8px !important;
        font-family: "Noto Serif SC", Georgia, serif !important;
        letter-spacing: .4px !important;
        transition: background .15s !important;
    }
    .sec-fold > button.label-wrap:hover { background: #ebe5d5 !important; }
    .sec-fold > button.label-wrap .icon {
        margin-left: auto !important; color: #c9a961 !important;
        font-size: 12px !important;
    }
    .sec-fold > [data-testid="accordion-content"] {
        padding: 16px 20px 20px !important;
    }
    .gradio-container h3 {
        color: #0f3d3e !important; font-size: 18px !important;
        margin-top: 20px !important;
    }
    .gradio-container .prose, .gradio-container .prose * {
        color: #1a1a1a !important;
    }

    #load_btn {
        width: 100% !important; max-width: 420px !important;
        height: 44px !important;
        font-size: 13.5px !important; font-weight: 600 !important;
        background: #f5f1e6 !important; color: #0f3d3e !important;
        border: 1px solid #e2dccb !important; border-radius: 10px !important;
        transition: background .15s, border-color .15s;
    }
    #load_btn:hover {
        background: #ebe5d5 !important; border-color: #c9a961 !important;
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
    """


def build_ui():
    # ★ 优雅退出注册移到这里：import hw_ui 无副作用
    atexit.register(_shutdown)
    try:
        signal.signal(signal.SIGINT, _sigint)
    except Exception:
        pass

    with gr.Blocks(
        title="PDF / Word / PPT 作业解题器",
    ) as demo:

        modal_html = gr.HTML(value="", elem_id="modal_host")

        gr.HTML("""
        <div class="card-head">
          <h1>📝 PDF / Word / PPT 作业解题器</h1>
          <div class="sub">
            答案位置可选<span class="dot">·</span>
            学科 Prompt 微调<span class="dot">·</span>
            仅解析模式<span class="dot">·</span>
            数学公式渲染<span class="dot">·</span>
            断点可续<span class="dot">·</span>
            试解页数可选
          </div>
        </div>
        """)

        with gr.Accordion("💡 使用说明（点击展开）", open=False):
            gr.HTML("""
            <div style="padding:14px 20px;font-size:13.5px;line-height:2;color:#444;
                        background:#fdfcf8;border:1px solid #ebe5d8;border-radius:12px">
              <div><b>🎯 答案位置</b>　三档：块内右下（默认）/ 块下方 / 块右侧留白（不够时退回下方）</div>
              <div><b>📚 学科</b>　通用 / 数学 / 英语 / 物理 / 化学 / 语文，会微调系统提示词</div>
              <div><b>📝 自定义要求</b>　会追加到系统提示词，比如「解答要简洁」</div>
              <div><b>📄 仅解析模式</b>　不改原 PDF，输出 <code>analysis.md</code> 题目→答案清单</div>
              <div><b>🔑 密钥</b>　留空取 <code>.env</code> 中的默认值；填写 → 临时覆盖</div>
              <div><b>📄 上传</b>　拖入 .pdf / .docx / .pptx 自动识别类型；<b>PDF 只处理含文字层的，扫描件无法识别</b></div>
              <div><b>🧪 试解</b>　勾选后 PDF 只处理前 N 页（3/5/10/20 可选）；满意后取消勾选再全量跑</div>
              <div><b>📁 输出</b>　结果在 <code>homework_result/&lt;文件名_hash&gt;/</code></div>
              <div><b>🔤 字体</b>　可设 <code>HOMEWORK_FONT_PATH</code> 或 <code>TRANSLATE_FONT_PATH</code> 环境变量</div>
              <div><b>🎉 完成提示</b>　所有页/段完成后会弹出居中卡片，列出结果文件与保存位置</div>
            </div>
            """)

        # ══════════════ 上：解题设置（每类一条，全宽长条，可折叠） ══════════════
        gr.HTML('<div class="section-title">⚙️ 解题设置</div>')

        with gr.Accordion("① 🔑 API 设置", open=True, elem_classes="sec-fold"):
            with gr.Row(equal_height=True):
                api_key = gr.Textbox(
                    label="🔑 DeepSeek API Key",
                    type="password",
                    placeholder="留空用 .env",
                    scale=3,
                )
                model = gr.Dropdown(
                    choices=["deepseek-chat", "deepseek-reasoner"],
                    value=hw_core.DEFAULT_MODEL,
                    label="🧠 模型",
                    scale=2,
                    elem_id="model_dd",
                )

        with gr.Accordion("② 📂 上传文档", open=True, elem_classes="sec-fold"):
            with gr.Row(equal_height=True):
                file_mode = gr.Radio(
                    choices=["📕 PDF 书籍", "📘 Word 文档", "📊 PPT 演示"],
                    value="📕 PDF 书籍",
                    label="📂 文档类型（仅参考，实际按后缀自动判断）",
                    elem_id="file_mode",
                    scale=1,
                )
                doc_file = gr.File(
                    label="📄 上传作业文档（.pdf / .docx / .pptx）",
                    file_types=[".pdf", ".docx", ".pptx"],
                    elem_id="pdf_upload",
                    scale=1,
                )

        with gr.Accordion("③ 🌐 解题设置", open=True, elem_classes="sec-fold"):
            with gr.Row(equal_height=True):
                subject_dd = gr.Dropdown(
                    choices=[(n, k) for n, k in hw_core.SUBJECT_CHOICES],
                    value="general",
                    label="📚 学科（微调 Prompt）",
                    scale=1,
                )
                ans_pos_radio = gr.Radio(
                    choices=[(n, k) for n, k in hw_core.ANS_POS_CHOICES],
                    value="inside",
                    label="🎯 答案位置（仅 PDF）",
                    scale=1,
                )

        with gr.Accordion("④ 📝 解题选项", open=True, elem_classes="sec-fold"):
            with gr.Row(equal_height=True):
                trial = gr.Checkbox(
                    label="🧪 试解模式：只处理 PDF 前 N 页",
                    value=False,
                    elem_id="trial_cb",
                    scale=1,
                )
                trial_pages_dd = gr.Dropdown(
                    choices=[(n, v) for n, v in hw_core.TRIAL_PAGE_CHOICES],
                    value=5,
                    label="试解页数",
                    scale=1,
                )
                parse_only_cb = gr.Checkbox(
                    label="📄 仅解析模式：不改 PDF，输出 analysis.md",
                    value=False,
                    elem_id="parse_only_cb",
                    scale=1,
                )
            custom_prompt_tb = gr.Textbox(
                label="📝 自定义要求（追加到系统提示词，可留空）",
                lines=2,
                placeholder="例如：解答简洁，每题不超过 3 句话；公式要用标准 LaTeX",
            )

        with gr.Accordion("⑤ 🚀 开始任务", open=True, elem_classes="sec-fold"):
            with gr.Column(elem_id="action_grid"):
                btn = gr.Button("▶ 开始解题", variant="primary",
                                elem_id="start_btn")
                with gr.Row(equal_height=True):
                    stop_btn = gr.Button("⏹ 全部停止", variant="secondary")
                    refresh_btn = gr.Button("🔄 刷新", variant="secondary")
                    open_btn = gr.Button("📁 打开目录", variant="secondary")
            open_hint = gr.Markdown(value="", elem_id="open_hint")

        # ══════════════ 中：任务与进度（全宽长条，可折叠） ══════════════
        gr.HTML('<div class="section-title" style="margin-top:28px">'
                '📋 任务与进度</div>')

        with gr.Accordion("📋 任务列表", open=True, elem_classes="sec-fold"):
            with gr.Row(elem_id="task_head_row"):
                gr.HTML('<div style="font-size:12px;color:#8b8578;'
                        'align-self:center">状态筛选：</div>')
                task_filter_radio = gr.Radio(
                    choices=[(n, k) for n, k in hw_core.TASK_FILTER_CHOICES],
                    value="all",
                    label="",
                    show_label=False,
                    container=False,
                    elem_id="task_filter",
                )
            task_list_html = gr.HTML(
                value='<div style="padding:18px;color:#a9a49a;font-size:13px;'
                      'text-align:center">暂无任务</div>',
                elem_id="task_list",
            )

        with gr.Accordion("🛑 停止任务", open=True, elem_classes="sec-fold"):
            with gr.Row(elem_id="stop_one_row"):
                stop_dd = gr.Dropdown(
                    label="选择要停止的任务（只列运行中）",
                    choices=[],
                    value=None,
                    interactive=True,
                    elem_id="stop_dd",
                )
                stop_one_btn = gr.Button("⏹ 停止选中", elem_id="stop_one_btn")

        with gr.Accordion("📊 当前进度", open=True, elem_classes="sec-fold"):
            progress_bar = gr.HTML(
                value=make_progress_html(0, 1, "等待开始"),
                elem_id="progress_host",
            )

        # ══════════════ 中部：任务日志（可折叠） ══════════════
        with gr.Accordion("📋 任务日志", open=True, elem_classes="sec-fold"):
            log = gr.Textbox(
                label="",
                lines=14,
                interactive=False,
                show_label=False,
                elem_id="task_log",
            )

        # ══════════════ 底部：效果预览（可折叠） ══════════════
        with gr.Accordion("👀 效果预览", open=False, elem_classes="sec-fold"):
            gallery = gr.HTML(
                value=build_preview_html([]),
                elem_id="preview_box",
            )
            load_btn = gr.Button(
                "🔍 加载当前任务的预览图和下载文件",
                variant="secondary",
                elem_id="load_btn",
            )

        # ══════════════ 底部：下载文件（可折叠） ══════════════
        with gr.Accordion("💾 下载文件", open=False, elem_classes="sec-fold"):
            out_files = gr.File(
                label="", file_count="multiple",
                interactive=False, show_label=False,
            )

        fast_outputs = [task_list_html, progress_bar, log, stop_dd, gallery, modal_html]
        full_outputs = [task_list_html, progress_bar, log, stop_dd,
                        gallery, modal_html, out_files, doc_file]

        btn.click(
            on_start,
            [api_key, model, doc_file, file_mode, trial, trial_pages_dd,
             subject_dd, custom_prompt_tb, ans_pos_radio, parse_only_cb,
             stop_dd],
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

        # ★ 新增：任务列表筛选
        task_filter_radio.change(
            on_filter_change, [task_filter_radio], fast_outputs,
            concurrency_limit=None, concurrency_id="filter",
        )

        try:
            timer = gr.Timer(3.0)
            timer.tick(on_refresh_fast, [stop_dd], fast_outputs,
                       concurrency_limit=None, concurrency_id="tick")
        except Exception:
            pass

        demo.load(on_refresh_fast, [stop_dd], fast_outputs,
                  concurrency_limit=None, concurrency_id="load")

    return demo

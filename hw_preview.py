# -*- coding: utf-8 -*-
"""作业解题器 · 预览：对比图渲染 + 预览/进度/完成弹窗 HTML 构建器。"""

import os
import io
import glob
import base64
import shutil
from collections import OrderedDict
from html import escape

import pymupdf as fitz
from PIL import Image

import hw_core

_RESAMPLE = getattr(Image, "Resampling", Image).LANCZOS
_DATAURL_CACHE = OrderedDict()
# ★ 修复：预览 dataURL 缓存上限 200 → 40（base64 常驻内存太大）
_DATAURL_CACHE_MAX = 40


# ============================================================
# 预览
# ============================================================
def _pix_to_pil(pix):
    """★ 修复：Pixmap samples 直读（移植自双语版，比 PNG 往返快数倍）。"""
    try:
        n = pix.n
        w, h_ = pix.width, pix.height
        if n == 3:
            return Image.frombytes("RGB", (w, h_), pix.samples)
        if n == 4:
            return Image.frombytes("RGBA", (w, h_), pix.samples).convert("RGB")
    except Exception:
        pass
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def render_preview_only(solved_path, paths, task, n_preview=hw_core.PREVIEW_PAGES):
    preview_dir = paths["preview_dir"]
    tmp_dir = preview_dir + ".new"

    if os.path.isdir(tmp_dir):
        for pat in ("*.png", "*.jpg", "*.jpeg"):
            for f in glob.glob(os.path.join(tmp_dir, pat)):
                try:
                    os.remove(f)
                except Exception:
                    pass
    os.makedirs(tmp_dir, exist_ok=True)

    if not os.path.exists(solved_path):
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        return []

    try:
        with open(solved_path, "rb") as f:
            data = f.read()
        solved = fitz.open(stream=data, filetype="pdf")
    except Exception:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        return []

    try:
        orig = fitz.open(task.src_path)
    except Exception:
        solved.close()
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        return []

    n = min(n_preview, len(orig), len(solved))
    preview_imgs = []
    for i in range(n):
        if task.stop_event.is_set():
            break
        try:
            o_pix = orig[i].get_pixmap(
                matrix=fitz.Matrix(hw_core.RENDER_ZOOM, hw_core.RENDER_ZOOM))
            s_pix = solved[i].get_pixmap(
                matrix=fitz.Matrix(hw_core.RENDER_ZOOM, hw_core.RENDER_ZOOM))
            o = _pix_to_pil(o_pix)
            s = _pix_to_pil(s_pix)
            hh = max(o.height, s.height)
            if o.height != hh:
                o = o.resize((int(o.width * hh / o.height), hh), _RESAMPLE)
            if s.height != hh:
                s = s.resize((int(s.width * hh / s.height), hh), _RESAMPLE)
            gap = 8
            canvas = Image.new("RGB",
                               (o.width + gap + s.width, hh), (40, 40, 40))
            canvas.paste(o, (0, 0))
            canvas.paste(s, (o.width + gap, 0))
            if canvas.width > hw_core.PREVIEW_MAX_WIDTH:
                ratio = hw_core.PREVIEW_MAX_WIDTH / canvas.width
                canvas = canvas.resize(
                    (hw_core.PREVIEW_MAX_WIDTH, int(canvas.height * ratio)),
                    _RESAMPLE,
                )
            p = os.path.join(tmp_dir, f"compare_{i:04d}.jpg")
            canvas.save(p, "JPEG", quality=hw_core.PREVIEW_JPEG_QUALITY, optimize=True)
            preview_imgs.append(p)
        except Exception:
            continue
    orig.close()
    solved.close()

    if not preview_imgs:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        return []

    old_dir = preview_dir + ".old"
    try:
        if os.path.isdir(old_dir):
            shutil.rmtree(old_dir, ignore_errors=True)
    except Exception:
        pass
    if os.path.isdir(preview_dir):
        try:
            os.rename(preview_dir, old_dir)
        except Exception:
            try:
                shutil.rmtree(preview_dir, ignore_errors=True)
            except Exception:
                pass
    try:
        os.rename(tmp_dir, preview_dir)
    except Exception:
        if not os.path.isdir(preview_dir) and os.path.isdir(old_dir):
            try:
                os.rename(old_dir, preview_dir)
            except Exception:
                pass
        return []
    try:
        if os.path.isdir(old_dir):
            shutil.rmtree(old_dir, ignore_errors=True)
    except Exception:
        pass

    new_imgs = []
    for i in range(len(preview_imgs)):
        new_p = os.path.join(preview_dir, f"compare_{i:04d}.jpg")
        if os.path.exists(new_p):
            new_imgs.append(new_p)
    return new_imgs


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
        st = os.stat(path)
        key = (path, st.st_mtime_ns, st.st_size)
    except Exception:
        return None

    cached = _DATAURL_CACHE.get(key)
    if cached is not None:
        _DATAURL_CACHE.move_to_end(key)
        return cached

    try:
        with open(path, "rb") as f:
            data = f.read()
        if not data:
            return None
        b64 = base64.b64encode(data).decode("ascii")
        low = path.lower()
        mime = "image/png" if low.endswith(".png") else "image/jpeg"
        url = f"data:{mime};base64,{b64}"
        _DATAURL_CACHE[key] = url
        while len(_DATAURL_CACHE) > _DATAURL_CACHE_MAX:
            _DATAURL_CACHE.popitem(last=False)
        return url
    except Exception:
        return None


def preview_signature(preview_imgs):
    parts = []
    for p in preview_imgs or []:
        try:
            st = os.stat(p)
            parts.append(
                f"{os.path.basename(p)}:{st.st_mtime_ns}:{st.st_size}"
            )
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
        cur_s = hw_core.fmt_size(current_size) if current_size else "0 B"
        est_s = hw_core.fmt_size(estimated_size) if estimated_size else "—"
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
            elif low.endswith(".md"):
                icon = "📝"
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
            f'📦 {hw_core.fmt_size(actual_size)}</span>'
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
        @keyframes doneFadeIn {{ from {{opacity:0}} to {{opacity:1}} }}
        @keyframes donePopIn {{
          from {{opacity:0;transform:scale(.92) translateY(10px)}}
          to {{opacity:1;transform:scale(1) translateY(0)}}
        }}
        #done_modal_card {{ animation:donePopIn .35s cubic-bezier(.2,1.2,.4,1); }}
        #done_modal_close_btn:hover {{
          transform:translateY(-1px);
          box-shadow:0 6px 20px rgba(15,61,62,.4) !important;
        }}
        #done_modal_close_btn:active {{ transform:translateY(0); }}
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

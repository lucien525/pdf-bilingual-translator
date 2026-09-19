# -*- coding: utf-8 -*-
"""预览 HTML 生成：图片 dataURL、左右对照预览、进度条、完成弹窗。"""

import os
import base64
from collections import OrderedDict
from html import escape

from core import config
from core.utils import fmt_size

_DATAURL_CACHE = OrderedDict()
# ★ 修复：预览 dataURL 缓存上限 200 → 40（base64 常驻内存太大）
_DATAURL_CACHE_MAX = 40


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
            st = os.stat(p)
            parts.append(f"{os.path.basename(p)}:{st.st_mtime_ns}:{st.st_size}")
        except Exception:
            parts.append(os.path.basename(p) if p else "")
    return tuple(parts)


def build_preview_html(preview_imgs, preview_html=None):
    if preview_html:
        return preview_html

    if not preview_imgs:
        return '''
        <div style="padding:48px 24px;color:#8b8578;text-align:center;font-size:13.5px;
                    background:#fdfcf9;border:1.5px dashed #ebe5d8;border-radius:16px;
                    line-height:1.9;min-height:300px;
                    display:flex;flex-direction:column;
                    align-items:center;justify-content:center">
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

    lang_label = config.LANG_NAMES.get(getattr(task, "target_lang", "zh-CN"), "简体中文")
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

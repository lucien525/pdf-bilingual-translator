# -*- coding: utf-8 -*-
"""
PDF 双语对照翻译器（网页版）
- 输出：result/<书名>/translated.pdf（中文版）+ bilingual.pdf（左右对照）
- 中间文件在 result/<书名>/_work/
- 断点续传：翻到哪页记在哪页，重新运行不重翻、不重复扣费
- 半成品也输出，日志明确提示
- 顶部实时进度条（插图页也算已处理）
"""

import os
import sys
import json
import time
import re
import hashlib

import pymupdf as fitz
from PIL import Image
from openai import OpenAI
import gradio as gr
from dotenv import load_dotenv

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

WORK_DIR = None
CACHE_FILE = None
OUTPUT_PDF = None
BILINGUAL_PDF = None
OUTPUT_HTML = None
IMG_DIR = None
PREVIEW_DIR = None
PROGRESS_FILE = None

SYSTEM_PROMPT = (
    "你是一位资深文学翻译家，精通中英双语，译笔力求神似而非字对字。"
    "请遵循：1) 译文必须符合中文母语者的阅读习惯和审美，流畅、有文采；"
    "2) 对话要自然生动，符合人物身份；3) 修辞、隐喻、双关尽量找到中文对应表达，"
    "实在无法对应则意译并保留神韵；4) 不遗漏任何内容，不总结，不输出任何解释。\n\n"
    "【格式要求】用户会给出一页英文的多个段落，每段以 [[B0]] [[B1]] [[B2]] ... 标记开头。"
    "你必须严格保留所有标记、保持顺序，标记后紧跟该段译文。"
    "除标记和译文外，不要输出任何其他文字、不加解释、不用代码块。"
)

STOP = False
# ==========================================


def safe_dirname(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip().rstrip(". ")
    if not name:
        name = "untitled"
    return name[:80]


def setup_output_dir(pdf_path):
    global WORK_DIR, CACHE_FILE, OUTPUT_PDF, OUTPUT_HTML
    global IMG_DIR, PREVIEW_DIR, BILINGUAL_PDF, PROGRESS_FILE

    base = os.path.basename(pdf_path)
    book = safe_dirname(os.path.splitext(base)[0])
    out_dir = os.path.abspath(os.path.join(RESULT_ROOT, book))
    os.makedirs(out_dir, exist_ok=True)

    work = os.path.join(out_dir, "_work")
    os.makedirs(work, exist_ok=True)
    WORK_DIR = work

    OUTPUT_PDF    = os.path.join(out_dir, "translated.pdf")
    BILINGUAL_PDF = os.path.join(out_dir, "bilingual.pdf")

    CACHE_FILE    = os.path.join(work, "translate_cache.json")
    OUTPUT_HTML   = os.path.join(work, "bilingual.html")
    IMG_DIR       = os.path.join(work, "bilingual_pages")
    PREVIEW_DIR   = os.path.join(work, "preview")
    PROGRESS_FILE = os.path.join(work, "progress.json")

    return out_dir


def h(s):
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:20]


def load_cache():
    if CACHE_FILE and os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_cache(c):
    if not CACHE_FILE:
        return
    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(c, f, ensure_ascii=False)


def load_progress():
    """已完成处理的页码集合（1-based），包括翻译成功的和插图跳过的"""
    if PROGRESS_FILE and os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return set(int(x) for x in data.get("done_pages", []))
        except Exception:
            return set()
    return set()


def save_progress(done_pages):
    if not PROGRESS_FILE:
        return
    data = {"done_pages": sorted(int(x) for x in done_pages)}
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


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


def call_api(client, model, user_content, retries=4):
    last_err = ""
    for attempt in range(retries):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
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


def translate_page(client, model, blocks, cache, pno):
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
        save_cache(cache)

    cache[key] = {str(k): v for k, v in parsed.items()}
    save_cache(cache)
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


def render_page_pngs(src_path, trans_path, num_pages):
    os.makedirs(IMG_DIR, exist_ok=True)
    orig = fitz.open(src_path)
    trans = fitz.open(trans_path)
    n = min(num_pages, len(orig), len(trans))
    orig_paths, trans_paths = [], []
    for i in range(n):
        o_path = os.path.join(IMG_DIR, f"orig_{i:04d}.png")
        t_path = os.path.join(IMG_DIR, f"trans_{i:04d}.png")
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


def generate_html(orig_paths, trans_paths, out_html):
    pairs = []
    for i, (o, t) in enumerate(zip(orig_paths, trans_paths)):
        o_rel = os.path.relpath(o, os.path.dirname(out_html)).replace("\\", "/")
        t_rel = os.path.relpath(t, os.path.dirname(out_html)).replace("\\", "/")
        pairs.append(
            f'<div class="pair">'
            f'<div class="labels"><div>原文 · Page {i+1}</div><div>译文 · Page {i+1}</div></div>'
            f'<div class="pages"><img src="{o_rel}"><img src="{t_rel}"></div>'
            f'</div>'
        )

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>双语对照阅读</title>
<style>
html,body{{margin:0;padding:0;background:#333;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}}
.bar{{position:sticky;top:0;z-index:100;background:#1c1c1c;color:#e6e6e6;padding:10px 18px;display:flex;gap:16px;align-items:center;box-shadow:0 2px 12px rgba(0,0,0,.5);font-size:13px}}
.bar b{{font-size:15px;font-weight:600}}
.bar .sp{{flex:1}}
.bar .hint{{color:#888}}
.bar input[type=range]{{width:150px;vertical-align:middle}}
.wrap{{max-width:1500px;margin:0 auto;padding:14px 8px 60px;transition:max-width .15s}}
.pair{{background:#fff;margin-bottom:14px;box-shadow:0 3px 16px rgba(0,0,0,.4);overflow:hidden}}
.labels{{display:flex;background:#1c1c1c;color:#999;font-size:11.5px;letter-spacing:.5px}}
.labels>div{{flex:1 1 50%;text-align:center;padding:6px 0;border-right:1px solid #333}}
.labels>div:last-child{{border-right:none}}
.pages{{display:flex;align-items:flex-start}}
.pages img{{width:50%;height:auto;display:block}}
.pages img:first-child{{border-right:1px solid #ccc}}
</style>
</head>
<body>
<div class="bar">
  <b>📖 双语对照阅读</b>
  <span class="hint">左：原文 ｜ 右：译文（逐页对齐）</span>
  <span class="sp"></span>
  <span class="hint">缩放</span>
  <input type="range" id="zoom" min="40" max="100" value="100" step="5">
  <span id="zval" class="hint" style="width:42px;display:inline-block">100%</span>
</div>
<div class="wrap" id="wrap">
{''.join(pairs)}
</div>
<script>
(function(){{
  var wrap=document.getElementById('wrap');
  var z=document.getElementById('zoom');
  var zv=document.getElementById('zval');
  z.addEventListener('input',function(){{
    var v=+z.value;
    zv.textContent=v+'%';
    wrap.style.maxWidth=(1500*v/100)+'px';
  }});
}})();
</script>
</body>
</html>"""
    with open(out_html, "w", encoding="utf-8") as f:
        f.write(html)


def stop_translate():
    global STOP
    STOP = True
    return "正在停止……当前页翻完就停，进度已保存，未翻完也会输出半成品。"


def open_result_folder():
    if not OUTPUT_PDF:
        return "⚠️ 还没有翻译结果。请先上传 PDF 并开始翻译。"
    folder = os.path.dirname(os.path.abspath(OUTPUT_PDF))
    if not os.path.exists(folder):
        return f"⚠️ 结果文件夹不存在：{folder}"
    try:
        if sys.platform.startswith("win"):
            os.startfile(folder)
        elif sys.platform == "darwin":
            import subprocess; subprocess.Popen(["open", folder])
        else:
            import subprocess; subprocess.Popen(["xdg-open", folder])
        return f"✅ 已打开文件夹：{folder}"
    except Exception as e:
        return f"❌ 打开失败：{e}\n路径：{folder}"


def run(api_key, model, pdf, trial):
    global STOP
    STOP = False
    log = []

    def say(m):
        log.append(m)
        return "\n".join(log[-40:])

    # ===== 前置检查 =====
    real_key = (api_key or "").strip() or DEFAULT_API_KEY
    if not real_key:
        yield say("❌ 未提供 API Key：在网页里填入，或在 .env 里设置 DEEPSEEK_API_KEY"), [], [], make_progress_html(0, 1, "未开始")
        return
    if not pdf:
        yield say("❌ 请先上传 PDF 文件"), [], [], make_progress_html(0, 1, "未开始")
        return
    if not os.path.exists(FONT_PATH):
        yield say(f"❌ 字体文件不存在：{FONT_PATH}"), [], [], make_progress_html(0, 1, "未开始")
        return

    if (api_key or "").strip():
        yield say("🔑 使用网页填写的 API Key"), [], [], make_progress_html(0, 1, "准备中")
    else:
        yield say("🔑 使用 .env 中的默认 API Key"), [], [], make_progress_html(0, 1, "准备中")

    model = (model or DEFAULT_MODEL).strip()
    if model not in ("deepseek-chat", "deepseek-reasoner"):
        model = "deepseek-chat"
    yield say(f"🧠 模型：{model}"), [], [], make_progress_html(0, 1, "准备中")

    src_path = pdf.name
    out_dir = setup_output_dir(src_path)
    yield say(f"📁 结果目录：{out_dir}"), [], [], make_progress_html(0, 1, "准备中")

    done_pages = load_progress()
    max_done = max(done_pages) if done_pages else 0

    doc = fitz.open(src_path)
    total = len(doc)
    limit = min(5, total) if trial else total

    if max_done > 0:
        done_in_scope_cnt = len([p for p in done_pages if p <= limit])
        yield (
            say(f"📚 检测到进度：已处理 {len(done_pages)} 页（最远到第 {max_done} 页）。"
                f"这些页会直接跳过，不重复请求 API、不重复扣费。"),
            [], [], make_progress_html(done_in_scope_cnt, limit, f"已有进度，本次目标前 {limit} 页")
        )
    else:
        yield (
            say("📚 全新开始，没有历史进度"),
            [], [], make_progress_html(0, limit, f"本次目标 {limit} 页")
        )

    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    cache = load_cache()

    def done_in_scope():
        return len([p for p in done_pages if p <= limit])

    yield (
        say(f"✅ PDF 共 {total} 页；本次处理上限第 {limit} 页"),
        [], [], make_progress_html(done_in_scope(), limit, f"本次目标 {limit} 页")
    )

    newly_translated = []
    stopped_by_user = False
    error_msg = None
    completed = False

    # ===== 主循环 =====
    for pno in range(total):
        page_num = pno + 1

        if STOP:
            stopped_by_user = True
            break

        if page_num > limit and page_num not in done_pages:
            break

        page = doc[pno]
        blocks = [b for b in page.get_text("blocks") if b[6] == 0 and b[4].strip()]

        # 已处理过的页：翻译的从缓存应用；插图的直接跳过
        if page_num in done_pages:
            if blocks:
                try:
                    trans = translate_page(client, model, blocks, cache, pno)
                    apply_translations(page, blocks, trans)
                except Exception:
                    pass
            continue

        # 插图页：标记为已处理
        if not blocks:
            done_pages.add(page_num)
            save_progress(done_pages)
            yield (
                say(f"⏭ 第 {page_num} 页无文字（插图页），跳过"),
                [], [], make_progress_html(done_in_scope(), limit, f"已跳过第 {page_num} 页")
            )
            continue

        yield (
            say(f"📄 正在翻译第 {page_num}/{limit} 页（{len(blocks)} 段）……"),
            [], [], make_progress_html(done_in_scope(), limit, f"正在翻译第 {page_num} 页")
        )

        try:
            trans = translate_page(client, model, blocks, cache, pno)
        except RuntimeError as e:
            error_msg = str(e)
            break
        except Exception as e:
            error_msg = f"未知错误：{e}"
            break

        try:
            apply_translations(page, blocks, trans)
        except Exception as e:
            error_msg = f"写入失败（第 {page_num} 页）：{e}"
            break

        done_pages.add(page_num)
        save_progress(done_pages)
        newly_translated.append(page_num)

        yield (
            say(f"✅ 第 {page_num} 页完成（本次新增 {len(newly_translated)} 页）"),
            [], [], make_progress_html(done_in_scope(), limit, f"已完成第 {page_num} 页")
        )
    else:
        completed = True

    # ===== 保存 PDF =====
    try:
        doc.save(OUTPUT_PDF, deflate=True)
    finally:
        doc.close()

    if not done_pages:
        yield (
            say("⚠️ 本次没有任何页面成功翻译，无法生成对照结果。\n"
                "可能原因：扫描版 PDF（需要先 OCR）；全部是插图页；或 API 请求全部失败。"),
            [], [], make_progress_html(0, 1, "无结果")
        )
        return

    render_up_to = max(done_pages)

    yield (
        say(f"🖼 正在渲染页面（到第 {render_up_to} 页）……"),
        [], [], make_progress_html(done_in_scope(), limit, "渲染对照图")
    )
    orig_paths, trans_paths = render_page_pngs(src_path, OUTPUT_PDF, render_up_to)
    generate_html(orig_paths, trans_paths, OUTPUT_HTML)

    yield (
        say("📄 正在合成左右对照 PDF……"),
        [], [], make_progress_html(done_in_scope(), limit, "合成对照 PDF")
    )
    make_bilingual_pdf(orig_paths, trans_paths, BILINGUAL_PDF)

    os.makedirs(PREVIEW_DIR, exist_ok=True)
    preview_imgs = []
    for i in range(len(orig_paths)):
        try:
            p = os.path.join(PREVIEW_DIR, f"compare_{i:04d}.png")
            make_side_by_side(orig_paths[i], trans_paths[i], p)
            preview_imgs.append(p)
        except Exception as e:
            yield (
                say(f"⚠️ 第 {i+1} 页预览生成失败：{e}"),
                [], preview_imgs,
                make_progress_html(done_in_scope(), limit, "生成预览")
            )

    # ===== 状态总结 =====
    total_done = len(done_pages)
    translated_cnt = len(newly_translated)
    if completed:
        status = f"🎉 全部完成！累计处理 {total_done} 页（本次新增翻译 {translated_cnt} 页）"
    elif stopped_by_user:
        status = (f"🛑 已暂停（半成品）。累计处理 {total_done} 页（本次新增翻译 {translated_cnt} 页）。\n"
                  f"下次点「开始翻译」会从第 {render_up_to + 1} 页继续，已翻过的页不会重复扣费。")
    elif error_msg:
        status = (f"❌ 中途出错（半成品）。累计处理 {total_done} 页（本次新增翻译 {translated_cnt} 页）。\n"
                  f"错误：{error_msg}\n"
                  f"下次点「开始翻译」会从第 {render_up_to + 1} 页继续。")
    else:
        status = f"⚠️ 未全部完成（半成品）。累计处理 {total_done} 页"

    final_label = "全部完成" if completed else "已暂停（半成品）"
    download_files = [os.path.abspath(OUTPUT_PDF), os.path.abspath(BILINGUAL_PDF)]
    yield (
        say(status + f"\n结果目录：{out_dir}\n"
            f"（点上方「📁 打开翻译结果文件夹」查看 translated.pdf 和 bilingual.pdf）"),
        download_files,
        preview_imgs,
        make_progress_html(done_in_scope(), limit, final_label)
    )


# ================= UI =================
with gr.Blocks(
    title="PDF 双语对照翻译器",
    theme=gr.themes.Base(
        primary_hue=gr.themes.colors.gray,
        neutral_hue=gr.themes.colors.gray,
        font=[gr.themes.GoogleFont("Noto Sans SC"), "system-ui", "sans-serif"],
    ),
    css="""
    /* ===== 全局 ===== */
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

    /* ===== 顶部卡片 ===== */
    .card-head {
        background: #f5f1e6;
        border: 1px solid #e2dccb;
        border-radius: 14px;
        padding: 28px 32px 24px;
        margin-bottom: 22px;
        position: relative;
        overflow: hidden;
    }
    .card-head::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
        background: linear-gradient(90deg, #2b2b2b, #8a7a4f, #2b2b2b);
    }
    .card-head h1 {
        font-size: 27px; font-weight: 700; color: #111; margin: 0;
        letter-spacing: .3px;
        font-family: "Noto Serif SC", Georgia, serif;
    }
    .card-head .sub {
        color: #555; font-size: 14.5px; margin-top: 12px;
        line-height: 1.95;
    }
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

    /* ===== 表单 ===== */
    label span, .gr-box > label > span {
        color: #1a1a1a !important;
        font-size: 14.5px !important;
        font-weight: 500 !important;
    }
    .gradio-container input:not([type="checkbox"]):not([type="radio"]),
    .gradio-container textarea,
    .gradio-container select {
        background: #ffffff !important;
        color: #111 !important;
        border: 1px solid #d8d2c0 !important;
        font-size: 14.5px !important;
    }
    .gradio-container .block {
        background: #ffffff !important;
        border: 1px solid #e6dfce !important;
        border-radius: 10px !important;
    }

    /* ===== 试翻 checkbox ===== */
    #trial_cb { background: transparent !important; border: none !important; padding: 8px 12px !important; }
    #trial_cb * { cursor: pointer !important; }
    #trial_cb input[type="checkbox"] {
        -webkit-appearance: checkbox !important;
        appearance: checkbox !important;
        width: 18px !important; height: 18px !important;
        min-width: 18px !important; max-width: 18px !important;
        display: inline-block !important;
        visibility: visible !important; opacity: 1 !important;
        accent-color: #2b2b2b !important;
        margin-right: 10px !important;
        border: 1px solid #888 !important;
    }
    #trial_cb label { cursor: pointer !important; font-size: 14.5px !important; }

    /* ===== 模型下拉框 ===== */
    #model_dd { min-width: 260px !important; }
    #model_dd input, #model_dd .wrap-inner { min-width: 240px !important; }

    /* ============================================================
       上传区：不依赖 elem_id，用 Gradio File 组件自带的 class 定位
       Gradio 4.x 的 File 组件外层一定是 .file-preview 的祖先
       用 :has() 选择器（现代浏览器都支持）
       ============================================================ */
    .gradio-container .block:has(> .file-preview),
    .gradio-container .block:has(> div > .file-preview),
    .gradio-container .block:has(> div > div > .file-preview) {
        height: 96px !important;
        min-height: 96px !important;
        max-height: 96px !important;
        overflow: hidden !important;
        box-sizing: border-box !important;
    }

    /* 内部所有层：清掉默认 min-height（这就是被撑大的元凶） */
    .gradio-container .block:has(.file-preview) * {
        min-height: 0 !important;
        box-sizing: border-box !important;
    }

    /* 上传后显示的文件名区域：一行高度，不撑大 */
    .gradio-container .file-preview {
        height: auto !important;
        max-height: 72px !important;
        overflow-y: auto !important;
        overflow-x: hidden !important;
        padding: 8px 12px !important;
        background: transparent !important;
        border: none !important;
    }
    .gradio-container .file-preview * {
        max-height: 60px !important;
        line-height: 1.4 !important;
        font-size: 13.5px !important;
        color: #1a1a1a !important;
    }

    /* 上传区内部提示文字 */
    .gradio-container .file-preview .empty,
    .gradio-container .file-preview span {
        font-size: 13.5px !important;
        color: #555 !important;
    }

    /* 上传按钮 */
    .gradio-container .file-preview button,
    .gradio-container button.upload-button {
        padding: 3px 10px !important;
        font-size: 12.5px !important;
    }
    .gradio-container .file-preview svg {
        max-height: 20px !important;
        max-width: 20px !important;
    }

    /* ===== 按钮 ===== */
    .gradio-container button.primary,
    .gradio-container .primary {
        background: #2b2b2b !important;
        color: #faf8f2 !important;
        border: none !important;
        font-size: 15px !important;
        font-weight: 600 !important;
    }
    .gradio-container button.primary:hover { background: #000 !important; }
    .gradio-container button.secondary {
        background: #f0ebdc !important;
        color: #222 !important;
        border: 1px solid #ddd5c0 !important;
        font-size: 14.5px !important;
    }

    .gradio-container .accordion-header {
        background: #f5f1e6 !important;
        color: #222 !important;
        font-size: 14.5px !important;
    }

    .gradio-container h3 { color: #111 !important; font-size: 18px !important; margin-top: 20px !important; }
    .gradio-container .prose, .gradio-container .prose * { color: #1a1a1a !important; }
    """,
) as demo:

    gr.HTML("""
    <div class="card-head">
      <h1>📖 PDF 双语对照翻译器</h1>
      <div class="sub">
        一页原文<span class="dot">·</span>一页译文
        <span class="dot">·</span>左右同行<span class="dot">·</span>读来无声
        <br>
        断点可续<span class="dot">·</span>半成品也留
      </div>
      <div class="meta">
        <span class="k">🔑 密钥</span>留空取 <code>.env</code> 中的默认值，亦可临时填入覆盖
        <br>
        <span class="k">📁 成果</span>归于 <code>result/&lt;书名&gt;/</code>：<code>translated.pdf</code>（纯中文） 与 <code>bilingual.pdf</code>（左右对照）
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

    pdf = gr.File(
        label="📄 上传文档（拖拽或点击）",
        file_types=[".pdf"],
        elem_id="pdf_upload",
    )
    trial = gr.Checkbox(
        label="🧪 试翻模式：只翻前 5 页，先看效果和费用",
        value=True,
        elem_id="trial_cb",
    )

    with gr.Row():
        btn = gr.Button("▶ 开始翻译", variant="primary", scale=3)
        stop_btn = gr.Button("⏸ 停止", scale=1)

    open_btn = gr.Button("📁 打开翻译结果文件夹", variant="secondary")

    progress_bar = gr.HTML(value=make_progress_html(0, 1, "等待开始"))

    with gr.Accordion("📋 运行日志（点击展开 / 收起）", open=False):
        log = gr.Textbox(label="", lines=14, interactive=False, show_label=False)

    gr.Markdown("### 👀 效果预览（左原文，右译文）")
    gallery = gr.Gallery(
        label="对照预览",
        columns=1,
        height=800,
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

    btn.click(
        run,
        [api_key, model, pdf, trial],
        [log, out_files, gallery, progress_bar],
    )
    stop_btn.click(stop_translate, None, [log])
    open_btn.click(open_result_folder, None, [log])


if __name__ == "__main__":
    if DEFAULT_API_KEY:
        print(f"✅ 已从 .env 读取默认 Key（{DEFAULT_API_KEY[:6]}...）")
    else:
        print("ℹ️  .env 中未找到 DEEPSEEK_API_KEY，需在网页里手动填写")
    print(f"ℹ️  默认模型：{DEFAULT_MODEL}")
    print(f"ℹ️  结果目录：{os.path.abspath(RESULT_ROOT)}")

    demo.launch(inbrowser=True, allowed_paths=[os.path.abspath(".")])

if __name__ == "__main__":
    if DEFAULT_API_KEY:
        print(f"✅ 已从 .env 读取默认 Key（{DEFAULT_API_KEY[:6]}...）")
    else:
        print("ℹ️  .env 中未找到 DEEPSEEK_API_KEY，需在网页里手动填写")
    print(f"ℹ️  默认模型：{DEFAULT_MODEL}")
    print(f"ℹ️  结果目录：{os.path.abspath(RESULT_ROOT)}")

    demo.launch(inbrowser=True, allowed_paths=[os.path.abspath(".")])
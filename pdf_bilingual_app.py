# -*- coding: utf-8 -*-
"""
PDF 双语对照翻译器（网页版）
- 保留原 PDF 版面和文字位置：白块遮盖原文，原位写入中文
- 生成左右对照 HTML：左原版、右译文，逐页严格对齐
- Gradio 网页界面：拖拽上传 PDF，可选择模型，可留空 Key 用 .env 默认
- 断点续传：翻译结果缓存在本地 JSON
"""

import os
import sys
import json
import time
import re
import hashlib

import fitz  # PyMuPDF
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
CACHE_FILE = "translate_cache.json"
OUTPUT_PDF = "translated.pdf"
OUTPUT_HTML = "bilingual.html"
IMG_DIR = "bilingual_pages"
RENDER_ZOOM = 2.0

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


def h(s):
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:20]


def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_cache(c):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(c, f, ensure_ascii=False)


def parse_marked(text, n):
    """把 [[B0]] xxx [[B1]] yyy ... 解析成 {0: 'xxx', 1: 'yyy'}"""
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
    """翻译一页的所有文本块，带缓存"""
    marked = "\n\n".join(f"[[B{i}]] {b[4].strip()}" for i, b in enumerate(blocks))
    key = "pg_" + h(marked)
    if key in cache:
        try:
            return {int(k): v for k, v in cache[key].items()}
        except Exception:
            pass

    raw = call_api(client, model, marked)
    parsed = parse_marked(raw, len(blocks))

    # 模型偶尔漏标记，逐块补救
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
    """粗略估计文本在给定宽度下占几行"""
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
    """找一个能塞进 rect 的字号"""
    for fs in [11, 10.5, 10, 9.5, 9, 8.5, 8, 7.5, 7, 6.5, 6, 5.5, 5, 4.5, 4]:
        if estimate_lines(rect.width, text, fs) * fs * 1.35 <= rect.height + 2:
            return fs
    return 4


def apply_translations(page, blocks, translations):
    """在原文位置遮盖 + 写入译文"""
    page.insert_font(fontname="cn", fontfile=FONT_PATH)
    for i, b in enumerate(blocks):
        text = translations.get(i, "").strip()
        if not text:
            continue
        rect = fitz.Rect(b[0], b[1], b[2], b[3])
        page.draw_rect(rect, color=None, fill=(1, 1, 1), overlay=True)
        fs = find_fontsize(rect, text)
        page.insert_textbox(
            rect, text,
            fontname="cn", fontsize=fs,
            color=(0, 0, 0), align=0, overlay=True,
        )


def generate_html(src_path, trans_path, num_pages):
    """渲染原版 + 译版每页为 PNG，生成左右对照 HTML"""
    os.makedirs(IMG_DIR, exist_ok=True)
    orig = fitz.open(src_path)
    trans = fitz.open(trans_path)
    n = min(num_pages, len(orig), len(trans))
    pairs = []
    for i in range(n):
        o_path = f"{IMG_DIR}/orig_{i:04d}.png"
        t_path = f"{IMG_DIR}/trans_{i:04d}.png"
        orig[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM)).save(o_path)
        trans[i].get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM)).save(t_path)
        pairs.append(
            f'<div class="pair">'
            f'<div class="labels"><div>原文 · Page {i+1}</div><div>译文 · Page {i+1}</div></div>'
            f'<div class="pages"><img src="{o_path}"><img src="{t_path}"></div>'
            f'</div>'
        )
    orig.close()
    trans.close()

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
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)


def stop_translate():
    global STOP
    STOP = True
    return "正在停止……当前页翻完就停，进度已保存。"


def run(api_key, model, pdf, trial):
    global STOP
    STOP = False
    log = []

    def say(m):
        log.append(m)
        return "\n".join(log[-40:])

    # 优先用网页里填的 Key，留空则回退到 .env
    real_key = (api_key or "").strip() or DEFAULT_API_KEY
    if not real_key:
        yield say("❌ 未提供 API Key：在网页里填入，或在 .env 里设置 DEEPSEEK_API_KEY"), None
        return
    if not pdf:
        yield say("❌ 请先上传 PDF 文件"), None
        return
    if not os.path.exists(FONT_PATH):
        yield say(f"❌ 字体文件不存在：\n{FONT_PATH}"), None
        return

    if (api_key or "").strip():
        yield say("🔑 使用网页填写的 API Key"), None
    else:
        yield say("🔑 使用 .env 中的默认 API Key"), None

    model = (model or DEFAULT_MODEL).strip()
    if model not in ("deepseek-chat", "deepseek-reasoner"):
        model = "deepseek-chat"
    yield say(f"🧠 模型：{model}"), None

    client = OpenAI(api_key=real_key, base_url="https://api.deepseek.com")
    src_path = pdf.name
    doc = fitz.open(src_path)
    total = len(doc)
    limit = min(5, total) if trial else total

    cache = load_cache()
    yield say(f"✅ PDF 共 {total} 页；本次处理前 {limit} 页；缓存已有 {len(cache)} 条翻译"), None

    for pno in range(limit):
        if STOP:
            yield say("🛑 已停止，进度已保存，可重新开始接着翻。"), None
            return

        page = doc[pno]
        blocks = [b for b in page.get_text("blocks") if b[6] == 0 and b[4].strip()]
        if not blocks:
            yield say(f"⏭ 第 {pno+1} 页无文字（可能是插图页），跳过"), None
            continue

        yield say(f"📄 正在翻译第 {pno+1}/{limit} 页（{len(blocks)} 段）……"), None
        try:
            trans = translate_page(client, model, blocks, cache, pno)
        except RuntimeError as e:
            yield say(f"❌ {e}"), None
            return
        except Exception as e:
            yield say(f"❌ 未知错误：{e}"), None
            return

        apply_translations(page, blocks, trans)
        yield say(f"✅ 第 {pno+1} 页完成"), None

    try:
        doc.save(OUTPUT_PDF, garbage=4, deflate=True)
    finally:
        doc.close()
    yield say(f"✅ 译文 PDF 已保存：{OUTPUT_PDF}"), None

    yield say("🖼 正在生成左右对照网页……"), None
    generate_html(src_path, OUTPUT_PDF, limit)
    yield say(f"🎉 全部完成！下载下方 {OUTPUT_HTML}，双击用浏览器打开。"), OUTPUT_HTML


# ================= UI =================
with gr.Blocks(title="PDF 双语对照翻译器") as demo:
    gr.Markdown(
        "# 📖 PDF 双语对照翻译器\n"
        "**保留原排版 · 左原版右译文 · 逐页严格对齐 · 断点续传**\n\n"
        "> Key 留空 → 使用 `.env` 里的默认 Key；填入 → 临时覆盖。"
    )
    with gr.Row():
        api_key = gr.Textbox(
            label="DeepSeek API Key（留空用 .env 默认）",
            type="password",
            placeholder="留空 → 使用 .env；或在此临时换一个 sk-...",
            scale=3,
        )
        model = gr.Dropdown(
            choices=["deepseek-chat", "deepseek-reasoner"],
            value=DEFAULT_MODEL,
            label="翻译模型",
            scale=1,
        )
    pdf = gr.File(label="上传英文 PDF（拖拽或点击选择）", file_types=[".pdf"])
    trial = gr.Checkbox(
        label="🧪 试翻模式：只翻前 5 页，先看效果和费用",
        value=True,
    )
    with gr.Row():
        btn = gr.Button("开始翻译", variant="primary", scale=3)
        stop_btn = gr.Button("停止", scale=1)
    log = gr.Textbox(label="运行日志", lines=14, interactive=False)
    out_file = gr.File(label="下载对照网页（HTML）")

    btn.click(run, [api_key, model, pdf, trial], [log, out_file])
    stop_btn.click(stop_translate, None, [log])


if __name__ == "__main__":
    # 启动时给一个提示，方便排查 .env 是否加载成功
    if DEFAULT_API_KEY:
        print(f"✅ 已从 .env 读取默认 Key（{DEFAULT_API_KEY[:6]}...）")
    else:
        print("ℹ️  .env 中未找到 DEEPSEEK_API_KEY，需在网页里手动填写")
    print(f"ℹ️  默认模型：{DEFAULT_MODEL}")

    demo.launch(inbrowser=True)
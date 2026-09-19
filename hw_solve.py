# -*- coding: utf-8 -*-
"""作业解题器 · 解题业务：提示词 / API / 解析响应 / 逐页解题 / 贴图 / 公式渲染。"""

import os
import re
import time
import io
import functools
import threading

import numpy as np
import pymupdf as fitz
from PIL import Image

import matplotlib
matplotlib.use("Agg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib import font_manager as mfont
matplotlib.rcParams["mathtext.fontset"] = "cm"
matplotlib.rcParams["axes.unicode_minus"] = False

import hw_core

ANS_COLOR = "#c0392b"
SOL_COLOR = "#1a3a8a"
SOL_BG = (0.94, 0.96, 1.0)

# ★ 优化：学科微调 Prompt
_SUBJECT_PROMPTS = {
    "general": "",
    "math": (
        "\n【数学特别注意】\n"
        "- 涉及计算时给出关键中间步骤，不要跳步；\n"
        "- 单位、符号、下标要完整；\n"
        "- 多解的情况要全部指出，并说明取舍。\n"
    ),
    "english": (
        "\n【英语特别注意】\n"
        "- 语法题指出考点（时态 / 从句 / 非谓语等）；\n"
        "- 完形填空结合上下文逻辑，不要只看单句；\n"
        "- 翻译要地道，符合英语母语者表达；\n"
        "- 阅读理解题给依据（原文哪句支持）。\n"
    ),
    "physics": (
        "\n【物理特别注意】\n"
        "- 先列公式、再代数值、最后给单位；\n"
        "- 说明物理过程（受力 / 运动 / 能量）；\n"
        "- 受力分析要完整，不要漏力。\n"
    ),
    "chemistry": (
        "\n【化学特别注意】\n"
        "- 方程式要配平，注明反应条件；\n"
        "- 有机题注意官能团和反应类型；\n"
        "- 计算题给摩尔比和单位。\n"
    ),
    "chinese": (
        "\n【语文特别注意】\n"
        "- 阅读题结合文本，不要空谈；\n"
        "- 古诗文先释义，再赏析；\n"
        "- 作文题给思路提纲和素材方向。\n"
    ),
}

_MPL_LOCK = threading.Lock()

CN_FONT = None
if hw_core.CN_FONT_PATH and os.path.exists(hw_core.CN_FONT_PATH):
    try:
        mfont.fontManager.addfont(hw_core.CN_FONT_PATH)
        CN_FONT = mfont.FontProperties(fname=hw_core.CN_FONT_PATH)
        print(f"[font] 已加载中文字体：{hw_core.CN_FONT_PATH}")
    except Exception as e:
        print(f"[warn] 中文字体注册失败：{e}")
else:
    print("[warn] 未找到中文字体，中文可能显示为方框")
    print(f"       扫描目录：{hw_core.FONT_ROOTS}")


# ============================================================
# matplotlib 渲染
# ============================================================
@functools.lru_cache(maxsize=512)
def render_rich_text_png(text, width_pt, fontsize=11, dpi=None,
                         color="#c0392b"):
    """
    渲染富文本（含 mathtext 公式）成 PNG。
    ★ 优化：按文本长度预估画布高度，避免每次开 40 英寸画布；
    ★ 优化：dpi 从 200 提到 300，公式更清晰；
    ★ 修复：bbox 溢出画布时自动加高重试，避免公式被截断；
    ★ 修复：按参数缓存渲染结果，相同答案不再重复渲染。
    """
    if dpi is None:
        dpi = hw_core._RENDER_DPI

    text = (text or "").strip()
    if not text:
        return None, 0.0, 0.0

    width_inch = max(0.5, float(width_pt) / 72.0)

    chars_per_line = max(6, int(width_pt / max(fontsize, 1)))
    total_lines = 0
    for line in text.split("\n"):
        if not line:
            total_lines += 1
        else:
            total_lines += max(
                1, (len(line) + chars_per_line - 1) // chars_per_line
            )
    if "$" in text:
        total_lines = int(total_lines * 1.5) + 1

    est_height_pt = max(
        fontsize * 1.55 * (total_lines + 4),
        fontsize * 4,
    )
    est_height_inch = max(0.3, est_height_pt / 72.0)

    with _MPL_LOCK:
        for _attempt in range(3):
            fig = Figure(figsize=(width_inch, est_height_inch), dpi=dpi)
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
                try:
                    fig.clear()
                except Exception:
                    pass
                return None, 0.0, 0.0

            buf = np.asarray(canvas.buffer_rgba())
            H, W = buf.shape[:2]

            overflow = (bbox.y0 < 1) or (bbox.x1 > W - 1)
            if overflow and _attempt < 2:
                est_height_inch *= 1.6
                try:
                    fig.clear()
                except Exception:
                    pass
                continue

            pad = 3
            x0 = max(0, int(bbox.x0) - pad)
            x1 = min(W, int(bbox.x1) + pad)
            y_top = max(0, H - int(bbox.y1) - pad)
            y_bot = min(H, H - int(bbox.y0) + pad)

            if x1 <= x0 or y_bot <= y_top:
                try:
                    fig.clear()
                except Exception:
                    pass
                return None, 0.0, 0.0

            crop = np.ascontiguousarray(buf[y_top:y_bot, x0:x1])
            img = Image.fromarray(crop, mode="RGBA")

            out = io.BytesIO()
            img.save(out, "PNG")

            w_pt = (x1 - x0) / dpi * 72.0
            h_pt = (y_bot - y_top) / dpi * 72.0

            try:
                fig.clear()
            except Exception:
                pass

            return out.getvalue(), w_pt, h_pt

        return None, 0.0, 0.0


# ============================================================
# Prompt
# ============================================================
def build_solve_system(subject="general", extra_prompt=""):
    """
    ★ 优化：新增 subject（学科微调）与 extra_prompt（用户自定义补充）
    """
    base = (
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

    if subject in _SUBJECT_PROMPTS:
        base += _SUBJECT_PROMPTS[subject]

    if extra_prompt and extra_prompt.strip():
        base += "\n【用户额外要求】\n" + extra_prompt.strip() + "\n"

    return base


# ============================================================
# API
# ============================================================
def _extract_status(err):
    for attr in ("status_code", "code"):
        v = getattr(err, attr, None)
        if isinstance(v, int):
            return v
    resp = getattr(err, "response", None)
    if resp is not None:
        v = getattr(resp, "status_code", None)
        if isinstance(v, int):
            return v
    return None


def _interruptible_sleep(seconds, stop_event):
    if stop_event is None:
        time.sleep(seconds)
        return False
    end = time.time() + seconds
    while True:
        remaining = end - time.time()
        if remaining <= 0:
            return False
        if stop_event.is_set():
            return True
        time.sleep(min(0.25, remaining))


def _api_call(client, model, system, user_content, retries=4, stop_event=None):
    last_err = ""
    for attempt in range(retries):
        if stop_event is not None and stop_event.is_set():
            raise RuntimeError("用户已停止")
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
            status = _extract_status(e)
            low = last_err.lower()
            if status == 402 or "insufficient" in low or "余额" in last_err:
                raise RuntimeError("账户余额不足，请去 DeepSeek 平台充值后再继续")
            if status == 401:
                raise RuntimeError("API Key 无效或已过期，请检查后重试")
            # ★ 修复：最后一次重试不再白睡一轮再抛错
            last_attempt = attempt >= retries - 1
            if status == 429:
                if last_attempt:
                    break
                _interruptible_sleep(10 * (attempt + 1), stop_event)
                continue
            if last_attempt:
                break
            _interruptible_sleep(5 * (attempt + 1), stop_event)
    raise RuntimeError(f"API 连续失败：{last_err}")


def call_solve_api(client, model, text, retries=4, stop_event=None,
                   subject="general", extra_prompt=""):
    system = build_solve_system(subject, extra_prompt)
    return _api_call(client, model, system, text, retries,
                     stop_event=stop_event)


# ============================================================
# 解析模型返回
# ============================================================
def _clean_tags(s):
    s = re.sub(r'</?ANS>', '', s or "", flags=re.I)
    s = re.sub(r'</?SOL>', '', s, flags=re.I)
    return s.strip()


def parse_solution_response(text, n_blocks=None):
    _ = n_blocks
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

        if not ans and not sol:
            stripped = _clean_tags(body)
            stripped = re.sub(r'\[\[B\d+\]\]', '', stripped).strip()
            if stripped and len(stripped) < 500:
                sol = stripped

        if ans or sol:
            result[idx] = {"ans": ans, "sol": sol}
    return result


# ============================================================
# PDF：解题 + 排版（支持三种答案位置）
# ============================================================
def solve_page(client, model, blocks, cache, cache_file, stop_event=None,
               subject="general", extra_prompt=""):
    marked = "\n\n".join(f"[[B{i}]] {b[4].strip()}"
                         for i, b in enumerate(blocks))
    key = "solve_" + hw_core.h(marked + "|" + subject + "|" + (extra_prompt or ""))
    if key in cache:
        try:
            cached = {int(k): v for k, v in cache[key].items()}
            if all(isinstance(v, dict) for v in cached.values()):
                return cached
        except Exception:
            pass

    raw = call_solve_api(client, model, marked, stop_event=stop_event,
                         subject=subject, extra_prompt=extra_prompt)
    parsed = parse_solution_response(raw, len(blocks))

    cache[key] = {str(k): v for k, v in parsed.items()}
    hw_core.save_json_file(cache_file, cache)
    return parsed


def solve_batch_office(client, model, texts, cache, cache_file, prefix,
                       stop_event=None, task=None,
                       subject="general", extra_prompt=""):
    """★ 修复：Word/PPT 批量解题（对齐双语版 translate_batch_office）。
    - 缓存键格式与旧逐段版完全一致（solve_<prefix>_ + h(text|subject)），兼容已有缓存；
    - 未命中段落按 OFFICE_BATCH_SIZE 一段次 API 调用，parse_solution_response 解析；
    - 任何 RuntimeError 直接上抛（与旧版"出错即终止任务"语义一致）。
    返回：results（与 texts 等长，None 表示因停止未处理）。"""
    results = [None] * len(texts)

    def ckey(t):
        return f"solve_{prefix}_" + hw_core.h(t + "|" + subject)

    pending = []
    for i, t in enumerate(texts):
        k = ckey(t)
        if k in cache and isinstance(cache[k], dict):
            results[i] = cache[k]
        else:
            pending.append((i, t))

    for start in range(0, len(pending), hw_core.OFFICE_BATCH_SIZE):
        if stop_event is not None and stop_event.is_set():
            break
        chunk = pending[start:start + hw_core.OFFICE_BATCH_SIZE]
        marked = "\n\n".join(f"[[B{j}]] {t}" for j, (_, t) in enumerate(chunk))
        raw = call_solve_api(client, model, marked, stop_event=stop_event,
                             subject=subject, extra_prompt=extra_prompt)
        parsed = parse_solution_response(raw, len(chunk))
        for j, (gi, t) in enumerate(chunk):
            res = parsed.get(j, {"ans": "", "sol": ""})
            results[gi] = res
            cache[ckey(t)] = res
        hw_core.save_json_file(cache_file, cache)

    return results


def apply_solution_to_page(page, blocks, results, ans_position="inside",
                           warn=None):
    """
    ★ 优化：ans_position 三档
      - inside: 答案贴块内右下（原逻辑）
      - below:  答案贴块下方
      - right:  答案贴块右侧（不够时退回下方）
    解析始终贴块下方（若答案已占块下方，则接在答案下面）。
    ★ 修复：贴下方前先量到下一个块的距离，不够就缩放；仍放不下则跳过
    并通过 warn 回调记日志，不再盖住下一题。
    """
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

        # ★ 修复：本块正下方最近的块顶（没有则取页底）
        next_y0 = page_rect.height
        for b2 in blocks[i + 1:]:
            ny0 = float(b2[1])
            if ny0 >= rect.y1 + 1:
                next_y0 = ny0
                break

        # 记录「解析的起始 y 位置」——默认为块底
        sol_top = rect.y1

        # ============ 答案 ============
        if ans:
            png, w_pt, h_pt = render_rich_text_png(
                ans, max(20.0, rect.width - 6),
                fontsize=11, color=ANS_COLOR,
            )
            if png and w_pt > 0 and h_pt > 0:
                dst = None

                # 右侧留白
                if ans_position == "right":
                    right_space = page_rect.width - 20.0 - rect.x1
                    if right_space >= 180:
                        max_w = min(right_space - 5, 400.0)
                        sc = min(1.0, max_w / w_pt)
                        w2 = max(20.0, w_pt * sc)
                        h2 = max(10.0, h_pt * sc)
                        y_center = (rect.y0 + rect.y1) / 2.0
                        dst = fitz.Rect(rect.x1 + 5, y_center - h2 / 2,
                                        rect.x1 + 5 + w2, y_center + h2 / 2)

                # 块内右下
                elif ans_position == "inside":
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

                # fallback（below / inside 失败 / right 无空间）
                if dst is None:
                    max_w = max(60.0, rect.width)
                    sc = min(1.0, max_w / w_pt)
                    w2 = max(20.0, w_pt * sc)
                    h2 = max(10.0, h_pt * sc)
                    sy0 = rect.y1 + 2
                    sy1 = sy0 + h2
                    # ★ 修复：先量到下一题的距离，不够就缩放，再不够跳过
                    avail = next_y0 - 2 - sy0
                    if avail < h2:
                        sc2 = avail / h2 if h2 > 0 and avail > 0 else 0.0
                        if sc2 >= 0.4:
                            w2 *= sc2
                            h2 = avail
                            sy1 = sy0 + h2
                        else:
                            if warn:
                                warn(f"⚠️ 第 {i + 1} 块下方空间不足，"
                                     f"答案/解析未贴入（避免盖住下一题）")
                            continue
                    if sy1 > page_rect.height - 8:
                        sy1 = page_rect.height - 8
                        sy0 = max(rect.y1 + 2, sy1 - h2)
                        sy1 = sy0 + h2
                    dst = fitz.Rect(rect.x0, sy0, rect.x0 + w2, sy1)
                    sol_top = sy1

                if dst is not None:
                    try:
                        page.insert_image(dst, stream=png, overlay=True)
                    except Exception:
                        pass

        # ============ 解析 ============
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

                sy0 = sol_top + 2
                sy1 = sy0 + h2

                # ★ 修复：先量到下一题的距离，不够就缩放，再不够跳过
                avail = next_y0 - 2 - sy0
                if avail < h2:
                    sc2 = avail / h2 if h2 > 0 and avail > 0 else 0.0
                    if sc2 >= 0.4:
                        w2 *= sc2
                        h2 = avail
                        sy1 = sy0 + h2
                    else:
                        if warn:
                            warn(f"⚠️ 第 {i + 1} 块下方空间不足，"
                                 f"解析未贴入（避免盖住下一题）")
                        continue

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
# 仅解析模式：生成 Markdown 报告
# ============================================================
def build_solution_markdown(book_title, page_results):
    """
    page_results: list of (page_no, [(src, ans, sol), ...])
    """
    lines = [f"# 《{book_title}》解题报告", ""]
    lines.append(f"- 生成时间：{time.strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"- 题目总数：{sum(len(it) for _, it in page_results)}")
    lines.append("")
    lines.append("> 本报告由 AI 生成，仅供参考。关键结论请自行核对。")
    lines.append("")

    for page_no, items in page_results:
        if not items:
            continue
        lines.append(f"## 第 {page_no} 页")
        lines.append("")
        for idx, (src, ans, sol) in enumerate(items, 1):
            src_clean = (src or "").strip().replace("\n", " ")
            if len(src_clean) > 160:
                src_clean = src_clean[:157] + "…"
            lines.append(f"### 第 {idx} 题")
            lines.append("")
            lines.append(f"**题目**　{src_clean}")
            lines.append("")
            if ans:
                lines.append(f"**答案**　{ans}")
                lines.append("")
            if sol:
                lines.append(f"**解析**　{sol}")
                lines.append("")
        lines.append("---")
        lines.append("")

    return "\n".join(lines)

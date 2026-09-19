# -*- coding: utf-8 -*-
"""PDF 翻译管线：解析标记、内容校验、排版写入、单页翻译、双语合成、预览图。"""

import os
import re
import io
import glob
import shutil
import time

import pymupdf as fitz
try:
    fitz.TOOLS.mupdf_display_errors(False)
except Exception:
    pass

from PIL import Image

from core import config
from core.utils import h, save_json_file
from core.api import cache_prefix, call_api, _is_fatal_api_error

_RESAMPLE = getattr(Image, "Resampling", Image).LANCZOS
_MARK_RE = re.compile(r'\[\[B(\d+)\]\]')


def safe_save_pdf(doc, out_path, retries=5, garbage=3):
    tmp_path = out_path + ".tmp.pdf"
    try:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    except Exception:
        pass

    doc.save(tmp_path, deflate=True, garbage=garbage)

    last_err = None
    for attempt in range(retries):
        try:
            os.replace(tmp_path, out_path)
            return out_path
        except (PermissionError, OSError) as e:
            last_err = e
            time.sleep(0.8 * (attempt + 1))

    if os.path.basename(out_path).endswith(".bak.pdf"):
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        raise last_err if last_err else RuntimeError("无法保存 PDF")

    fallback = out_path[:-4] + ".bak.pdf"
    try:
        if os.path.exists(fallback):
            try:
                os.remove(fallback)
            except Exception:
                pass
        os.replace(tmp_path, fallback)
        return fallback
    except Exception:
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        raise last_err if last_err else RuntimeError("无法保存 PDF")


def _pix_to_pil(pix):
    try:
        n = pix.n
        w, h_ = pix.width, pix.height
        if n == 3:
            return Image.frombytes("RGB", (w, h_), pix.samples)
        if n == 4:
            return Image.frombytes("RGBA", (w, h_), pix.samples).convert("RGB")
        if n == 1:
            return Image.frombytes("L", (w, h_), pix.samples).convert("RGB")
    except Exception:
        pass
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def parse_marked(text, _n=None):
    result = {}
    if not text:
        return result
    if config.HAS_NOTES and config.NB is not None:
        idx = text.rfind(config.NB.TERM_MARK)
        if idx >= 0:
            text = text[:idx]
    matches = list(_MARK_RE.finditer(text))
    for i, m in enumerate(matches):
        idx = int(m.group(1))
        if idx in result:
            continue
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        result[idx] = text[start:end].strip()
    return result


# ============================================================
# 内容校验
# ============================================================

def _has_target_chars(text, target_lang):
    if not text:
        return False
    if target_lang in ("zh-CN", "zh-TW"):
        return any('一' <= c <= '鿿' for c in text)
    elif target_lang == "ja":
        return any(
            ('぀' <= c <= 'ヿ') or ('一' <= c <= '鿿')
            for c in text
        )
    elif target_lang == "ko":
        return any('가' <= c <= '힣' for c in text)
    elif target_lang == "ru":
        return any('Ѐ' <= c <= 'ӿ' for c in text)
    elif target_lang == "ar":
        return any('؀' <= c <= 'ۿ' for c in text)
    return True


def page_is_translated(page, target_lang="zh-CN", src_page=None):
    if target_lang not in config.NON_LATIN_SCRIPT_LANGS:
        return True
    try:
        text = page.get_text()
    except Exception:
        return False

    if not text or not text.strip():
        if src_page is not None:
            try:
                if not src_page.get_text().strip():
                    return True
            except Exception:
                pass
        return False

    if src_page is None:
        return True

    try:
        src_text = src_page.get_text()
        if src_text.strip() == text.strip():
            return False
    except Exception:
        pass

    total_nonspace = sum(1 for c in text if not c.isspace())
    if total_nonspace < 5:
        if src_page is not None:
            try:
                src_ns = sum(
                    1 for c in src_page.get_text() if not c.isspace())
            except Exception:
                src_ns = 0
            if src_ns < 5:
                return True
        return False

    if target_lang in ("zh-CN", "zh-TW"):
        rng = ('一', '鿿')
    elif target_lang == "ja":
        cnt_kana = 0
        for c in text:
            if '぀' <= c <= 'ヿ':
                cnt_kana += 1
        if cnt_kana >= max(3, config.PAGE_CN_THRESHOLD // 4):
            return True
        rng = ('一', '鿿')
    elif target_lang == "ko":
        rng = ('가', '힣')
    elif target_lang == "ru":
        rng = ('Ѐ', 'ӿ')
    elif target_lang == "ar":
        rng = ('؀', 'ۿ')
    else:
        return True

    lo, hi = rng
    cnt = 0
    for c in text:
        if lo <= c <= hi:
            cnt += 1

    need_chars = max(3, min(config.PAGE_CN_THRESHOLD, total_nonspace // 3))
    if cnt < need_chars:
        return False
    return (cnt / total_nonspace) >= config.PAGE_CN_RATIO


def translations_look_valid(translations, blocks, target_lang="zh-CN"):
    if not translations or not blocks:
        return False
    valid = 0
    for i, b in enumerate(blocks):
        v = (translations.get(i) or "").strip()
        src = (b[4] or "").strip()
        if not v:
            continue
        if v.lower() == src.lower():
            continue
        if target_lang in config.NON_LATIN_SCRIPT_LANGS:
            if not _has_target_chars(v, target_lang) and len(v) > 5:
                continue
        valid += 1
    need = max(1, int(len(blocks) * config.VALID_RATIO_THRESHOLD))
    return valid >= need


# ============================================================
# PDF 字体与排版
# ============================================================

def _resolve_font(font_path):
    if font_path and os.path.exists(font_path):
        return font_path, "user_font"
    return None, "china-s"


def _pdf_font_setup(page, font_path=None):
    fp, name = _resolve_font(font_path)
    if fp:
        try:
            page.insert_font(fontname=name, fontfile=fp, set_simple=False)
            return name
        except Exception as e:
            # ★ 修复：字体注册失败至少打印，排障时可见
            print(f"[warn] 字体注册失败 {fp}: {e}")
    return "china-s"


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


def find_fontsize(rect, text, max_size=11.0):
    fs = float(max_size)
    min_fs = 4.5
    while fs >= min_fs:
        if estimate_lines(rect.width, text, fs) * fs * 1.35 <= rect.height + 2:
            return fs
        fs -= 0.5
    return min_fs


def _clean_block_text(t):
    return re.sub(r'[ \t]*\n[ \t]*', ' ', t or "").strip()


def apply_translations(page, blocks, translations,
                       font_path=None, max_font_size=11.0,
                       target_lang="zh-CN"):
    items = []
    for i, b in enumerate(blocks):
        text = (translations.get(i) or "").strip()
        if not text:
            continue
        try:
            rect = fitz.Rect(b[0], b[1], b[2], b[3])
        except Exception:
            continue
        if rect.is_empty or rect.width <= 0 or rect.height <= 0:
            continue
        items.append((rect, text))

    if not items:
        return 0, 0

    for rect, _ in items:
        try:
            page.add_redact_annot(rect)
        except Exception:
            pass

    try:
        page.apply_redactions(
            images=fitz.PDF_REDACT_IMAGE_NONE,
            graphics=fitz.PDF_REDACT_LINE_ART_NONE,
            text=fitz.PDF_REDACT_TEXT_REMOVE,
        )
    except Exception:
        try:
            page.apply_redactions()
        except Exception:
            pass

    fontname = _pdf_font_setup(page, font_path)
    align = 2 if target_lang in config.RTL_LANGS else 0
    written = 0
    failed = 0

    for rect, text in items:
        fs = find_fontsize(rect, text, max_font_size)

        def _try(r, size):
            try:
                return page.insert_textbox(
                    r, text, fontname=fontname, fontsize=size,
                    color=(0, 0, 0), align=align, overlay=True,
                )
            except Exception:
                return -1

        rc = _try(rect, fs)
        while rc < 0 and fs > 4.5:
            fs -= 0.5
            rc = _try(rect, fs)

        if rc < 0:
            ext = fitz.Rect(
                rect.x0, rect.y0,
                rect.x1,
                max(rect.y1, page.rect.y1 - 20),
            )
            rc = _try(ext, max(fs, 5.0))

        if rc < 0:
            try:
                page.insert_text(
                    fitz.Point(rect.x0, rect.y1 - 2),
                    text, fontname=fontname, fontsize=max(fs, 5.0),
                    color=(0, 0, 0), overlay=True,
                )
            except Exception:
                failed += 1
                continue

        written += len(text)

    return written, failed


def translate_page(client, model, blocks, cache, cache_file,
                   target_lang="zh-CN", stop_event=None,
                   reader_profile="", want_terms=True,
                   domain="general", extra_prompt=""):
    pref = cache_prefix(target_lang)
    cleaned = [_clean_block_text(b[4]) for b in blocks]
    marked = "\n\n".join(f"[[B{i}]] {t}" for i, t in enumerate(cleaned))

    terms_suffix = "|T" if (want_terms and config.HAS_NOTES) else ""
    # ★ 优化：缓存 key 加上 domain + extra_prompt
    domain_suffix = f"|D{domain}" if domain and domain != "general" else ""
    extra_suffix = ("|X" + h(extra_prompt)) if (extra_prompt and extra_prompt.strip()) else ""
    key = f"pg_{pref}" + h(marked + terms_suffix + domain_suffix + extra_suffix)

    if key in cache:
        try:
            entry = cache[key]
            if isinstance(entry, dict) and "paragraphs" in entry:
                cached = {int(k): v for k, v in entry["paragraphs"].items()}
                cached_terms = entry.get("terms", [])
            else:
                cached = {int(k): v for k, v in entry.items()}
                cached_terms = []
            cached_complete = all(
                (cached.get(i) or "").strip() for i in range(len(blocks))
            )
            if cached_complete and translations_look_valid(
                    cached, blocks, target_lang):
                return cached, cached_terms
        except Exception:
            pass

    raw = call_api(client, model, marked, target_lang,
                   stop_event=stop_event,
                   reader_profile=reader_profile,
                   want_terms=(want_terms and config.HAS_NOTES),
                   domain=domain, extra_prompt=extra_prompt)

    if config.HAS_NOTES and config.NB is not None:
        body, terms_raw = config.NB.split_translation_and_terms(raw)
    else:
        body, terms_raw = raw, ""

    parsed = parse_marked(body, len(blocks))
    terms = config.NB.parse_terms_block(terms_raw) \
        if (config.HAS_NOTES and config.NB is not None and want_terms) else []

    missing = [i for i in range(len(blocks))
               if i not in parsed or not parsed[i].strip()]

    if missing and not (stop_event is not None and stop_event.is_set()):
        try:
            missing_texts = [cleaned[i] for i in missing]
            batch_marked = "\n\n".join(
                f"[[B{j}]] {t}" for j, t in enumerate(missing_texts)
            )
            r = call_api(client, model, batch_marked, target_lang,
                         stop_event=stop_event,
                         reader_profile=reader_profile, want_terms=False,
                         domain=domain, extra_prompt=extra_prompt)
            if config.HAS_NOTES and config.NB is not None:
                body2, _ = config.NB.split_translation_and_terms(r)
            else:
                body2 = r
            sub = parse_marked(body2, len(missing))
            for j, i in enumerate(missing):
                v = (sub.get(j) or "").strip()
                if v:
                    parsed[i] = v
        except RuntimeError as e:
            if _is_fatal_api_error(str(e)):
                raise
        except Exception:
            pass

        still_missing = [i for i in missing
                         if not (parsed.get(i) or "").strip()]
        for i in still_missing:
            if stop_event is not None and stop_event.is_set():
                break
            bk = f"bk_{pref}" + h(cleaned[i])
            if bk in cache and isinstance(cache[bk], str) and cache[bk].strip():
                parsed[i] = cache[bk]
                continue
            tr = ""
            try:
                r = call_api(client, model, f"[[B0]] {cleaned[i]}",
                             target_lang, stop_event=stop_event,
                             reader_profile=reader_profile, want_terms=False,
                             domain=domain, extra_prompt=extra_prompt)
                if config.HAS_NOTES and config.NB is not None:
                    body2, _ = config.NB.split_translation_and_terms(r)
                else:
                    body2 = r
                sub = parse_marked(body2, 1)
                tr = (sub.get(0) or "").strip()
            except RuntimeError as e:
                if _is_fatal_api_error(str(e)):
                    raise
                tr = ""
            except Exception:
                tr = ""
            parsed[i] = tr
            if tr:
                cache[bk] = tr
                save_json_file(cache_file, cache)

    parsed_complete = all(
        (parsed.get(i) or "").strip() for i in range(len(blocks))
    )
    if parsed_complete and translations_look_valid(parsed, blocks, target_lang):
        cache[key] = {
            "paragraphs": {str(k): v for k, v in parsed.items()},
            "terms": terms,
        }
        save_json_file(cache_file, cache)

    return parsed, terms


def render_preview_only(trans_path, paths, task, n_preview=config.PREVIEW_PAGES):
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

    if not os.path.exists(trans_path):
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        return []

    try:
        with open(trans_path, "rb") as f:
            data = f.read()
        trans = fitz.open(stream=data, filetype="pdf")
    except Exception as e:
        print(f"[preview] open trans failed: {e}")
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        return []

    try:
        orig = fitz.open(task.src_path)
    except Exception as e:
        print(f"[preview] open src failed: {e}")
        trans.close()
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        return []

    n = min(n_preview, len(orig), len(trans))
    preview_imgs = []
    zoom_mat = fitz.Matrix(config.RENDER_ZOOM, config.RENDER_ZOOM)
    for i in range(n):
        if task.stop_event.is_set():
            break
        try:
            o_pix = orig[i].get_pixmap(matrix=zoom_mat)
            t_pix = trans[i].get_pixmap(matrix=zoom_mat)
            o = _pix_to_pil(o_pix)
            t = _pix_to_pil(t_pix)
            hh = max(o.height, t.height)
            if o.height != hh:
                o = o.resize((int(o.width * hh / o.height), hh), _RESAMPLE)
            if t.height != hh:
                t = t.resize((int(t.width * hh / t.height), hh), _RESAMPLE)
            gap = 8
            canvas = Image.new("RGB", (o.width + gap + t.width, hh), (40, 40, 40))
            canvas.paste(o, (0, 0))
            canvas.paste(t, (o.width + gap, 0))
            if canvas.width > config.PREVIEW_MAX_WIDTH:
                ratio = config.PREVIEW_MAX_WIDTH / canvas.width
                canvas = canvas.resize(
                    (config.PREVIEW_MAX_WIDTH, int(canvas.height * ratio)),
                    _RESAMPLE,
                )
            p = os.path.join(tmp_dir, f"compare_{i:04d}.jpg")
            canvas.save(p, "JPEG", quality=config.PREVIEW_JPEG_QUALITY, optimize=True)
            preview_imgs.append(p)
        except Exception as e:
            print(f"[preview] page {i} render failed: {e}")
            continue
    orig.close()
    trans.close()

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


def make_bilingual_pdf(trans_path, paths, task,
                       n_pages=None, done_pages=None,
                       zoom=None, jpeg_quality=None, garbage=3):
    if not os.path.exists(trans_path):
        task.log_msg(f"⚠️ 双语 PDF 源不存在：{trans_path}")
        return
    try:
        with open(trans_path, "rb") as f:
            data = f.read()
        trans = fitz.open(stream=data, filetype="pdf")
    except Exception as e:
        task.log_msg(f"⚠️ 双语 PDF 打开源失败：{e}")
        return

    try:
        orig = fitz.open(task.src_path)
        total = min(len(orig), len(trans))

        if done_pages:
            pages = sorted(p for p in done_pages if 1 <= p <= total)
        else:
            cap = total if n_pages is None else min(n_pages, total)
            pages = list(range(1, cap + 1))

        if not pages:
            task.log_msg("⚠️ 双语 PDF 无有效页")
            orig.close()
            trans.close()
            return

        out_doc = fitz.open()
        done_count = 0
        n_total = len(pages)

        # ★ 修复：每 BILINGUAL_FLUSH_EVERY 页检查点落盘并重开 out_doc，
        # 释放已插入页的内存（300+ 页印刷级原本峰值可达数百 MB ~ 1GB）
        ckpt_a = os.path.join(task.work_dir, "_bilingual_ckpt.1.pdf")
        ckpt_b = os.path.join(task.work_dir, "_bilingual_ckpt.2.pdf")
        out_doc_holder = [out_doc]
        ckpt_toggle = 0

        def _checkpoint_flush():
            """落盘并重开 out_doc；失败时保证 out_doc_holder[0] 仍可用。"""
            nonlocal ckpt_toggle
            doc = out_doc_holder[0]
            ckpt = ckpt_a if ckpt_toggle % 2 == 0 else ckpt_b
            ckpt_toggle += 1
            prev = ckpt_b if ckpt is ckpt_a else ckpt_a
            try:
                doc.save(ckpt + ".new", deflate=True, garbage=0)
                doc.close()
                out_doc_holder[0] = None
            except Exception:
                return False
            try:
                os.replace(ckpt + ".new", ckpt)
                out_doc_holder[0] = fitz.open(ckpt)
            except Exception:
                # replace 失败：改从 .new 继续（内容完整）
                out_doc_holder[0] = fitz.open(ckpt + ".new")
                return False
            try:
                if os.path.exists(prev):
                    os.remove(prev)
            except Exception:
                pass
            return True

        zoom_val = float(zoom) if zoom else config.BILINGUAL_ZOOM
        jpeg_val = int(jpeg_quality) if jpeg_quality else config.BILINGUAL_JPEG_QUALITY
        zoom_mat = fitz.Matrix(zoom_val, zoom_val)

        for idx, pno in enumerate(pages):
            if task.stop_event.is_set():
                task.log_msg(
                    f"⏸ 双语 PDF 生成时收到停止信号，已生成 {done_count} 页")
                break
            i = pno - 1
            try:
                o_pix = orig[i].get_pixmap(matrix=zoom_mat)
                t_pix = trans[i].get_pixmap(matrix=zoom_mat)
                o = _pix_to_pil(o_pix)
                t = _pix_to_pil(t_pix)

                hh = max(o.height, t.height)
                if o.height != hh:
                    o = o.resize((int(o.width * hh / o.height), hh),
                                 _RESAMPLE)
                if t.height != hh:
                    t = t.resize((int(t.width * hh / t.height), hh),
                                 _RESAMPLE)

                gap = 10
                canvas = Image.new(
                    "RGB", (o.width + gap + t.width, hh), (255, 255, 255))
                canvas.paste(o, (0, 0))
                canvas.paste(t, (o.width + gap, 0))

                buf = io.BytesIO()
                canvas.save(buf, "JPEG",
                            quality=jpeg_val, optimize=False)
                img_bytes = buf.getvalue()
                buf.close()

                w, ph = canvas.size
                page = out_doc_holder[0].new_page(width=w, height=ph)
                page.insert_image(fitz.Rect(0, 0, w, ph), stream=img_bytes)

                canvas.close()
                o.close()
                t.close()
                done_count += 1

                if done_count % 10 == 0 or done_count == n_total:
                    task.log_msg(f"📐 双语 PDF {done_count}/{n_total} 页")

                if (done_count % config.BILINGUAL_FLUSH_EVERY == 0
                        and done_count < n_total):
                    try:
                        if not _checkpoint_flush():
                            task.log_msg("⚠️ 双语 PDF 检查点异常（继续内存生成）")
                    except Exception as e:
                        task.log_msg(f"⚠️ 双语 PDF 检查点失败：{e}")
                    if out_doc_holder[0] is None:
                        # 双保险：从最新的检查点文件恢复，避免丢已生成页
                        cand = None
                        for c in (ckpt_a + ".new", ckpt_a,
                                  ckpt_b + ".new", ckpt_b):
                            if os.path.exists(c):
                                cand = c
                                break
                        out_doc_holder[0] = fitz.open(cand) if cand \
                            else fitz.open()
            except Exception as e:
                task.log_msg(f"⚠️ 双语 PDF 第 {pno} 页失败：{e}")
                continue

        orig.close()
        trans.close()

        if done_count > 0:
            try:
                saved = safe_save_pdf(out_doc_holder[0], paths["bilingual_pdf"],
                                      garbage=garbage)
                if saved != paths["bilingual_pdf"]:
                    task.log_msg(
                        f"⚠️ 双语 PDF 保存到备用路径：{os.path.basename(saved)}"
                    )
                    paths["bilingual_pdf"] = saved
                task.log_msg(
                    f"✅ 双语 PDF 生成完成：{done_count} 页 → "
                    f"{os.path.basename(paths['bilingual_pdf'])}"
                )
                # 成功落盘后清理检查点文件
                for ck in (ckpt_a, ckpt_b):
                    for ext in ("", ".new"):
                        try:
                            if os.path.exists(ck + ext):
                                os.remove(ck + ext)
                        except Exception:
                            pass
            except Exception as e:
                task.log_msg(f"⚠️ 双语 PDF 保存失败：{e}")
        else:
            task.log_msg("⚠️ 双语 PDF 生成 0 页")
        try:
            out_doc_holder[0].close()
        except Exception:
            pass
    except Exception as e:
        import traceback
        task.log_msg(f"⚠️ 双语 PDF 生成异常：{e}")
        task.log_msg(traceback.format_exc()[:500])

# -*- coding: utf-8 -*-
"""三个 worker：pdf_worker / docx_worker / pptx_worker。

worker 只操作传入的 TaskState 对象，不碰任何 UI 全局状态。
"""

import os

import pymupdf as fitz
from openai import OpenAI

from core import config
from core.utils import (load_json_file, load_progress_file, save_progress_file,
                        fmt_size, estimate_output_size, h)
from core.tasks import save_state, _add_output, refresh_task_size
from core.api import cache_prefix
from core.preview import _build_docx_preview_html, _build_pptx_preview_html
from core.pdf_pipeline import (translate_page, apply_translations,
                               translations_look_valid, page_is_translated,
                               safe_save_pdf, render_preview_only,
                               make_bilingual_pdf)
from core.notes import insert_notes_into_pdf
from core.office import (translate_batch_office, _docx_replace_para_text,
                         _docx_insert_after, _pptx_set_para_text,
                         _iter_pptx_shapes, Document, Presentation)


# ============================================================
# PDF Worker
# ============================================================

def pdf_worker(task, paths, real_key, model, trial, target_lang="zh-CN",
               reader_profile="", want_terms=True,
               font_path="", max_font_size=config.DEFAULT_FONT_SIZE,
               domain="general", extra_prompt="", trial_pages=5,
               regen_bilingual=True):
    # ★ 修复：base_url 进 .env；设 timeout + max_retries=0（避免 SDK
    # 默认 600s 超时 × 内置重试 × 自写重试叠加，停止信号无法打断）
    client = OpenAI(api_key=real_key, base_url=config.DEEPSEEK_BASE_URL,
                    timeout=config.API_TIMEOUT, max_retries=0)
    cache = load_json_file(paths["cache_file"], {})
    done_pages = load_progress_file(paths["progress_file"])

    preset = config.PDF_QUALITY_PRESETS.get(
        task.pdf_quality, config.PDF_QUALITY_PRESETS[config.DEFAULT_PDF_QUALITY]
    )
    if task.make_bilingual:
        task.log_msg(
            f"🎨 双语 PDF 质量：{task.pdf_quality}"
            f"（zoom={preset['zoom']}，JPEG={preset['jpeg']}）"
        )
    else:
        task.log_msg("🎨 已关闭双语 PDF 生成（仅输出译文）")

    enable_terms = bool(want_terms and config.HAS_NOTES and config.NB is not None)
    terms_file = os.path.join(task.work_dir, "terms.json")
    global_terms = config.NB.load_terms(terms_file) if enable_terms else {"terms": {}}

    chapters = []
    if enable_terms:
        try:
            with fitz.open(task.src_path) as _d:
                chapters = config.NB.detect_chapters(_d)
            if chapters:
                task.log_msg(f"📚 识别到 {len(chapters)} 个章节，术语将按章归类")
            else:
                task.log_msg("📚 未检测到 PDF 目录，术语不按章节归类")
        except Exception as e:
            task.log_msg(f"⚠️ 章节识别失败：{e}")

    if font_path and os.path.exists(font_path):
        task.log_msg(f"🔤 正文字体：{os.path.basename(font_path)}")
    else:
        task.log_msg("🔤 正文字体：PyMuPDF 内置宋体（china-s）")
    task.log_msg(f"📏 正文字号上限：{max_font_size}pt")
    if target_lang in config.RTL_LANGS:
        task.log_msg(f"↔️ 目标语言为 RTL，使用右对齐")
    if domain != "general":
        task.log_msg(f"📚 领域微调：{domain}")
    if extra_prompt and extra_prompt.strip():
        task.log_msg(f"📝 已启用自定义要求（{len(extra_prompt)} 字）")

    try:
        with fitz.open(task.src_path) as _tmp_doc:
            total_src_pages = len(_tmp_doc)
    except Exception:
        total_src_pages = 0
    est = estimate_output_size(
        "pdf", task.src_path, total_src_pages,
        target_lang=target_lang, want_terms=enable_terms,
        pdf_quality=task.pdf_quality,
        make_bilingual=task.make_bilingual,
    )
    task.estimated_size = est.get("total", 0)
    if task.estimated_size:
        task.log_msg(
            f"📦 预估总产出 ≈ {fmt_size(task.estimated_size)}"
            f"（译文 ≈ {fmt_size(est.get('main', 0))} · "
            f"双语 ≈ {fmt_size(est.get('bilingual', 0))} · "
            f"笔记 ≈ {fmt_size(est.get('notes', 0))}）"
        )
    save_state(task, force=True)

    pdf_exists = os.path.exists(paths["output_pdf"])
    if done_pages and pdf_exists:
        try:
            with open(paths["output_pdf"], "rb") as f:
                data = f.read()
            doc = fitz.open(stream=data, filetype="pdf")
            task.log_msg(f"📂 从已翻译 PDF 续传（进度记录说已翻 {len(done_pages)} 页）")
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

    if pdf_exists and target_lang in config.NON_LATIN_SCRIPT_LANGS:
        try:
            orig_check = fitz.open(task.src_path)
            actually_done = set()
            checked = 0
            n_pages = min(len(doc), len(orig_check))
            # ★ 修复：抽样校验（最多 ~40 页），全量逐页 get_text 大书要几十秒。
            # 只有抽中且确认已翻的页才进 actually_done（无假阳性）；
            # 未抽中页走正常流程，缓存命中不重复扣费。
            stride = max(1, n_pages // 40)
            for i in range(0, n_pages, stride):
                if task.stop_event.is_set():
                    break
                checked += 1
                if page_is_translated(doc[i], target_lang, src_page=orig_check[i]):
                    actually_done.add(i + 1)
            orig_check.close()

            reported = len(done_pages)
            actual = len(actually_done)
            task.log_msg(
                f"🔎 内容校验（抽样 {checked}/{n_pages} 页）：确认已翻 {actual} 页"
            )
            if checked and actual == 0:
                task.log_msg(
                    "⚠️ 抽样页全部未翻译，将整体重翻（缓存命中不重复扣费）"
                )
            elif actual < checked:
                task.log_msg(
                    f"ℹ️ 抽样中发现 {checked - actual} 页未翻译，"
                    f"进度记录原报 {reported} 页，已按实际修正"
                )

            done_pages = actually_done
            save_progress_file(paths["progress_file"], done_pages)
        except Exception as e:
            task.log_msg(f"⚠️ 内容校验失败（不影响继续）：{e}")

    total = len(doc)
    # ★ 优化：试翻页数可配置
    limit = min(trial_pages, total) if trial else total

    task.status = "running"
    task.total = limit
    task.label = f"PDF · 目标前 {limit} 页"
    task.current = sum(1 for p in done_pages if p <= limit)
    task.log_msg(f"✅ PDF 共 {total} 页，本次目标 {limit} 页，已翻 {task.current} 页")
    save_state(task, force=True)

    if pdf_exists:
        try:
            task.preview_images = render_preview_only(paths["output_pdf"], paths, task)
            _add_output(task, paths["output_pdf"])
            refresh_task_size(task, force=True)
            save_state(task, force=True)
        except Exception as e:
            task.log_msg(f"⚠️ 初始预览生成失败：{e}")

    orig_doc_for_check = None
    try:
        orig_doc_for_check = fitz.open(task.src_path)
    except Exception as e:
        task.log_msg(f"⚠️ 无法打开原 PDF 用于校验/回滚：{e}（本次运行跳过回滚保护）")

    error_msg = None
    all_pages_ok = True

    for pno in range(total):
        if task.stop_event.is_set():
            task.log_msg("⏸ 检测到停止信号，结束当前循环")
            all_pages_ok = False
            break
        page_num = pno + 1
        if page_num > limit:
            break
        if page_num in done_pages:
            continue

        page = doc[pno]
        blocks = [b for b in page.get_text("blocks") if b[6] == 0 and b[4].strip()]

        if not blocks:
            done_pages.add(page_num)
            save_progress_file(paths["progress_file"], done_pages)
            if page_num <= limit:
                task.current += 1
            task.label = f"已跳过插图页 {page_num}"
            save_state(task)
            continue

        try:
            trans, page_terms = translate_page(
                client, model, blocks, cache,
                paths["cache_file"], target_lang,
                stop_event=task.stop_event,
                reader_profile=reader_profile,
                want_terms=enable_terms,
                domain=domain, extra_prompt=extra_prompt,
            )
            trans_complete = all(
                (trans.get(i) or "").strip() for i in range(len(blocks))
            )
            if not trans_complete or not translations_look_valid(
                    trans, blocks, target_lang):
                task.log_msg(f"⚠️ 第 {page_num} 页翻译结果不完整，跳过，稍后重试")
                all_pages_ok = False
                continue

            written, failed = apply_translations(
                page, blocks, trans,
                font_path=font_path,
                max_font_size=max_font_size,
                target_lang=target_lang,
            )
            if failed:
                task.log_msg(
                    f"⚠️ 第 {page_num} 页有 {failed}/{len(blocks)} 段写入失败"
                    f"（该页将进行字符校验，可能回滚重试）"
                )
                all_pages_ok = False

            if target_lang in config.NON_LATIN_SCRIPT_LANGS:
                src_check_page = None
                if (orig_doc_for_check is not None
                        and pno < len(orig_doc_for_check)):
                    src_check_page = orig_doc_for_check[pno]

                if not page_is_translated(page, target_lang,
                                          src_page=src_check_page):
                    try:
                        after_text = page.get_text() or ""
                    except Exception:
                        after_text = ""
                    total_ns = sum(1 for c in after_text if not c.isspace())
                    if target_lang in ("zh-CN", "zh-TW"):
                        cnt_target = sum(
                            1 for c in after_text if '一' <= c <= '鿿')
                    elif target_lang == "ja":
                        cnt_target = sum(
                            1 for c in after_text
                            if '぀' <= c <= 'ヿ')
                    elif target_lang == "ko":
                        cnt_target = sum(
                            1 for c in after_text if '가' <= c <= '힣')
                    elif target_lang == "ru":
                        cnt_target = sum(
                            1 for c in after_text if 'Ѐ' <= c <= 'ӿ')
                    elif target_lang == "ar":
                        cnt_target = sum(
                            1 for c in after_text if '؀' <= c <= 'ۿ')
                    else:
                        cnt_target = 0
                    ratio = (cnt_target / total_ns) if total_ns else 0.0
                    n_translated = sum(
                        1 for v in trans.values() if (v or "").strip())
                    task.log_msg(
                        f"⚠️ 第 {page_num} 页写入后未检出目标语言字符 "
                        f"(块={len(blocks)}, 译段={n_translated}, "
                        f"目标字符={cnt_target}, 非空白={total_ns}, "
                        f"占比={ratio:.1%}, "
                        f"阈值需≥{config.PAGE_CN_THRESHOLD}且≥{config.PAGE_CN_RATIO:.0%})，"
                        f"回滚该页"
                    )
                    if (orig_doc_for_check is not None
                            and pno < len(orig_doc_for_check)):
                        try:
                            doc.insert_pdf(
                                orig_doc_for_check,
                                from_page=pno, to_page=pno, start_at=pno,
                            )
                            doc.delete_page(pno + 1)
                            page = doc[pno]
                        except Exception as _e:
                            error_msg = (
                                f"回滚第 {page_num} 页失败，"
                                f"PDF 结构可能已损坏：{_e}"
                            )
                            task.log_msg(f"❌ {error_msg}")
                            all_pages_ok = False
                            break
                    all_pages_ok = False
                    continue

            if enable_terms and page_terms:
                ch_title = config.NB.guess_chapter_for_page(page_num, chapters)
                config.NB.merge_page_terms(global_terms, page_terms, page_num, ch_title)
                config.NB.save_terms(terms_file, global_terms)
                task.log_msg(
                    f"📝 第 {page_num} 页抽取术语 {len(page_terms)} 个"
                    f"（全局 {len(global_terms.get('terms', {}))} 个）"
                )
        except RuntimeError as e:
            error_msg = str(e)
            break
        except Exception as e:
            error_msg = f"未知错误：{e}"
            break

        done_pages.add(page_num)
        save_progress_file(paths["progress_file"], done_pages)
        if page_num <= limit:
            task.current += 1
        task.label = f"已翻 {task.current}/{limit} 页"
        task.log_msg(f"✅ 第 {page_num} 页完成")
        save_state(task, force=(page_num % 5 == 0))

        should_save = (page_num <= config.PREVIEW_PAGES) or (page_num % config.CHECKPOINT_EVERY == 0)
        if should_save:
            try:
                saved = safe_save_pdf(doc, paths["output_pdf"], garbage=1)
                if saved != paths["output_pdf"]:
                    task.log_msg(
                        f"⚠️ 原文件被占用，已保存到备用路径："
                        f"{os.path.basename(saved)}"
                    )
                    paths["output_pdf"] = saved
                task.log_msg(f"💾 已落盘（前 {page_num} 页）")
                if page_num <= config.PREVIEW_PAGES:
                    task.preview_images = render_preview_only(
                        paths["output_pdf"], paths, task)
                _add_output(task, paths["output_pdf"])
                refresh_task_size(task, force=True)
                save_state(task, force=True)
            except Exception as e:
                task.log_msg(f"⚠️ 落盘失败：{e}")

    save_failed = False
    try:
        saved = safe_save_pdf(doc, paths["output_pdf"], garbage=3)
        if saved != paths["output_pdf"]:
            task.log_msg(
                f"⚠️ 原文件被占用，已保存到备用路径："
                f"{os.path.basename(saved)}"
            )
            paths["output_pdf"] = saved
    except Exception as e:
        save_failed = True
        error_msg = error_msg or f"保存译文 PDF 失败：{e}"
    finally:
        try:
            doc.close()
        except Exception:
            pass
        try:
            if orig_doc_for_check is not None:
                orig_doc_for_check.close()
        except Exception:
            pass

    if save_failed or not os.path.exists(paths["output_pdf"]):
        task.status = "error"
        task.error = error_msg or "译文 PDF 保存失败，无法继续生成双语 / 笔记"
        task.label = "出错（保存失败）"
        task.log_msg(f"❌ {task.error}")
        refresh_task_size(task, force=True)
        save_state(task, force=True)
        return

    target_pages = set(range(1, limit + 1))
    done_in_target = target_pages & done_pages
    completed = (done_in_target == target_pages) and all_pages_ok and not error_msg

    if not done_pages and not pdf_exists:
        if task.stop_event.is_set():
            task.status = "paused"
            task.error = ""
            task.label = "已暂停（未翻译任何页）"
            task.log_msg("🛑 已暂停：尚未开始翻译")
        elif error_msg:
            task.status = "error"
            task.error = error_msg
            task.log_msg(f"❌ {error_msg}")
        else:
            task.status = "error"
            task.error = "本次没有任何页面成功翻译"
            task.log_msg("⚠️ " + task.error)
        save_state(task, force=True)
        return

    try:
        dp = {p for p in done_pages if p <= limit}
        if not dp:
            task.log_msg("ℹ️ 无已翻译页，跳过双语 PDF 生成")
        elif not task.make_bilingual:
            task.log_msg("ℹ️ 用户设置不生成双语 PDF，跳过")
            _add_output(task, paths["output_pdf"])
        elif (not regen_bilingual
                and os.path.exists(paths["bilingual_pdf"])):
            # ★ 新增：清晰度升级时用户选择不覆盖 → 保留旧双语 PDF
            task.log_msg(
                f"ℹ️ 保留旧双语 PDF（未重新生成）："
                f"{os.path.basename(paths['bilingual_pdf'])}")
            _add_output(task, paths["output_pdf"], paths["bilingual_pdf"])
        else:
            task.log_msg(
                f"🖼 生成左右对照双语 PDF（{len(dp)} 页，"
                f"质量「{task.pdf_quality}」，请稍候）……")
            make_bilingual_pdf(
                paths["output_pdf"], paths, task,
                n_pages=None, done_pages=dp,
                zoom=preset["zoom"],
                jpeg_quality=preset["jpeg"],
                garbage=preset["garbage"],
            )
            if os.path.exists(paths["bilingual_pdf"]):
                _add_output(task, paths["output_pdf"], paths["bilingual_pdf"])
            else:
                _add_output(task, paths["output_pdf"])
    except Exception as e:
        task.log_msg(f"⚠️ 收尾生成双语 PDF 失败：{e}")

    if enable_terms and global_terms.get("terms"):
        try:
            book_title = os.path.splitext(task.src_name)[0]
            lang_label = config.LANG_NAMES.get(target_lang, target_lang)

            notes_dir = paths["notes_dir"]
            os.makedirs(notes_dir, exist_ok=True)

            notes_md   = os.path.join(notes_dir, "notes.md")
            notes_html = os.path.join(notes_dir, "notes.html")
            terms_csv  = os.path.join(notes_dir, "terms.csv")

            md = config.NB.build_notes_markdown(
                book_title, global_terms,
                reader_profile=reader_profile,
                total_pages=limit, lang_label=lang_label,
            )
            with open(notes_md, "w", encoding="utf-8") as f:
                f.write(md)

            try:
                html = config.NB.build_notes_html(
                    book_title, global_terms,
                    reader_profile=reader_profile,
                    total_pages=limit, lang_label=lang_label,
                    pdf_rel_path="../translated.pdf",
                )
            except TypeError:
                html = config.NB.build_notes_html(
                    book_title, global_terms,
                    reader_profile=reader_profile,
                    total_pages=limit, lang_label=lang_label,
                )
            with open(notes_html, "w", encoding="utf-8") as f:
                f.write(html)

            config.NB.build_terms_csv(global_terms, terms_csv)

            _add_output(task, notes_md, notes_html, terms_csv)

            task.log_msg(
                f"📓 阅读笔记已生成（{len(global_terms['terms'])} 个术语）"
                f" → {os.path.basename(notes_dir)}/"
            )
        except Exception as e:
            task.log_msg(f"⚠️ 生成阅读笔记失败：{e}")

    if enable_terms and global_terms.get("terms"):
        try:
            book_title = os.path.splitext(task.src_name)[0]
            lang_label = config.LANG_NAMES.get(target_lang, target_lang)

            # ★ 保留旧双语 PDF 时不再插入术语页（旧文件上次已插过，避免重复）
            if (task.make_bilingual and regen_bilingual
                    and os.path.exists(paths["bilingual_pdf"])):
                if insert_notes_into_pdf(
                    paths["bilingual_pdf"], global_terms, book_title,
                    reader_profile, limit, lang_label,
                    font_path=font_path, position="front",
                ):
                    task.log_msg("📎 术语表已插入 bilingual.pdf 开头（带书签）")

            try:
                with open(paths["output_pdf"], "rb") as f:
                    base_data = f.read()
                base_doc = fitz.open(stream=base_data, filetype="pdf")
                trans_with_notes = paths["trans_with_notes_pdf"]
                base_doc.save(trans_with_notes, deflate=True, garbage=3)
                base_doc.close()

                if insert_notes_into_pdf(
                    trans_with_notes, global_terms, book_title,
                    reader_profile, limit, lang_label,
                    font_path=font_path, position="front",
                ):
                    task.log_msg(
                        "📎 术语表已插入 translated_with_notes.pdf 开头（带书签）")
                    _add_output(task, trans_with_notes)
            except Exception as e:
                task.log_msg(f"⚠️ 生成 translated_with_notes.pdf 失败：{e}")
        except Exception as e:
            task.log_msg(f"⚠️ 插入术语页失败：{e}")

    try:
        task.preview_images = render_preview_only(paths["output_pdf"], paths, task)
    except Exception:
        pass

    if completed:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 全部完成！共 {len(done_pages)} 页，最大页码 {max(done_pages)}")
    elif task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停（半成品），共翻 {len(done_pages)} 页"
        if done_pages:
            task.log_msg(f"🛑 已暂停。下次点开始会从第 {max(done_pages) + 1} 页继续")
        else:
            task.log_msg("🛑 已暂停。")
    elif error_msg:
        task.status = "error"
        task.error = error_msg
        task.label = "出错（半成品）"
        task.log_msg(f"❌ {error_msg}")
    else:
        task.status = "paused"
        task.label = "半成品（部分页未成功）"
    refresh_task_size(task, force=True)
    save_state(task, force=True)


def docx_worker(task, paths, real_key, model, target_lang="zh-CN",
                domain="general", extra_prompt=""):
    # ★ 修复：base_url 进 .env；设 timeout + max_retries=0（避免 SDK
    # 默认 600s 超时 × 内置重试 × 自写重试叠加，停止信号无法打断）
    client = OpenAI(api_key=real_key, base_url=config.DEEPSEEK_BASE_URL,
                    timeout=config.API_TIMEOUT, max_retries=0)
    cache = load_json_file(paths["cache_file"], {})

    est = estimate_output_size("docx", task.src_path, 0,
                               target_lang=target_lang, want_terms=False)
    task.estimated_size = est.get("total", 0)
    if task.estimated_size:
        task.log_msg(f"📦 预估总产出 ≈ {fmt_size(task.estimated_size)}")

    try:
        wdoc = Document(task.src_path)
    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        task.status = "error"
        task.error = f"打开 Word 失败：{e}"
        task.log_msg(f"❌ 打开 Word 失败：{e}")
        if "Package not found" in err_detail or "not a zip" in err_detail.lower():
            task.log_msg("   ⚠️ 这通常意味着 .docx 其实是从 .doc 改后缀来的。")
        elif "PermissionError" in err_detail:
            task.log_msg("   ⚠️ 文件被其他程序占用。")
        save_state(task, force=True)
        return

    para_targets = [p for p in wdoc.paragraphs if p.text.strip()]
    cell_targets = []
    seen = set()
    for table in wdoc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                for p in cell.paragraphs:
                    if p.text.strip():
                        cell_targets.append(p)

    all_targets = para_targets + cell_targets
    total_targets = len(all_targets)

    task.status = "running"
    task.total = total_targets
    task.current = 0
    task.label = f"Word · 共 {total_targets} 段"
    task.log_msg(f"📘 Word 已打开，共 {total_targets} 段，将批量翻译")
    save_state(task, force=True)

    all_texts = [p.text for p in all_targets]

    try:
        translated, fail_count = translate_batch_office(
            client, model, all_texts, cache, paths["cache_file"],
            target_lang, stop_event=task.stop_event, task=task,
            domain=domain, extra_prompt=extra_prompt,
        )
    except RuntimeError as e:
        task.status = "error"
        task.error = str(e)
        task.log_msg(f"❌ {e}")
        save_state(task, force=True)
        return

    preview_pairs = []
    for i, (para, tr) in enumerate(zip(all_targets, translated)):
        if task.stop_event.is_set():
            break
        if tr is None:
            # ★ 修复：翻译失败的段落不把原文写回译文
            task.current = i + 1
            task.label = f"Word 应用 {i+1}/{total_targets}"
            continue
        _docx_replace_para_text(para, tr)
        if len(preview_pairs) < config.PREVIEW_PARAS:
            preview_pairs.append((all_texts[i], tr))
            task.preview_html = _build_docx_preview_html(preview_pairs)
        task.current = i + 1
        task.label = f"Word 应用 {i+1}/{total_targets}"
        if (i + 1) % 5 == 0:
            save_state(task)

    try:
        wdoc.save(paths["office_cn"])
        task.log_msg(f"✅ 译文版已保存：{paths['office_cn']}")
    except Exception as e:
        task.status = "error"
        task.error = f"保存 Word 失败：{e}"
        task.log_msg("❌ " + task.error)
        save_state(task, force=True)
        return

    bi_ok = False
    if not task.stop_event.is_set():
        try:
            doc_bi = Document(task.src_path)
            para_bi = [p for p in doc_bi.paragraphs if p.text.strip()]
            cell_bi = []
            seen_bi = set()
            for table in doc_bi.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell._tc in seen_bi:
                            continue
                        seen_bi.add(cell._tc)
                        for p in cell.paragraphs:
                            if p.text.strip():
                                cell_bi.append(p)
            all_bi = para_bi + cell_bi
            fail_bi = 0
            for i, para in enumerate(all_bi):
                if task.stop_event.is_set():
                    break
                src = para.text
                key = f"t_{cache_prefix(target_lang)}" + h(src)
                tr = cache.get(key)
                if not tr:
                    # ★ 修复：翻译失败的段落插入失败标记，不再原文冒充译文
                    tr = "⚠ 翻译失败（无译文）"
                    fail_bi += 1
                _docx_insert_after(para, tr)
            if fail_bi:
                task.log_msg(
                    f"⚠️ 双语版有 {fail_bi} 段翻译失败，已插入失败标记"
                )
            doc_bi.save(paths["office_bi"])
            task.log_msg(f"✅ 双语版已保存：{paths['office_bi']}")
            bi_ok = True
        except Exception as e:
            task.log_msg(f"⚠️ 保存双语版失败：{e}")

    _add_output(task, paths["office_cn"])
    if bi_ok:
        _add_output(task, paths["office_bi"])

    if task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停 {task.current}/{total_targets}"
        task.log_msg(f"🛑 已暂停（半成品），共处理 {task.current}/{total_targets} 段")
    elif fail_count > 0:
        task.status = "paused"
        task.error = f"{fail_count} 段翻译失败，未替换为译文"
        task.label = f"半成品（{fail_count} 段失败）"
        task.log_msg(f"⚠️ {task.error}（可再次点击开始以重试）")
    else:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 Word 翻译完成！共 {total_targets} 段")
    refresh_task_size(task, force=True)
    save_state(task, force=True)


def pptx_worker(task, paths, real_key, model, target_lang="zh-CN",
                domain="general", extra_prompt=""):
    # ★ 修复：base_url 进 .env；设 timeout + max_retries=0（避免 SDK
    # 默认 600s 超时 × 内置重试 × 自写重试叠加，停止信号无法打断）
    client = OpenAI(api_key=real_key, base_url=config.DEEPSEEK_BASE_URL,
                    timeout=config.API_TIMEOUT, max_retries=0)
    cache = load_json_file(paths["cache_file"], {})

    est = estimate_output_size("pptx", task.src_path, 0,
                               target_lang=target_lang, want_terms=False)
    task.estimated_size = est.get("total", 0)
    if task.estimated_size:
        task.log_msg(f"📦 预估总产出 ≈ {fmt_size(task.estimated_size)}")

    try:
        prs = Presentation(task.src_path)
    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        task.status = "error"
        task.error = f"打开 PPT 失败：{e}"
        task.log_msg(f"❌ 打开 PPT 失败：{e}")
        if "Package not found" in err_detail or "not a zip" in err_detail.lower():
            task.log_msg("   ⚠️ 这通常意味着 .pptx 其实是从 .ppt 改后缀来的。")
        save_state(task, force=True)
        return

    targets = []
    texts = []
    for si, slide in enumerate(prs.slides):
        for shape in _iter_pptx_shapes(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            for para in shape.text_frame.paragraphs:
                text = "".join(r.text for r in para.runs)
                if text.strip():
                    targets.append((si, para))
                    texts.append(text)

    total_targets = len(targets)
    task.status = "running"
    task.total = total_targets
    task.current = 0
    task.label = f"PPT · 共 {total_targets} 段"
    task.log_msg(f"📊 PPT 已打开，共 {total_targets} 段")
    save_state(task, force=True)

    try:
        translated, fail_count = translate_batch_office(
            client, model, texts, cache, paths["cache_file"],
            target_lang, stop_event=task.stop_event, task=task,
            domain=domain, extra_prompt=extra_prompt,
        )
    except RuntimeError as e:
        task.status = "error"
        task.error = str(e)
        task.log_msg(f"❌ {e}")
        save_state(task, force=True)
        return

    preview_pairs = []
    for i, ((si, para), tr) in enumerate(zip(targets, translated)):
        if task.stop_event.is_set():
            break
        if tr is None:
            # ★ 修复：翻译失败的段落不把原文写回译文
            task.current = i + 1
            task.label = f"PPT {i+1}/{total_targets}（第 {si+1} 张）"
            continue
        _pptx_set_para_text(para, tr)
        if len(preview_pairs) < config.PREVIEW_PARAS:
            preview_pairs.append((si + 1, texts[i], tr))
            task.preview_html = _build_pptx_preview_html(preview_pairs)
        task.current = i + 1
        task.label = f"PPT {i+1}/{total_targets}（第 {si+1} 张）"
        if (i + 1) % 5 == 0:
            save_state(task)

    try:
        prs.save(paths["office_cn"])
        _add_output(task, paths["office_cn"])
    except Exception as e:
        task.status = "error"
        task.error = f"保存 PPT 失败：{e}"
        task.log_msg("❌ " + task.error)
        save_state(task, force=True)
        return

    if task.stop_event.is_set():
        task.status = "paused"
        task.label = f"已暂停 {task.current}/{total_targets}"
        task.log_msg(f"🛑 已暂停（半成品），共处理 {task.current}/{total_targets} 段")
    elif fail_count > 0:
        task.status = "paused"
        task.error = f"{fail_count} 段翻译失败，未替换为译文"
        task.label = f"半成品（{fail_count} 段失败）"
        task.log_msg(f"⚠️ {task.error}（可再次点击开始以重试）")
    else:
        task.status = "done"
        task.label = "全部完成"
        task.log_msg(f"🎉 PPT 翻译完成！共 {total_targets} 段")
    refresh_task_size(task, force=True)
    save_state(task, force=True)

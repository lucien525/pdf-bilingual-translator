# -*- coding: utf-8 -*-
"""术语页 PDF：生成术语笔记、插入 PDF 开头。"""

import os
import time

import pymupdf as fitz

from core.pdf_pipeline import _pdf_font_setup


# ============================================================
# 术语页 PDF
# ============================================================

def build_notes_pdf(global_terms, book_title, reader_profile="",
                    total_pages=0, lang_label="", font_path=None):
    items = list(global_terms.get("terms", {}).values())
    if not items:
        return None
    items.sort(key=lambda x: (x.get("first_page", 999999), x.get("term", "")))

    A4_W, A4_H = 595, 842
    MARGIN = 44
    BOTTOM = A4_H - MARGIN

    doc = fitz.open()
    state = {"page": None, "fn": None, "y": MARGIN}

    def new_page():
        p = doc.new_page(width=A4_W, height=A4_H)
        fn = _pdf_font_setup(p, font_path)
        state["page"] = p
        state["fn"] = fn
        state["y"] = MARGIN

    new_page()

    def write(text, size, color=(0, 0, 0), gap=4, indent=0):
        page = state["page"]
        fn = state["fn"]
        y = state["y"]

        rect = fitz.Rect(MARGIN + indent, y, A4_W - MARGIN, BOTTOM)
        rc = -1
        try:
            rc = page.insert_textbox(
                rect, text, fontname=fn, fontsize=size,
                color=color, align=0, overlay=True,
            )
        except Exception:
            rc = -1

        if rc < 0:
            new_page()
            page = state["page"]
            fn = state["fn"]
            y = MARGIN
            rect = fitz.Rect(MARGIN + indent, y, A4_W - MARGIN, BOTTOM)
            try:
                rc = page.insert_textbox(
                    rect, text, fontname=fn, fontsize=size,
                    color=color, align=0, overlay=True,
                )
            except Exception:
                rc = -1

        if rc < 0:
            try:
                page.insert_text(
                    fitz.Point(MARGIN + indent, y + size),
                    text, fontname=fn, fontsize=size,
                    color=color, overlay=True,
                )
            except Exception:
                pass
            state["y"] = y + size + gap
            return

        used = rect.height - rc
        state["y"] = y + used + gap

    write(f"《{book_title}》阅读笔记", 20, color=(0.08, 0.08, 0.08), gap=14)

    meta = f"术语 {len(items)} 个"
    if total_pages:
        meta += f"    |    全书 {total_pages} 页"
    if lang_label:
        meta += f"    |    译文 {lang_label}"
    write(meta, 9.5, color=(0.42, 0.42, 0.42), gap=4)

    if reader_profile:
        write(f"读者背景：{reader_profile}", 9.5,
              color=(0.42, 0.42, 0.42), gap=4)

    write("注：本笔记由 AI 在翻译过程中同步生成，仅供参考。"
          "关键概念请务必核对原书定义。",
          9, color=(0.65, 0.25, 0.25), gap=18)

    by_chapter = {}
    for t in items:
        ch = t.get("chapter") or "术语总表（未识别章节）"
        by_chapter.setdefault(ch, []).append(t)

    for ch, ch_items in by_chapter.items():
        write(f"§ {ch}", 14, color=(0.42, 0.36, 0.2), gap=12)

        for t in ch_items:
            term = t.get("term", "")
            trans = t.get("translation", "")
            note = t.get("note", "")
            first = t.get("first_page", "?")
            cnt = t.get("count", 1)

            write(f"· {term}  —  {trans}", 11,
                  color=(0.08, 0.08, 0.08), gap=2)
            write(f"    首次出现 p.{first} · 全书出现 {cnt} 次", 8.5,
                  color=(0.55, 0.55, 0.55), gap=4)

            if note:
                write(f"    {note}", 10, color=(0.28, 0.28, 0.28), gap=6)
            else:
                state["y"] += 2

    state["y"] += 10
    write("—— 术语表结束 · 正文自此开始 ——", 11,
          color=(0.55, 0.42, 0.2), gap=8)

    return doc


def insert_notes_into_pdf(pdf_path, global_terms, book_title,
                          reader_profile="", total_pages=0, lang_label="",
                          font_path=None, position="front"):
    if not pdf_path or not os.path.exists(pdf_path):
        return False
    if not global_terms or not global_terms.get("terms"):
        return False

    notes_doc = build_notes_pdf(
        global_terms, book_title,
        reader_profile=reader_profile,
        total_pages=total_pages,
        lang_label=lang_label,
        font_path=font_path,
    )
    if notes_doc is None or len(notes_doc) == 0:
        if notes_doc:
            notes_doc.close()
        return False

    notes_pages = len(notes_doc)
    main_doc = None
    try:
        with open(pdf_path, "rb") as f:
            data = f.read()
        main_doc = fitz.open(stream=data, filetype="pdf")

        orig_toc = []
        try:
            orig_toc = main_doc.get_toc(simple=True) or []
        except Exception:
            pass

        if position == "front":
            main_doc.insert_pdf(notes_doc,
                                from_page=0,
                                to_page=len(notes_doc) - 1,
                                start_at=0)
        else:
            main_doc.insert_pdf(notes_doc)

        try:
            new_toc = []
            if position == "front":
                new_toc.append([1, "📖 术语表", 1])
                for item in orig_toc:
                    if not isinstance(item, (list, tuple)) or len(item) < 3:
                        continue
                    level, title, page = item[0], item[1], item[2]
                    try:
                        page = int(page) + notes_pages
                    except Exception:
                        continue
                    if page < 1:
                        continue
                    new_toc.append([level, title, page])
            else:
                for item in orig_toc:
                    if isinstance(item, (list, tuple)) and len(item) >= 3:
                        new_toc.append(list(item))
                new_toc.append([1, "📖 术语表",
                                max(1, len(main_doc) - notes_pages + 1)])

            if new_toc:
                main_doc.set_toc(new_toc)
        except Exception:
            pass

        tmp = pdf_path + ".tmp.pdf"
        main_doc.save(tmp, deflate=True, garbage=3)

        last_err = None
        for attempt in range(5):
            try:
                os.replace(tmp, pdf_path)
                return True
            except (PermissionError, OSError) as e:
                last_err = e
                time.sleep(0.8 * (attempt + 1))
        try:
            os.remove(tmp)
        except Exception:
            pass
        print(f"插入术语页失败（无法替换文件）：{last_err}")
        return False
    except Exception as e:
        print(f"插入术语页失败：{e}")
        return False
    finally:
        try:
            if main_doc is not None:
                main_doc.close()
        except Exception:
            pass
        try:
            notes_doc.close()
        except Exception:
            pass

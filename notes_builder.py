# -*- coding: utf-8 -*-
"""
notes_builder.py
术语表 + 阅读笔记生成器（技术书专用）
被 bilingual_app.py 引用，也可独立使用。
"""

import os
import re
import csv
import json
import threading
from datetime import datetime
from html import escape

TERM_MARK = "[[TERMS]]"
_TERMS_IO_LOCK = threading.Lock()
_NOTE_MAX_LEN = 120
_TERM_MAX_LEN = 80

_HTML_LANG_MAP = {
    "简体中文": "zh-CN", "繁体中文": "zh-TW",
    "英语": "en", "日语": "ja", "韩语": "ko",
    "法语": "fr", "德语": "de", "西班牙语": "es",
    "葡萄牙语": "pt", "俄语": "ru", "阿拉伯语": "ar",
    "意大利语": "it",
}


def _to_html_lang(lang_label):
    if not lang_label:
        return "zh-CN"
    s = str(lang_label).strip()
    if re.fullmatch(r'[a-z]{2,3}(-[A-Za-z0-9]{2,8})?', s):
        return s
    return _HTML_LANG_MAP.get(s, "zh-CN")


def _escape_md_cell(s):
    return (s or "").replace("|", "\\|").replace("\n", " ").strip()


def _escape_md_inline(s):
    return (s or "").replace("*", "\\*").replace("_", "\\_").strip()


def build_terms_instruction(reader_profile="", enabled=True):
    if not enabled:
        return ""
    profile_line = ""
    if reader_profile and reader_profile.strip():
        profile_line = (
            "\n读者背景：" + reader_profile.strip()
            + "。对于该背景的读者难以理解的概念，请在说明中多花一句话解释。"
        )
    return (
        "\n\n【术语抽取】正文翻译完成后，追加一段术语块，格式如下：\n"
        "[[TERMS]]\n"
        "原语术语 | 译名 | 一句话说明\n"
        "原语术语 | 译名 | 一句话说明\n"
        "（每行一个术语，用竖线分隔；没有术语就只写 [[TERMS]] 四个字符）\n\n"
        "要求：\n"
        "1) 只列【本页真正出现的】专业术语，不要列举常见词；\n"
        "2) 术语列保留原文形式（英文保留英文、日文保留日文、法文保留法文），"
        "不要统一改写为英文；\n"
        "3) 译名用目标语言书写；\n"
        "4) 说明控制在一句话（不超过 60 字），只讲\"是什么\"，不要编造出处；\n"
        "5) 若不确定术语的准确含义，说明写「原文未明确定义」；\n"
        "6) 术语块必须放在所有 [[B#]] 段落之后，不要插在段落中间。"
        + profile_line
    )


def split_translation_and_terms(api_text):
    if not api_text:
        return "", ""
    idx = api_text.rfind(TERM_MARK)
    if idx < 0:
        return api_text, ""
    return api_text[:idx], api_text[idx + len(TERM_MARK):]


def parse_terms_block(terms_raw):
    if not terms_raw or not terms_raw.strip():
        return []
    out = []
    seen = set()
    for raw_line in terms_raw.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("```"):
            continue
        line = re.sub(r'^[-*•]\s*', '', line)
        line = line.replace("｜", "|")
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 2:
            continue
        term = parts[0]
        translation = parts[1]
        # ★ 优化：多竖线 note 保留原始 "|"，不合成 " | "
        note = "|".join(parts[2:]).strip() if len(parts) > 2 else ""
        if not term or not translation:
            continue
        if len(term) > _TERM_MAX_LEN or len(translation) > _TERM_MAX_LEN:
            continue
        if len(note) > _NOTE_MAX_LEN:
            note = note[:_NOTE_MAX_LEN - 1] + "…"
        key = term.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append((term, translation, note))
    return out


def load_terms(path):
    if path and os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and "terms" in data:
                return data
        except Exception:
            pass
    return {"terms": {}, "meta": {"created_at": datetime.now().isoformat()}}


def save_terms(path, data):
    if not path:
        return
    data.setdefault("meta", {})["updated_at"] = datetime.now().isoformat()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with _TERMS_IO_LOCK:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)


def merge_page_terms(global_terms, page_terms, page_num,
                     chapter_title="", note_policy="first"):
    bucket = global_terms.setdefault("terms", {})
    for term, translation, note in page_terms:
        key = term.lower().strip()
        if not key:
            continue
        new_note = (note or "").strip()
        if key in bucket:
            item = bucket[key]
            if page_num not in item["pages"]:
                item["pages"].append(page_num)
                item["count"] = len(item["pages"])
            old_note = (item.get("note") or "").strip()
            if new_note:
                if note_policy == "longest":
                    if len(new_note) > len(old_note):
                        item["note"] = new_note
                elif note_policy == "latest":
                    item["note"] = new_note
                else:
                    if not old_note:
                        item["note"] = new_note
            if (item.get("translation") and translation
                    and item["translation"] != translation):
                alts = item.setdefault("alt_translations", [])
                if translation not in alts:
                    alts.append(translation)
            if not item.get("chapter") and chapter_title:
                item["chapter"] = chapter_title
        else:
            bucket[key] = {
                "term": term,
                "translation": translation,
                "note": new_note,
                "first_page": page_num,
                "pages": [page_num],
                "count": 1,
                "chapter": chapter_title or "",
            }
    return global_terms


def guess_chapter_for_page(pno, chapters):
    if not chapters:
        return ""
    s = sorted(chapters, key=lambda x: x[1])
    current = s[0][0]
    for title, start in s:
        if pno >= start:
            current = title
        else:
            break
    return current


def detect_chapters(doc):
    try:
        toc = doc.get_toc(simple=True)
    except Exception:
        return []
    if not toc:
        return []
    by_level = {}
    for item in toc:
        if not isinstance(item, (list, tuple)) or len(item) < 3:
            continue
        level, title, page = item[0], item[1], item[2]
        try:
            level = int(level)
        except Exception:
            continue
        title = (title or "").strip()
        if not title:
            continue
        try:
            page = int(page)
        except Exception:
            continue
        if page < 1:
            continue
        by_level.setdefault(level, []).append((title, page))
    if not by_level:
        return []
    if 1 in by_level:
        chapters = by_level[1]
    else:
        min_level = min(by_level.keys())
        chapters = by_level[min_level]
    chapters.sort(key=lambda x: x[1])
    return chapters


def _sorted_terms(global_terms):
    items = list(global_terms.get("terms", {}).values())
    items.sort(key=lambda x: (x.get("first_page", 999999), x.get("term", "")))
    return items


def build_notes_markdown(book_title, global_terms, reader_profile="",
                         total_pages=0, lang_label=""):
    terms = _sorted_terms(global_terms)
    lines = []
    lines.append(f"# 《{book_title}》阅读笔记")
    lines.append("")
    lines.append(f"- 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    if lang_label:
        lines.append(f"- 译文语言：{lang_label}")
    if total_pages:
        lines.append(f"- 总页数：{total_pages}")
    if reader_profile:
        lines.append(f"- 读者背景：{reader_profile}")
    lines.append(f"- 术语数量：{len(terms)}")
    lines.append("")
    lines.append("> 注意：本笔记由 AI 在翻译过程中同步生成，"
                 "术语说明仅供参考，不能替代教材 / 原书定义。"
                 "遇到关键概念请务必核对原书。")
    lines.append("")
    if not terms:
        lines.append("_本次翻译未抽取到术语。_")
        return "\n".join(lines)
    lines.append("## 📖 术语总表（按首次出现排序）")
    lines.append("")
    lines.append("| 术语 | 译名 | 说明 | 首现页 | 出现次数 |")
    lines.append("|---|---|---|---|---|")
    for t in terms:
        lines.append(
            "| {} | {} | {} | p.{} | {} |".format(
                _escape_md_cell(t.get("term", "")),
                _escape_md_cell(t.get("translation", "")),
                _escape_md_cell(t.get("note", "")),
                t.get("first_page", "?"),
                t.get("count", 1),
            )
        )
    lines.append("")
    lines.append("## 📚 分章术语")
    lines.append("")
    by_chapter = {}
    for t in terms:
        ch = t.get("chapter") or "（未识别章节）"
        by_chapter.setdefault(ch, []).append(t)
    for ch, items in by_chapter.items():
        lines.append(f"### {_escape_md_inline(ch)}")
        lines.append("")
        for t in items:
            term = t.get("term", "")
            trans = t.get("translation", "")
            note = t.get("note", "")
            pages = t.get("pages", [])
            lines.append(f"**{_escape_md_inline(term)}** — {_escape_md_inline(trans)}")
            if note:
                lines.append(f"  - {note}")
            if pages:
                head = pages[:20]
                page_str = ", ".join(f"p.{p}" for p in head)
                if len(pages) > 20:
                    page_str += f" …（共 {len(pages)} 处）"
                lines.append(f"  - 出现位置：{page_str}")
            alts = t.get("alt_translations") or []
            if alts:
                lines.append(f"  - 其他译法：{' / '.join(alts)}")
            lines.append("")
    return "\n".join(lines)


def build_terms_csv(global_terms, path):
    if not path:
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    items = _sorted_terms(global_terms)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["term", "translation", "note",
                    "first_page", "count", "chapter", "pages",
                    "alt_translations"])
        for t in items:
            w.writerow([
                t.get("term", ""),
                t.get("translation", ""),
                t.get("note", ""),
                t.get("first_page", ""),
                t.get("count", 1),
                t.get("chapter", ""),
                ";".join(str(p) for p in t.get("pages", [])),
                " / ".join(t.get("alt_translations") or []),
            ])


def build_notes_html(book_title, global_terms, reader_profile="",
                     total_pages=0, lang_label="",
                     pdf_rel_path="../translated.pdf"):
    items = _sorted_terms(global_terms)
    html_lang = _to_html_lang(lang_label)

    table_rows = []
    for t in items:
        term = t.get("term", "")
        trans = t.get("translation", "")
        note = t.get("note", "")
        first = int(t.get("first_page", 1) or 1)
        cnt = int(t.get("count", 1) or 1)
        pages = t.get("pages", []) or []
        ch = t.get("chapter", "") or ""
        alts = t.get("alt_translations") or []

        safe_term = escape(term)
        safe_trans = escape(trans)
        safe_note = escape(note)
        safe_ch = escape(ch)

        page_str = ", ".join(f"p.{p}" for p in pages[:12])
        if len(pages) > 12:
            page_str += f" …（共 {len(pages)} 处）"

        link = f"{pdf_rel_path}#page={first}" if pdf_rel_path else "#"

        alts_html = ""
        if alts:
            alts_html = (
                f'<div class="tip-line tip-alt">'
                f'其他译法：{escape(" / ".join(alts))}</div>'
            )

        ch_html = (
            f'<div class="tip-line">📚 章节：{safe_ch}</div>' if ch else ""
        )
        note_html = (
            f'<div class="tip-note">{safe_note}</div>'
            if safe_note else
            '<div class="tip-note tip-note-empty">（暂无说明）</div>'
        )
        pages_html = (
            f'<div class="tip-line tip-pages">📍 {escape(page_str)}</div>'
            if page_str else ""
        )

        table_rows.append(f'''
        <tr>
          <td class="term-cell">
            <a class="term-link" href="{escape(link)}" target="_blank"
               rel="noopener">{safe_term}</a>
            <div class="term-tip" role="tooltip">
              <div class="tip-arrow"></div>
              <div class="tip-head">
                <span class="tip-term">{safe_term}</span>
                <span class="tip-arrow-icon">→</span>
                <span class="tip-trans">{safe_trans}</span>
              </div>
              <div class="tip-line">
                首次出现 <b>p.{first}</b> · 全书 <b>{cnt}</b> 次
              </div>
              {ch_html}
              {pages_html}
              {alts_html}
              {note_html}
              <div class="tip-hint">🔗 点击术语打开译文 PDF 第 {first} 页</div>
            </div>
          </td>
          <td class="trans-cell">{safe_trans}</td>
          <td class="note-cell">{safe_note}</td>
          <td class="page-cell">p.{first}</td>
          <td class="cnt-cell">{cnt}</td>
        </tr>''')

    rows_html = "".join(table_rows) if table_rows else (
        '<tr><td colspan="5" style="text-align:center;color:#a9a49a;'
        'padding:30px">本次翻译未抽取到术语</td></tr>'
    )

    by_chapter = {}
    for t in items:
        ch = t.get("chapter") or "（未识别章节）"
        by_chapter.setdefault(ch, []).append(t)

    chapter_blocks = []
    for ch, ch_items in by_chapter.items():
        li = []
        for t in ch_items:
            term = t.get("term", "")
            trans = t.get("translation", "")
            note = t.get("note", "")
            first = int(t.get("first_page", 1) or 1)
            cnt = int(t.get("count", 1) or 1)
            pages = t.get("pages", []) or []
            link = f"{pdf_rel_path}#page={first}" if pdf_rel_path else "#"
            page_str = ", ".join(f"p.{p}" for p in pages[:10])
            if len(pages) > 10:
                page_str += f" …(共 {len(pages)})"

            pages_line = (
                f'<div class="ch-pages">📍 {escape(page_str)}</div>'
                if page_str else ""
            )

            li.append(f'''
            <li class="ch-item">
              <a class="ch-term" href="{escape(link)}" target="_blank"
                 rel="noopener">
                <span class="ch-term-name">{escape(term)}</span>
                <span class="ch-term-trans">{escape(trans)}</span>
              </a>
              <div class="ch-meta">
                <span class="ch-badge">p.{first}</span>
                <span class="ch-badge">{cnt} 次</span>
              </div>
              {f'<div class="ch-note">{escape(note)}</div>' if note else ''}
              {pages_line}
            </li>''')
        chapter_blocks.append(f'''
        <section class="ch-section">
          <h3 class="ch-title">{escape(ch)}</h3>
          <ul class="ch-list">{''.join(li)}</ul>
        </section>''')

    chapters_html = "".join(chapter_blocks) if chapter_blocks else (
        '<div class="empty-hint">（未识别章节）</div>'
    )

    meta_bits = []
    meta_bits.append(
        f'<span class="meta-pill">📚 {len(items)} 个术语</span>')
    if total_pages:
        meta_bits.append(
            f'<span class="meta-pill">📄 全书 {total_pages} 页</span>')
    if lang_label:
        meta_bits.append(
            f'<span class="meta-pill">🌐 译文 {escape(lang_label)}</span>')
    if reader_profile:
        meta_bits.append(
            f'<span class="meta-pill">👤 {escape(reader_profile)}</span>')
    meta_html = "".join(meta_bits)

    return f"""<!DOCTYPE html>
<html lang="{html_lang}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>《{escape(book_title)}》阅读笔记</title>
<style>
  :root {{
    --ink:      #0f3d3e;
    --ink-2:    #1f5b5c;
    --gold:     #c9a961;
    --gold-2:   #a98a45;
    --bg:       #faf8f2;
    --card:     #ffffff;
    --line:     #ebe5d8;
    --muted:    #8b8578;
    --muted-2:  #a9a49a;
    --text:     #1a1a1a;
    --text-2:   #5a5a5a;
  }}
  * {{ box-sizing: border-box; }}
  html, body {{
    margin: 0; padding: 0;
    background: var(--bg);
    color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Noto Sans SC",
                 "PingFang SC", "Microsoft YaHei", "Segoe UI", sans-serif;
    font-size: 15px;
    line-height: 1.75;
  }}
  .wrap {{
    max-width: 980px;
    margin: 0 auto;
    padding: 56px 28px 80px;
  }}

  .hero {{
    background: linear-gradient(135deg, #f7f3e8 0%, #f0ebdc 100%);
    border: 1px solid var(--line);
    border-radius: 20px;
    padding: 36px 40px 30px;
    margin-bottom: 32px;
    position: relative;
    overflow: hidden;
    box-shadow: 0 4px 24px rgba(15,61,62,.06);
  }}
  .hero::before {{
    content: ""; position: absolute; top: 0; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg, var(--ink), var(--gold), var(--ink));
  }}
  .hero h1 {{
    margin: 0 0 14px;
    font-family: "Noto Serif SC", Georgia, serif;
    font-size: 30px;
    color: var(--ink);
    letter-spacing: .5px;
    font-weight: 700;
  }}
  .hero .sub {{
    color: var(--text-2);
    font-size: 13.5px;
    margin: 0 0 18px;
  }}
  .meta {{ display: flex; flex-wrap: wrap; gap: 8px; }}
  .meta-pill {{
    display: inline-block;
    background: #fff;
    border: 1px solid var(--line);
    color: var(--ink);
    padding: 5px 14px;
    border-radius: 999px;
    font-size: 12.5px;
    font-weight: 500;
  }}

  .callout {{
    background: #fff6f6;
    border-left: 4px solid #d9534f;
    border-radius: 8px;
    padding: 14px 18px;
    margin: 22px 0;
    font-size: 13.5px;
    color: #8a3a3a;
    line-height: 1.8;
  }}

  .section-title {{
    display: flex; align-items: center; gap: 10px;
    margin: 44px 0 20px;
    font-family: "Noto Serif SC", Georgia, serif;
    font-size: 20px;
    color: var(--ink);
    font-weight: 700;
    letter-spacing: .3px;
  }}
  .section-title::after {{
    content: ""; flex: 1; height: 1px;
    background: linear-gradient(90deg, var(--gold), transparent);
  }}

  .table-wrap {{
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 16px;
    overflow: visible;
    box-shadow: 0 2px 14px rgba(15,61,62,.05);
  }}
  table.term-table {{
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    font-size: 13.5px;
  }}
  table.term-table thead th {{
    background: #f5f1e6;
    color: var(--ink);
    font-weight: 600;
    text-align: left;
    padding: 12px 16px;
    font-size: 12px;
    letter-spacing: .8px;
    text-transform: uppercase;
    border-bottom: 1px solid var(--line);
  }}
  table.term-table thead th:first-child {{ border-top-left-radius: 16px; }}
  table.term-table thead th:last-child  {{ border-top-right-radius: 16px; }}
  table.term-table tbody td {{
    padding: 13px 16px;
    border-bottom: 1px solid var(--line);
    vertical-align: top;
  }}
  table.term-table tbody tr:last-child td {{ border-bottom: none; }}
  table.term-table tbody tr:hover {{ background: #fbf9f3; }}

  .term-cell {{
    position: relative;
    min-width: 180px;
    white-space: nowrap;
  }}
  .term-link {{
    color: var(--ink);
    text-decoration: none;
    font-weight: 600;
    border-bottom: 1px dashed var(--gold);
    padding-bottom: 1px;
    transition: color .15s, border-color .15s;
    cursor: pointer;
  }}
  .term-link:hover {{
    color: var(--gold-2);
    border-bottom-style: solid;
    border-bottom-color: var(--gold-2);
  }}

  .term-tip {{
    visibility: hidden;
    opacity: 0;
    pointer-events: none;
    position: absolute;
    left: 8px;
    top: calc(100% + 10px);
    z-index: 999;
    min-width: 320px;
    max-width: 460px;
    background: var(--ink);
    color: #f5f1e6;
    padding: 16px 20px;
    border-radius: 14px;
    box-shadow: 0 16px 44px rgba(15,61,62,.35),
                0 4px 12px rgba(0,0,0,.18);
    font-size: 12.5px;
    line-height: 1.75;
    white-space: normal;
    transition: opacity .16s ease, visibility .16s, transform .18s ease;
    transform: translateY(-4px);
  }}
  /* ★ 优化：右侧列的浮层自动翻转，避免溢出屏幕 */
  .term-cell:nth-last-child(-n+2) .term-tip {{
    left: auto;
    right: 8px;
  }}
  .term-cell:nth-last-child(-n+2) .tip-arrow {{
    left: auto;
    right: 22px;
  }}

  .term-cell:hover .term-tip,
  .term-cell:focus-within .term-tip {{
    visibility: visible;
    opacity: 1;
    pointer-events: auto;
    transform: translateY(0);
  }}
  .tip-arrow {{
    position: absolute;
    top: -7px; left: 22px;
    width: 14px; height: 14px;
    background: var(--ink);
    transform: rotate(45deg);
    border-radius: 2px;
  }}
  .tip-head {{
    display: flex; align-items: center; flex-wrap: wrap;
    gap: 8px;
    font-size: 14.5px;
    padding-bottom: 10px;
    margin-bottom: 10px;
    border-bottom: 1px solid rgba(201,169,97,.35);
  }}
  .tip-term {{ color: #f5f1e6; font-weight: 700; }}
  .tip-arrow-icon {{ color: var(--gold); font-weight: 400; }}
  .tip-trans {{ color: var(--gold); font-weight: 600; }}
  .tip-line {{
    color: #d6cfc0;
    font-size: 12px;
    line-height: 1.85;
  }}
  .tip-line b {{ color: var(--gold); font-weight: 600; }}
  .tip-alt {{ color: #bfb6a3; font-style: italic; }}
  .tip-pages {{ color: #bfb6a3; word-break: break-word; }}
  .tip-note {{
    margin-top: 10px;
    padding-top: 10px;
    border-top: 1px dashed rgba(201,169,97,.28);
    color: #ece5d8;
    font-size: 12.5px;
    line-height: 1.85;
    font-style: italic;
  }}
  .tip-note-empty {{ color: #8b8578; }}
  .tip-hint {{
    margin-top: 12px;
    padding-top: 10px;
    border-top: 1px solid rgba(201,169,97,.2);
    color: var(--gold);
    font-size: 11.5px;
    letter-spacing: .3px;
  }}

  .trans-cell {{ color: var(--ink-2); font-weight: 500; }}
  .note-cell  {{ color: var(--text-2); font-size: 13px; }}
  .page-cell, .cnt-cell {{
    font-family: "SF Mono", "Consolas", monospace;
    color: var(--muted);
    font-size: 12.5px;
    white-space: nowrap;
  }}

  .ch-section {{
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 22px 26px 18px;
    margin-bottom: 18px;
    box-shadow: 0 2px 10px rgba(15,61,62,.04);
  }}
  .ch-title {{
    margin: 0 0 16px;
    font-family: "Noto Serif SC", Georgia, serif;
    font-size: 16px;
    color: var(--ink);
    font-weight: 700;
    padding-bottom: 10px;
    border-bottom: 1px solid var(--line);
    letter-spacing: .3px;
  }}
  .ch-list {{ list-style: none; padding: 0; margin: 0; }}
  .ch-item {{
    padding: 12px 0;
    border-bottom: 1px solid #f2ede0;
  }}
  .ch-item:last-child {{ border-bottom: none; }}
  .ch-term {{
    display: flex; align-items: baseline; gap: 10px;
    text-decoration: none;
    flex-wrap: wrap;
    transition: color .15s;
  }}
  .ch-term-name {{
    color: var(--ink);
    font-weight: 600;
    font-size: 14.5px;
    border-bottom: 1px dashed transparent;
    transition: border-color .15s;
  }}
  .ch-term:hover .ch-term-name {{
    border-bottom-color: var(--gold);
  }}
  .ch-term-trans {{
    color: var(--gold-2);
    font-size: 13.5px;
    font-weight: 500;
  }}
  .ch-meta {{ margin-top: 4px; display: flex; gap: 6px; flex-wrap: wrap; }}
  .ch-badge {{
    background: #f5f1e6;
    color: var(--ink);
    padding: 2px 10px;
    border-radius: 999px;
    font-size: 11.5px;
    font-family: "SF Mono", "Consolas", monospace;
    font-weight: 500;
  }}
  .ch-note {{
    color: var(--text-2);
    font-size: 13px;
    margin-top: 6px;
    line-height: 1.75;
  }}
  .ch-pages {{
    color: var(--muted-2);
    font-size: 11.5px;
    margin-top: 5px;
    font-family: "SF Mono", "Consolas", monospace;
    word-break: break-word;
  }}

  .empty-hint {{
    text-align: center; color: var(--muted-2);
    padding: 40px 20px; font-size: 13px;
  }}

  .footer {{
    margin-top: 60px;
    padding-top: 24px;
    border-top: 1px solid var(--line);
    text-align: center;
    color: var(--muted-2);
    font-size: 12px;
    line-height: 1.9;
  }}
  .footer .dot {{ color: var(--gold); margin: 0 6px; }}

  @media (max-width: 720px) {{
    .wrap {{ padding: 30px 16px 60px; }}
    .hero {{ padding: 24px 22px 22px; }}
    .hero h1 {{ font-size: 22px; }}
    .term-tip {{
      min-width: 260px; max-width: 90vw;
      left: 0;
    }}
    table.term-table thead th,
    table.term-table tbody td {{ padding: 10px 12px; }}
  }}
</style>
</head>
<body>
<div class="wrap">

  <div class="hero">
    <h1>《{escape(book_title)}》</h1>
    <div class="sub">阅读笔记 · 术语速查</div>
    <div class="meta">{meta_html}</div>
  </div>

  <div class="callout">
    <b>使用提示</b>　鼠标悬停术语 → 弹出术语卡片（译名 / 首现页 / 出现次数 / 说明）；
    点击术语 → 打开译文 PDF 的对应页（浏览器原生 PDF 阅读器支持）。
    <br>
    <b>注意</b>　本笔记由 AI 在翻译过程中同步生成，术语说明仅供参考，
    不能替代原书定义，关键概念请务必核对原文。
  </div>

  <h2 class="section-title">📖 术语总表</h2>
  <div class="table-wrap">
    <table class="term-table">
      <thead>
        <tr>
          <th>术语</th>
          <th>译名</th>
          <th>说明</th>
          <th>首现页</th>
          <th>次数</th>
        </tr>
      </thead>
      <tbody>
        {rows_html}
      </tbody>
    </table>
  </div>

  <h2 class="section-title">📚 分章术语</h2>
  {chapters_html}

  <div class="footer">
    由 PDF / Word / PPT 翻译器自动生成
    <span class="dot">·</span>
    {datetime.now().strftime('%Y-%m-%d %H:%M')}
  </div>

</div>
</body>
</html>
"""
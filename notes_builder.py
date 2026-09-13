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
        note = " | ".join(p for p in parts[2:] if p).strip() if len(parts) > 2 else ""
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


def _md_to_simple_html(md):
    html_lines = []
    in_table = False
    for line in md.splitlines():
        s = line.rstrip()
        if s.startswith("### "):
            if in_table:
                html_lines.append("</table>")
                in_table = False
            html_lines.append(f"<h3>{escape(s[4:])}</h3>")
        elif s.startswith("## "):
            if in_table:
                html_lines.append("</table>")
                in_table = False
            html_lines.append(f"<h2>{escape(s[3:])}</h2>")
        elif s.startswith("# "):
            if in_table:
                html_lines.append("</table>")
                in_table = False
            html_lines.append(f"<h1>{escape(s[2:])}</h1>")
        elif s.startswith("|"):
            if not in_table:
                html_lines.append(
                    '<table style="border-collapse:collapse;width:100%;'
                    'margin:12px 0;background:#fff">'
                )
                in_table = True
            cells = [c.strip() for c in s.strip("|").split("|")]
            if all(re.fullmatch(r':?-{3,}:?', c) for c in cells if c):
                continue
            row = "".join(
                '<td style="border:1px solid #ddd;padding:6px 10px;'
                'font-size:13.5px;vertical-align:top">'
                + escape(c) + "</td>"
                for c in cells
            )
            html_lines.append(f"<tr>{row}</tr>")
        else:
            if in_table:
                html_lines.append("</table>")
                in_table = False
            if s.strip():
                html_lines.append(f"<p>{escape(s)}</p>")
            else:
                html_lines.append("")
    if in_table:
        html_lines.append("</table>")
    return "\n".join(html_lines)


def build_notes_html(book_title, global_terms, reader_profile="",
                     total_pages=0, lang_label=""):
    md = build_notes_markdown(
        book_title, global_terms, reader_profile, total_pages, lang_label
    )
    body = _md_to_simple_html(md)
    html_lang = _to_html_lang(lang_label)
    return f"""<!DOCTYPE html>
<html lang="{html_lang}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>《{escape(book_title)}》阅读笔记</title>
<style>
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Noto Sans SC",
                 "PingFang SC", "Microsoft YaHei", sans-serif;
    max-width: 880px; margin: 40px auto; padding: 0 24px;
    line-height: 1.85; color: #222; background: #faf8f2;
  }}
  h1 {{
    font-family: "Noto Serif SC", Georgia, serif;
    border-bottom: 2px solid #b09b63; padding-bottom: 8px;
    color: #111;
  }}
  h2 {{ color: #8a7a4f; margin-top: 36px; }}
  h3 {{ color: #444; margin-top: 24px; }}
  table {{ background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,.06); }}
  tr:nth-child(even) td {{ background: #fafaf7; }}
  p {{ margin: 6px 0; }}
  code {{ background: #ebe5d5; padding: 1px 6px; border-radius: 3px; }}
  a {{ color: #8a7a4f; }}
</style>
</head>
<body>
{body}
</body>
</html>
"""
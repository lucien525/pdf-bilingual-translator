# -*- coding: utf-8 -*-
import os
import pymupdf as fitz

# ★ 优化：优先读 .env 里的 TRANSLATE_FONT_PATH，没有再用默认路径
try:
    from dotenv import load_dotenv
    load_dotenv()
    FONT_PATH = os.getenv("TRANSLATE_FONT_PATH", "").strip()
except Exception:
    FONT_PATH = ""

if not FONT_PATH:
    FONT_PATH = r"D:\file\translate\word_type\09_SourceHanSerifSC\OTF\SimplifiedChinese\SourceHanSerifSC-Regular.otf"

print("=" * 50)
print("1. 字体文件是否存在：", os.path.exists(FONT_PATH))
print("   使用的字体路径：", FONT_PATH)
if os.path.exists(FONT_PATH):
    print("   文件大小：", os.path.getsize(FONT_PATH), "字节")
else:
    print("   ❌ 路径不对，请检查：")
    d = os.path.dirname(FONT_PATH)
    if os.path.exists(d):
        for f in os.listdir(d):
            print("      文件夹里实际是：", f)
    else:
        print("      连文件夹都不存在：", d)
    print()
    print("   提示：可以在 .env 里设置 TRANSLATE_FONT_PATH 指向正确字体")

print("=" * 50)
print("2. 测试字体嵌入……")
doc = fitz.open()
page = doc.new_page(width=400, height=200)

try:
    page.insert_font(fontname="cn", fontfile=FONT_PATH, set_simple=False)
    print("   insert_font 成功")
except Exception as e:
    print("   ❌ insert_font 失败：", e)

rc = page.insert_textbox(
    fitz.Rect(20, 20, 380, 180),
    "测试中文：计算机网络 自顶向下方法 第九版",
    fontname="cn",
    fontsize=16,
    color=(0, 0, 0),
)
print("   insert_textbox 返回码：", rc, "（正数=成功，负数=放不下）")

doc.save("font_test.pdf")
doc.close()
print("=" * 50)
print("3. 已保存 font_test.pdf，双击打开看看中文是否显示")
print("   绝对路径：", os.path.abspath("font_test.pdf"))
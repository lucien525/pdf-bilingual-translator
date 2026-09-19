# -*- coding: utf-8 -*-
"""
PDF / Word / PPT 翻译器（全功能优化版）主程序入口
- 领域 Prompt 微调（通用/技术/文学/新闻/法律/医学）
- 自定义 Prompt 追加
- 试翻页数可选（3/5/10/20）
- 任务列表可按状态筛选
- 日志带时间戳
- 端口 7860~7869，与 homework 7870~7899 完全分开
"""

import sys
sys.dont_write_bytecode = True

import os
import socket

from core import config
from core import fonts
from core.tasks import MANAGER, cleanup_orphan_files, scan_all_states
from ui import build_ui, UI_CSS, UI_THEME


# ============================================================
# main
# ============================================================

def _find_free_port(start=7860, end=7869):
    for p in range(start, end + 1):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                s.bind(("127.0.0.1", p))
                return p
        except OSError:
            continue
    return start


if __name__ == "__main__":
    PORT = _find_free_port(7860, 7869)
    URL = "http://127.0.0.1:" + str(PORT)

    try:
        with open("访问网址.txt", "w", encoding="utf-8") as f:
            f.write("PDF / Word / PPT 翻译器\n\n")
            f.write("浏览器访问网址：\n")
            f.write(URL + "\n\n")
            f.write("（这个文件由程序自动生成，改动此文件无效）\n")
            f.write("（如果打不开，说明程序已停止，请双击 重启翻译器.bat）\n")
    except Exception:
        pass

    print()
    print("=" * 64)
    print("  PDF / Word / PPT 翻译器 已启动")
    print()
    print("  在浏览器输入以下网址进入：")
    print()
    print("      " + URL)
    print()
    print("  别关这个终端窗口，否则程序停止")
    print("  网址也保存在 访问网址.txt 中，随时可查")
    print("=" * 64)
    print()

    print(f"   结果目录：{os.path.abspath(config.RESULT_ROOT)}")
    if config.DEFAULT_API_KEY:
        print(f"   API Key：已从 .env 读取（{config.DEFAULT_API_KEY[:6]}...）")
    else:
        print("   API Key：未配置，需在网页填写")

    print(f"   字体扫描目录：")
    for _root in config.FONT_ROOTS:
        exists = "v" if os.path.isdir(_root) else "x"
        print(f"      [{exists}] {_root}")
    print(f"   已找到字体：{len(fonts.FONTS_MAP)} 个")
    if fonts.FONTS_MAP:
        _default = fonts._default_font_display()
        print(f"   默认字体：{_default}")

    print(f"   双语 PDF 质量预设：{len(config.PDF_QUALITY_PRESETS)} 档")
    for _name, _cfg in config.PDF_QUALITY_PRESETS.items():
        print(f"      · {_name}（zoom={_cfg['zoom']}, JPEG={_cfg['jpeg']}）")

    print(f"   领域数量：{len(config.DOMAIN_CHOICES)}（支持自定义 Prompt）")
    print(f"   试翻页数：{', '.join(str(v) for _, v in config.TRIAL_PAGE_CHOICES)}")

    if not config.HAS_OFFICE:
        print("   [!] 未装 python-docx / python-pptx，Word / PPT 不可用")
    if not config.HAS_NOTES:
        print("   [!] 未找到 notes_builder.py，术语表功能不可用")
    else:
        print("   术语表模块：notes_builder.py 已加载")
    print()

    n_removed = cleanup_orphan_files()
    if n_removed:
        print(f"   已清理 {n_removed} 个冗余/临时文件")
    scan_all_states()
    n = len(MANAGER.all_sorted())
    if n:
        print(f"   已恢复 {n} 个历史任务")
    print()

    demo = build_ui()

    demo.queue(default_concurrency_limit=None)
    demo.launch(
        server_name="127.0.0.1",
        server_port=PORT,
        inbrowser=True,
        show_error=False,
        quiet=True,
        css=UI_CSS,
        theme=UI_THEME,
        allowed_paths=[os.path.abspath(config.RESULT_ROOT)],
    )

# -*- coding: utf-8 -*-
"""
PDF / Word / PPT 作业解题器（全功能优化版）主程序入口
- 答案可贴：块内右下 / 块下方 / 块右侧
- 解析贴在题目下方（浅蓝底）
- 数学公式用 matplotlib mathtext 渲染成 PNG 贴入（dpi=300）
- 支持学科微调 Prompt + 自定义补充要求
- 支持「仅解析模式」（不改原 PDF，输出 analysis.md）
- 任务列表可按状态筛选
- 断点续传 / 试解 5 页 / 每页落盘
- 独立字体变量 HOMEWORK_FONT_PATH
"""

import sys
sys.dont_write_bytecode = True

import os
import socket

import hw_ui
import hw_core
import hw_tasks
import hw_solve

# ================= 薄壳 re-export =================
# test_pure_functions.py 通过 `import homework_app as hw` 使用以下名字
safe_dirname = hw_core.safe_dirname
parse_solution_response = hw_solve.parse_solution_response

DEFAULT_API_KEY = hw_core.DEFAULT_API_KEY
RESULT_ROOT = hw_core.RESULT_ROOT
HAS_OFFICE = hw_core.HAS_OFFICE
SUBJECT_CHOICES = hw_core.SUBJECT_CHOICES
_RENDER_DPI = hw_core._RENDER_DPI
CN_FONT_PATH = hw_core.CN_FONT_PATH
MANAGER = hw_tasks.MANAGER
scan_all_states = hw_tasks.scan_all_states
cleanup_orphan_files = hw_tasks.cleanup_orphan_files


# ============================================================
# main
# ============================================================
def _find_free_port(start=7870, end=7899):
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
    PORT = _find_free_port(7870, 7899)
    URL = "http://127.0.0.1:" + str(PORT)

    try:
        with open("访问网址_解题器.txt", "w", encoding="utf-8") as f:
            f.write("PDF / Word / PPT 作业解题器\n\n")
            f.write("浏览器访问网址：\n" + URL + "\n\n")
            f.write("（如果打不开，说明程序已停止，请双击 启动解题器.bat）\n")
    except Exception:
        pass

    print()
    print("=" * 64)
    print("  PDF / Word / PPT 作业解题器 已启动")
    print()
    print("  在浏览器输入以下网址进入：")
    print()
    print("      " + URL)
    print()
    print("  别关这个终端窗口，否则程序停止")
    print("=" * 64)
    print()

    print(f"   结果目录：{os.path.abspath(RESULT_ROOT)}")
    if DEFAULT_API_KEY:
        print(f"   API Key：已从 .env 读取（{DEFAULT_API_KEY[:6]}...）")
    else:
        print("   API Key：未配置，需在网页填写")
    print(f"   中文字体：{CN_FONT_PATH or '（未找到，中文可能显示为方框）'}")
    print(f"   公式渲染 dpi：{_RENDER_DPI}")
    print(f"   学科数量：{len(SUBJECT_CHOICES)}（支持自定义 Prompt）")
    if not HAS_OFFICE:
        print("   ⚠️ 未装 python-docx / python-pptx，Word / PPT 不可用")
    print()

    scan_all_states()
    n = len(MANAGER.all_sorted())
    if n:
        print(f"   已恢复 {n} 个历史任务")
        print()

    removed = cleanup_orphan_files()
    if removed:
        print(f"   清理遗留临时文件 {removed} 个")

    demo = hw_ui.build_ui()

    demo.queue(default_concurrency_limit=None)
    demo.launch(
        server_name="127.0.0.1",
        server_port=PORT,
        inbrowser=True,
        show_error=False,
        quiet=True,
        css=hw_ui.UI_CSS,
        theme=hw_ui.UI_THEME,
        allowed_paths=[os.path.abspath(RESULT_ROOT)],
    )

# -*- coding: utf-8 -*-
"""纯函数测试：main + homework_app（import 不会启动 Gradio）。

运行：python -m unittest test_pure_functions -v
"""
import os
import sys
import unittest

sys.dont_write_bytecode = True
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import pymupdf as fitz  # noqa: E402

from core import utils as bi          # safe_dirname
from core import pdf_pipeline as bip  # parse_marked / page_is_translated
import homework_app as hw  # noqa: E402


class TestSafeDirname(unittest.TestCase):
    def test_illegal_chars_replaced(self):
        self.assertEqual(bi.safe_dirname('a<b>c:d"e/f\\g|h?i*j'),
                         "a_b_c_d_e_f_g_h_i_j")

    def test_dotdot(self):
        self.assertEqual(bi.safe_dirname("..secret"), "_secret")

    def test_trailing_dots_spaces(self):
        # 注：先 replace("..", "_") 再 rstrip(". ") 的现有行为
        self.assertEqual(bi.safe_dirname("name... "), "name_")

    def test_long_truncated(self):
        self.assertEqual(len(bi.safe_dirname("x" * 200)), 80)

    def test_empty(self):
        self.assertEqual(bi.safe_dirname(""), "untitled")

    def test_same_in_both_apps(self):
        self.assertEqual(bi.safe_dirname('a/b:c'), hw.safe_dirname('a/b:c'))


class TestParseMarked(unittest.TestCase):
    def test_normal(self):
        self.assertEqual(bip.parse_marked("[[B0]] 你好\n\n[[B1]] world"),
                         {0: "你好", 1: "world"})

    def test_out_of_order(self):
        self.assertEqual(bip.parse_marked("[[B2]] x\n[[B0]] y"),
                         {2: "x", 0: "y"})

    def test_duplicate_keeps_first(self):
        self.assertEqual(bip.parse_marked("[[B0]] a\n[[B0]] b"), {0: "a"})

    def test_empty(self):
        self.assertEqual(bip.parse_marked(""), {})

    def test_no_markers(self):
        self.assertEqual(bip.parse_marked("plain text"), {})


class TestParseSolutionResponse(unittest.TestCase):
    def test_tags(self):
        r = hw.parse_solution_response(
            "[[B0]] <ANS>42</ANS><SOL>因为……</SOL>")
        self.assertEqual(r[0]["ans"], "42")
        self.assertEqual(r[0]["sol"], "因为……")

    def test_no_tags_fallback(self):
        r = hw.parse_solution_response("[[B0]] 直接给解析", 1)
        self.assertEqual(r[0]["ans"], "")
        self.assertEqual(r[0]["sol"], "直接给解析")

    def test_empty(self):
        self.assertEqual(hw.parse_solution_response(""), {})

    def test_multiple_blocks(self):
        r = hw.parse_solution_response(
            "[[B0]] <ANS>a0</ANS><SOL>s0</SOL>\n\n"
            "[[B1]] <ANS>a1</ANS><SOL>s1</SOL>")
        self.assertEqual(r[0]["ans"], "a0")
        self.assertEqual(r[0]["sol"], "s0")
        self.assertEqual(r[1]["ans"], "a1")
        self.assertEqual(r[1]["sol"], "s1")


class TestPageIsTranslated(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.doc = fitz.open()
        cls.p_cn = cls.doc.new_page()
        rc = cls.p_cn.insert_textbox(
            fitz.Rect(50, 50, 500, 200),
            "这是中文内容测试段落。这是中文内容测试段落。"
            "这是中文内容测试段落。这是中文内容测试段落。",
            fontname="china-s", fontsize=14)
        assert rc > 0, "中文插入失败"
        cls.p_en = cls.doc.new_page()
        cls.p_en.insert_textbox(
            fitz.Rect(50, 50, 500, 200),
            "This is an English test paragraph with enough text "
            "to pass the threshold check.",
            fontname="helv", fontsize=14)
        cls.p_blank = cls.doc.new_page()
        # PyMuPDF：新建页面后旧 Page 引用会失效，需从文档重新取
        cls.p_cn = cls.doc[0]
        cls.p_en = cls.doc[1]
        cls.p_blank = cls.doc[2]

    @classmethod
    def tearDownClass(cls):
        cls.doc.close()

    def test_chinese_page_is_translated(self):
        self.assertTrue(bip.page_is_translated(self.p_cn, "zh-CN",
                                              src_page=self.p_en))

    def test_english_page_not_translated_to_zh(self):
        self.assertFalse(bip.page_is_translated(self.p_en, "zh-CN",
                                               src_page=self.p_cn))

    def test_blank_vs_blank(self):
        self.assertTrue(bip.page_is_translated(self.p_blank, "zh-CN",
                                              src_page=self.p_blank))

    def test_latin_target_always_true(self):
        self.assertTrue(bip.page_is_translated(self.p_en, "en",
                                              src_page=self.p_cn))

    def test_no_src_page_true(self):
        self.assertTrue(bip.page_is_translated(self.p_en, "zh-CN"))


if __name__ == "__main__":
    unittest.main(verbosity=2)

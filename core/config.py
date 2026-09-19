# -*- coding: utf-8 -*-
"""全局配置：环境变量、路径、质量预设、领域词表、可选依赖。"""

import os
from pathlib import Path

from dotenv import load_dotenv

# ================= 可选依赖：Word / PPT =================
# ★ 这些导入用于探测依赖是否安装（HAS_OFFICE），名字由 core/office.py 自行导入
try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    from pptx import Presentation
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
    except Exception:
        MSO_SHAPE_TYPE = None
    HAS_OFFICE = True
except ImportError:
    Document = qn = OxmlElement = Paragraph = Presentation = None
    MSO_SHAPE_TYPE = None
    HAS_OFFICE = False

# 标记为已使用（仅探测，实际使用见 core/office.py）
_ = (Document, qn, OxmlElement, Paragraph, Presentation, MSO_SHAPE_TYPE)

try:
    import notes_builder as NB
    HAS_NOTES = True
except ImportError:
    NB = None
    HAS_NOTES = False

# ================= .env =================
load_dotenv()
DEFAULT_API_KEY = os.getenv("DEEPSEEK_API_KEY", "").strip()
DEFAULT_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat").strip()
if DEFAULT_MODEL not in ("deepseek-chat", "deepseek-reasoner"):
    DEFAULT_MODEL = "deepseek-chat"
# ★ 修复：base_url / 超时进 .env，不再硬编码
DEEPSEEK_BASE_URL = os.getenv(
    "DEEPSEEK_BASE_URL", "https://api.deepseek.com"
).rstrip("/")
try:
    API_TIMEOUT = float(os.getenv("DEEPSEEK_TIMEOUT", "180"))
except Exception:
    API_TIMEOUT = 180.0

# ================= 字体扫描 =================
# ★ 本文件在 core/ 子目录下，项目根目录取上一级
_HERE = str(Path(__file__).resolve().parent.parent)

# ★ 修复：env 字体路径的目录也加入扫描根，换机器不用改代码
_ENV_FONT = os.getenv("TRANSLATE_FONT_PATH", "").strip()

FONT_ROOTS = [
    r"D:\file\translate\word_type",
    os.path.join(_HERE, "fonts"),
]
if _ENV_FONT and os.path.isdir(os.path.dirname(_ENV_FONT)):
    FONT_ROOTS.insert(0, os.path.dirname(_ENV_FONT))

_DEFAULT_FONT_CANDIDATES = [
    _ENV_FONT,
    r"D:\file\translate\word_type\09_SourceHanSerifSC\OTF\SimplifiedChinese\SourceHanSerifSC-Regular.otf",
    os.path.join(_HERE, "fonts", "SourceHanSerifSC-Regular.otf"),
]
FONT_PATH = next((p for p in _DEFAULT_FONT_CANDIDATES
                  if p and os.path.exists(p)), "")

# ================= 字号 =================
FONT_SIZE_CHOICES = [
    ("很小 · 8pt", 8.0),
    ("小 · 9pt", 9.0),
    ("中 · 10pt", 10.0),
    ("标准 · 11pt（默认）", 11.0),
    ("大 · 12pt", 12.0),
    ("很大 · 14pt", 14.0),
    ("特大 · 16pt", 16.0),
]
DEFAULT_FONT_SIZE = 11.0

# ================= 双语 PDF 质量预设 =================
PDF_QUALITY_PRESETS = {
    "省流 · 小体积（~1/3）": {
        "zoom": 1.0, "jpeg": 60, "garbage": 3,
        "hint": "文字仍清晰，适合自己看 · 300 页 ≈ 60 MB",
    },
    "标准 · 推荐": {
        "zoom": 1.4, "jpeg": 72, "garbage": 3,
        "hint": "清晰度与体积均衡（默认）· 300 页 ≈ 180 MB",
    },
    "清晰": {
        "zoom": 1.8, "jpeg": 85, "garbage": 2,
        "hint": "文字边缘锐利，适合放平板 · 300 页 ≈ 330 MB",
    },
    "印刷级 · 大体积": {
        "zoom": 2.5, "jpeg": 92, "garbage": 1,
        "hint": "接近无损，可打印 · 300 页 ≈ 720 MB",
    },
}
DEFAULT_PDF_QUALITY = "标准 · 推荐"

# ★ 预设按清晰度从低到高排列，顺序即等级（用于升级清晰度检测）
PDF_QUALITY_ORDER = list(PDF_QUALITY_PRESETS.keys())


def quality_rank(name):
    return PDF_QUALITY_ORDER.index(name) if name in PDF_QUALITY_ORDER else -1

# ================= 其他配置 =================
RESULT_ROOT = "trans_result"

RENDER_ZOOM = 2.0
BILINGUAL_ZOOM = 1.4
BILINGUAL_JPEG_QUALITY = 72

PREVIEW_PAGES = 5
PREVIEW_MAX_WIDTH = 1400
PREVIEW_JPEG_QUALITY = 88
PREVIEW_PARAS = 5

CHECKPOINT_EVERY = 10
PAGE_CN_THRESHOLD = 20
PAGE_CN_RATIO = 0.2
VALID_RATIO_THRESHOLD = 0.5

OFFICE_BATCH_SIZE = 20
STATE_SAVE_INTERVAL = 1.5
SIZE_REFRESH_INTERVAL = 5.0

# ★ 修复：双语 PDF 每 N 页检查点落盘并重开，避免整本攒内存
BILINGUAL_FLUSH_EVERY = 20

# ★ 新增：任务列表筛选
TASK_FILTER_CHOICES = [
    ("全部", "all"),
    ("运行中", "running"),
    ("已完成", "done"),
    ("出错", "error"),
]

# ★ 新增：试翻页数可选
TRIAL_PAGE_CHOICES = [
    ("3 页", 3),
    ("5 页（默认）", 5),
    ("10 页", 10),
    ("20 页", 20),
]

# ★ 新增：领域 Prompt 微调
DOMAIN_CHOICES = [
    ("通用（默认）", "general"),
    ("技术 / 计算机", "tech"),
    ("文学 / 小说", "literature"),
    ("新闻 / 报道", "news"),
    ("法律 / 合同", "law"),
    ("医学 / 生物", "medical"),
]

_DOMAIN_PROMPTS = {
    "general": "",
    "tech": (
        "\n【技术文本特别注意】\n"
        "- 术语首次出现时可在括号内保留原文；\n"
        "- 代码、命令、变量名保持原样不译；\n"
        "- 准确性优先于文采。\n"
    ),
    "literature": (
        "\n【文学文本特别注意】\n"
        "- 优先保留原文节奏与韵味；\n"
        "- 比喻、双关尽量在目标语言中找到对等表达；\n"
        "- 对话符合人物身份，不要书面化。\n"
    ),
    "news": (
        "\n【新闻文本特别注意】\n"
        "- 保持客观中立的新闻语气；\n"
        "- 数字、日期、机构名准确；\n"
        "- 不要添加主观评论。\n"
    ),
    "law": (
        "\n【法律文本特别注意】\n"
        "- 术语必须精确，不要意译；\n"
        "- 句式严格对应原文结构；\n"
        "- 不确定的术语保留原文并加括号说明。\n"
    ),
    "medical": (
        "\n【医学文本特别注意】\n"
        "- 医学术语使用标准译名；\n"
        "- 药物名、剂量单位必须准确；\n"
        "- 不确定的保留原文。\n"
    ),
}

# ★ 修复：领域角色设定（避免"文学翻译家"与"术语必须精确"自相矛盾）
_DOMAIN_ROLES = {
    "general": "你是一位资深文学翻译家，精通多国语言，译笔力求神似而非字对字。",
    "literature": "你是一位资深文学翻译家，精通多国语言，译笔力求神似而非字对字。",
    "tech": "你是一位资深技术文档翻译，术语准确、表达严谨，准确性优先于文采。",
    "law": "你是一位资深法律翻译，术语必须精确对应，句式严格对应原文结构。",
    "medical": "你是一位资深医学翻译，使用标准医学译名，术语必须准确。",
    "news": "你是一位资深新闻翻译，译笔准确客观、简洁严谨。",
}

LANG_NAMES = {
    "zh-CN": "简体中文", "zh-TW": "繁体中文",
    "en": "英语", "ja": "日语", "ko": "韩语",
    "fr": "法语", "de": "德语", "es": "西班牙语",
    "pt": "葡萄牙语", "ru": "俄语", "ar": "阿拉伯语", "it": "意大利语",
}

NON_LATIN_SCRIPT_LANGS = ("zh-CN", "zh-TW", "ja", "ko", "ru", "ar")
CJK_LIKE_LANGS = NON_LATIN_SCRIPT_LANGS
RTL_LANGS = ("ar",)

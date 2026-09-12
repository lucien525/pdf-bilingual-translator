# 📖 PDF 双语对照翻译器

把英文 PDF 翻译成中文，保留原版面和文字位置，输出两个文件：

- **`translated.pdf`** —— 纯中文版（中文直接盖在原文位置）
- **`bilingual.pdf`** —— 左右对照版（左英文、右中文，逐页严格对齐）

基于 **PyMuPDF + DeepSeek API + Gradio**，完全本地运行，只连接 DeepSeek。

---

## ✨ 特性

- **保留原排版**：用白色矩形遮盖原文，在相同坐标写入中文译文，段落位置、页眉页脚、页边距全部保持原样
- **左右对照**：`bilingual.pdf` 每页左边原文、右边译文，页高严格一致，滚动不会错位
- **断点续传**：翻到哪页记在哪页，中途关掉窗口重开，会从断点继续，已翻过的页不重复扣费
- **半成品也保存**：中途暂停或出错，已翻的部分也会输出成 PDF，不浪费 token
- **网页界面**：拖拽上传 PDF，实时进度条 + 效果预览 + 一键打开结果文件夹
- **模型可选**：`deepseek-chat`（便宜）或 `deepseek-reasoner`（质量更高）
- **API Key 灵活**：`.env` 里配默认 Key，网页里可临时覆盖

---

## 📦 环境要求

- **Python 3.10+**
- **Conda**（推荐）或系统 Python
- **中文字体**：思源宋体（Source Han Serif SC），免费开源
  - 下载：https://github.com/adobe-fonts/source-han-serif/releases
  - 拿简体中文的 `SourceHanSerifSC-Regular.otf`
- **DeepSeek API Key**：去 https://platform.deepseek.com 申请

---

## 🚀 快速开始

### 1. 创建环境并安装依赖

```bash
conda create -n pdf_trans python=3.10 -y
conda activate pdf_trans
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

或手动装：

```bash
pip install pymupdf openai gradio python-dotenv pillow
```

### 2. 配置 API Key

复制模板：

```bash
cp .env.example .env
```

编辑 `.env`，填入你的真实 Key：

```
DEEPSEEK_API_KEY=sk-你的密钥
DEEPSEEK_MODEL=deepseek-chat
```

> `.env` 已被 `.gitignore` 忽略，不会上传到 Git。

### 3. 指定中文字体路径

打开 `pdf_bilingual_app.py`，找到这一行，改成你电脑上字体文件的**完整路径**：

```python
FONT_PATH = r"D:\your\path\to\SourceHanSerifSC-Regular.otf"
```

> 路径里用 `r"..."`（原始字符串）或双反斜杠 `\\`。

### 4. 启动

```bash
python pdf_bilingual_app.py
```

浏览器自动打开 `http://127.0.0.1:7860`。

### 5. 使用

1. 拖一个英文 PDF 进去
2. 先用**试翻模式**（只翻前 5 页）跑一遍，看效果和费用
3. 满意后取消试翻模式，点「开始翻译」全量翻

---

## 📁 输出目录结构

翻完之后，结果在 `result/<书名>/` 下：

```
result/
└── 你的书名/
    ├── translated.pdf       ← 中文版（中文盖原文）
    ├── bilingual.pdf        ← 左右对照（左英右中）
    └── _work/               ← 中间文件，不用管
        ├── translate_cache.json   断点缓存
        ├── progress.json          进度记录
        ├── bilingual.html         对照网页
        ├── bilingual_pages/       页面图片
        └── preview/               预览拼接图
```

**同一本书重复翻译**：复用同一个目录，断点续传。  
**不同书**：自动新建一个以书名命名的子文件夹。

---

## 🎯 工作流程

```
上传 PDF
  │
  ├─ 从 PDF 提取每页的文字块（带坐标）
  │
  ├─ 逐页调用 DeepSeek 翻译（带缓存，重复跑不扣费）
  │
  ├─ 在原 PDF 副本上：
  │    · 用白色矩形盖住原文块
  │    · 在相同坐标写入中文译文（字号自适应）
  │
  ├─ 生成 translated.pdf
  │
  ├─ 渲染每页原图 + 译图
  │
  └─ 拼接成 bilingual.pdf
```

---

## 💰 关于费用

- DeepSeek 价格非常低，一本 300 页的书全量翻译大约 **几块钱**
- **建议先用试翻模式**：翻前 5 页，看 DeepSeek 后台的 token 消耗，乘以 `总页数 ÷ 5` 就是全书估算
- 已翻过的页缓存在 `_work/translate_cache.json` 里，重复运行不重复扣费
- 暂停、断电、崩溃后重跑，缓存还在，从断点继续

---

## ⚠️ 已知限制

- **扫描版 PDF 不支持**：如果 PDF 里的文字是图片（不能选中文字），本工具无法提取文本。需要先用 OCR（如 ABBYY、Adobe、百度 OCR）转成可选中文字的 PDF
- **左对齐为主**：中文写入时按原文块左上角对齐，居中的标题、特殊排版可能视觉上略有偏移
- **字号自适应**：中文比英文短很多，一般能塞进原框；如果塞不进会自动降字号，最小 4pt

---

## 🛠️ 常见问题

**Q：警告 "The `fitz` API is deprecated"？**  
A：PyMuPDF 新版改了包名。本项目已用 `import pymupdf as fitz` 兼容，如果还看到这条，检查 PyMuPDF 版本是否 ≥ 1.24。

**Q：字体嵌入失败 / 中文显示为方框？**  
A：检查 `FONT_PATH` 是否指向真实存在的 `.otf` 文件。

**Q：翻译出来的 PDF 是空白？**  
A：旧版本的 `doc.save(..., garbage=4)` 会把 CJK 字体当垃圾清理掉。本项目已去掉该参数，如果还空白，删掉 `translated.pdf` 重跑。

**Q：进度条一直 60%？**  
A：已修复。现在插图页（无文字的页）也会计入进度。

**Q：`git add` 报 `LF will be replaced by CRLF`？**  
A：Windows 上正常提示，不影响。想消掉可以跑：
```bash
git config --global core.autocrlf true
```

---

## 📂 项目结构

```
pdf_bilingual/
├── pdf_bilingual_app.py    主程序（Gradio 界面 + 翻译逻辑）
├── requirements.txt        依赖列表
├── .env                    私密配置（不上传）
├── .env.example            配置模板
├── .gitignore              Git 忽略规则
├── README.md               本文件
└── result/                 翻译结果（不上传）
```

---

## 🔧 可调参数（在 `pdf_bilingual_app.py` 顶部）

| 参数 | 默认值 | 说明 |
|---|---|---|
| `FONT_PATH` | （必填） | 中文字体路径 |
| `RESULT_ROOT` | `"result"` | 结果根目录 |
| `RENDER_ZOOM` | `2.0` | 图片渲染缩放（越大越清晰，文件越大） |

---

## 📜 License

本项目仅供学习和个人使用。翻译内容的版权归原作者所有，请勿用于商业传播。

---

## 🙏 致谢

- [PyMuPDF](https://pymupdf.readthedocs.io/) —— PDF 处理
- [DeepSeek](https://platform.deepseek.com/) —— 翻译 API
- [Gradio](https://gradio.app/) —— 网页界面
- [思源宋体](https://github.com/adobe-fonts/source-han-serif) —— 中文字体
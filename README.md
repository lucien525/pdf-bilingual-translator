# 📖 PDF / Word / PPT 翻译器

把英文文档翻译成中文，**保留原排版**，输出译文和左右对照两个版本。

- **PDF** → `translated.pdf`（纯中文版）+ `bilingual.pdf`（左英右中对照）
- **Word** → `xxx_cn.docx`（纯中文）+ `xxx_bilingual.docx`（原文段 + 中文段交错）
- **PPT** → `xxx_cn.pptx`（文本框替换为中文，样式保留）

基于 **PyMuPDF + python-docx + python-pptx + DeepSeek API + Gradio**，完全本地运行，只连接 DeepSeek。

---

## ✨ 特性

- **保留原排版**
  - PDF：白块遮盖原文 + 原坐标写入译文，段落、页眉页脚、页边距全保持
  - Word：段落级替换 + 双语版原文段下插入译文段（米色底区分）
  - PPT：文本框内替换，字体大小颜色位置全保留
- **左右对照**：PDF 版每页左原文、右译文，页高严格一致
- **多任务并行**：可以同时跑 PDF / Word / PPT 多个任务，互不干扰
- **单独停止**：可以只停某一个任务，其它继续跑
- **断点续传**：翻到哪页记在哪页，中途关窗口重开继续，已翻过的不重复扣费
- **半成品也保存**：中途暂停或出错，已翻部分照常输出成 PDF
- **网页界面**：拖拽上传，进度条 + 前 5 页预览 + 一键打开结果文件夹
- **模型可选**：`deepseek-chat`（便宜）或 `deepseek-reasoner`（质量更高）
- **API Key 灵活**：`.env` 里配默认 Key，网页里可临时覆盖

---

## 📦 环境要求

- **Python 3.10+**
- **Conda**（推荐）
- **中文字体**：思源宋体（Source Han Serif SC）
  - 下载：https://github.com/adobe-fonts/source-han-serif/releases
  - 拿简体中文的 `SourceHanSerifSC-Regular.otf`
- **DeepSeek API Key**：https://platform.deepseek.com

---

## 🚀 快速开始

### 1. 安装

```bash
conda create -n pdf_trans python=3.10 -y
conda activate pdf_trans
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 2. 配置 API Key

```bash
cp .env.example .env
```

编辑 `.env`：

```
DEEPSEEK_API_KEY=sk-你的密钥
DEEPSEEK_MODEL=deepseek-chat
```

### 3. 指定中文字体路径

打开 `bilingual_app.py`，找到：

```python
FONT_PATH = r"D:\your\path\to\SourceHanSerifSC-Regular.otf"
```

改成你电脑上字体文件的完整路径。

### 4. 启动

```bash
python bilingual_app.py
```

浏览器自动打开 `http://127.0.0.1:7860`。

**Windows 用户**：也可以直接双击 `重启翻译器.bat` 一键启动。

### 5. 使用

1. 选择文档类型（PDF / Word / PPT）
2. 上传文件
3. 先用**试翻模式**（PDF 只翻前 5 页）跑一遍，看效果和费用
4. 满意后取消试翻模式，点「创建新任务并开始」全量翻

---

## 📁 输出目录结构

```
result/
└── 你的书名/                          ← 以 PDF/Word/PPT 文件名命名
    ├── translated.pdf                 ← PDF 纯中文版
    ├── bilingual.pdf                  ← PDF 左右对照
    ├── 书名_cn.docx / 书名_cn.pptx    ← Word / PPT 纯中文版
    ├── 书名_bilingual.docx            ← Word 双语版
    └── _work/                         ← 中间文件，不用管
        ├── state.json                 ← 任务状态（多任务并行用）
        ├── progress.json              ← 页码进度
        ├── translate_cache.json       ← 断点缓存（**唯一不能删的文件**）
        ├── input/                     ← 上传文件的副本
        ├── preview/                   ← 前 5 页预览图
        └── ...
```

**同一本书重复跑** → 复用同目录，断点续传。  
**不同书** → 各自独立子目录。

---

## 💰 关于费用

- DeepSeek 价格非常低，一本 300 页的书全量翻译大约 **几块钱**
- **建议先用试翻模式**：翻前 5 页，看 DeepSeek 后台 token 消耗 × (总页数 ÷ 5) 就是全书估算
- 已翻过的页缓存在 `_work/translate_cache.json`，重复运行不重复扣费
- 暂停、断电、崩溃后重跑，从断点继续

**会重新扣费的唯一情况**：删掉了 `translate_cache.json`，或者修改了原始 PDF 的文字内容。

---

## ⚠️ 已知限制

- **扫描版 PDF 不支持**：文字是图片（不能选中文字），需先 OCR
- **`.doc` / `.ppt` 老格式不支持**：需要先用 Office 另存为 `.docx` / `.pptx`
- **Word 里的文本框、页眉页脚、批注**：`python-docx` 默认访问不到
- **PPT 里的 SmartArt、图表、备注页**：`python-pptx` 不直接支持
- **左对齐为主**：居中的标题、特殊排版可能视觉上略有偏移
- **字号自适应**：中文比英文短，一般能塞进原框；塞不进会自动降字号，最小 4pt

---

## 🛠️ 常见问题

**Q：警告 "The `fitz` API is deprecated"？**  
A：PyMuPDF 新版改了包名。已用 `import pymupdf as fitz` 兼容。

**Q：中文显示为方框 / 字体嵌入失败？**  
A：检查 `FONT_PATH` 是否指向真实存在的 `.otf` 文件。

**Q：PDF 是空白？**  
A：删掉 `translated.pdf` 重跑。旧版本曾用 `garbage=4` 把 CJK 字体当垃圾清理掉，本项目已去掉。

**Q：`translated_new_时间戳.pdf` 是什么？**  
A：保存时目标文件被占用（杀毒软件、PDF 阅读器等），自动另存的兜底文件。程序启动时会自动清理。不影响正常文件。

**Q：`git add` 报 `LF will be replaced by CRLF`？**  
A：Windows 上正常提示，不影响。想消掉：
```bash
git config --global core.autocrlf true
```

---

## 📂 项目结构

```
pdf_bilingual/
├── bilingual_app.py       主程序（Gradio 界面 + 翻译逻辑）
├── requirements.txt       依赖列表
├── .env                   私密配置（不上传）
├── .env.example           配置模板
├── .gitignore             Git 忽略规则
├── README.md              本文件
├── 重启翻译器.bat          一键启动（Windows 可选）
└── result/                翻译结果（不上传）
```

---

## 🔧 可调参数（`bilingual_app.py` 顶部）

| 参数 | 默认 | 说明 |
|---|---|---|
| `FONT_PATH` | （必填） | 中文字体路径 |
| `RESULT_ROOT` | `"result"` | 结果根目录 |
| `RENDER_ZOOM` | `2.0` | PDF 渲染缩放，越大越清晰 |
| `PREVIEW_PAGES` | `5` | 预览显示前几页 |
| `PREVIEW_MAX_WIDTH` | `1400` | 预览图最大宽度 |
| `PREVIEW_JPEG_QUALITY` | `88` | 预览图 JPEG 质量 |

---

## 📜 License

本项目仅供学习和个人使用。翻译内容的版权归原作者所有，请勿用于商业传播。

---

## 🙏 致谢

- [PyMuPDF](https://pymupdf.readthedocs.io/) —— PDF 处理
- [python-docx](https://python-docx.readthedocs.io/) —— Word 处理
- [python-pptx](https://python-pptx.readthedocs.io/) —— PPT 处理
- [DeepSeek](https://platform.deepseek.com/) —— 翻译 API
- [Gradio](https://gradio.app/) —— 网页界面
- [思源宋体](https://github.com/adobe-fonts/source-han-serif) —— 中文字体
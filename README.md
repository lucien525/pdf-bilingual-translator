# 📖 PDF / Word / PPT 翻译器

把任意文档**保留原排版**地翻译成 **12 种目标语言**，输出纯译文和左右对照两个版本。

- **PDF** → `translated.pdf`（纯译文）+ `bilingual.pdf`（左原文右译文）
- **Word** → `xxx_cn.docx`（纯译文）+ `xxx_bilingual.docx`（原文段 + 译文段交错）
- **PPT** → `xxx_cn.pptx`（文本框替换，样式保留）

基于 **PyMuPDF + python-docx + python-pptx + DeepSeek API + Gradio**，完全本地运行，只连接 DeepSeek。

---

## ✨ 特性

### 翻译能力

- **12 种目标语言**：简体中文、繁体中文、英语、日语、韩语、法语、德语、西班牙语、葡萄牙语、俄语、阿拉伯语、意大利语
- **保留原排版**
  - **PDF**：白块遮盖原文 + 原坐标写入译文，段落、页眉页脚、页边距全保持
  - **Word**：段落级替换 + 双语版原文段下插入译文段（米色底区分）
  - **PPT**：文本框内替换，字体大小颜色位置全保留
- **左右对照**：PDF 版每页左原文、右译文，页高严格一致
- **字号自适应**：中文比英文短，一般能塞进原框；塞不进会自动降字号，最小 4pt
- **模型可选**：`deepseek-chat`（便宜）或 `deepseek-reasoner`（质量更高）

### 任务管理

- **多任务并行**：可以同时跑 PDF / Word / PPT 多个任务，互不干扰
- **单独停止**：可以只停某一个任务，其它继续跑
- **断点续传**：翻到哪页记在哪页，中途关窗口重开继续，已翻过的不重复扣费
- **内容自愈**：续传时自动扫描译文 PDF 每一页，**发现缺页自动补翻**（即使进度记录说"已完成"）
- **每 10 页落盘**：翻译过程每 10 页保存一次译文，崩溃最多丢 10 页进度
- **半成品也保存**：中途暂停或出错，已翻部分照常输出成 PDF
- **空翻译拦截**：API 返回空内容时**不会标记为完成**，下次续传自动重试

### 网页界面

- 拖拽上传，自动识别文件类型
- 实时进度条
- 前 5 页预览（PDF 图片对照 · Word/PPT 文本对照）
- 任务完成时**绿色大横幅**，显示源文件、目标语言、页数、结果文件列表、保存路径
- 一键打开结果文件夹

---

## 🌐 支持的目标语言

| 代码 | 语言 | 代码 | 语言 |
|:---:|:---:|:---:|:---:|
| `zh-CN` | 简体中文 | `zh-TW` | 繁体中文 |
| `en` | 英语 | `ja` | 日语 |
| `ko` | 韩语 | `fr` | 法语 |
| `de` | 德语 | `es` | 西班牙语 |
| `pt` | 葡萄牙语 | `ru` | 俄语 |
| `ar` | 阿拉伯语 | `it` | 意大利语 |

### 关于语言的两点说明

**① 内容自愈的适用范围**

续传时的"自动扫描缺页"能力，只对**可以用字符类型区分**的语言生效：

| 支持自愈 | 不支持自愈（仍能正常翻译） |
|---|---|
| 中、日、韩、俄、阿 | 英、法、德、西、葡、意 |

原因：翻成英文/法文等拉丁系语言时，无法仅凭字符判断"这一页到底是原版还是译文"。

**② PDF 输出的字体限制**

PDF 输出固定使用**思源宋体**渲染译文。如果翻译成阿拉伯语或俄语，CJK 字体可能无法完整显示那些字符——**PDF 模式目前最适合中/日/韩目标语言**。

Word / PPT 使用原文档字体，没有这个限制。

---

## 📦 环境要求

- **Python 3.10+**
- **Conda**（推荐，也支持普通 venv）
- **中文字体**：思源宋体（Source Han Serif SC）
  - 下载：https://github.com/adobe-fonts/source-han-serif/releases
  - 拿简体中文的 `SourceHanSerifSC-Regular.otf`
- **DeepSeek API Key**：https://platform.deepseek.com

---

## 🚀 快速开始

### 1. 安装依赖

```bash
conda create -n trans python=3.10 -y
conda activate trans
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

> 也可以留空 `.env`，启动后在网页里临时填 Key。

### 3. 指定中文字体路径

打开 `bilingual_app.py`，找到第 45 行左右：

```python
FONT_PATH = r"D:\your\path\to\SourceHanSerifSC-Regular.otf"
```

改成你电脑上字体文件的**完整路径**。

### 4. 启动

**方式 A：命令行**

```bash
python bilingual_app.py
```

浏览器自动打开 `http://127.0.0.1:7860`。

**方式 B：双击启动（Windows）**

双击 `重启翻译器.bat`。它会**自动探测** conda 安装位置并激活 `trans` 环境。

> 如果探测失败，脚本会提示你手动编辑 `CONDA_ACTIVATE` 路径。

### 5. 使用

1. **选目标语言**（默认简体中文）
2. **上传文件**（拖入即可，自动识别 PDF / Word / PPT）
3. **先试翻**（勾选试翻模式，PDF 只翻前 5 页），看效果和费用
4. 满意后**取消试翻**，点「创建新任务并开始」全量翻

---

## 📁 输出目录结构

```
result/
└── 你的书名/                              ← 以源文件名命名
    ├── translated.pdf                     ← PDF 纯译文版
    ├── bilingual.pdf                      ← PDF 左右对照
    ├── 书名_cn.docx                       ← Word 纯译文版
    ├── 书名_bilingual.docx                ← Word 双语版
    ├── 书名_cn.pptx                       ← PPT 译文版
    └── _work/                             ← 中间文件
        ├── state.json                     ← 任务状态（多任务并行用）
        ├── progress.json                  ← 页码进度
        ├── translate_cache.json           ← ⚠️ 断点缓存（唯一不能删的文件）
        ├── input/                         ← 上传文件的副本
        └── preview/                       ← 前 5 页预览图
```

**同一本书重复跑** → 复用同目录，断点续传。
**不同书** → 各自独立子目录。

---

## 💰 关于费用

- DeepSeek 价格非常低，一本 300 页的书全量翻译大约 **几块钱**
- **建议先用试翻模式**：翻前 5 页，看 DeepSeek 后台 token 消耗 × (总页数 ÷ 5) 就是全书估算
- 已翻过的页缓存在 `_work/translate_cache.json`，重复运行不重复扣费
- 暂停、断电、崩溃后重跑，从断点继续

**会重新扣费的唯一情况**：
- 删掉了 `translate_cache.json`
- 修改了原始文件的文字内容
- **换了目标语言**（每种语言有独立缓存，互不影响）

---

## ⚠️ 已知限制

### 文件格式

- **扫描版 PDF 不支持**：文字是图片（不能选中文字），需先 OCR
- **`.doc` / `.ppt` 老格式不支持**：需要先用 Office 另存为 `.docx` / `.pptx`

### 内容覆盖范围

- **Word 里的文本框、页眉页脚、批注**：`python-docx` 默认访问不到
- **PPT 里的 SmartArt、图表、备注页**：`python-pptx` 不直接支持

### 排版

- **左对齐为主**：居中的标题、特殊排版可能视觉上略有偏移
- **字号最小 4pt**：极端情况下（原框太小、译文太长）可能溢出框外

### 多语言

- **PDF 输出固定用思源宋体**：翻成阿拉伯语/俄语可能显示不全，**PDF 模式推荐用中/日/韩**
- **内容自愈只对 CJK 系语言生效**（见上方说明）

---

## 🛠️ 常见问题

**Q：警告 "The `fitz` API is deprecated"？**
A：PyMuPDF 新版改了包名。已用 `import pymupdf as fitz` 兼容，无需理会。

**Q：中文显示为方框 / 字体嵌入失败？**
A：检查 `FONT_PATH` 是否指向真实存在的 `.otf` 文件。可以跑 `test.py` 快速验证：
```bash
python test.py
```

**Q：PDF 是空白？**
A：删掉 `translated.pdf` 重跑。旧版本曾用 `garbage=4` 把 CJK 字体当垃圾清理掉，本项目已去掉。

**Q：`translated_new_时间戳.pdf` 是什么？**
A：保存时目标文件被占用（杀毒软件、PDF 阅读器等），自动另存的兜底文件。程序启动时会自动清理。

**Q：程序中途崩了，怎么恢复？**
A：重新上传**同一个文件**（文件名要一致），点开始。程序会自动扫描已翻部分，从断点继续，**不重复扣费**。

**Q：翻译完成但发现某页还是原文？**
A：新版有**内容自愈**功能，重新上传同一文件点开始即可，会自动检测并补翻。

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
├── test.py                字体测试工具（可选）
├── requirements.txt       依赖列表
├── .env                   私密配置（不上传 git）
├── .env.example           配置模板
├── .gitignore             Git 忽略规则
├── README.md              本文件
├── 重启翻译器.bat          一键启动（Windows，自动探测 conda）
├── 打开网页.bat            打开本地网页
└── result/                翻译结果（不上传 git）
```

---

## 🔧 可调参数

全部在 `bilingual_app.py` 顶部的配置区：

| 参数 | 默认 | 说明 |
|---|---|---|
| `FONT_PATH` | **（必填）** | 中文字体文件路径 |
| `RESULT_ROOT` | `"result"` | 结果根目录 |
| `RENDER_ZOOM` | `2.0` | PDF 渲染缩放，越大越清晰（也越慢、越大） |
| `PREVIEW_PAGES` | `5` | 预览显示前几页 |
| `PREVIEW_MAX_WIDTH` | `1400` | 预览图最大宽度 |
| `PREVIEW_JPEG_QUALITY` | `88` | 预览图 JPEG 质量 |
| `PREVIEW_PARAS` | `5` | Word / PPT 预览显示前几段 |
| `CHECKPOINT_EVERY` | `10` | 每多少页落盘一次 |
| `PAGE_CN_THRESHOLD` | `20` | 判定一页"翻译过"的最少中文字符数 |
| `VALID_RATIO_THRESHOLD` | `0.3` | 判定一次 API 返回有效的最少段落占比 |

**调优建议**：
- 书的页数多、PDF 太大 → 把 `RENDER_ZOOM` 降到 `1.5`
- 预览图糊 → 把 `RENDER_ZOOM` 提到 `2.5`
- 磁盘 IO 太频繁 → 把 `CHECKPOINT_EVERY` 提到 `20`

---

## 📜 License

本项目仅供学习和个人使用。翻译内容的版权归原作者所有，请勿用于商业传播。

---

## 🙏 致谢

- [PyMuPDF](https://pymupdf.readthedocs.io/) — PDF 处理
- [python-docx](https://python-docx.readthedocs.io/) — Word 处理
- [python-pptx](https://python-pptx.readthedocs.io/) — PPT 处理
- [DeepSeek](https://platform.deepseek.com/) — 翻译 API
- [Gradio](https://gradio.app/) — 网页界面
- [思源宋体](https://github.com/adobe-fonts/source-han-serif) — 中文字体
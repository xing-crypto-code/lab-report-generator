# 🧪 无痛实验报告生成器 (Lab Report Generator)

> **解放工科生双手**：上传实验指导手册 + 老师模板，AI 一键生成标准实验报告

[![GitHub](https://img.shields.io/badge/GitHub-xing--crypto--code-blue)](https://github.com/xing-crypto-code/lab-report-generator)

## ✨ 特性

- 📷 **智能 OCR**：EasyOCR 中文识别 + 图像预处理（去噪/增强/二值化）
- 🧠 **多模型支持**：硅基流动 SiliconFlow 11 款模型可选（DeepSeek V3/V4/Qwen/GLM）
- 📚 **工科知识库**：6 大实验类型（物理/化学/电路/力学/光学/计算机）标准结构
- 🗂️ **可扩展范例库**：把自己的范文/指导书丢进 `kb/`，生成时自动检索注入（检索用 BM25 + 字符二元组，纯标准库，不装额外依赖）
- 📝 **格式稳定**：python-docx 保持原模板格式，宋体正文 + 黑体标题，不乱码
- 🎨 **精致 UI**：Gradio 现代界面，支持内置标准模板
- 🆓 **免费使用**：硅基流动新用户 14 元额度，够生成数百份报告

## 🚀 快速开始

### ⭐ 一键启动（推荐，给不会命令行的同学）

安装好 [Python](https://www.python.org/downloads/) 后，**双击 `start.bat`**：
自动检查依赖 → 启动网页 → 自动打开浏览器 → 填 API Key 即可用。

- 依赖是装在 `libs/` 里的（本仓库自带的三方库），**无需全局安装**。
- 没有 API Key 也能进界面，用内置模板 + 知识库先试效果。
- 首次若提示缺依赖：`python install_deps.py` 装全。

### 手动启动（开发用）

```bash
pip install -r requirements.txt
# 或手动：pip install gradio PyMuPDF easyocr torch python-docx openai pillow
cp .env.example .env    # 编辑填 API Key
python app.py           # 浏览器打开 http://127.0.0.1:7860
```

### 分享给同学/别人（局域网）

`app.py` 已监听 `0.0.0.0:7860`。在同一局域网内，让别人浏览器访问
`http://你的IP:7860` 即可（`ipconfig` 查 IP）。注意这是**你自己的 API Key**，
别把 `.env` 发给别人。

想让好友用已有的 `libs/`、不装依赖？直接把整个文件夹打包发走，对方装好
Python 后双击 `start.bat` 即可（体积较大，含 OCR 引擎）。

## 📂 项目结构

```
├── app.py                  # Gradio Web 主入口
├── config.py               # 配置中心
├── modules/
│   ├── parser.py           # PDF/图片 OCR 解析
│   ├── preprocessor.py     # 图像预处理（去噪/增强/纠偏）
│   ├── template.py         # .docx 模板结构解析
│   ├── filler.py           # AI 填充 + 多模型对比
│   ├── knowledge.py        # 工科实验知识库 + 标准模板生成
│   ├── retrieval.py        # 范例库检索（kb/ 入库 · 切块 · BM25）
│   └── generator.py        # 文档生成（保持格式）
├── kb/                     # 你的范例库：范文/ 指导书/ 规范/
├── templates/              # 内置模板目录
├── outputs/                # 生成结果
└── libs/                   # pip --target 安装目录
```

## 🔬 知识库覆盖

| 实验类型 | 标准章节 |
|----------|----------|
| 物理实验 | 目的 → 原理 → 器材 → 步骤 → 数据 → 处理 → 误差 → 结论 |
| 化学实验 | 目的 → 反应方程式 → 试剂仪器 → 步骤 → 数据 → 讨论 → 结论 |
| 电路实验 | 目的 → 电路图 → 仪器器件 → 步骤 → 波形/数据 → 误差分析 |
| 力学实验 | 目的 → 公式推导 → 仪器 → 步骤 → 数据图表 → 误差 → 结论 |
| 光学实验 | 目的 → 光路图 → 仪器 → 步骤 → 数据 → 处理 → 误差 |
| 计算机实验 | 目的 → 环境 → 算法说明 → 步骤与代码 → 结果分析 → 总结 |

## 🗂️ 用你自己的范文提升生成效果

模型权重改不了，但**每次生成时检索几段本校的范文塞进 prompt**，效果差别很大。
把文件按身份放好，跑一条命令建索引即可：

```bash
kb/范文/    ← 同学/学长写得不错的报告（最值钱，决定"好报告长什么样"）
kb/指导书/  ← 课程实验指导书、讲义（实验内容以此为准）
kb/规范/    ← 写作规范、评分标准
kb/草稿/    ← 自己以前的稿子：默认不参与检索，KB_DRAFTS=1 才放行

python -m modules.retrieval ingest     # 增量入库
python -m modules.retrieval stats -v   # 看看库里有啥（逐份列出分类）
python -m modules.retrieval search "霍尔效应 螺线管 磁场" -k 3   # 验证检索效果
```

仓库是公开的，入库前先清个人信息（姓名、学号、教师名）：

```bash
python scripts/clean_pii.py scan 张三 李四        # 体检
python scripts/clean_pii.py scrub 王老师 --teacher  # 清理并复核
```

生成时会自动检索 3 段最相关的片段（限 1500 字）注入 system prompt，
并明确要求**只学写法、不得抄用范例里的数据**。没有入库时整个功能静默跳过，不影响原有流程。

可调环境变量：`KB_ENABLE=0` 关掉检索｜`KB_TOP_K` 片段数｜`KB_MAX_CHARS` 字数上限。
PDF 需要 `pymupdf`（`pip install pymupdf`），扫描件请先 OCR。

细节见 [docs/知识库使用指南.md](docs/知识库使用指南.md)。

## 🤖 可用模型

通过硅基流动 SiliconFlow，11+ 模型可选：

- **DeepSeek V3 / V3.2 / V4 Pro / R1** — 综合能力优秀，中英文俱佳
- **Qwen 3.5 27B / 35B / Qwen 3 32B** — 阿里通义千问，中文理解强
- **GLM-5.3 / GLM-5.3 Flash / GLM-4 32B** — 智谱系列，中文优异
- **DeepSeek-OCR** — 专用 OCR 模型（可选替代 EasyOCR）

## 📄 License

MIT
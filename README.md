# 🧪 无痛实验报告生成器 (Lab Report Generator)

> **解放工科生双手**：上传实验指导手册 + 老师模板，AI 一键生成标准实验报告

[![GitHub](https://img.shields.io/badge/GitHub-xing--crypto--code-blue)](https://github.com/xing-crypto-code/lab-report-generator)

## ✨ 特性

- 📷 **智能 OCR**：EasyOCR 中文识别 + 图像预处理（去噪/增强/二值化）
- 🧠 **多模型支持**：硅基流动 SiliconFlow 11 款模型可选（DeepSeek V3/V4/Qwen/GLM）
- 📚 **工科知识库**：6 大实验类型（物理/化学/电路/力学/光学/计算机）标准结构
- 📝 **格式稳定**：python-docx 保持原模板格式，宋体正文 + 黑体标题，不乱码
- 🎨 **精致 UI**：Gradio 现代界面，支持内置标准模板
- 🆓 **免费使用**：硅基流动新用户 14 元额度，够生成数百份报告

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
# 或手动：pip install gradio PyMuPDF easyocr torch python-docx openai pillow
```

### 2. 配置 API Key

```bash
cp .env.example .env
# 编辑 .env，填入硅基流动 API Key（注册: https://siliconflow.cn）
```

或直接在 Web 界面填写。

### 3. 启动

```bash
python app.py
# 浏览器打开 http://127.0.0.1:7860
```

### 4. 使用

1. 上传实验指导手册（PDF / 照片）
2. 上传老师的 .docx 模板（或勾选内置模板）
3. 选择 AI 模型 → 点击「一键生成」
4. 下载 .docx，WPS/Office 打开微调 → 交作业

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
│   └── generator.py        # 文档生成（保持格式）
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

## 🤖 可用模型

通过硅基流动 SiliconFlow，11+ 模型可选：

- **DeepSeek V3 / V3.2 / V4 Pro / R1** — 综合能力优秀，中英文俱佳
- **Qwen 3.5 27B / 35B / Qwen 3 32B** — 阿里通义千问，中文理解强
- **GLM-5.3 / GLM-5.3 Flash / GLM-4 32B** — 智谱系列，中文优异
- **DeepSeek-OCR** — 专用 OCR 模型（可选替代 EasyOCR）

## 📄 License

MIT
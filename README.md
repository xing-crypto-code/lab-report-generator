# 无痛实验报告生成器 (Lab Report Generator)

> 解放工科生双手：做完实验，一键生成可直接交作业的实验报告。

## 痛点

工科生每天满课，做完实验后还要花大量时间写实验报告。本项目的目标是把"写实验报告"从几小时压缩到几分钟：

- 上传实验指导手册（PDF / 照片 / 截图）→ 自动识别文字、图片、表格等数据信息，提取实验内容、步骤、数据表
- 上传老师发的实验报告模板（Word / PDF）→ 自动理解模板结构
- 智能生成 → 自动填入模板对应位置，润色、去重
- 输出可编辑文档 → Word / PDF 互转，不乱码，可在 WPS / Office Word 中直接编辑微调后交作业
- **格式稳定**：改一个地方不会导致全局错乱

## 技术选型（规划中）

| 模块 | 方案 |
|------|------|
| 文档解析 | pdfplumber / PyMuPDF / 表格识别 |
| OCR | PaddleOCR / Tesseract |
| 文档生成 | python-docx（保留模板格式） |
| 转换 | LibreOffice / docx2pdf |
| 智能生成 | 大模型 API（OpenAI 兼容 / 本地） |

## 当前进度

- [x] 创建 GitHub 仓库
- [ ] 确定实现形态（Web 应用 / 桌面 EXE / VS Code 插件）
- [ ] 搭建项目骨架
- [ ] 文档解析模块（PDF / 图片 OCR / 表格）
- [ ] 模板解析模块
- [ ] 内容生成与填充模块
- [ ] 格式保持与转换模块
- [ ] 测试与验收

## 仓库信息

- 仓库地址：https://github.com/xing-crypto-code/lab-report-generator
- 作者：xing-crypto-code

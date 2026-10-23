"""
无痛实验报告生成器 · Gradio Web 界面
=====================================
主入口：运行 `python app.py` 即可启动 Web 服务
在浏览器中打开 http://127.0.0.1:7860 使用
"""
import gradio as gr
import os
from pathlib import Path
import shutil

from config import UPLOAD_DIR, OUTPUT_DIR
from modules.parser import parse_document, parse_pdf, parse_image
from modules.template import analyze_template, get_template_summary
from modules.filler import generate_report
from modules.generator import generate_report_document


def process_files(manual_file, template_file, api_key):
    """
    核心处理流水线：
    1. 解析上传的实验指导手册 → OCR 提取内容
    2. 解析上传的模板 → 理解结构
    3. 调 AI 填入模板 → 润色去重
    4. 生成 .docx → 返回下载
    """
    if manual_file is None:
        return None, None, "❌ 请上传实验指导手册（PDF/照片）"
    if template_file is None:
        return None, None, "❌ 请上传老师发的实验报告模板（.docx）"

    # 设置 API Key（如果用户填了）
    if api_key and api_key.strip() and not api_key.startswith("sk-your"):
        os.environ["LLM_API_KEY"] = api_key.strip()

    status_lines = []

    try:
        # ---- 第1步：解析实验指导手册 ----
        status_lines.append("🔍 正在解析实验指导手册...")
        manual_path = str(manual_file.name) if hasattr(manual_file, 'name') else manual_file
        doc_result = parse_document(manual_path)
        ocr_text = doc_result["text"]
        ocr_tables = doc_result["tables"]
        status_lines.append(f"   ✓ 识别 {len(ocr_text)} 字符，{len(ocr_tables)} 个表格")
        status_lines.append("")

        # ---- 第2步：解析模板 ----
        status_lines.append("📋 正在解析实验报告模板结构...")
        template_path = str(template_file.name) if hasattr(template_file, 'name') else template_file
        template_analysis = analyze_template(template_path)
        template_summary = get_template_summary(template_analysis)
        status_lines.append(f"   ✓ 识别 {template_analysis['total_paragraphs']} 个段落，"
                           f"{template_analysis['total_tables']} 个表格")
        status_lines.append(f"   ✓ 检测到章节: {', '.join(template_analysis['sections_found'][:8])}")
        status_lines.append("")

        # ---- 第3步：AI 填充 ----
        status_lines.append("🤖 正在调用 AI 生成实验报告内容...")
        filled = generate_report(ocr_text, ocr_tables, template_summary)
        filled_sections = filled.get("sections", {})
        filled_count = sum(1 for v in filled_sections.values() if v and v.strip())
        status_lines.append(f"   ✓ AI 已填充 {filled_count} 个章节")
        if filled.get("notes"):
            status_lines.append(f"   ℹ️ {filled['notes']}")
        status_lines.append("")

        # ---- 第4步：生成文档 ----
        status_lines.append("📝 正在生成最终文档...")

        # 从 manual 文件名推断输出名称
        manual_name = Path(manual_path).stem
        output_name = f"{manual_name}_实验报告"

        result_paths = generate_report_document(
            template_path=template_path,
            filled_data=filled,
            output_name=output_name,
            to_pdf=False  # 默认只生成 .docx（PDF 需要装 Word）
        )

        docx_path = result_paths.get("docx", "")
        if docx_path and Path(docx_path).exists():
            status_lines.append(f"   ✓ 文档已生成: {Path(docx_path).name}")
            status_lines.append("")
            status_lines.append("---")
            status_lines.append("✅ **完成！请下载下方的 .docx 文件，用 WPS/Office Word 打开微调后直接交作业。**")

            # 显示填充内容预览
            preview = _build_preview(filled_sections, template_analysis)

            return (
                str(docx_path),  # 下载文件
                preview,         # 预览区
                "\n".join(status_lines)  # 状态日志
            )
        else:
            status_lines.append("   ❌ 文档生成失败")
            return None, None, "\n".join(status_lines)

    except Exception as e:
        import traceback
        status_lines.append(f"\n❌ 处理出错: {e}")
        status_lines.append(f"```\n{traceback.format_exc()[:500]}\n```")
        return None, None, "\n".join(status_lines)


def _build_preview(filled_sections: dict, template_analysis: dict) -> str:
    """生成填充内容预览的 Markdown 文本"""
    lines = ["# 📄 实验报告内容预览\n"]

    for section_name, content in filled_sections.items():
        if content and content.strip() and content.strip() != "[待补充]":
            # 限制预览长度
            preview = content[:300] + ("..." if len(content) > 300 else "")
            lines.append(f"## {section_name}")
            lines.append(preview)
            lines.append("")

    if len(lines) == 1:
        lines.append("*（AI 未生成内容，请检查 API 配置或重试）*")

    return "\n".join(lines)


# ==================== Gradio 界面 ====================

def build_ui():
    """构建 Gradio Web 界面"""
    with gr.Blocks(
        title="无痛实验报告生成器",
        theme=gr.themes.Soft(primary_hue="blue"),
    ) as demo:

        # ---- 标题 ----
        gr.Markdown("""
        # 🧪 无痛实验报告生成器
        > **解放工科生双手** — 上传实验指导手册 + 老师模板，AI 自动生成可交作业的实验报告

        """)

        # ---- 输入区 ----
        with gr.Row():
            with gr.Column(scale=1):
                gr.Markdown("### 📥 第1步：上传实验指导手册")
                manual_input = gr.File(
                    label="实验指导手册",
                    file_types=[".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".webp"],
                    type="filepath",
                )
                gr.Markdown("*支持 PDF、照片、截图*")

            with gr.Column(scale=1):
                gr.Markdown("### 📋 第2步：上传实验报告模板")
                template_input = gr.File(
                    label="老师发的模板 (.docx)",
                    file_types=[".docx"],
                    type="filepath",
                )
                gr.Markdown("*支持 .docx 格式（WPS/Word 模板）*")

        # ---- API Key ----
        with gr.Accordion("⚙️ API 设置（可选，也可在 .env 中配置）", open=False):
            api_key_input = gr.Textbox(
                label="API Key",
                placeholder="sk-... （DeepSeek 注册即送 500万 tokens）",
                type="password"
            )
            gr.Markdown("""
            > **没有 API Key？** [注册 DeepSeek](https://platform.deepseek.com) 新用户免费获得 500万 tokens。
            > 也支持 OpenAI / Qwen / 其他兼容 API，在 `.env` 文件里修改 `LLM_BASE_URL` 即可。
            """)

        # ---- 生成按钮 ----
        with gr.Row():
            generate_btn = gr.Button("🚀 一键生成实验报告", variant="primary", size="lg")

        # ---- 输出区 ----
        gr.Markdown("---")

        with gr.Row():
            with gr.Column(scale=1):
                gr.Markdown("### 📥 下载文件")
                download_output = gr.File(label="生成的实验报告 (.docx)")

            with gr.Column(scale=2):
                gr.Markdown("### 📄 内容预览")
                preview_output = gr.Markdown(value="*等待生成...*")

        # ---- 处理日志 ----
        with gr.Accordion("🔍 处理日志", open=False):
            status_output = gr.Textbox(label="详细日志", lines=12, max_lines=30)

        # ---- 绑定事件 ----
        generate_btn.click(
            fn=process_files,
            inputs=[manual_input, template_input, api_key_input],
            outputs=[download_output, preview_output, status_output],
        )

        # ---- 使用说明 ----
        gr.Markdown("""
        ---
        ### 📖 使用说明

        1. **准备文件**：把实验指导手册（PDF或拍照）和老师发的模板（.docx）准备好
        2. **上传文件**：分别拖入上方两个上传框
        3. **配置 API**：填入 DeepSeek API Key（或在项目目录创建 `.env` 文件配置）
        4. **一键生成**：点击按钮，等待 10-30 秒
        5. **下载文档**：下载生成的 .docx，用 WPS / Office Word 打开微调即可交作业

        ### ⚡ 技术栈

        | 环节 | 技术 |
        |------|------|
        | OCR 识别 | PaddleOCR（中文最佳） |
        | PDF 解析 | PyMuPDF |
        | AI 润色填充 | DeepSeek / GPT / Qwen 等 OpenAI 兼容 API |
        | 文档生成 | python-docx（保持原格式） |

        ### 🔗 链接

        - 📦 [GitHub 仓库](https://github.com/xing-crypto-code/lab-report-generator)
        - 🔑 [获取 DeepSeek API Key](https://platform.deepseek.com/api_keys)
        """)

    return demo


# ==================== 启动入口 ====================

if __name__ == "__main__":
    demo = build_ui()
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=False,  # 改为 True 可生成公网链接
        show_error=True,
    )
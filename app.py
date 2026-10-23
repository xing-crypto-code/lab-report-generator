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

# 先把本地 libs 目录加入 sys.path（免 venv 方案）
import sys
_libs = str(Path(__file__).resolve().parent / "libs")
if _libs not in sys.path:
    sys.path.insert(0, _libs)

from config import UPLOAD_DIR, OUTPUT_DIR
from modules.parser import parse_document, parse_pdf, parse_image
from modules.template import analyze_template, get_template_summary
from modules.filler import generate_report
from modules.generator import generate_report_document


def process_files(manual_file, template_file, api_key, api_base, model):
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

    # 设置 API 配置（优先用界面上填的）
    if api_key and api_key.strip():
        os.environ["LLM_API_KEY"] = api_key.strip()
    if api_base and api_base.strip():
        os.environ["LLM_BASE_URL"] = api_base.strip()
    if model and model.strip():
        os.environ["LLM_MODEL"] = model.strip()

    status_lines = []

    try:
        # ---- 第1步：解析实验指导手册 ----
        status_lines.append("🔍 正在解析实验指导手册...")
        manual_path = str(manual_file)
        doc_result = parse_document(manual_path)
        ocr_text = doc_result["text"]
        ocr_tables = doc_result["tables"]
        status_lines.append(f"   ✓ 识别 {len(ocr_text)} 字符，{len(ocr_tables)} 个表格")
        status_lines.append("")

        # ---- 第2步：解析模板 ----
        status_lines.append("📋 正在解析实验报告模板结构...")
        template_path = str(template_file)
        template_analysis = analyze_template(template_path)
        template_summary = get_template_summary(template_analysis)
        status_lines.append(f"   ✓ 识别 {template_analysis['total_paragraphs']} 个段落，"
                           f"{template_analysis['total_tables']} 个表格")
        sections = template_analysis.get("sections_found", [])
        status_lines.append(f"   ✓ 检测到章节: {', '.join(sections[:8])}")
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
        manual_name = Path(manual_path).stem
        output_name = f"{manual_name}_实验报告"

        result_paths = generate_report_document(
            template_path=template_path,
            filled_data=filled,
            output_name=output_name,
            to_pdf=False
        )

        docx_path = result_paths.get("docx", "")
        if docx_path and Path(docx_path).exists():
            status_lines.append(f"   ✓ 文档已生成: {Path(docx_path).name}")
            status_lines.append("")
            status_lines.append("✅ **完成！请下载下方的 .docx 文件，用 WPS/Office Word 微调后直接交作业。**")

            preview = _build_preview(filled_sections, template_analysis)
            return (str(docx_path), preview, "\n".join(status_lines))
        else:
            status_lines.append("   ❌ 文档生成失败")
            return None, None, "\n".join(status_lines)

    except Exception as e:
        import traceback
        status_lines.append(f"\n❌ 处理出错: {e}")
        status_lines.append(f"```\n{traceback.format_exc()[:800]}\n```")
        return None, None, "\n".join(status_lines)


def _build_preview(filled_sections: dict, template_analysis: dict) -> str:
    """生成填充内容预览的 Markdown 文本"""
    lines = ["# 📄 实验报告内容预览\n"]
    for section_name, content in filled_sections.items():
        if content and content.strip() and content.strip() != "[待补充]":
            preview = content[:300] + ("..." if len(content) > 300 else "")
            lines.append(f"## {section_name}")
            lines.append(preview)
            lines.append("")
    if len(lines) == 1:
        lines.append("*（AI 未生成内容，请检查 API 配置或重试）*")
    return "\n".join(lines)


# ==================== Gradio 界面 ====================

def build_ui():
    with gr.Blocks(
        title="无痛实验报告生成器",
        theme=gr.themes.Soft(primary_hue="blue"),
    ) as demo:

        gr.Markdown("""
        # 🧪 无痛实验报告生成器
        > **解放工科生双手** — 上传实验指导手册 + 老师模板，AI 一键生成可交作业的实验报告
        """)

        with gr.Row():
            with gr.Column(scale=1):
                gr.Markdown("### 📥 第1步：上传实验指导手册")
                manual_input = gr.File(
                    label="实验指导手册",
                    file_types=[".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".webp"],
                    type="filepath",
                )
                gr.Markdown("*支持 PDF / 照片 / 截图*")

            with gr.Column(scale=1):
                gr.Markdown("### 📋 第2步：上传实验报告模板")
                template_input = gr.File(
                    label="老师发的模板 (.docx)",
                    file_types=[".docx"],
                    type="filepath",
                )
                gr.Markdown("*支持 .docx 格式（WPS/Word 模板）*")

        # API 配置面板
        with gr.Accordion("⚙️ API 设置（免费方案见下方说明）", open=True):
            with gr.Row():
                api_key_input = gr.Textbox(
                    label="API Key",
                    placeholder="sk-... 支持 DeepSeek / Qwen / GPT 等",
                    type="password",
                    value=os.getenv("LLM_API_KEY", "")
                )
            with gr.Row():
                api_base_input = gr.Textbox(
                    label="API 地址",
                    value="https://api.deepseek.com",
                    placeholder="https://api.deepseek.com"
                )
                model_input = gr.Textbox(
                    label="模型名称",
                    value="deepseek-chat",
                    placeholder="deepseek-chat / qwen-turbo / gpt-4o-mini"
                )
            gr.Markdown("""
            ### 🆓 免费 API 获取方式

            | 平台 | 免费额度 | 获取方式 |
            |------|----------|----------|
            | **DeepSeek** | 注册送 500万 tokens | [platform.deepseek.com](https://platform.deepseek.com) → API Keys |
            | **阿里百炼 Qwen** | 100万 tokens/月 免费 | [bailian.console.aliyun.com](https://bailian.console.aliyun.com) |
            | **硅基流动 SiliconFlow** | 14元免费额度 | [siliconflow.cn](https://siliconflow.cn) → 支持 DeepSeek/Qwen 等 |
            | **Cloudflare Workers AI** | 10万次/天 | 需配置 Worker，支持 Llama 等开源模型 |

            > ⚡ **推荐**：DeepSeek 注册即用，API 地址填 `https://api.deepseek.com`，模型 `deepseek-chat`。
            > 也可在项目目录创建 `.env` 文件永久保存配置，无需每次填入。
            """)

        with gr.Row():
            generate_btn = gr.Button("🚀 一键生成实验报告", variant="primary", size="lg")

        gr.Markdown("---")

        with gr.Row():
            with gr.Column(scale=1):
                gr.Markdown("### 📥 下载文件")
                download_output = gr.File(label="生成的实验报告 (.docx)")
            with gr.Column(scale=2):
                gr.Markdown("### 📄 内容预览")
                preview_output = gr.Markdown(value="*等待生成...*")

        with gr.Accordion("🔍 处理日志", open=False):
            status_output = gr.Textbox(label="详细日志", lines=12, max_lines=30)

        generate_btn.click(
            fn=process_files,
            inputs=[manual_input, template_input, api_key_input, api_base_input, model_input],
            outputs=[download_output, preview_output, status_output],
        )

        gr.Markdown("""
        ---
        ### 📖 使用说明

        1. **准备文件**：实验指导手册（PDF/照片）+ 老师发的 .docx 模板
        2. **获取 API Key**：上方表格任选一个免费平台注册获取
        3. **上传 + 配置**：拖入文件，填入 API Key
        4. **一键生成**：等待 10-30 秒，下载 .docx
        5. **微调交作业**：WPS / Office Word 打开微调格式即可

        ### ⚡ 技术栈

        | 环节 | 技术 |
        |------|------|
        | OCR 识别 | EasyOCR（中文+英文） |
        | PDF 解析 | PyMuPDF |
        | AI 润色填充 | DeepSeek / Qwen / GPT（OpenAI 兼容） |
        | 文档生成 | python-docx（保持原模板格式） |

        ### 🔗 链接
        - 📦 [GitHub 仓库](https://github.com/xing-crypto-code/lab-report-generator)
        """)

    return demo


if __name__ == "__main__":
    demo = build_ui()
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=False,
        show_error=True,
    )
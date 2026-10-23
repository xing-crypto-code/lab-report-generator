"""
无痛实验报告生成器 · Gradio Web 界面 v2.0
=========================================
硅基流动 SiliconFlow 多模型支持 · 工科知识库 · 精致 UI

启动：python app.py  →  浏览器打开 http://127.0.0.1:7860
"""
# ⚠️  必须最先：把 libs/ 加入 sys.path，否则找不到 pip 安装的包
import sys
from pathlib import Path
_libs = str(Path(__file__).resolve().parent / "libs")
if _libs not in sys.path:
    sys.path.insert(0, _libs)

# ⚠️  清理代理环境变量：本机代理(127.0.0.1:7890)与 NO_PROXY 中的 [::1] 语法
# 会让 httpx/requests 在 gradio 内部构建 Client 时解析 URL 崩溃（Invalid port: ':1]'）。
# 程序集 API 调用走内部配置，不受影响；HTTP(S)_PROXY 由 gen 子进程另行设置。
import os as _os
for _k in (
    "HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy",
    "ALL_PROXY", "all_proxy", "NO_PROXY", "no_proxy",
):
    _os.environ.pop(_k, None)

import gradio as gr
import os

from config import UPLOAD_DIR, OUTPUT_DIR
from modules.parser import parse_document
from modules.template import analyze_template, get_template_summary
from modules.filler import generate_report, compare_models, list_available_models, detect_experiment_type
from modules.generator import generate_report_document
from modules.knowledge import build_standard_template, EXPERIMENT_TYPES

# ─── 模型列表（延迟计算，避免导入时序问题） ───

def _get_models_sorted():
    """返回按推荐排序的模型"""
    from modules.filler import SILICONFLOW_MODELS
    order = [
        "deepseek-ai/DeepSeek-V4.1-Flash",
        "deepseek-ai/DeepSeek-V4-Flash",
        "deepseek-ai/DeepSeek-V4-Pro",
        "deepseek-ai/DeepSeek-V3",
        "deepseek-ai/DeepSeek-V3.2",
        "Qwen/Qwen3.5-35B-A3B",
        "Qwen/Qwen3.5-27B",
        "zai-org/GLM-5.3",
        "zai-org/GLM-5.3-Flash",
        "deepseek-ai/DeepSeek-R1",
    ]
    seen = set()
    result = {}
    for mid in order:
        if mid in SILICONFLOW_MODELS:
            result[mid] = SILICONFLOW_MODELS[mid]
            seen.add(mid)
    for mid, info in SILICONFLOW_MODELS.items():
        if mid not in seen:
            result[mid] = info
    return result

def _get_model_choices():
    return [(v["label"], k) for k, v in _get_models_sorted().items()]


def process_files(manual_file, template_file, api_key, api_base, model, use_builtin_template):
    """核心处理流水线"""
    if manual_file is None:
        return None, None, "❌ 请上传实验指导手册（PDF/照片）"

    if api_key and api_key.strip():
        os.environ["LLM_API_KEY"] = api_key.strip()
    if api_base and api_base.strip():
        os.environ["LLM_BASE_URL"] = api_base.strip()

    status = []

    try:
        # ── 第1步：解析实验指导手册 + 图像预处理 ──
        status.append("🔍 正在解析实验指导手册…")
        manual_path = str(manual_file)
        doc_result = parse_document(manual_path)
        ocr_text = doc_result["text"]
        ocr_tables = doc_result["tables"]
        status.append(f"   ✓ 识别 {len(ocr_text):,} 字符，{len(ocr_tables)} 个表格")

        # 自动检测实验类型
        exp_type = detect_experiment_type(ocr_text)
        status.append(f"   🔬 检测实验类型：{exp_type}")
        status.append("")

        # ── 第2步：获取模板 ──
        if template_file is not None:
            status.append("📋 正在解析上传的模板…")
            template_path = str(template_file)
            template_analysis = analyze_template(template_path)
            t_summary = get_template_summary(template_analysis)
            status.append(f"   ✓ {template_analysis['total_paragraphs']} 段落，{template_analysis['total_tables']} 表格")
        else:
            status.append("📋 使用内置标准模板…")
            tpl_bytes = build_standard_template(exp_type)
            tpl_path = Path("temp") / "_builtin_template.docx"
            tpl_path.write_bytes(tpl_bytes)
            template_path = str(tpl_path)
            template_analysis = analyze_template(template_path)
            t_summary = get_template_summary(template_analysis)
            status.append(f"   ✓ 已生成 {exp_type} 标准模板")
        status.append("")

        # ── 第3步：AI 填充 ──
        status.append(f"🤖 正在调用 {model} 生成实验报告…")
        filled = generate_report(ocr_text, ocr_tables, t_summary, model=model)
        sections = filled.get("sections", {})
        count = sum(1 for v in sections.values() if v and v.strip() and "[待补充]" not in v)
        status.append(f"   ✓ AI 已填充 {count} 个章节")
        status.append(f"   📝 模型：{filled.get('model_used', model)}")
        if filled.get("notes"):
            status.append(f"   ℹ️ {filled['notes'][:100]}")
        status.append("")

        # ── 第4步：生成文档 ──
        status.append("📝 正在生成最终文档…")
        manual_name = Path(manual_path).stem
        output_name = f"{manual_name}_实验报告"

        result_paths = generate_report_document(
            template_path=template_path,
            filled_data=filled,
            output_name=output_name,
        )
        docx_path = result_paths.get("docx", "")
        if docx_path and Path(docx_path).exists():
            status.append(f"   ✓ 文档已生成：{Path(docx_path).name}")
            status.append("")
            status.append("✅ 完成！请下载 .docx 用 WPS/Office 打开微调后交作业。")
            preview = _build_preview(sections, template_analysis)
            return (str(docx_path), preview, "\n".join(status))
        else:
            return None, None, "\n".join(status + ["❌ 生成失败"])

    except Exception as e:
        import traceback
        status.append(f"\n❌ 出错：{e}")
        status.append(f"```\n{traceback.format_exc()[:600]}\n```")
        return None, None, "\n".join(status)


def _build_preview(sections, analysis):
    lines = ["# 📄 实验报告预览\n"]
    for name, text in sections.items():
        if text and text.strip() and "[待补充]" not in text:
            p = text[:250] + ("…" if len(text) > 250 else "")
            lines.append(f"## {name}")
            lines.append(p)
            lines.append("")
    return "\n".join(lines) if len(lines) > 1 else "*（等待生成…）*"


# ═══════════════════  UI  ═══════════════════

CSS = """
.gradio-container { max-width: 900px !important; margin: 0 auto; }
.header-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 24px; color: white; margin-bottom: 20px; }
.header-card h1 { font-size: 1.8em; margin: 0; }
.header-card p { opacity: 0.9; margin: 8px 0 0 0; }
.step-label { font-weight: 600; color: #374151; margin-bottom: 4px; }
footer { visibility: hidden; }
"""


def build_ui():
    with gr.Blocks(title="无痛实验报告生成器", theme=gr.themes.Soft(primary_hue="indigo"), css=CSS) as demo:

        gr.HTML("""
        <div class="header-card">
            <h1>🧪 无痛实验报告生成器</h1>
            <p>上传实验指导手册（PDF/照片）+ 老师模板（可选），AI 一键生成标准实验报告</p>
        </div>
        """)

        # ── 上传区 ──
        with gr.Row(equal_height=True):
            with gr.Column(scale=1):
                gr.Markdown("#### 📥 实验指导手册")
                manual_in = gr.File(
                    label="支持 PDF · Word · 图片 · 截图",
                    file_types=[".pdf", ".docx", ".doc", ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp"],
                    type="filepath",
                )

            with gr.Column(scale=1):
                gr.Markdown("#### 📋 实验报告模板（可选）")
                template_in = gr.File(
                    label="支持 .docx / .doc 模板",
                    file_types=[".docx", ".doc"],
                    type="filepath",
                )
                use_builtin = gr.Checkbox(label="☑ 没有模板 → 使用内置标准模板", value=True)

        # ── 模型 & API ──
        with gr.Accordion("⚙️ AI 模型 & API 设置", open=True):
            with gr.Row():
                model_sel = gr.Dropdown(
                    label="🤖 AI 模型",
                    choices=_get_model_choices(),
                    value="deepseek-ai/DeepSeek-V4.1-Flash",
                    interactive=True,
                    scale=3,
                )
                with gr.Column(scale=1):
                    api_key_in = gr.Textbox(
                        label="API Key（留空=用 .env）",
                        type="password",
                        placeholder="sk-…",
                    )
                    api_base_in = gr.Textbox(
                        label="API 地址",
                        value="https://api.siliconflow.cn/v1",
                    )

        # ── 生成按钮 ──
        with gr.Row():
            btn = gr.Button("🚀 一键生成实验报告", variant="primary", size="lg", scale=1)

        gr.Markdown("---")

        # ── 输出区 ──
        with gr.Row():
            with gr.Column(scale=1):
                gr.Markdown("#### 📥 下载")
                download_out = gr.File(label="生成的实验报告 (.docx)")
            with gr.Column(scale=2):
                gr.Markdown("#### 📄 内容预览")
                preview_out = gr.Markdown("*等待生成…*")

        with gr.Accordion("🔍 处理日志", open=False):
            log_out = gr.Textbox(label="日志", lines=10, max_lines=25)

        btn.click(
            fn=process_files,
            inputs=[manual_in, template_in, api_key_in, api_base_in, model_sel, use_builtin],
            outputs=[download_out, preview_out, log_out],
        )

        # ── 说明 ──
        gr.Markdown("""
        ---
        ### 📖 快速上手

        | 步骤 | 操作 |
        |:---:|---|
        | ① | 上传 **实验指导手册**（PDF 或拍照） |
        | ② | 上传老师的 **报告模板** .docx（或勾选内置模板） |
        | ③ | 选择 **AI 模型**（推荐 DeepSeek V3） |
        | ④ | 点击 **一键生成**，下载 .docx 用 WPS 微调 |

        ### 🆓 免费使用

        API 走 **硅基流动 SiliconFlow**，新用户送 14 元额度，够生成数百份报告。

        ### 🔧 内置知识库

        | 实验类型 | 标准章节 |
        |---|---|
        | 物理实验 | 目的 → 原理 → 器材 → 步骤 → 数据 → 处理 → 误差 → 结论 |
        | 化学实验 | 目的 → 反应方程式 → 试剂仪器 → 步骤 → 数据 → 讨论 → 结论 |
        | 电路实验 | 目的 → 电路图 → 仪器器件 → 步骤 → 波形/数据 → 误差分析 |
        | 力学实验 | 目的 → 公式推导 → 仪器 → 步骤 → 数据图表 → 误差 → 结论 |

        > AI 会自动识别实验类型并匹配相应的标准结构。OCR 前会进行图像去噪/增强/二值化预处理，提升识别率。
        """)

    return demo


if __name__ == "__main__":
    demo = build_ui()
    demo.launch(server_name="0.0.0.0", server_port=7860, share=False, show_error=True)
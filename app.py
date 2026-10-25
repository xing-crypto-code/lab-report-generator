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

from config import UPLOAD_DIR, OUTPUT_DIR, LLM_MODEL, LLM_BASE_URL, BASE_DIR, TEMPLATE_DIR
from modules.parser import parse_document
from modules.template import analyze_template, get_template_summary
from modules.filler import generate_report, compare_models, list_available_models, detect_experiment_type
from modules.generator import generate_report_document
from modules.knowledge import build_standard_template, EXPERIMENT_TYPES
from modules.pdf_template import convert_pdf_to_docx
from modules.render import render_docx
from modules import history as history_mod
from modules import chat as chat_mod

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

def _default_model() -> str:
    """返回当前生效的默认模型（优先 .env 的 LLM_MODEL）"""
    return (LLM_MODEL or "deepseek-chat").strip() or "deepseek-chat"


def _get_model_choices():
    """返回带标签的模型选择。DeepSeek 官方模型 + 硅基流动模型都会出现在列表，
    .env 当前默认模型不在列表中时自动补上。"""
    # DeepSeek 官方模型
    deepseek_official = {
        "deepseek-chat":     {"label": "DeepSeek-chat（V3 对话）", "desc": "DeepSeek 官方对话模型"},
        "deepseek-reasoner": {"label": "DeepSeek-reasoner（R1 推理）", "desc": "DeepSeek 官方推理模型"},
    }
    from modules.filler import SILICONFLOW_MODELS

    # 合并：官方 + 硅基
    merged = dict(deepseek_official)
    merged.update(SILICONFLOW_MODELS)

    # 排序：DeepSeek 官方优先，其次 .env 默认，再硅基推荐
    order = [
        "deepseek-chat", "deepseek-reasoner",
        _default_model(),
    ]
    # 再加之前硅基推荐顺序
    order += [
        "deepseek-ai/DeepSeek-V4.1-Flash",
        "deepseek-ai/DeepSeek-V4-Flash",
        "deepseek-ai/DeepSeek-V4-Pro",
        "deepseek-ai/DeepSeek-V3",
        "deepseek-ai/DeepSeek-V3.2",
    ]

    seen = set()
    result = {}
    for mid in order:
        if mid in merged and mid not in seen:
            result[mid] = merged[mid]
            seen.add(mid)
    for mid, info in merged.items():
        if mid not in seen:
            result[mid] = info
            seen.add(mid)
    return [(v["label"], k) for k, v in result.items()]


def process_files(manual_file, template_file, api_key, api_base, model, use_builtin_template,
                  requirement=""):
    """核心处理流水线，返回 (下载文件路径, 下载状态文本, 预览HTML, 日志)
    manual_file / template_file 现在支持多文件（传入列表）。
    requirement: 用户在同一个输入框里顺带写的要求（会并入 AI 提示）
    """

    def _norm(f):
        """把单文件名或多文件列表统一成列表"""
        if f is None:
            return []
        if isinstance(f, (list, tuple)):
            return [str(x) for x in f]
        return [str(f)]

    manuals = _norm(manual_file)
    templates = _norm(template_file)

    if not manuals:
        return gr.update(), "*请上传文件后生成*", "*请上传实验指导手册（可多个）*", "❌ 请上传实验指导手册（PDF/图片/Word/Markdown，可多选）"

    if api_key and api_key.strip():
        os.environ["LLM_API_KEY"] = api_key.strip()
    if api_base and api_base.strip():
        os.environ["LLM_BASE_URL"] = api_base.strip()

    status = []

    try:
        # ── 第1步：解析实验指导手册（支持多文件）──
        status.append(f"🔍 正在解析 {len(manuals)} 个实验指导手册…")
        merged_text_parts = []
        merged_md_parts = []
        merged_tables = []
        for idx, mf in enumerate(manuals, 1):
            name = Path(mf).stem
            ext = Path(mf).suffix.lower()
            _is_img = ext in ('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.webp')
            if _is_img:
                status.append(f"   🖼️ [{idx}/{len(manuals)}] 《{Path(mf).name}》 OCR 识别中…"
                              f"（CPU 模式较慢，单张最多 90 秒）")
            else:
                status.append(f"   📄 [{idx}/{len(manuals)}] 《{Path(mf).name}》 解析中…")
            try:
                doc_result = parse_document(mf)
            except ValueError as e:
                status.append(f"   ⚠️ 《{name}》解析失败：{e}")
                continue
            ocr_text = doc_result["text"]
            ocr_tables = doc_result.get("tables", [])
            merged_text_parts.append(f"【来源：{Path(mf).name}】\n{ocr_text}")
            merged_tables.extend(ocr_tables)
            # 结构化 Markdown（PDF 经 pymupdf4llm 抽取）——让 AI 能原样沿用
            # 指导书里的「实验目的/实验仪器」小节
            _md = (doc_result.get("markdown") or "").strip()
            if _md:
                merged_md_parts.append(f"【来源：{Path(mf).name} 结构化原文】\n{_md}")
            status.append(f"   ✓ 《{Path(mf).name}》识别 {len(ocr_text):,} 字符，{len(ocr_tables)} 个表格"
                          + (f"，结构化 {len(_md)} 字符" if _md else ""))
        ocr_text = "\n\n".join(merged_text_parts)
        source_markdown = "\n\n".join(merged_md_parts)
        if not ocr_text.strip():
            return gr.update(), "*解析失败*", "*未能从文件中识别出文字*", "❌ 未能从上传文件中识别到文字内容"

        # 自动检测实验类型
        exp_type = detect_experiment_type(ocr_text)
        status.append(f"   🔬 检测实验类型：{exp_type}")
        status.append("")

        # ── 第2步：获取模板（支持多文件，优先 .docx，PDF 自动转 docx）──
        if templates:
            # 优先级：.docx > .doc > .md > .pdf（PDF 会转成 docx 再填）
            tplate_rank = {".docx": 0, ".doc": 1, ".md": 2, ".pdf": 3}
            templates_sorted = sorted(
                templates,
                key=lambda p: tplate_rank.get(Path(p).suffix.lower(), 9),
            )
            template_path = templates_sorted[0]
            status.append(f"📋 正在解析上传的模板（{Path(template_path).name}）…")
            # PDF 模板 → 转成可填充的 docx 线框（pdfplumber 提取表格+文本重建）
            if Path(template_path).suffix.lower() == ".pdf":
                try:
                    template_path = convert_pdf_to_docx(
                        template_path, str(Path("temp") / f"_pdf_tpl_{os.getpid()}.docx")
                    )
                    status.append(f"   🔁 PDF 模板已转成可编辑 docx 线框（{Path(template_path).name}）")
                except Exception as pe:
                    status.append(f"   ⚠️ PDF 模板解析失败：{pe}")
                    template_path = ""
            template_analysis = analyze_template(template_path) if template_path else {"total_paragraphs": 0, "total_tables": 0, "paragraphs": [], "tables": [], "sections_found": [], "placeholder_count": 0}
            t_summary = get_template_summary(template_analysis)
            status.append(f"   ✓ {template_analysis['total_paragraphs']} 段落，{template_analysis['total_tables']} 表格")
            if len(templates) > 1:
                status.append(f"   ℹ️ 上传了 {len(templates)} 个模板，已选用：{Path(templates_sorted[0]).name}")
            if not template_path:
                status.append("   ⚠️ 模板为空，回退使用内置模板")
                if Path(TEMPLATE_DIR / "内置实验报告模板.docx").exists():
                    template_path = str(TEMPLATE_DIR / "内置实验报告模板.docx")
                    template_analysis = analyze_template(template_path)
                    t_summary = get_template_summary(template_analysis)
        else:
            # 优先使用内置的成都大学模板（已优化、质量高）；不存在则回退到通用生成模板
            builtin = Path(TEMPLATE_DIR) / "内置实验报告模板.docx"
            if builtin.exists():
                status.append("📋 使用内置系统模板（成都大学标准线框模板）…")
                template_path = str(builtin)
                # 自动校验内置模板是否已固化最新修复；未固化则就地烘焙
                try:
                    from modules.template_fix import verify_template, rebake_template
                    _probs = verify_template(str(builtin))
                    if _probs:
                        status.append(f"   🔧 内置模板有 {len(_probs)} 处未固化，正在自动烘焙…")
                        rebake_template(str(builtin))
                        _probs2 = verify_template(str(builtin))
                        if _probs2:
                            status.append(f"   ⚠️ 烘焙后仍有 {len(_probs2)} 处问题")
                        else:
                            status.append("   ✓ 模板已烘焙到最新版")
                except Exception as _tfe:
                    status.append(f"   ⚠️ 模板自检跳过：{_tfe}")
            else:
                status.append("📋 使用内置标准模板…")
                tpl_bytes = build_standard_template(exp_type)
                tpl_path = Path("temp") / "_builtin_template.docx"
                tpl_path.write_bytes(tpl_bytes)
                template_path = str(tpl_path)
            template_analysis = analyze_template(template_path)
            t_summary = get_template_summary(template_analysis)
            status.append(f"   ✓ {template_analysis['total_paragraphs']} 段落，{template_analysis['total_tables']} 表格")
        status.append("")

        # ── 第3步：AI 填充 ──
        status.append(f"🤖 正在调用 {model} 生成实验报告…")
        if requirement and requirement.strip():
            status.append(f"   📌 附加要求：{requirement.strip()[:80]}")
        # 显式传入 UI 的 api_base/api_key（或 None → 按模型自动路由到对应 provider）
        filled = generate_report(
            ocr_text, ocr_tables, t_summary, model=model,
            api_key=api_key, api_base=api_base,
            requirement=requirement,
            source_markdown=source_markdown,
        )
        sections = filled.get("sections", {})
        count = sum(1 for v in sections.values() if v and v.strip() and "[待补充]" not in v)
        status.append(f"   ✓ AI 已填充 {count} 个章节")
        status.append(f"   📝 模型：{filled.get('model_used', model)}")
        if filled.get("notes"):
            status.append(f"   ℹ️ {filled['notes'][:120]}")
        status.append("")

        # ── 关键兜底：AI 完全没填出任何章节 → 中止，不要产出空白模板副本 ──
        if count == 0:
            err_note = filled.get("notes") or "AI 未能生成内容"
            status.append("🚫 未生成任何有效内容，已中止，避免输出空模板。")
            return (
                gr.update(),
                "❌ 生成失败：AI 没有返回内容",
                "*AI 未返回可填写的内容，请查看日志调整模型或 API*",
                "\n".join(status),
            )

        # ── 第4步：生成文档 ──
        status.append("📝 正在生成最终文档…")
        manual_name = Path(manuals[0]).stem
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
            status.append("✅ 完成！点击下载按钮用 WPS/Office 打开微调后交作业。")
            # 渲染真实 docx 内容（含表格结构）供网页预览
            try:
                preview = render_docx(docx_path)
                preview = (
                    f"<div class='r-wrap'><div class='r-toolbar'>"
                    f"📄 实时预览 <span style='opacity:.6'>· 由生成的 {Path(docx_path).name} 渲染</span>"
                    f"</div>{preview}</div>"
                )
            except Exception as _re:
                print("[render] 预览渲染失败:", _re)
                preview = _build_preview(sections, template_analysis)
            # 保存生成上下文供"对话修改"复用
            try:
                chat_mod.save_context(
                    ocr_text=ocr_text, ocr_tables=merged_tables,
                    template_summary=t_summary, sections=sections,
                    model=model, experiment_type=exp_type,
                    template_path=template_path,
                    output_base=output_name,
                )
            except Exception as _ce:
                print("[chat] 保存上下文失败:", _ce)
            # ── 写入历史记录（供日后免费用重新下载，不耗 token）──
            try:
                history_mod.add_record({
                    "title": f"{Path(docx_path).stem}",
                    "model": filled.get("model_used", model),
                    "docx": str(docx_path),
                })
            except Exception as _he:
                print("[history] 记录失败:", _he)
            safe_name = Path(docx_path).stem.replace(" ", "_")
            return (str(docx_path), f"✅ 已生成，点击上方按钮下载 **{safe_name}.docx**（保存位置可自选）", preview, "\n".join(status))
        else:
            return gr.update(), "❌ 生成失败", "*生成失败，请查看下方日志*", "\n".join(status + ["❌ 生成失败"])

    except Exception as e:
        import traceback
        status.append(f"\n❌ 出错：{e}")
        status.append(f"```\ntraceback.format_exc()[:600]```")
        err = f"*❌ 生成出错：{e}*"
        return gr.update(), "❌ 出现错误", err, "\n".join(status)


def refresh_history_ui():
    """返回最新的历史面板 HTML，供刷新按钮/生成后调用"""
    return history_mod.render_history_html()


def composer_submit(message, template_file, api_key, api_base, model, use_builtin):
    """统一入口（DeepSeek-Harness 式单框）：
    · 带附件  → 走生成流程（附件 = 实验指导手册，可多选）
    · 不带附件 → 走对话修订（已生成过报告时，按文字要求修改并重新出稿）
    返回 (聊天HTML, 下载文件, 下载状态, 预览HTML, 日志)
    """
    _nochange = gr.update()
    msg = message or {}
    if isinstance(msg, dict):
        text = (msg.get("text") or "").strip()
        raw_files = msg.get("files") or []
    else:
        text, raw_files = str(msg or "").strip(), []

    files = []
    for f in raw_files:
        p = None
        if isinstance(f, str):
            p = f
        elif isinstance(f, dict):
            p = f.get("path") or f.get("name")
        else:
            p = (getattr(f, "path", None) or getattr(f, "name", None)
                 or str(f))
        if p:
            files.append(str(p))

    # ── A. 带文件 → 生成报告 ──
    if files:
        chat_mod._session.append((
            "user",
            (text + "\n" if text else "") + f"📎 已上传 {len(files)} 个文件：" +
            "、".join(Path(p).name for p in files),
        ))
        try:
            dl, dl_status, preview, log = process_files(
                files, template_file, api_key, api_base, model, use_builtin,
                requirement=text,
            )
        except Exception as e:
            chat_mod._session.append(("ai", f"❌ 生成出错：{e}"))
            return chat_mod.render_chat_preview(chat_mod._session), _nochange, "❌ 出错", _nochange, _nochange
        if dl:
            chat_mod._session.append(("ai", "✅ 报告已生成，可在上方预览；点「下载 .docx」保存。\n"
                                           "想改？在输入框继续提要求即可。"))
        else:
            chat_mod._session.append(("ai", f"❌ 生成失败：{dl_status}"))
        return chat_mod.render_chat_preview(chat_mod._session), dl, dl_status, preview, log

    # ── B. 无文件且无文字 → 忽略 ──
    if not text:
        return chat_mod.render_chat_preview(chat_mod._session), _nochange, _nochange, _nochange, _nochange

    # ── C. 无文件、有文字 → 对话修订 ──
    if not chat_mod.get_context().get("sections"):
        chat_mod._session.append(("user", text))
        chat_mod._session.append(("ai", "⚠️ 还没有报告可以改。请先用输入框左侧 📎 上传实验指导书，"
                                        "我会先生成第一版。"))
        return chat_mod.render_chat_preview(chat_mod._session), _nochange, _nochange, _nochange, _nochange

    chat_mod._session.append(("user", text))
    res = chat_mod.run_chat_revision(
        text, template_path=chat_mod.get_context().get("template_path", "")
    )
    if not res["ok"]:
        chat_mod._session.append(("ai", "❌ " + (res.get("msg") or "修改失败，请查看日志")))
        return chat_mod.render_chat_preview(chat_mod._session), _nochange, _nochange, _nochange, _nochange

    docx = res.get("docx", "")
    chat_mod._session.append(("ai", f"✅ {res.get('msg', '已重新生成')}"))
    return (
        chat_mod.render_chat_preview(chat_mod._session),
        docx or _nochange,
        f"✅ 已按你的要求重新生成 **{Path(docx).name}**" if docx else _nochange,
        res.get("preview", "") or _nochange,
        "💬 对话修订完成，已重新出稿",
    )


def _build_preview(sections, analysis) -> str:
    """将 AI 填充内容渲染为带样式的 HTML 预览"""
    if not sections:
        return "<p style='color:#888'>（无内容，等待生成…）</p>"
    parts = ["<div style='font-family:Songti,SimSun,serif'>"]
    has_any = False
    for name, text in sections.items():
        if text and text.strip() and "[待补充]" not in text:
            has_any = True
            p = text[:260] + ("…" if len(text) > 260 else "")
            # 转义 HTML 特殊字符，避免渲染串位
            import html
            p = html.escape(p).replace("\n", "<br/>")
            parts.append(f"<h3 style='color:#4f46e5;margin:14px 0 6px 0;border-bottom:1px solid #e5e7eb'>{html.escape(name)}</h3>")
            parts.append(f"<p style='margin:4px 0 10px 12px;line-height:1.7;color:#374151'>{p}</p>")
    parts.append("</div>")
    if not has_any:
        return "<p style='color:#888'>（AI 暂未填充到有效内容，请查看日志）</p>"
    return "".join(parts)


# ═══════════════════  UI  ═══════════════════

# ── 像素风 / 精简终端风（星露谷式像素字体 + 亮/暗双主题一键切换）──
CSS = r"""
/* ── 浅色（默认）── */
:root {
  --bg0: #ffffff;          /* 页面底 */
  --bg1: #fafafa;          /* 卡片 */
  --bg2: #f4f4f5;          /* hover / 次级块 */
  --line: #e4e4e7;         /* 常规边框 */
  --line2: #d4d4d8;        /* 强调边框 */
  --acc: #2563eb;          /* 主色（克制蓝） */
  --acc-hover: #1d4ed8;
  --ink: #18181b;          /* 正文 */
  --ink2: #3f3f46;         /* 次级正文 */
  --dim: #71717a;          /* 说明文字 */
  --err: #dc2626;
  --ok: #16a34a;
  --soft: #ffffff;         /* 输入底 */
  --radius: 10px;
  --radius-sm: 8px;
  --shadow: 0 1px 2px rgba(0,0,0,.04);
  --font: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC',
          'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
  --mono: ui-monospace, SFMono-Regular, 'SF Mono', Consolas, monospace;
  --song: 'SimSun', '宋体', serif;
}

/* ── 深色 ── */
:root[data-ui-theme="dark"] {
  --bg0: #09090b;
  --bg1: #111113;
  --bg2: #18181b;
  --line: #27272a;
  --line2: #3f3f46;
  --acc: #3b82f6;
  --acc-hover: #60a5fa;
  --ink: #fafafa;
  --ink2: #d4d4d8;
  --dim: #a1a1aa;
  --err: #f87171;
  --ok: #4ade80;
  --soft: #0c0c0e;
  --shadow: none;
}

/* ── 基础排版 ── */
html, body, .gradio-container {
  background: var(--bg0) !important;
  color: var(--ink);
  font-family: var(--font) !important;
  -webkit-font-smoothing: antialiased;
}
.gradio-container {
  max-width: 900px !important;
  margin: 0 auto !important;
  padding: 0 20px 64px !important;
  font-size: 14px;
}
:root[data-ui-theme="dark"] .gradio-container { color-scheme: dark; }

/* ── 顶栏：简洁标题 + 副标题 ── */
.pix-header {
  padding: 40px 0 28px;
  margin: 0;
  border: none;
  background: none;
  box-shadow: none;
  display: block;
}
.pix-header h1 {
  font-family: var(--font);
  font-size: 22px;
  font-weight: 600;
  letter-spacing: -0.01em;
  color: var(--ink);
  margin: 0 0 6px;
  text-shadow: none;
}
.pix-header .sub {
  font-family: var(--font);
  font-size: 13px;
  color: var(--dim);
  margin: 0;
  letter-spacing: 0;
  font-weight: 400;
}
.pix-header .pix-dot { color: var(--line2); margin: 0 6px; }
.pix-caret { display: none; }   /* 去掉闪烁光标 */
.pix-dot { color: var(--line2); }

/* ── 主题切换：右上角一个极简图标按钮 ── */
#themeBtn {
  position: fixed !important;
  top: 18px; right: 20px;
  z-index: 50;
  width: 34px !important; height: 34px;
  min-width: 34px !important;
  padding: 0 !important;
  font-size: 15px !important;
  line-height: 1;
  display: flex; align-items: center; justify-content: center;
  background: var(--bg1) !important;
  color: var(--dim) !important;
  border: 1px solid var(--line) !important;
  border-radius: 999px !important;
  box-shadow: var(--shadow) !important;
  cursor: pointer;
  transition: all .15s;
}
#themeBtn:hover {
  color: var(--ink) !important;
  border-color: var(--line2) !important;
  transform: none !important;
  box-shadow: var(--shadow) !important;
}

/* ── 通用块：去掉所有像素装饰 ── */
.block, .form, .gr-panel, .panel,
.gradio-container .block {
  background: transparent !important;
  border: none !important;
  border-radius: 0 !important;
  box-shadow: none !important;
  padding: 0 !important;
}

/* ── 输入框 ── */
textarea, input[type=text], input[type=password], input[type=email],
.gradio-container input, .gradio-container textarea, .gradio-container select {
  background: var(--soft) !important;
  border: 1px solid var(--line) !important;
  border-radius: var(--radius-sm) !important;
  color: var(--ink) !important;
  font-family: var(--font) !important;
  font-size: 14px !important;
  padding: 9px 12px !important;
  box-shadow: none !important;
  transition: border-color .15s;
}
textarea:focus, input:focus {
  border-color: var(--acc) !important;
  box-shadow: 0 0 0 3px rgba(37,99,235,.10) !important;
  outline: none !important;
}

/* ── 按钮：Codex 式扁平 ── */
button:not(#themeBtn), .gradio-button, [type=button] {
  font-family: var(--font) !important;
  font-size: 13.5px !important;
  font-weight: 500 !important;
  letter-spacing: 0 !important;
  text-transform: none !important;
  border-radius: var(--radius-sm) !important;
  border: 1px solid var(--line) !important;
  background: var(--bg1) !important;
  color: var(--ink) !important;
  box-shadow: none !important;
  padding: 9px 16px !important;
  transition: background .15s, border-color .15s, color .15s;
  transform: none !important;
}
button:not(#themeBtn):hover {
  background: var(--bg2) !important;
  border-color: var(--line2) !important;
  transform: none !important;
  box-shadow: none !important;
}
button.primary {
  background: var(--ink) !important;
  color: var(--bg0) !important;
  border-color: var(--ink) !important;
  font-weight: 500 !important;
}
button.primary:hover {
  background: var(--ink2) !important;
  border-color: var(--ink2) !important;
  filter: none !important;
}

/* ── 主输入框（composer）：Codex 输入条 ── */
#composerBox {
  border: 1px solid var(--line2) !important;
  border-radius: 14px !important;
  background: var(--soft) !important;
  box-shadow: 0 2px 8px rgba(0,0,0,.05) !important;
  padding: 2px !important;
  transition: border-color .15s, box-shadow .15s;
}
:root[data-ui-theme="dark"] #composerBox { box-shadow: none !important; }
#composerBox:focus-within {
  border-color: var(--acc) !important;
  box-shadow: 0 0 0 3px rgba(37,99,235,.10) !important;
}
/* 内部所有层都要透明白底，避免 gradio 默认深色块 */
#composerBox * { background-color: transparent !important; }
#composerBox textarea, #composerBox input {
  font-family: var(--font) !important;
  font-size: 14.5px !important;
  line-height: 1.6 !important;
  border: none !important;
  background: transparent !important;
  box-shadow: none !important;
  color: var(--ink) !important;
  padding: 8px 10px !important;
}
#composerBox textarea:focus { box-shadow: none !important; }
/* 发送按钮（真实类名 .submit-button）：圆形实心 */
#composerBox button.submit-button,
button.submit-button {
  border-radius: 999px !important;
  width: 32px !important; min-width: 32px !important; height: 32px !important;
  padding: 0 !important;
  background: var(--ink) !important;
  color: var(--bg0) !important;
  border: none !important;
  box-shadow: none !important;
}
#composerBox button.submit-button:hover,
button.submit-button:hover {
  background: var(--ink2) !important;
}
#composerBox button.submit-button svg { color: var(--bg0) !important; fill: var(--bg0) !important; }

/* 左侧附件按钮（真实类名 .upload-button.icon-button）：无底色图标 */
#composerBox button.icon-button,
button.icon-button {
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
  color: var(--dim) !important;
  width: 32px !important; min-width: 32px !important; height: 32px !important;
  padding: 0 !important;
  border-radius: 8px !important;
}
#composerBox button.icon-button:hover,
button.icon-button:hover {
  background: var(--bg2) !important;
  color: var(--ink) !important;
}

/* ── 上传区/标签：去掉黑底反色 ── */
.gradio-container label,
.gradio-container .block-title,
.gradio-container span[data-testid="block-info"],
[data-testid="block-info"],
label > span {
  background: transparent !important;
  color: var(--dim) !important;
  font-size: 12.5px !important;
  font-weight: 500 !important;
  padding: 0 2px !important;
}
/* gradio 有时给 label 套一层深色胶囊 */
.gradio-container .label-wrap,
.gradio-container .label-wrap > span {
  background: transparent !important;
  color: var(--dim) !important;
  border: none !important;
  box-shadow: none !important;
}

/* ── 帮助/说明文字：提高对比度 ── */
.gradio-container .info,
.gradio-container .description,
.gradio-container p.info,
div[class*="info"] {
  color: var(--dim) !important;
  font-size: 12.5px !important;
}

/* ── 折叠面板：极简行 ── */
details, .accordion, [data-testid="accordion"] {
  border: none !important;
  border-top: 1px solid var(--line) !important;
  border-radius: 0 !important;
  background: transparent !important;
  margin-top: 24px !important;
  padding-top: 4px !important;
}
.accordion-header, .category, [data-testid="accordion"] summary {
  font-family: var(--font) !important;
  font-size: 13px !important;
  font-weight: 500 !important;
  letter-spacing: 0 !important;
  color: var(--dim) !important;
  padding: 10px 0 !important;
  transition: color .15s;
}
.accordion-header:hover, [data-testid="accordion"] summary:hover {
  color: var(--ink) !important;
}

/* ── 上传控件：细虚线、小高度 ── */
.file-preview, [data-testid="file"], .upload-button, [data-testid="file-upload"] {
  border: 1px dashed var(--line2) !important;
  border-radius: var(--radius-sm) !important;
  background: transparent !important;
  min-height: 64px !important;
  box-shadow: none !important;
}
.upload-button {
  display: flex; align-items: center; justify-content: center;
  padding: 8px 12px !important;
  color: var(--dim) !important;
  font-size: 12.5px !important;
}
.upload-button:hover { border-color: var(--acc) !important; }
[data-testid="file-upload"] { font-size: 12.5px !important; color: var(--dim) !important; }
[data-testid="file"] > div > div { padding: 2px 4px !important; }

/* ── 下拉 ── */
.gradio-dropdown, .wrap, .wrap-inner {
  font-family: var(--font) !important;
  border-radius: var(--radius-sm) !important;
}

/* ── 分隔线 / 说明 ── */
hr { border: none !important; border-top: 1px solid var(--line) !important; opacity: 1; margin: 32px 0 !important; }
.hint { font-family: var(--font); font-size: 12px; color: var(--dim); }

/* ── 标题层级 ── */
.gradio-container h3, .gradio-container h4, .gradio-container h5 {
  font-family: var(--font) !important;
  font-size: 13px !important;
  font-weight: 600 !important;
  color: var(--ink) !important;
  letter-spacing: 0 !important;
  margin: 20px 0 10px !important;
}

/* ── markdown 表格（说明区）── */
.gradio-markdown table {
  border-collapse: collapse; width: 100%;
  font-size: 13px;
  border: 1px solid var(--line);
  border-radius: var(--radius-sm);
  overflow: hidden;
}
.gradio-markdown th, .gradio-markdown td {
  border-bottom: 1px solid var(--line);
  border-left: none; border-right: none;
  padding: 8px 12px;
  text-align: left;
}
.gradio-markdown th {
  background: var(--bg2);
  color: var(--ink);
  font-family: var(--font);
  font-size: 12.5px;
  font-weight: 600;
}
.gradio-markdown td { color: var(--ink2); }

/* ── 历史记录卡片 ── */
.hist-card {
  border: 1px solid var(--line);
  border-radius: var(--radius-sm);
  background: var(--bg1);
  padding: 12px 14px;
  margin-bottom: 8px;
  border-left: 1px solid var(--line);
  box-shadow: var(--shadow);
  transition: border-color .15s;
}
.hist-card:hover { border-color: var(--line2); }
.hist-card .h-title { font-size: 14px; font-weight: 500; color: var(--ink); }
.hist-card .h-meta { font-size: 12px; color: var(--dim); font-family: var(--font); }
.hist-card a { color: var(--acc); text-decoration: none; font-weight: 500; }
.hist-card a:hover { color: var(--acc-hover); }

/* ── 状态栏 ── */
.pix-statusbar {
  border-top: none;
  margin-top: 0;
  padding: 8px 2px;
  font-family: var(--font);
  font-size: 13px;
  color: var(--dim);
  min-height: 0;
  line-height: 1.6;
}
/* 预览列在空状态下不要占高度 */
.out-row .html-container:empty,
.out-row > div { min-height: 0 !important; }
.gradio-container .html-container { min-height: 0 !important; padding: 0 !important; }
.gradio-container .html-container > div:empty { display: none; }

/* ── 报告实时预览 ── */
.r-wrap {
  max-height: 620px; overflow: auto;
  background: #ffffff; color: #222;
  border: 1px solid var(--line2);
  border-radius: var(--radius);
  padding: 12px 16px;
  box-shadow: var(--shadow);
  font-family: var(--song);
}
:root[data-ui-theme="dark"] .r-wrap { background: #fafafa; }
.r-toolbar {
  font-family: var(--font); font-size: 12px; color: var(--dim);
  border-bottom: 1px solid #e4e4e7; padding-bottom: 8px; margin-bottom: 10px;
  position: sticky; top: 0; background: inherit; z-index: 2;
}
.r-page { line-height: 1.5; }
.r-page .r-para { margin: 3px 0; font-size: 12px; }
.r-table {
  width: 100%; border-collapse: collapse; margin: 6px 0;
  font-size: 11px; page-break-inside: avoid;
}
.r-table td {
  border: 1px solid #a1a1aa; padding: 4px 6px; vertical-align: top;
  background: #fff;
}
.r-table .r-para { margin: 1px 0; }
.r-nested td { border: 1px solid #d4d4d8; }

/* 预览虚线占位框 */
.r-ph {
  border: 1.5px dashed #b4b4bb;
  border-radius: 6px;
  background: #fafafa;
  margin: 10px auto;
  display: flex;
  align-items: center;
  justify-content: center;
  box-sizing: border-box;
}
.r-ph span {
  color: #9ca3af;
  font-family: var(--song);
  font-size: 12.5px;
  letter-spacing: .5px;
}

/* ── 对话气泡：Codex 式 ── */
.chat-empty {
  color: var(--dim);
  font-style: normal;
  padding: 4px 2px 20px;
  font-family: var(--font);
  font-size: 13.5px;
  line-height: 1.7;
}
.chat-user, .chat-ai { display: flex; gap: 10px; margin: 14px 0; align-items: flex-start; }
.chat-user { flex-direction: row-reverse; }
.chat-ico { flex: 0 0 auto; font-size: 15px; line-height: 1.6; opacity: .7; }
.chat-bubble {
  max-width: 84%;
  padding: 10px 14px;
  border-radius: 12px;
  line-height: 1.65;
  font-size: 13.5px;
  white-space: pre-wrap;
  font-family: var(--font);
}
.chat-user .chat-bubble {
  background: var(--bg2);
  color: var(--ink);
  border: 1px solid var(--line);
  border-bottom-right-radius: 4px;
}
.chat-ai .chat-bubble {
  background: transparent;
  color: var(--ink2);
  border: 1px solid var(--line);
  border-bottom-left-radius: 4px;
}

/* ── composer 工具行 ── */
.composer-tools { margin-top: 12px; align-items: flex-start; gap: 10px; }
.composer-tools .mini-upload { min-width: 0; }
.out-row { margin-top: 20px; align-items: flex-start; gap: 12px; }
.chat-shell { display: block; margin-top: 24px; }

/* 下载按钮：紧跟内容，不撑满整行 */
.dl-btn {
  width: fit-content !important;
  max-width: 100% !important;
  min-width: 0 !important;
  white-space: nowrap;
  padding: 8px 18px !important;
}
/* 下载状态说明：小字灰色 */
.out-row .gradio-markdown {
  font-size: 12px !important;
  color: var(--dim) !important;
  line-height: 1.55 !important;
}
.out-row .gradio-markdown p { margin: 6px 0 0 !important; }

/* 上传框内的标签：去掉深色底胶囊（真实类名 .label-wrap） */
.gradio-container .label-wrap,
.gradio-container button.label-wrap {
  background: transparent !important;
  color: var(--dim) !important;
  border: none !important;
  box-shadow: none !important;
  backdrop-filter: none !important;
  font-size: 12.5px !important;
  font-weight: 500 !important;
  padding: 0 2px !important;
}
.gradio-container .label-wrap > span {
  background: transparent !important;
  color: inherit !important;
}
/* 通用：所有 label 都不要深色底 */
.gradio-container label > span,
.gradio-container label {
  background-color: transparent !important;
}

/* ── 选中文字 ── */
::selection { background: var(--acc); color: #fff; }
a { color: var(--acc); }

/* ── 隐藏 footer ── */
.gradio-container footer { display: none !important; }

/* ── 滚动条：细 ── */
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--line2); border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: var(--dim); }
"""


# ── 亮/暗主题切换（客户端 JS，走 gradio 的 js= 回调，避免被 HTML 清洗掉）──
TOGGLE_JS = r"""
function(b){
  const r=document.documentElement;
  const next=r.getAttribute('data-ui-theme')==='dark' ? '' : 'dark';
  r.setAttribute('data-ui-theme', next);
  const el=document.getElementById('themeBtn');
  if(el) el.innerHTML = next==='dark' ? '☼ 明' : '☾ 暗';
  try{ localStorage.setItem('uiTheme', next); }catch(e){}
  return [];
}
"""
RESTORE_JS = r"""
function(b){
  let s='';
  try{ s = localStorage.getItem('uiTheme') || ''; }catch(e){}
  const r=document.documentElement;
  const el=document.getElementById('themeBtn');
  if(s==='dark'){ r.setAttribute('data-ui-theme','dark'); if(el) el.innerHTML='☼ 明'; }
  else if(el){ el.innerHTML='☾ 暗'; }
  return [];
}
"""


def build_ui():
    with gr.Blocks(title="实验报告生成器", theme=gr.themes.Base(primary_hue="blue", neutral_hue="zinc"), css=CSS) as demo:

        # ── 顶栏（纯 HTML/CSS，不含需执行的内联 JS）──
        gr.HTML("""
        <div class="pix-header">
            <div class="hd-left">
                <h1>实验报告生成器</h1>
                <p class="sub">上传实验指导书，生成符合学校格式的报告<span class="pix-dot">·</span>OCR + AI</p>
            </div>
        </div>
        """)

        # ── 主题切换按钮（gradio Button + js 回调，客户端切换 CSS 变量）──
        theme_btn = gr.Button("☾", elem_id="themeBtn", variant="secondary")

        # ══════════════════════════════════════════════
        #  主区（Codex 式）：输入框是视觉核心，结果区紧随其后
        # ══════════════════════════════════════════════

        # ── 单框输入：附件 + 文字 + 发送 ──
        composer = gr.MultimodalTextbox(
            show_label=False,
            placeholder="上传实验指导手册（可多选），或输入修改要求…",
            file_count="multiple",
            file_types=[".pdf", ".docx", ".doc", ".md", ".markdown", ".txt",
                        ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp"],
            sources=["upload"],
            lines=2,
            max_lines=8,
            submit_btn=True,
            elem_id="composerBox",
        )

        # ── 模板选项：一行紧凑 ──
        with gr.Row(elem_classes="composer-tools"):
            with gr.Column(scale=3):
                template_in = gr.File(
                    label="老师模板（可选）",
                    file_types=[".docx", ".doc", ".pdf", ".md"],
                    file_count="multiple",
                    type="filepath",
                    height=64,
                    elem_classes="mini-upload",
                )
            with gr.Column(scale=1, min_width=130):
                use_builtin = gr.Checkbox(label="用内置模板", value=True)

        # ── 结果区 ──
        with gr.Row(elem_classes="out-row"):
            with gr.Column(scale=1, min_width=170):
                download_btn = gr.DownloadButton(
                    "下载 .docx", variant="primary", interactive=True, visible=True,
                    elem_classes="dl-btn",
                )
                download_status = gr.Markdown("生成后在此下载")
            with gr.Column(scale=4):
                preview_out = gr.HTML(
                    '<div class="pix-statusbar">生成后在这里预览报告</div>'
                )

        # ── 对话流 ──
        with gr.Column(elem_classes="chat-shell"):
            chat_display = gr.HTML(chat_mod.render_chat_preview([]))

        with gr.Accordion("模型 / API", open=False):
            with gr.Row():
                model_sel = gr.Dropdown(
                    label="AI 模型",
                    choices=_get_model_choices(),
                    value=_default_model(),
                    interactive=True,
                    scale=3,
                )
                with gr.Column(scale=1):
                    api_key_in = gr.Textbox(label="API Key", type="password", placeholder="留空则用 .env")
                    api_base_in = gr.Textbox(label="API 地址", value=LLM_BASE_URL or "https://api.deepseek.com")

        with gr.Accordion("处理日志", open=False):
            log_out = gr.Textbox(label="日志", lines=10, max_lines=25, show_label=False)

        # ── 历史记录区 ──
        with gr.Accordion("历史记录", open=False):
            with gr.Row():
                refresh_btn = gr.Button("刷新")
            history_out = gr.HTML(history_mod.render_history_html())

        # ── 发送（Box 内回车 / ➤ 按钮 都走同一入口）──
        composer.submit(
            fn=composer_submit,
            inputs=[composer, template_in, api_key_in, api_base_in, model_sel, use_builtin],
            outputs=[chat_display, download_btn, download_status, preview_out, log_out],
        ).then(
            fn=refresh_history_ui, inputs=[], outputs=[history_out],
        )

        refresh_btn.click(
            fn=refresh_history_ui,
            inputs=[],
            outputs=[history_out],
        )

        # ── 明暗主题切换（客户端 JS，不入库、不耗 token）──
        theme_btn.click(fn=None, inputs=[], outputs=[], js=TOGGLE_JS)
        demo.load(fn=None, inputs=[], outputs=[], js=RESTORE_JS)

        # ── 说明（简洁两段，收进折叠区）──
        with gr.Accordion("使用说明", open=False):
            gr.Markdown("""
**三步出稿**

1. 点输入框左侧回形针，上传**实验指导书**（PDF / Word / 图片，可多选）
2. 想用老师的模板就上传到「老师模板」，不传则用内置成都大学线框
3. 点发送，生成后可下载 `.docx`

**改稿**：生成后在同一个输入框继续打字，例如「真值表第三行重算」
「实验小结写具体些」，每次生成新版本（`_改稿v1`、`_改稿v2`…）。

**模型**：选 `deepseek-chat` 走官方；选带前缀的（`deepseek-ai/…`、`Qwen/…`）
走硅基流动，需在「模型 / API」里填 SiliconFlow 的 Key。

**识别实验类型**：物理 / 化学 / 电路 / 力学 / 光学各自匹配标准章节结构。
""")

    return demo


if __name__ == "__main__":
    demo = build_ui()
    demo.launch(
        server_name="0.0.0.0", server_port=7860, share=False, show_error=True,
        allowed_paths=[str(UPLOAD_DIR), str(OUTPUT_DIR), str(BASE_DIR)],
    )
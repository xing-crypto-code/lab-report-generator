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
                  requirement="", on_progress=None):
    """核心处理流水线，返回 (下载文件路径, 下载状态文本, 预览HTML, 日志)
    manual_file / template_file 现在支持多文件（传入列表）。
    requirement: 用户在同一个输入框里顺带写的要求（会并入 AI 提示）
    on_progress: 可选回调，接收进度文本，用于 UI 实时刷新。
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
    # on_progress 回调：每次都追加到 status 并尽量实时外发（梯进式刷新）
    def report(t):
        report(t)
        if on_progress:
            try:
                on_progress(t)
            except Exception:
                pass

    try:
        # ── 第1步：解析实验指导手册（支持多文件）──
        report(f"🔍 正在解析 {len(manuals)} 个实验指导手册…")
        merged_text_parts = []
        merged_md_parts = []
        merged_tables = []
        for idx, mf in enumerate(manuals, 1):
            name = Path(mf).stem
            ext = Path(mf).suffix.lower()
            _is_img = ext in ('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.webp')
            if _is_img:
                report(f"   🖼️ [{idx}/{len(manuals)}] 《{Path(mf).name}》 OCR 识别中…"
                              f"（CPU 模式较慢，单张最多 90 秒）")
            else:
                report(f"   📄 [{idx}/{len(manuals)}] 《{Path(mf).name}》 解析中…")
            try:
                doc_result = parse_document(mf)
            except ValueError as e:
                report(f"   ⚠️ 《{name}》解析失败：{e}")
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
            report(f"   ✓ 《{Path(mf).name}》识别 {len(ocr_text):,} 字符，{len(ocr_tables)} 个表格"
                          + (f"，结构化 {len(_md)} 字符" if _md else ""))
        ocr_text = "\n\n".join(merged_text_parts)
        source_markdown = "\n\n".join(merged_md_parts)
        if not ocr_text.strip():
            return gr.update(), "*解析失败*", "*未能从文件中识别出文字*", "❌ 未能从上传文件中识别到文字内容"

        # 自动检测实验类型
        exp_type = detect_experiment_type(ocr_text)
        report(f"   🔬 检测实验类型：{exp_type}")
        report("")

        # ── 第2步：获取模板（支持多文件，优先 .docx，PDF 自动转 docx）──
        if templates:
            # 优先级：.docx > .doc > .md > .pdf（PDF 会转成 docx 再填）
            tplate_rank = {".docx": 0, ".doc": 1, ".md": 2, ".pdf": 3}
            templates_sorted = sorted(
                templates,
                key=lambda p: tplate_rank.get(Path(p).suffix.lower(), 9),
            )
            template_path = templates_sorted[0]
            report(f"📋 正在解析上传的模板（{Path(template_path).name}）…")
            # PDF 模板 → 转成可填充的 docx 线框（pdfplumber 提取表格+文本重建）
            if Path(template_path).suffix.lower() == ".pdf":
                try:
                    template_path = convert_pdf_to_docx(
                        template_path, str(Path("temp") / f"_pdf_tpl_{os.getpid()}.docx")
                    )
                    report(f"   🔁 PDF 模板已转成可编辑 docx 线框（{Path(template_path).name}）")
                except Exception as pe:
                    report(f"   ⚠️ PDF 模板解析失败：{pe}")
                    template_path = ""
            template_analysis = analyze_template(template_path) if template_path else {"total_paragraphs": 0, "total_tables": 0, "paragraphs": [], "tables": [], "sections_found": [], "placeholder_count": 0}
            t_summary = get_template_summary(template_analysis)
            report(f"   ✓ {template_analysis['total_paragraphs']} 段落，{template_analysis['total_tables']} 表格")
            if len(templates) > 1:
                report(f"   ℹ️ 上传了 {len(templates)} 个模板，已选用：{Path(templates_sorted[0]).name}")
            if not template_path:
                report("   ⚠️ 模板为空，回退使用内置模板")
                if Path(TEMPLATE_DIR / "内置实验报告模板.docx").exists():
                    template_path = str(TEMPLATE_DIR / "内置实验报告模板.docx")
                    template_analysis = analyze_template(template_path)
                    t_summary = get_template_summary(template_analysis)
        else:
            # 优先使用内置的成都大学模板（已优化、质量高）；不存在则回退到通用生成模板
            builtin = Path(TEMPLATE_DIR) / "内置实验报告模板.docx"
            if builtin.exists():
                report("📋 使用内置系统模板（成都大学标准线框模板）…")
                template_path = str(builtin)
                # 自动校验内置模板是否已固化最新修复；未固化则就地烘焙
                try:
                    from modules.template_fix import verify_template, rebake_template
                    _probs = verify_template(str(builtin))
                    if _probs:
                        report(f"   🔧 内置模板有 {len(_probs)} 处未固化，正在自动烘焙…")
                        rebake_template(str(builtin))
                        _probs2 = verify_template(str(builtin))
                        if _probs2:
                            report(f"   ⚠️ 烘焙后仍有 {len(_probs2)} 处问题")
                        else:
                            report("   ✓ 模板已烘焙到最新版")
                except Exception as _tfe:
                    report(f"   ⚠️ 模板自检跳过：{_tfe}")
            else:
                report("📋 使用内置标准模板…")
                tpl_bytes = build_standard_template(exp_type)
                tpl_path = Path("temp") / "_builtin_template.docx"
                tpl_path.write_bytes(tpl_bytes)
                template_path = str(tpl_path)
            template_analysis = analyze_template(template_path)
            t_summary = get_template_summary(template_analysis)
            report(f"   ✓ {template_analysis['total_paragraphs']} 段落，{template_analysis['total_tables']} 表格")
        report("")

        # ── 第3步：AI 填充 ──
        report(f"🤖 正在调用 {model} 生成实验报告…")
        if requirement and requirement.strip():
            report(f"   📌 附加要求：{requirement.strip()[:80]}")
        # 显式传入 UI 的 api_base/api_key（或 None → 按模型自动路由到对应 provider）
        filled = generate_report(
            ocr_text, ocr_tables, t_summary, model=model,
            api_key=api_key, api_base=api_base,
            requirement=requirement,
            source_markdown=source_markdown,
        )
        sections = filled.get("sections", {})
        count = sum(1 for v in sections.values() if v and v.strip() and "[待补充]" not in v)
        report(f"   ✓ AI 已填充 {count} 个章节")
        report(f"   📝 模型：{filled.get('model_used', model)}")
        if filled.get("notes"):
            report(f"   ℹ️ {filled['notes'][:120]}")
        report("")

        # ── 关键兜底：AI 完全没填出任何章节 → 中止，不要产出空白模板副本 ──
        if count == 0:
            err_note = filled.get("notes") or "AI 未能生成内容"
            report("🚫 未生成任何有效内容，已中止，避免输出空模板。")
            return (
                gr.update(),
                "❌ 生成失败：AI 没有返回内容",
                "*AI 未返回可填写的内容，请查看日志调整模型或 API*",
                "\n".join(status),
            )

        # ── 第4步：生成文档 ──
        report("📝 正在生成最终文档…")
        manual_name = Path(manuals[0]).stem
        output_name = f"{manual_name}_实验报告"

        result_paths = generate_report_document(
            template_path=template_path,
            filled_data=filled,
            output_name=output_name,
        )
        docx_path = result_paths.get("docx", "")
        if docx_path and Path(docx_path).exists():
            report(f"   ✓ 文档已生成：{Path(docx_path).name}")
            report("")
            report("✅ 完成！点击下载按钮用 WPS/Office 打开微调后交作业。")
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
        report(f"\n❌ 出错：{e}")
        report(f"```\ntraceback.format_exc()[:600]```")
        err = f"*❌ 生成出错：{e}*"
        return gr.update(), "❌ 出现错误", err, "\n".join(status)


def refresh_history_ui():
    """返回最新的历史面板 HTML，供刷新按钮/生成后调用"""
    return history_mod.render_history_html()


def composer_submit(message, template_file, api_key, api_base, model, use_builtin,
                    progress=gr.Progress()):
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
            def _on_p(msg):
                # 实时进度：去掉前缀 emoji 更清爽，只留说明文本
                progress(msg.replace("\n", " "))
            dl, dl_status, preview, log = process_files(
                files, template_file, api_key, api_base, model, use_builtin,
                requirement=text,
                on_progress=_on_p,
            )
        except Exception as e:
            chat_mod._session.append(("ai", f"❌ 生成出错：{e}"))
            return chat_mod.render_chat_preview(chat_mod._session), _nochange, "❌ 出错", _nochange, _nochange
        if dl:
            gr.Info(f"✅ 报告已生成！点「下载 .docx」保存")
            chat_mod._session.append(("ai", "✅ 报告已生成，可在上方预览；点「下载 .docx」保存。\n"
                                           "想改？在输入框继续提要求即可。"))
        else:
            gr.Warning(f"❌ 生成失败：{dl_status[:120]}")
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
/* ═══════════════════════════════════════════════════════════
   实验报告工坊 · 界面样式
   设计语言：中性极简（参考 shadcn/ui · UI-TARS Desktop）
   —— 纯白底、近黑主按钮、1px 细边、10px 圆角、系统字体
   双主题：亮色（默认） / 暗色（.data-ui-theme="dark"）
   ═══════════════════════════════════════════════════════════ */

/* ─────────── 设计令牌 ─────────── */
:root {
  --app-bg:            #ffffff;
  --app-bg-subtle:     #fafafa;
  --app-fg:            #0a0a0a;
  --app-fg-muted:      #737373;
  --app-card:          #ffffff;
  --app-muted:         #f5f5f5;
  --app-muted-hover:   #ededed;
  --app-border:        #e5e5e5;
  --app-border-strong: #d4d4d4;
  --app-primary:       #171717;
  --app-primary-fg:    #fafafa;
  --app-primary-hover: #262626;
  --app-accent:        #2563eb;   /* 仅用于链接/强调，不抢主色 */
  --app-danger:        #dc2626;
  --app-success:       #16a34a;
  --app-ring:          rgba(23,23,23,.16);
  --app-radius:        10px;
  --app-radius-lg:     14px;
  --app-shadow-xs:     0 1px 2px rgba(0,0,0,.04);
  --app-shadow-sm:     0 1px 2px rgba(0,0,0,.05), 0 1px 3px rgba(0,0,0,.04);
  --app-shadow-md:     0 4px 12px rgba(0,0,0,.06), 0 1px 3px rgba(0,0,0,.04);
  --app-font: ui-sans-serif, system-ui, -apple-system, "Segoe UI", "PingFang SC",
          "Hiragino Sans GB", "Microsoft YaHei", "Noto Sans SC", sans-serif;
  --app-mono: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;
}

:root[data-ui-theme="dark"] {
  --app-bg:            #0a0a0a;
  --app-bg-subtle:     #111111;
  --app-fg:            #fafafa;
  --app-fg-muted:      #a3a3a3;
  --app-card:          #171717;
  --app-muted:         #262626;
  --app-muted-hover:   #303030;
  --app-border:        #262626;
  --app-border-strong: #404040;
  --app-primary:       #fafafa;
  --app-primary-fg:    #171717;
  --app-primary-hover: #e5e5e5;
  --app-accent:        #60a5fa;
  --app-danger:        #f87171;
  --app-success:       #4ade80;
  --app-ring:          rgba(250,250,250,.22);
  --app-shadow-xs:     0 1px 2px rgba(0,0,0,.5);
  --app-shadow-sm:     0 1px 2px rgba(0,0,0,.6);
  --app-shadow-md:     0 4px 12px rgba(0,0,0,.6);
}

/* ─────────── 基础 ─────────── */
html, body, .gradio-container {
  background: var(--app-bg) !important;
  color: var(--app-fg);
  font-family: var(--app-font) !important;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-rendering: optimizeLegibility;
}

/* gradio 会按系统深色模式给 body 加 .dark，并在 gradio-app 上刷深色底；
   本应用主题由 html[data-ui-theme] 控制，这里统一压平，避免出现深色"边框沟" */
body.dark, html.dark { background: var(--app-bg) !important; }
gradio-app {
  background: var(--app-bg) !important;
  display: flex !important;
  flex-direction: column !important;
  align-items: center !important;
}

/* 让 gradio 原生组件跟随我们的中性色 */
.gradio-container {
  --body-text-color: var(--app-fg);
  --body-text-color-subdued: var(--app-fg-muted);
  --background-fill-primary: var(--app-bg);
  --background-fill-secondary: var(--app-muted);
  --block-background-fill: var(--app-card);
  --block-border-color: var(--app-border);
  --block-label-text-color: var(--app-fg);
  --block-title-text-color: var(--app-fg);
  --border-color-primary: var(--app-border);
  --input-background-fill: var(--app-bg);
  --input-border-color: var(--app-border);
  --input-placeholder-color: var(--app-fg-muted);
  --button-primary-background-fill: var(--app-primary);
  --button-primary-background-fill-hover: var(--app-primary-hover);
  --button-primary-text-color: var(--app-primary-fg);
  --button-primary-border-color: var(--app-primary);
  --button-secondary-background-fill: var(--app-bg);
  --button-secondary-background-fill-hover: var(--app-muted);
  --button-secondary-text-color: var(--app-fg);
  --button-secondary-border-color: var(--app-border);
  --link-text-color: var(--app-accent);
  --color-accent: var(--app-primary);

  max-width: 1180px !important;
  width: 100% !important;
  margin: 0 auto !important;
  padding: 0 24px 56px !important;
  font-size: 14px;
  line-height: 1.6;
}

/* 去掉 gradio 默认的块状容器外观，交给组件自己 */
.block, .form, .gr-panel, .panel, .gradio-container .block,
.gradio-container .form {
  background: transparent !important;
  border: none !important;
  border-radius: 0 !important;
  box-shadow: none !important;
  padding: 0 !important;
}

/* ─────────── 顶栏 ─────────── */
.app-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  padding: 22px 2px 18px;
  border-bottom: 1px solid var(--app-border);
  margin-bottom: 22px;
}
.app-header .brand { display: flex; align-items: center; gap: 12px; min-width: 0; }
.app-header .brand-mark {
  width: 34px; height: 34px; flex: 0 0 auto;
  display: flex; align-items: center; justify-content: center;
  background: var(--app-primary); color: var(--app-primary-fg);
  border-radius: var(--app-radius);
  font-size: 15px; font-weight: 700;
}
.app-header h1 {
  font-size: 17px; font-weight: 650; letter-spacing: -.01em;
  margin: 0; color: var(--app-fg); line-height: 1.25;
}
.app-header .brand-text p {
  margin: 2px 0 0; font-size: 12.5px; color: var(--app-fg-muted); line-height: 1.3;
}
.app-header .pill {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 5px 10px; font-size: 12px; color: var(--app-fg-muted);
  background: var(--app-muted); border: 1px solid var(--app-border);
  border-radius: 999px; white-space: nowrap;
}
.app-header .pill b { color: var(--app-fg); font-weight: 600; }

/* ─────────── 三步提示条 ─────────── */
.steps-bar {
  display: flex; flex-wrap: wrap; align-items: center; gap: 6px 20px;
  padding: 11px 14px; margin: 0 0 18px;
  background: var(--app-bg-subtle); border: 1px solid var(--app-border);
  border-radius: var(--app-radius); font-size: 12.5px; color: var(--app-fg-muted);
}
.steps-bar .step { display: inline-flex; align-items: center; gap: 7px; }
.steps-bar .step b {
  display: inline-flex; align-items: center; justify-content: center;
  width: 18px; height: 18px; flex: 0 0 auto;
  background: var(--app-fg); color: var(--app-bg);
  border-radius: 999px; font-size: 11px; font-weight: 700;
}

/* ─────────── 按钮 ─────────── */
button, .gradio-button, [type=button] {
  font-family: var(--app-font) !important;
  font-size: 13.5px !important;
  font-weight: 550 !important;
  letter-spacing: 0 !important;
  border-radius: var(--app-radius) !important;
  border: 1px solid var(--app-border) !important;
  background: var(--app-card) !important;
  background-image: none !important;
  color: var(--app-fg) !important;
  box-shadow: var(--app-shadow-xs) !important;
  padding: 8px 14px !important;
  min-height: 36px;
  transition: background-color .15s ease, border-color .15s ease,
              box-shadow .15s ease, color .15s ease !important;
  text-shadow: none !important;
}
button:hover, .gradio-button:hover {
  background: var(--app-muted) !important;
  border-color: var(--app-border-strong) !important;
}
button:active, .gradio-button:active { transform: none !important; }
button:focus-visible, .gradio-button:focus-visible {
  outline: none !important;
  box-shadow: 0 0 0 3px var(--app-ring) !important;
}

/* 主按钮：近黑填充 */
button.primary, .gradio-button.primary, button[variant="primary"] {
  background: var(--app-primary) !important;
  border-color: var(--app-primary) !important;
  color: var(--app-primary-fg) !important;
  box-shadow: var(--app-shadow-sm) !important;
}
button.primary:hover, .gradio-button.primary:hover {
  background: var(--app-primary-hover) !important;
  border-color: var(--app-primary-hover) !important;
}
button.primary:disabled {
  opacity: .45 !important;
}

/* 下载按钮占满卡片宽度 */
.out-row .dl-btn {
  width: 100% !important;
  max-width: 100% !important;
  white-space: nowrap;
  justify-content: center !important;
}

/* ─────────── 输入框 / 文本域 ─────────── */
textarea, input[type=text], input[type=password], input[type=email], input[type=number],
.gradio-container input, .gradio-container textarea, .gradio-container select {
  background: var(--app-bg) !important;
  border: 1px solid var(--app-border) !important;
  border-radius: var(--app-radius) !important;
  color: var(--app-fg) !important;
  font-family: var(--app-font) !important;
  font-size: 14px !important;
  padding: 9px 12px !important;
  box-shadow: none !important;
  transition: border-color .15s ease, box-shadow .15s ease !important;
}
textarea:focus, input:focus {
  border-color: var(--app-border-strong) !important;
  box-shadow: 0 0 0 3px var(--app-ring) !important;
  outline: none !important;
}
::placeholder { color: var(--app-fg-muted) !important; opacity: 1; }

/* ─────────── Composer（主输入区）─────────── */
#composerBox {
  border: 1px solid var(--app-border) !important;
  border-radius: var(--app-radius-lg) !important;
  background: var(--app-card) !important;
  background-color: var(--app-card) !important;
  background-image: none !important;
  box-shadow: var(--app-shadow-sm) !important;
  padding: 6px 8px 8px !important;
  transition: border-color .15s ease, box-shadow .15s ease !important;
}
.gradio-container #composerBox { background-color: var(--app-card) !important; }
#composerBox:focus-within {
  border-color: var(--app-border-strong) !important;
  box-shadow: 0 0 0 3px var(--app-ring), var(--app-shadow-sm) !important;
}
#composerBox .form,
#composerBox .full-container,
#composerBox .input-container,
#composerBox .input-wrapper,
#composerBox .input-row,
#composerBox > div,
#composerBox > div > div {
  background-color: transparent !important;
  background-image: none !important;
}
#composerBox .upload-container.hidden { display: none !important; }
#composerBox .input-row { align-items: flex-end !important; gap: 6px !important; }
#composerBox .input-wrapper { flex: 1 1 auto !important; min-width: 0 !important; }
#composerBox textarea {
  width: 100% !important;
  min-height: 54px !important;
  border: none !important;
  background: transparent !important;
  box-shadow: none !important;
  padding: 8px 8px !important;
  font-size: 14px !important;
  line-height: 1.65 !important;
  color: var(--app-fg) !important;
}
#composerBox textarea:focus { box-shadow: none !important; }
#composerBox textarea::placeholder { color: var(--app-fg-muted) !important; }

/* 附件按钮：幽灵按钮 */
#composerBox button[data-testid="upload-button"],
#composerBox button.icon-button {
  background: transparent !important;
  border: 1px solid transparent !important;
  border-radius: var(--app-radius) !important;
  color: var(--app-fg-muted) !important;
  width: 34px !important; min-width: 34px !important;
  height: 34px !important; padding: 0 !important;
  box-shadow: none !important;
  margin-bottom: 2px;
}
#composerBox button[data-testid="upload-button"]:hover,
#composerBox button.icon-button:hover {
  background: var(--app-muted) !important;
  color: var(--app-fg) !important;
}

/* 发送按钮：主色圆形 */
#composerBox button.submit-button, button.submit-button {
  width: 34px !important; min-width: 34px !important;
  height: 34px !important; padding: 0 !important;
  border-radius: 999px !important;
  background: var(--app-primary) !important;
  border: 1px solid var(--app-primary) !important;
  color: var(--app-primary-fg) !important;
  box-shadow: var(--app-shadow-sm) !important;
  margin-bottom: 2px;
}
#composerBox button.submit-button:hover, button.submit-button:hover {
  background: var(--app-primary-hover) !important;
  border-color: var(--app-primary-hover) !important;
}
#composerBox button.submit-button svg { color: var(--app-primary-fg) !important; }

/* ─────────── 模板上传 ─────────── */
.composer-tools { margin-top: 12px; gap: 16px !important; align-items: flex-start !important; }
#tplUpload {
  border: 1px solid var(--app-border) !important;
  border-radius: var(--app-radius) !important;
  background: var(--app-bg) !important;
  background-image: none !important;
  box-shadow: none !important;
  overflow: hidden;
}
#tplUpload, #tplUpload * {
  color: var(--app-fg) !important;
  text-shadow: none !important;
}
#tplUpload .upload-container,
#tplUpload [data-testid="file-upload"] {
  border: 1px dashed var(--app-border-strong) !important;
  border-radius: calc(var(--app-radius) - 2px) !important;
  background: var(--app-bg-subtle) !important;
  background-image: none !important;
  box-shadow: none !important;
  min-height: 46px !important;
  display: flex !important; align-items: center !important; justify-content: center !important;
  cursor: pointer !important;
  color: var(--app-fg-muted) !important;
  font-family: var(--app-font) !important;
  font-size: 13px !important;
  font-weight: 400 !important;
  transition: background-color .15s ease, border-color .15s ease !important;
}
#tplUpload .upload-container:hover,
#tplUpload [data-testid="file-upload"]:hover {
  background: var(--app-muted) !important;
  border-color: var(--app-fg-muted) !important;
  color: var(--app-fg) !important;
}
#tplUpload .upload-container > div,
#tplUpload [data-testid="upload-text"] {
  color: var(--app-fg-muted) !important;
  font-family: var(--app-font) !important;
  font-size: 13px !important;
  font-weight: 400 !important;
}
#tplUpload .upload-container:hover > div,
#tplUpload .upload-container:hover [data-testid="upload-text"] { color: var(--app-fg) !important; }
#tplUpload svg, #tplUpload [class*="icon"] {
  color: var(--app-fg-muted) !important; fill: var(--app-fg-muted) !important;
  width: 18px !important; height: 18px !important;
}
#tplUpload label, #tplUpload .block-label {
  font-size: 12px !important; font-weight: 500 !important;
  color: var(--app-fg-muted) !important;
}

/* ─────────── 勾选框 ─────────── */
input[type="checkbox"],
#useBuiltin input[type="checkbox"],
.gradio-container input[type="checkbox"] {
  -webkit-appearance: none !important;
  appearance: none !important;
  box-sizing: border-box !important;
  width: 17px !important; height: 17px !important;
  min-width: 17px !important; min-height: 17px !important;
  max-width: 17px !important; max-height: 17px !important;
  margin: 0 8px 0 0 !important;
  padding: 0 !important;
  border: 1px solid var(--app-border-strong) !important;
  border-radius: 4px !important;
  background: var(--app-bg) !important;
  box-shadow: var(--app-shadow-xs) !important;
  cursor: pointer !important;
  position: relative !important;
  flex: 0 0 auto !important;
  transition: background-color .15s ease, border-color .15s ease,
              box-shadow .15s ease !important;
}
input[type="checkbox"]:hover { border-color: var(--app-fg-muted) !important; }
input[type="checkbox"]:checked {
  background: var(--app-primary) !important;
  border-color: var(--app-primary) !important;
}
input[type="checkbox"]:checked::after {
  content: "" !important;
  position: absolute !important;
  left: 5px !important; top: 2px !important;
  width: 4px !important; height: 8px !important;
  border: solid var(--app-primary-fg) !important;
  border-width: 0 1.8px 1.8px 0 !important;
  transform: rotate(42deg) !important;
}
input[type="checkbox"]:focus-visible {
  outline: none !important;
  box-shadow: 0 0 0 3px var(--app-ring) !important;
}
#useBuiltin label, #useBuiltin .label-text {
  display: inline-flex !important; align-items: center !important;
  font-size: 13px !important; font-weight: 500 !important;
  color: var(--app-fg) !important; cursor: pointer !important;
}

/* ─────────── 结果区 ─────────── */
.out-row { margin-top: 18px !important; gap: 16px !important; align-items: stretch !important; }
.out-row .gradio-markdown {
  font-family: var(--app-font) !important;
  font-size: 12.5px !important;
  color: var(--app-fg-muted) !important;
  line-height: 1.5 !important;
}
.out-row .gradio-markdown p { margin: 6px 0 0 !important; }

.preview-panel {
  border: 1px solid var(--app-border) !important;
  border-radius: var(--app-radius-lg) !important;
  background: var(--app-card) !important;
  box-shadow: var(--app-shadow-sm) !important;
  padding: 0 !important;
  min-height: 120px;
  display: flex; align-items: center; justify-content: center;
  overflow: hidden;
}
.pix-statusbar,
.preview-panel .pix-statusbar {
  font-family: var(--app-font) !important;
  font-size: 13px !important;
  font-weight: 400 !important;
  color: var(--app-fg-muted) !important;
  text-align: center;
  padding: 36px 20px !important;
  line-height: 1.6;
  letter-spacing: 0;
  text-shadow: none !important;
}
.preview-panel .pix-statusbar * { color: var(--app-fg-muted) !important; }
/* 有内容时取消居中，让预览铺开 */
.preview-panel:has(.r-wrap) { display: block; padding: 12px !important; }

/* ─────────── 文档预览（纸张）─────────── */
.r-wrap {
  max-height: 640px; overflow: auto;
  background: #ffffff; color: #18181b;
  border: 1px solid var(--app-border);
  border-radius: var(--app-radius);
  padding: 18px 22px;
  box-shadow: none;
  font-family: "Songti SC", "SimSun", "宋体", serif;
}
:root[data-ui-theme="dark"] .r-wrap { background: #ffffff; color: #18181b; }
.r-toolbar {
  font-family: var(--app-font); font-size: 12px; color: #71717a;
  border-bottom: 1px solid #e4e4e7; padding-bottom: 8px; margin-bottom: 12px;
  position: sticky; top: 0; background: #ffffff; z-index: 2;
}
.r-page { line-height: 1.6; }
.r-page .r-para { margin: 4px 0; font-size: 12.5px; }
.r-table {
  width: 100%; border-collapse: collapse; margin: 8px 0;
  font-size: 11.5px; page-break-inside: avoid;
}
.r-table td { border: 1px solid #a1a1aa; padding: 5px 7px; vertical-align: top; background: #fff; }
.r-table th { border: 1px solid #a1a1aa; padding: 5px 7px; background: #f4f4f5; font-weight: 600; }
.r-ph { border: 1px dashed #d4d4d8; display: inline-block; }

/* ─────────── 对话流 ─────────── */
.chat-shell { margin-top: 18px !important; }
.chat-empty {
  font-size: 13px; color: var(--app-fg-muted); text-align: center;
  padding: 26px 16px; border: 1px dashed var(--app-border);
  border-radius: var(--app-radius); background: var(--app-bg-subtle);
}
.chat-user, .chat-ai {
  display: flex; gap: 10px; align-items: flex-start; margin: 10px 0;
}
.chat-user { flex-direction: row-reverse; }
.chat-ico {
  flex: 0 0 auto; width: 26px; height: 26px;
  display: flex; align-items: center; justify-content: center;
  border-radius: 999px; font-size: 13px;
  background: var(--app-muted); color: var(--app-fg);
  border: 1px solid var(--app-border);
}
.chat-bubble {
  max-width: 76%; padding: 9px 13px; font-size: 13.5px; line-height: 1.6;
  border-radius: var(--app-radius); border: 1px solid var(--app-border);
  background: var(--app-card); color: var(--app-fg);
  box-shadow: var(--app-shadow-xs);
  white-space: pre-wrap; word-break: break-word;
}
.chat-user .chat-bubble {
  background: var(--app-primary); color: var(--app-primary-fg);
  border-color: var(--app-primary);
}

/* ─────────── 历史记录 ─────────── */
.hist-card {
  border: 1px solid var(--app-border); border-radius: var(--app-radius);
  background: var(--app-card); padding: 12px 14px; margin-bottom: 10px;
  box-shadow: var(--app-shadow-xs);
  transition: border-color .15s ease, box-shadow .15s ease;
}
.hist-card:hover { border-color: var(--app-border-strong); box-shadow: var(--app-shadow-sm); }
.hist-card .h-title { font-size: 13.5px; font-weight: 600; color: var(--app-fg); }
.hist-card .h-meta { font-size: 12px; color: var(--app-fg-muted); margin: 2px 0 8px; }
.hist-card a {
  display: inline-flex; align-items: center; gap: 5px;
  font-size: 12.5px; font-weight: 550; color: var(--app-primary-fg);
  background: var(--app-primary); border: 1px solid var(--app-primary);
  border-radius: 8px; padding: 6px 11px; text-decoration: none;
  transition: background-color .15s ease;
}
.hist-card a:hover { background: var(--app-primary-hover); border-color: var(--app-primary-hover); }

/* ─────────── 折叠面板（卡片化）─────────── */
details, .accordion, [data-testid="accordion"], .gr-accordion {
  border: 1px solid var(--app-border) !important;
  border-radius: var(--app-radius) !important;
  background: var(--app-card) !important;
  box-shadow: none !important;
  overflow: hidden !important;
  margin-top: 14px !important;
  padding: 0 !important;
}
details > summary, .label-wrap, [data-testid="accordion"] > button {
  padding: 12px 16px !important;
  font-family: var(--app-font) !important;
  font-size: 13.5px !important;
  font-weight: 550 !important;
  color: var(--app-fg) !important;
  background: var(--app-card) !important;
  border: none !important;
  cursor: pointer;
  transition: background-color .15s ease !important;
}
details > summary:hover, .label-wrap:hover { background: var(--app-muted) !important; }
details[open] > summary, .label-wrap.open {
  border-bottom: 1px solid var(--app-border) !important;
}
details > *:not(summary) { margin-top: 0 !important; }
.gradio-container .html-container { min-height: 0 !important; padding: 0 !important; }
.accordion > div, [data-testid="accordion-content"] { padding: 14px 16px !important; }

/* ─────────── 下拉 ─────────── */
.gradio-dropdown, .wrap, .wrap-inner {
  font-family: var(--app-font) !important;
  border-radius: var(--app-radius) !important;
}
ul.options, .options {
  border: 1px solid var(--app-border) !important;
  border-radius: var(--app-radius) !important;
  box-shadow: var(--app-shadow-md) !important;
  background: var(--app-card) !important;
  overflow: hidden;
}
ul.options li, .options .item {
  font-size: 13px !important;
  padding: 8px 12px !important;
}
ul.options li:hover, .options .item:hover { background: var(--app-muted) !important; }

/* ─────────── 文字排版 ─────────── */
.gradio-container h1, .gradio-container h2, .gradio-container h3,
.gradio-container h4, .gradio-container h5 {
  font-family: var(--app-font) !important;
  color: var(--app-fg) !important;
  letter-spacing: -.01em;
  font-weight: 600 !important;
}
.gradio-container h3 { font-size: 15px !important; margin: 16px 0 8px !important; }
.gradio-container h4 { font-size: 14px !important; margin: 14px 0 6px !important; }
.gradio-container p, .prose p { color: var(--app-fg); }
.prose { font-size: 13.5px !important; line-height: 1.65 !important; }
.prose code {
  background: var(--app-muted); color: var(--app-fg);
  padding: 1px 5px; border-radius: 5px; font-size: 12.5px;
  font-family: var(--app-mono); border: 1px solid var(--app-border);
}
.prose a { color: var(--app-accent); text-decoration: none; }
.prose a:hover { text-decoration: underline; }
.prose ul, .prose ol { padding-left: 20px; }
.prose li { margin: 3px 0; }
hr { border: none !important; border-top: 1px solid var(--app-border) !important; margin: 20px 0 !important; }

/* ─────────── 主题切换按钮 ─────────── */
#themeBtn {
  position: fixed !important;
  top: 18px; right: 20px; z-index: 60;
  width: 36px !important; height: 36px !important;
  min-width: 36px !important; padding: 0 !important;
  display: flex; align-items: center; justify-content: center;
  background: var(--app-card) !important;
  color: var(--app-fg) !important;
  border: 1px solid var(--app-border) !important;
  border-radius: var(--app-radius) !important;
  box-shadow: var(--app-shadow-xs) !important;
  font-size: 14px !important;
  transition: background-color .15s ease !important;
}
#themeBtn:hover { background: var(--app-muted) !important; }

/* ─────────── 滚动条 ─────────── */
* { scrollbar-width: thin; scrollbar-color: var(--app-border-strong) transparent; }
*::-webkit-scrollbar { width: 10px; height: 10px; }
*::-webkit-scrollbar-track { background: transparent; }
*::-webkit-scrollbar-thumb {
  background: var(--app-border-strong); border-radius: 999px;
  border: 3px solid transparent; background-clip: content-box;
}
*::-webkit-scrollbar-thumb:hover { background: var(--app-fg-muted); background-clip: content-box; }

/* ─────────── 响应式 ─────────── */
@media (max-width: 860px) {
  .gradio-container { padding: 0 14px 48px !important; }
  .app-header { padding: 16px 0 14px; }
  .app-header .pill { display: none; }
  .chat-bubble { max-width: 88%; }
}
@media (max-width: 640px) {
  .app-header h1 { font-size: 15.5px; }
  .app-header .brand-text p { font-size: 11.5px; }
  .steps-bar { font-size: 11.5px; padding: 9px 12px; gap: 5px 14px; }
  .out-row { flex-direction: column; }
}
"""


# ── 亮/暗主题切换（客户端 JS，走 gradio 的 js= 回调）──
TOGGLE_JS = r"""
function(b){
  const r = document.documentElement;
  const next = r.getAttribute('data-ui-theme') === 'dark' ? '' : 'dark';
  r.setAttribute('data-ui-theme', next);

  // 主题只由 data-ui-theme 决定：压掉 gradio 依系统深色加的 .dark
  r.classList.remove('dark');
  if (document.body) document.body.classList.remove('dark');

  const el = document.getElementById('themeBtn');
  if (el) {
    const ico = el.querySelector('span') || el;
    ico.textContent = next === 'dark' ? '☀' : '☾';
  }
  try { localStorage.setItem('uiTheme', next); } catch(e) {}
  return [];
}
"""
RESTORE_JS = r"""
function(b){
  let s = '';
  try { s = localStorage.getItem('uiTheme') || ''; } catch(e) {}
  const r = document.documentElement;
  r.setAttribute('data-ui-theme', s === 'dark' ? 'dark' : '');

  const strip = () => {
    r.classList.remove('dark');
    if (document.body) document.body.classList.remove('dark');
  };
  strip();
  setTimeout(strip, 250);
  setTimeout(strip, 1000);

  const el = document.getElementById('themeBtn');
  if (el) {
    const ico = el.querySelector('span') || el;
    ico.textContent = s === 'dark' ? '☀' : '☾';
  }
  return [];
}
"""


def build_ui():
    with gr.Blocks(title="实验报告工坊") as demo:

        # ── 顶栏 ──
        gr.HTML("""
        <header class="app-header">
          <div class="brand">
            <div class="brand-mark">◆</div>
            <div class="brand-text">
              <h1>实验报告工坊</h1>
              <p>上传指导书 · OCR + AI · 一键成稿</p>
            </div>
          </div>
          <span class="pill">本地运行 · <b>密钥自配</b></span>
        </header>
        """)

        # ── 主题切换按钮（gradio Button + js 回调）──
        theme_btn = gr.Button("☾", elem_id="themeBtn", variant="secondary")

        # ── 新手三步引导条（一眼看懂的流程，取代藏在折叠区的说明）──
        gr.HTML("""
        <div class="steps-bar">
            <span class="step"><b>1</b> 回形针上传实验指导书（PDF/Word/图片，可多选）</span>
            <span class="step"><b>2</b> 可选：传老师模板，不传用内置模板</span>
            <span class="step"><b>3</b> 回车 → 生成 → 下载 .docx</span>
        </div>
        """)

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
                    elem_id="tplUpload",
                )
            with gr.Column(scale=1, min_width=130):
                use_builtin = gr.Checkbox(label="用内置模板", value=True, elem_id="useBuiltin")

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
                    '<div class="pix-statusbar">生成后在这里预览报告</div>',
                    elem_classes="preview-panel",
                )

        # ── 对话流 ──
        with gr.Column(elem_classes="chat-shell"):
            chat_display = gr.HTML(chat_mod.render_chat_preview([]))

        with gr.Accordion("模型 / API", open=True):
            gr.Markdown("**选一个模型；没填 Key 则用 `.env` 里的**（DeepSeek 官方 `deepseek-chat` 最省心）")
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
        theme=gr.themes.Base(primary_hue="neutral", neutral_hue="neutral"),
        css=CSS,
        allowed_paths=[str(UPLOAD_DIR), str(OUTPUT_DIR), str(BASE_DIR)],
    )
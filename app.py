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
            chat_mod._session.append(("ai", "✅ 报告已生成，见下方预览；点「⬇ 下载」保存 .docx。\n"
                                           "还想改？直接在下面继续输入要求即可。"))
        else:
            chat_mod._session.append(("ai", f"❌ 生成失败：{dl_status}"))
        return chat_mod.render_chat_preview(chat_mod._session), dl, dl_status, preview, log

    # ── B. 无文件且无文字 → 忽略 ──
    if not text:
        return chat_mod.render_chat_preview(chat_mod._session), _nochange, _nochange, _nochange, _nochange

    # ── C. 无文件、有文字 → 对话修订 ──
    if not chat_mod.get_context().get("sections"):
        chat_mod._session.append(("user", text))
        chat_mod._session.append(("ai", "⚠️ 还没有报告可以修改。请先在下方输入框用 📎 上传实验指导手册，"
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
@import url('https://fonts.googleapis.com/css2?family=Pixelify+Sans:wght@400;600;700&family=Silkscreen&family=VT323&display=swap');

/* ── 浅色主题（默认）· 宇宙废土沙盒像素风（Minecraft 材质/光影灵感）── */
:root {
  --bg0: #e9e2d5;   /* 页面底（废土沙黄） */
  --bg1: #fbf7ef;   /* 卡片（沙色白） */
  --bg2: #e4d9c2;   /* hover / 次级块（干草色） */
  --line: #b9ab8f;  /* 边框（沙岩） */
  --line2: #8f7d5e; /* 强调边框（焦褐） */
  --acc: #6f8f4e;   /* 主色 矿坑苔绿 */
  --acc2: #b97a3d;  /* 副色 废土橙/熔岩 */
  --ink: #3c3a33;   /* 正文（深土） */
  --dim: #7d735f;   /* 次级文字（灰土） */
  --err: #c0392b;
  --deep: #433c2c;  /* 深描边（像素黑边·岩漆黑） */
  --soft: #fffdf7;  /* 输入底（沙白） */
  --pix: 'Pixelify Sans','Silkscreen','Press Start 2P','SimSun','宋体',cursive,monospace;
  --term: 'VT323','Courier New',monospace;
  --song: 'SimSun','宋体',sans-serif;
  --fade: linear-gradient(180deg, rgba(111,143,78,.10), rgba(185,122,61,.10));
}

/* ── 深色主题（一键切换 · 宇宙废土暗夜）── */
:root[data-ui-theme="dark"] {
  --bg0: #181410;   /* 岩洞黑 */
  --bg1: #241e16;   /* 卡片·焦木 */
  --bg2: #2f271d;   /* hover·熔岩暗层 */
  --line: #3d3327;  /* 岩缝 */
  --line2: #5c4a33; /* 焦岩 */
  --acc: #7fae56;   /* 末地苔绿 */
  --acc2: #d98a4b;  /* 熔岩橙 */
  --ink: #e6dcc8;   /* 沙白正文 */
  --dim: #9a8c72;   /* 灰土 */
  --err: #ff6b5e;
  --deep: #000000;  /* 星际黑 */
  --soft: #1c1712;  /* 岩灰输入底 */
}

.gradio-container {
  max-width: 980px !important;
  margin: 0 auto;
  background: var(--bg0) !important;
  font-family: var(--song);
  color: var(--ink);
}
body, .gradio-container { background: var(--bg0) !important; }

/* —— 宇宙废土沙盒氛围：背景像素噪点 + 暗物质星云(深色) —— */
body::before {
  content: "";
  position: fixed; inset: 0; z-index: -1; pointer-events: none;
  background-image:
    radial-gradient(circle at 18% 12%, var(--fade), transparent 42%),
    radial-gradient(circle at 82% 86%, var(--fade), transparent 45%);
}
/* 前景像素噪点（早沙滩/石块颗粒），营造材质包颗粒感 */
.gradio-container {
  background-image:
    linear-gradient(45deg, rgba(0,0,0,.015) 25%, transparent 25%, transparent 75%, rgba(0,0,0,.015) 75%),
    linear-gradient(45deg, rgba(0,0,0,.015) 25%, transparent 25%, transparent 75%, rgba(0,0,0,.015) 75%);
  background-size: 6px 6px;
  background-position: 0 0, 3px 3px;
}
:root[data-ui-theme="dark"] .gradio-container {
  background-image:
    radial-gradient(1px 1px at 20px 30px, rgba(255,255,255,.18), transparent),
    radial-gradient(1px 1px at 90px 120px, rgba(255,255,255,.12), transparent),
    radial-gradient(1.2px 1.2px at 200px 60px, rgba(255,255,255,.15), transparent),
    radial-gradient(1px 1px at 320px 180px, rgba(255,255,255,.10), transparent),
    radial-gradient(1px 1px at 420px 90px, rgba(255,255,255,.16), transparent);
  background-size: 520px 320px, 620px 420px, 700px 400px, 800px 500px, 900px 600px;
}

/* —— 顶部像素横幅 —— */
.pix-header {
  position: relative;
  background: var(--bg1);
  border: 3px solid var(--deep);
  box-shadow: inset 0 0 0 2px var(--line), 0 6px 0 0 var(--line);
  padding: 18px 22px;
  margin: 14px 0 18px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
}
.pix-header .hd-left { flex: 1; }
.pix-header h1 {
  font-family: var(--pix);
  font-size: 19px;
  letter-spacing: 1px;
  color: var(--acc);
  margin: 0 0 8px;
  text-shadow: 2px 0 0 var(--bg1), -2px 0 0 var(--bg1), 0 2px 0 var(--bg1), 0 -2px 0 var(--bg1);
}
.pix-header .sub {
  font-family: var(--term);
  font-size: 17px;
  color: var(--dim);
  margin: 0;
  letter-spacing: 0.5px;
}
.pix-header .pix-dot { color: var(--acc2); }

/* —— 主题切换按钮（星露谷式像素按钮）—— */
#themeRow { margin-top: -6px; margin-bottom: 8px; }
#themeBtn {
  font-family: var(--pix) !important;
  font-size: 14px !important;
  background: var(--acc) !important;
  color: #fff !important;
  border: 3px solid var(--deep) !important;
  border-radius: 2px !important;
  box-shadow: 3px 3px 0 0 var(--line2) !important;
  padding: 8px 16px !important;
  cursor: pointer;
  white-space: nowrap;
  line-height: 1;
  width: auto !important;
  min-width: 92px;
}
#themeBtn:hover { transform: translate(1px,1px); box-shadow: 2px 2px 0 0 var(--line2) !important; }

/* 区块标签（步骤） */
.sec-tag {
  font-family: var(--pix);
  font-size: 12px;
  color: #fff;
  background: var(--acc);
  display: inline-block;
  padding: 5px 9px;
  margin-bottom: 10px;
  box-shadow: 3px 3px 0 0 var(--line2);
  letter-spacing: .5px;
}
.sec-tag.purple { background: var(--acc2); }

/* 通用块覆盖 */
.block, .form, .gr-panel, .panel {
  background: var(--bg1) !important;
  border: 1px solid var(--line) !important;
  border-radius: 2px !important;
  box-shadow: inset 0 0 0 1px var(--bg0);
}
.gradio-container .block { padding: 12px; }

/* 输入框 */
textarea, input[type=text], input[type=password], input[type=email],
.gradio-container input, .gradio-container textarea, .gradio-container select {
  background: var(--soft) !important;
  border: 1px solid var(--line) !important;
  border-radius: 2px !important;
  color: var(--ink) !important;
  font-family: var(--term) !important;
  font-size: 15px !important;
}
textarea:focus, input:focus {
  border-color: var(--acc) !important;
  box-shadow: 0 0 0 1px var(--acc) !important;
}

/* 上传框虚线像素感 —— 小而美，压缩成紧凑的单栏条 */
.file-preview, [data-testid="file"], .upload-button, [data-testid="file-upload"] {
  border: 2px dashed var(--line2) !important;
  border-radius: 2px !important;
  background: var(--soft) !important;
  min-height: 78px !important;
}
.upload-button { display: flex; align-items: center; justify-content: center; padding: 6px 10px !important; }
.upload-button:hover { border-color: var(--acc) !important; }
/* 上传区内文字压小 */
[data-testid="file-upload"] {
  font-size: 12px !important;
  color: var(--dim) !important;
}
/* 上传可用区域压扁（去掉多余大留白） */
[data-testid="file"] > div > div { padding: 4px 6px !important; }

/* 下拉 */
.gradio-dropdown, .wrap, .wrap-inner {
  font-family: var(--term) !important;
}

/* 按钮——像素块 */
button:not(.theme-toggle), .gradio-button, [type=button] {
  font-family: var(--pix) !important;
  font-size: 12px !important;
  letter-spacing: 1px !important;
  border: 3px solid var(--deep) !important;
  border-radius: 2px !important;
  box-shadow: 4px 4px 0 0 var(--line), inset 0 -2px 0 0 rgba(0,0,0,.15);
  transition: none !important;
  text-transform: uppercase;
}
button:not(.theme-toggle):hover { transform: translate(2px,2px); box-shadow: 2px 2px 0 0 var(--line); }
button.primary {
  background: var(--acc) !important;
  color: #fff !important;
  border-color: var(--deep) !important;
}
button.primary:hover { filter: brightness(1.05); }

/* 折叠面板头 */
details, .accordion, [data-testid="accordion"] {
  border: 1px solid var(--line) !important;
  border-radius: 2px !important;
  background: var(--bg1) !important;
}
.accordion-header, .category, [data-testid="accordion"] summary {
  font-family: var(--pix) !important;
  font-size: 12px !important;
  letter-spacing: .5px !important;
  color: var(--ink) !important;
}

/* 分隔线 */
hr { border-color: var(--line) !important; opacity: .6; }

/* markdown 表格（上手说明） */
.gradio-markdown table { border-collapse: collapse; width: 100%; }
.gradio-markdown th, .gradio-markdown td { border: 1px solid var(--line); padding: 6px 10px; }
.gradio-markdown th { background: var(--bg2); color: var(--acc); font-family: var(--pix); font-size: 12px; }

/* 历史记录卡片 */
.hist-card {
  border: 1px solid var(--line);
  background: var(--bg1);
  padding: 10px 12px;
  margin-bottom: 8px;
  border-left: 4px solid var(--acc);
}
.hist-card .h-title { font-size: 15px; color: var(--ink); }
.hist-card .h-meta { font-size: 13px; color: var(--dim); font-family: var(--term); }
.hist-card a { color: var(--acc); text-decoration: none; font-weight: 600; }

/* 状态光标动画 */
@keyframes blink { 50% { opacity: 0; } }
.pix-caret { animation: blink 1.1s step-end infinite; color: var(--acc); }

/* 预览框 */
.preview-box {
  border: 1px solid var(--line);
  background: var(--soft);
  padding: 10px;
  font-family: var(--term);
  color: var(--ink);
}

/* 内容区标题 */
.gradio-container h3, .gradio-container h4, .gradio-container h5 {
  font-family: var(--pix) !important;
  font-size: 13px !important;
  color: var(--acc2) !important;
  letter-spacing: .5px;
}

/* 状态栏 */
.pix-statusbar {
  border-top: 2px solid var(--line);
  margin-top: 18px;
  padding-top: 10px;
  font-family: var(--term);
  font-size: 14px;
  color: var(--dim);
}

.gradio-container footer { visibility: hidden; }

/* ── 报告实时预览（docx→HTML 渲染）── */
.r-wrap {
  max-height: 620px; overflow: auto;
  background: #fff; color: #222;
  border: 2px solid var(--line2);
  border-radius: 6px; padding: 10px 14px;
  box-shadow: inset 0 0 0 3px var(--soft);
  font-family: 'SimSun','宋体',serif;
}
:root[data-ui-theme="dark"] .r-wrap { background: #fdf6e9; }
.r-toolbar {
  font-family: var(--term); font-size: 13px; color: var(--acc);
  border-bottom: 2px solid var(--line); padding-bottom: 6px; margin-bottom: 8px;
  position: sticky; top: 0; background: inherit; z-index: 2;
}
.r-page { line-height: 1.5; }
.r-page .r-para { margin: 3px 0; font-size: 12px; }
.r-table {
  width: 100%; border-collapse: collapse; margin: 6px 0;
  font-size: 11px; page-break-inside: avoid;
}
.r-table td {
  border: 1px solid #777; padding: 3px 5px; vertical-align: top;
  background: #fff;
}
.r-table .r-para { margin: 1px 0; }
.r-nested td { border: 1px solid #999; }

/* 预览里的虚线插图占位框（对应 Word 里的虚线框） */
.r-ph {
  border: 2px dashed #9aa0a6;
  border-radius: 4px;
  background: #fafafa;
  margin: 8px auto;
  display: flex;
  align-items: center;
  justify-content: center;
  box-sizing: border-box;
}
.r-ph span {
  color: #9aa0a6;
  font-family: 'SimSun','宋体',serif;
  font-size: 13px;
  letter-spacing: 1px;
}

/* ── AI 对话区（DeepSeek Harness 风格气泡）── */
.chat-empty { color: var(--dim); font-style: italic; padding: 10px 4px; font-family: var(--term); }
.chat-user, .chat-ai { display: flex; gap: 8px; margin: 8px 0; align-items: flex-start; }
.chat-user { flex-direction: row-reverse; }
.chat-ico { flex: 0 0 auto; font-size: 18px; line-height: 1.4; }
.chat-bubble {
  max-width: 82%; padding: 8px 12px; border-radius: 10px;
  line-height: 1.5; font-size: 13px; white-space: pre-wrap;
}
.chat-user .chat-bubble { background: var(--acc); color: #fff; border: 1px solid var(--line2); }
.chat-ai .chat-bubble { background: var(--bg2); color: var(--ink); border: 1px solid var(--line); }

/* ── 单框输入（DeepSeek-Harness 式 composer）布局 ── */
.chat-shell {
  background: var(--bg1);
  border: 2px solid var(--line2);
  border-radius: 10px;
  padding: 10px 14px;
  min-height: 96px;
  max-height: 420px;
  overflow-y: auto;
  box-shadow: inset 0 0 0 3px var(--soft);
  margin-bottom: 10px;
}
#composerBox textarea {
  font-family: var(--song);
  font-size: 15px;
  line-height: 1.6;
}
#composerBox {
  border: 2px solid var(--line2) !important;
  border-radius: 12px !important;
  background: var(--bg1) !important;
  box-shadow: inset 0 0 0 3px var(--soft);
}
#composerBox:focus-within { border-color: var(--acc) !important; }
.composer-tools { margin-top: 8px; align-items: flex-end; }
.composer-tools .mini-upload { min-width: 0; }
.hint {
  font-family: var(--term);
  font-size: 12px;
  color: var(--dim);
}

/* 链接与选中文字主题色统一 */
::selection { background: var(--acc); color: #fff; }
a { color: var(--acc); }
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
    with gr.Blocks(title="无痛实验报告生成器 ▣ PIXEL LAB", theme=gr.themes.Base(primary_hue="green", neutral_hue="gray"), css=CSS) as demo:

        # ── 顶栏（纯 HTML/CSS，不含需执行的内联 JS）──
        gr.HTML("""
        <div class="pix-header">
            <div class="hd-left">
                <h1>▣ 无痛实验报告生成器 <span class="pix-caret">▊</span></h1>
                <p class="sub">> PIXEL LAB REPORT GENERATOR <span class="pix-dot">·</span> OCR + AI → 一键成稿 <span class="pix-dot">·</span> v2.0</p>
            </div>
        </div>
        """)

        # ── 主题切换按钮（gradio Button + js 回调，客户端切换 CSS 变量）──
        theme_btn = gr.Button("☼ 明", elem_id="themeBtn", variant="primary")

        # ══════════════════════════════════════════════
        #  主区：上方对话流，下方 DeepSeek-Harness 式单框输入
        #         一个输入框搞定「上传文件 + 提要求 + 发送」
        # ══════════════════════════════════════════════
        with gr.Column(elem_classes="chat-shell"):
            chat_display = gr.HTML(chat_mod.render_chat_preview([]))

        # ── 输出区（生成/修订后就地出现）──
        with gr.Row(elem_classes="out-row"):
            with gr.Column(scale=1):
                download_btn = gr.DownloadButton(
                    "⬇ 下载报告 (.docx)", variant="primary", interactive=True, visible=True
                )
                download_status = gr.Markdown(
                    "*上传指导手册 → 点 ➤ 发送，我会先生成第一版；"
                    "之后在同一输入框继续提要求即可改稿。*"
                )
            with gr.Column(scale=3):
                preview_out = gr.HTML(
                    '<div class="pix-statusbar">> 等待生成…<span class="pix-caret">▊</span></div>'
                )

        # ── 单框输入：附件 + 文字 + 发送（模仿 DeepSeek Harness 输入条）──
        composer = gr.MultimodalTextbox(
            show_label=False,
            placeholder="上传实验指导手册（可多选），或直接输入修改要求…",
            file_count="multiple",
            file_types=[".pdf", ".docx", ".doc", ".md", ".markdown", ".txt",
                        ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp"],
            sources=["upload"],
            lines=2,
            max_lines=8,
            submit_btn=True,
            elem_id="composerBox",
        )
        with gr.Row(elem_classes="composer-tools"):
            template_in = gr.File(
                label="📄 老师模板（可选）",
                file_types=[".docx", ".doc", ".pdf", ".md"],
                file_count="multiple",
                type="filepath",
                height=78,
                elem_classes="mini-upload",
            )
            with gr.Column(min_width=120):
                use_builtin = gr.Checkbox(label="内置线框模板", value=True)
                gr.Markdown(
                    "<span class='hint'>不传模板就用内置成都大学线框</span>"
                )

        with gr.Accordion("⚙ 模型 / API 设置", open=False):
            with gr.Row():
                model_sel = gr.Dropdown(
                    label="AI 模型",
                    choices=_get_model_choices(),
                    value=_default_model(),
                    interactive=True,
                    scale=3,
                )
                with gr.Column(scale=1):
                    api_key_in = gr.Textbox(label="API Key（留空=用 .env）", type="password", placeholder="sk-…")
                    api_base_in = gr.Textbox(label="API 地址", value=LLM_BASE_URL or "https://api.deepseek.com")

        with gr.Accordion("🔍 处理日志 / LOG", open=False):
            log_out = gr.Textbox(label="日志", lines=10, max_lines=25)

        # ── 历史记录区 ──
        with gr.Accordion("📚 历史生成记录（可重复下载，不耗 Token）", open=False):
            with gr.Row():
                refresh_btn = gr.Button("↻ 刷新历史")
                gr.Markdown("&nbsp;")
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

        # ── 说明 ──
        gr.Markdown("""
        ---
        ### ▸ 快速上手

        | 步骤 | 操作 |
        |:---:|---|
        | ① | 上传 **实验指导手册**（PDF 或拍照） |
        | ② | 上传老师 **模板** .docx（或勾选内置模板） |
        | ③ | 选择 **AI 模型**（推荐 DeepSeek V3） |
        | ④ | 点击 **一键生成**，下载 .docx 微调 |

        ### ▸ 免费使用

        API 走 **硅基流动 SiliconFlow** 或 **DeepSeek 官方**，自动匹配。

        **模型 ↔ 服务商自动匹配**：选「DeepSeek-chat / DeepSeek-reasoner」→ 自动走官方便用 `.env` 里的官方 Key；
        选带前缀的模型（`deepseek-ai/…`、`Qwen/…`、`zai-org/…`）→ 自动走硅基流动 SiliconFlow，需在「引擎」填入
        SiliconFlow 的 API Key（或在 `.env` 设置 `SILICONFLOW_API_KEY`），否则会提示鉴权失败。

        ### ▸ 内置知识库

        | 实验类型 | 标准章节 |
        |---|---|
        | 物理 | 目的 → 原理 → 器材 → 步骤 → 数据 → 处理 → 误差 → 结论 |
        | 化学 | 目的 → 反应方程式 → 试剂仪器 → 步骤 → 数据 → 讨论 → 结论 |
        | 电路 | 目的 → 电路图 → 仪器器件 → 步骤 → 波形/数据 → 误差分析 |
        | 力学 | 目的 → 公式推导 → 仪器 → 步骤 → 数据图表 → 误差 → 结论 |

        > AI 自动识别实验类型并匹配标准结构；OCR 前做图像去噪/增强/二值化预处理。
        """)

    return demo


if __name__ == "__main__":
    demo = build_ui()
    demo.launch(
        server_name="0.0.0.0", server_port=7860, share=False, show_error=True,
        allowed_paths=[str(UPLOAD_DIR), str(OUTPUT_DIR), str(BASE_DIR)],
    )
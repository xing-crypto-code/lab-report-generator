"""
AI 对话修改模块 · chat.py
=======================
在生成完成后，用户可输入补充/修改要求，AI 基于【上次生成内容 + 实验上下文】给出
修订后的 sections，并重新落盘生成 docx。让用户能"提要求改报告"而不是只能整份重来。
"""
import os
import json
from config import LLM_API_KEY, LLM_BASE_URL, LLM_MODEL

# 保存最近一次的生成上下文（进程内），供后续对话修订引用
_last_context = {
    "ocr_text": "",
    "ocr_tables": [],
    "template_summary": "",
    "sections": {},
    "model": "",
    "experiment_type": "",
    "template_path": "",
}

# 会话历史（供 chat 展示）
_session = []


def save_context(ocr_text="", ocr_tables=None, template_summary="",
                 sections=None, model="", experiment_type="", template_path="",
                 output_base=""):
    """生成成功后保存上下文，供对话修订使用。
    注意：不清空 _session —— 生成那一刻用户刚发的消息气泡要保留。"""
    _last_context.clear()
    _last_context.update({
        "ocr_text": ocr_text,
        "ocr_tables": ocr_tables or [],
        "template_summary": template_summary,
        "sections": sections or {},
        "model": model or LLM_MODEL,
        "experiment_type": experiment_type or "电路/电子实验",
        "template_path": template_path,
        "output_base": output_base or "实验报告",
        "_rev": 0,
    })


def get_context():
    return dict(_last_context)


def chat_continue(user_msg: str) -> dict:
    """用户追加要求 → 调 AI 修订 sections。

    返回 {"sections": {...}, "notes": "...", "ok": bool, "error": str}
    """
    ctx = _last_context
    if not ctx["sections"]:
        return {"sections": {}, "notes": "尚未生成报告，无法对话修改", "ok": False,
                "error": "先点『一键生成』生成一份报告，再在这里提修改要求。"}

    from modules.filler import _get_client
    model = ctx["model"] or LLM_MODEL
    try:
        client = _get_client(model)
    except Exception as e:
        return {"sections": {}, "ok": False, "error": f"客户端初始化失败：{e}",
                "notes": ""}

    # 组装 prompt：展示上次的各章节内容 + 用户的新要求
    prev_txt = "\n\n".join(
        f"### {k}\n{(v or '（空）')[:1500]}"
        for k, v in ctx["sections"].items() if v and v.strip()
    ) or "（上次报告内容为空）"

    system = f"""你是实验报告修改助手。用户已生成了一份实验报告，现在提出修改要求。
你需要基于【上次报告内容】+【用户要求】输出【修订后的完整 sections】。
原则：
- 只按用户要求修改对应章节，未提及的章节保持原样。
- 实验数据(真值表/测量值)必须准确，务必每行验算，保持与逻辑定义一致。
- 不编造数据；实测值留空或 [待补充]。
- 输出严格 JSON：{{"sections":{{"<章节名>":"<修订后全文>"}},"notes":"说明"}}。
- sections 的 key 必须与下方【上次报告内容】的章节名一致，不要新增/删除章节。"""

    user = f"""【上次报告内容】
{prev_txt}

【用户修改要求】
{user_msg}

请输出修订后的完整 sections JSON。"""

    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "system", "content": system},
                      {"role": "user", "content": user}],
            temperature=0.3, max_tokens=8192,
        )
        content = resp.choices[0].message.content
        import re
        m = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', content, re.DOTALL)
        if m:
            content = m.group(1).strip()
        result = json.loads(content)
        new_sections = result.get("sections", {})
        # 与上下文合并：保留未在输出中的旧章节
        merged = dict(ctx["sections"])
        merged.update({k: v for k, v in new_sections.items() if v})
        ctx["sections"] = merged
        return {"sections": merged, "notes": result.get("notes", ""), "ok": True, "error": ""}
    except Exception as e:
        return {"sections": {}, "ok": False, "error": str(e),
                "notes": f"对话修订失败：{e}"}


def run_chat_revision(user_msg: str, template_path: str) -> dict:
    """对话修订并重新生成 docx，返回 {docx, preview, sections, ok, msg}"""
    from modules.generator import generate_report_document
    from modules.render import render_docx
    ctx = _last_context

    res = chat_continue(user_msg)
    if not res["ok"]:
        return {"ok": False, "msg": res.get("error") or res.get("notes"), "docx": "", "sections": {}}

    sections = res["sections"]
    # 版本化输出名：沿用上次的基准名 + _vN，避免覆盖与"报告.docx"这种泛名
    ctx["_rev"] = int(ctx.get("_rev", 0)) + 1
    base = ctx.get("output_base") or "实验报告"
    output_name = f"{base}_改稿v{ctx['_rev']}"
    try:
        result_paths = generate_report_document(
            template_path=ctx["template_path"] or template_path,
            filled_data={"sections": sections, "tables_fill": []},
            output_name=output_name,
        )
        docx_path = result_paths.get("docx", "")
        if not docx_path or not os.path.exists(docx_path):
            return {"ok": False, "msg": "重新生成失败", "docx": "", "sections": sections}
        preview = render_docx(docx_path)
        preview = f"<div class='r-wrap'><div class='r-toolbar'>📄 对话后预览 · {os.path.basename(docx_path)}</div>{preview}</div>"
        return {"ok": True, "msg": f"已按你的要求重新生成：{os.path.basename(docx_path)}",
                "docx": docx_path, "sections": sections, "preview": preview}
    except Exception as e:
        return {"ok": False, "msg": f"重新生成失败：{e}", "docx": "", "sections": sections}


def render_chat_preview(history: list) -> str:
    """把聊天历史渲染为 HTML 气泡列表（DeepSeek Harness 风格）"""
    import html as _h
    parts = []
    for role, text in history:
        cls = "chat-user" if role == "user" else "chat-ai"
        icon = "🧑" if role == "user" else "🤖"
        safe = _h.escape(str(text))
        parts.append(f"<div class='{cls}'><span class='chat-ico'>{icon}</span>"
                     f"<div class='chat-bubble'>{safe}</div></div>")
    if not parts:
        parts.append("<div class='chat-empty'>生成报告后，在这里输入修改要求（如：真值表重算、实验过程补充小结）</div>")
    return "".join(parts)
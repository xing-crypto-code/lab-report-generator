"""
AI 填充模块 · filler.py
=======================
职责：调用大模型 API，将 OCR 识别出的实验内容智能填入模板

功能：
- 硅基流动 SiliconFlow 多模型支持 → 可对比不同模型的生成效果
- 集成工科实验知识库 → 确保输出符合高校学术格式标准
- 智能润色去重 → 不编造数据，学术规范语言
"""
from openai import OpenAI
import time
import os
import json
import re
from config import (
    LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, MAX_RETRY,
    SILICONFLOW_API_KEY, SILICONFLOW_BASE_URL,
)
from modules.knowledge import (
    EXPERIMENT_TYPES, FORMAT_RULES, POLISH_GUIDELINES, DOMAIN_SPEC, ANTI_AI_STYLE,
)

_client_cache = {}

# ─── 硅基流动可用模型列表 ───

SILICONFLOW_MODELS = {
    "deepseek-ai/DeepSeek-V4.1-Flash": {"label": "DeepSeek V4.1 Flash ⭐最新", "desc": "2025-10 最新，速度/质量最佳平衡"},
    "deepseek-ai/DeepSeek-V4-Flash":   {"label": "DeepSeek V4 Flash ⚡",     "desc": "高性价比，生成快"},
    "deepseek-ai/DeepSeek-V4-Pro":     {"label": "DeepSeek V4 Pro",         "desc": "最新旗舰，质量最高"},
    "deepseek-ai/DeepSeek-V3":         {"label": "DeepSeek V3",             "desc": "经典稳定版，中英文俱佳"},
    "deepseek-ai/DeepSeek-V3.2":       {"label": "DeepSeek V3.2",           "desc": "V3 迭代，推理更强"},
    "deepseek-ai/DeepSeek-R1":         {"label": "DeepSeek R1",             "desc": "深度推理，适合复杂分析"},
    "Qwen/Qwen3.5-35B-A3B":            {"label": "Qwen 3.5 35B",            "desc": "阿里大杯，中文理解强"},
    "Qwen/Qwen3.5-27B":                {"label": "Qwen 3.5 27B",            "desc": "阿里通义千问，中文优秀"},
    "zai-org/GLM-5.3":                 {"label": "GLM-5.3",                 "desc": "智谱最新，中文优异"},
    "zai-org/GLM-5.3-Flash":           {"label": "GLM-5.3 Flash ⚡",        "desc": "智谱轻量版，速度快"},
}


def _load_key(base_url_default, key_default):
    """按 base url 运行时从环境重新读取（覆盖 import 时固化的问题）"""
    if "siliconflow" in (base_url_default or "").lower():
        return os.getenv("SILICONFLOW_API_KEY") or os.getenv("SF_API_KEY") or key_default
    return os.getenv("LLM_API_KEY") or key_default


def _get_client(model: str = None, api_key: str = None, api_base: str = None) -> OpenAI:
    """获取 OpenAI 客户端（按 model 缓存，支持切换不同模型）。

    自动按模型路由 provider：
    - 模型名带硅基前缀（deepseek-ai/、Qwen/、zai-org/ 等）→ SiliconFlow
    - 否则（deepseek-chat / deepseek-reasoner / 简短模型名）→ DeepSeek 官方
    显式传入的 api_base / api_key 优先（来自 UI 输入）。
    """
    global _client_cache
    model = model or LLM_MODEL

    # 确定 base url
    explicit_base = api_base and api_base.strip()
    explicit_key = api_key and api_key.strip()

    if explicit_base:
        base_url = explicit_base
        key = explicit_key or _load_key(base_url, LLM_API_KEY)
    elif "/" in model:  # 硅基流动风格的模型全名 → 走硅基
        base_url = SILICONFLOW_BASE_URL
        key = (explicit_key
               or os.getenv("SILICONFLOW_API_KEY")
               or os.getenv("SF_API_KEY")
               or SILICONFLOW_API_KEY
               or os.getenv("LLM_API_KEY"))  # 兜底用官方 key，让 API 返回真实鉴权错误而非客户端崩溃
    else:  # deepseek 官方短模型名
        base_url = LLM_BASE_URL
        key = explicit_key or os.getenv("LLM_API_KEY") or LLM_API_KEY

    # OpenAI() 要求非空 key；为空时用一个占位，让 provider 返回 HTTP 4xx 而不是本地抛异常
    if not key:
        key = "missing-key-placeholder"
    cache_key = f"{model}|{base_url}|{key}"
    if cache_key not in _client_cache:
        _client_cache[cache_key] = OpenAI(api_key=key, base_url=base_url)
    return _client_cache[cache_key]


def _parse_json_robust(content: str):
    """尝试解析模型返回的 JSON。若整体失败（如被 max_tokens 截断），
    尽力抢救已完整输出到的一半 sections。返回 dict 或 None。"""
    if not content:
        return None
    # 常规解析
    try:
        return json.loads(content)
    except Exception:
        pass
    # 抢救：模型常先输出 {"sections":{...}}，截断在尾部（notes 或不完整章节）。
    # 尝试补全闭合括号，取出能解析的最长前缀。
    candidates = []
    # 去掉末尾可能的不完整片段，逐级减少长度尝试
    trimmed = content.rstrip()
    for cut in range(len(trimmed), 0, -1):
        frag = trimmed[:cut]
        # 只尝试以 } 结尾或 }," 结尾的片段
        if not (frag.endswith("}") or frag.endswith("}}")):
            continue
        for closing in (frag + "}", frag + "]}"):
            try:
                obj = json.loads(closing)
                if isinstance(obj, dict):
                    return obj
            except Exception:
                continue
    return None


def detect_experiment_type(ocr_text: str) -> str:
    """
    自动判断实验类型（物理/化学/电路/力学/光学/计算机）
    """
    ocr_lower = ocr_text.lower()
    scores = {}
    keywords = {
        # 电路/电子 优先命中：数电、模电、信号与系统相关术语
        "电路/电子实验": [
            "电路", "电阻", "电压", "示波器", "万用表", "焊接", "波形", "频率",
            "信号", "接线", "门电路", "真值表", "与非", "与非门", "逻辑", "引脚",
            "74LS", "7400", "芯片", "集成", "运算放大器", "运放", "同相", "反相",
            "电压增益", "反馈", "静态工作点", "输出", "卷积", "傅里叶", "变换",
            "冲激响应", "线性时不变", "差分方程", "频谱",
        ],
        "物理实验": [
            "物理", "速度", "加速度", "力学", "电磁", "波动", "共振", "驻波",
            "波节", "波腹", "弦线", "声速", "张力", "波长", "温度",
            "热学", "等效电阻", "伏安特性",
        ],
        "化学实验": ["化学", "滴定", "溶液", "反应", "浓度", "沉淀", "pH", "mol", "mL", "试剂", "酸碱"],
        "力学实验": ["力学", "牛顿", "质量", "重力", "摩擦", "弹性", "转动", "力矩", "弹簧"],
        "光学实验": ["光学", "透镜", "折射", "反射", "偏振", "光栅", "激光", "衍射", "干涉"],
        "计算机/编程实验": ["编程", "算法", "代码", "Python", "C++", "数据结构", "编译", "调试", "程序"],
    }
    for exp_type, kws in keywords.items():
        scores[exp_type] = sum(1 for kw in kws if kw in ocr_lower)
    # 电路/电子 相对宽松命中时优先，且有逻辑/信号/运放等强信号词时加成
    circ_score = scores["电路/电子实验"]
    strong_circuit = any(k in ocr_lower for k in
                         ["门电路", "真值表", "与非", "运算放大器", "运放", "卷积",
                          "傅里叶", "冲激响应", "频谱", "电压增益"])
    if strong_circuit:
        scores["电路/电子实验"] = circ_score + 4
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "物理实验"


def _sections_from_template_summary(summary: str) -> list[str]:
    """从模板摘要里抠出"识别到的章节"，当作本次要输出的键。

    以前这里写死 5 个键（目的/仪器/原理/过程/小结），模板一旦多出「误差分析」
    「思考题」这类栏位，模型也只回 5 栏，多出来的小节永远是空的。
    """
    m = re.search(r"识别到的章节[:：]\s*([^\n]+)", summary or "")
    if not m:
        return []
    skip = ("课程名称", "实验名称", "教师评阅", "评阅", "班", "学号", "姓名",
            "学院", "成绩", "实验日期", "指导教师")
    out = []
    for raw in re.split(r"[,，、]", m.group(1)):
        name = re.sub(r"\s+", " ", raw).strip()
        name = re.sub(r"[（(][^）)]*[)）]", "", name)      # 实验过程（可另附页）
        name = re.sub(r"\s*(与建议|及建议|意见及建议).*$", "", name)
        name = name.strip()
        if not name or any(name.startswith(s) for s in skip):
            continue
        if name not in out:
            out.append(name)
    return out


def generate_report(
    ocr_text: str,
    ocr_tables: list,
    template_summary: str,
    model: str = None,
    experiment_type: str = None,
    api_key: str = None,
    api_base: str = None,
    requirement: str = "",
    source_markdown: str = "",
) -> dict:
    """
    核心方法：根据 OCR 内容和模板结构，调用 AI 生成填入方案

    参数:
        ocr_text: 从指导手册识别出的全文
        ocr_tables: 识别出的表格数据
        template_summary: 模板结构摘要
        model: 指定模型（None 则用 config 默认）
        experiment_type: 实验类型（None 则自动检测）
        api_key / api_base: 显式覆盖 provider（None 则自动按模型路由）
        requirement: 用户在输入框里顺带写的额外要求（会并入提示词）

    返回:
        {
            "sections": {"实验目的": "...", ...},
            "tables_fill": [...],
            "model_used": "deepseek-ai/DeepSeek-V3",
            "notes": "..."
        }
    """
    model = model or LLM_MODEL
    client = _get_client(model, api_key=api_key, api_base=api_base)

    # 自动检测实验类型
    if experiment_type is None:
        experiment_type = detect_experiment_type(ocr_text)

    # 获取该实验类型的标准结构
    exp_info = EXPERIMENT_TYPES.get(experiment_type, EXPERIMENT_TYPES["物理实验"])
    standard_sections = exp_info["sections"]
    type_requirements = exp_info["requirements"]

    # 领域专项规范：数电/模电/信号与系统/大学物理 注入强写作模板（避免照抄指导书）
    domain_codes = ["电路/电子实验", "物理实验"]
    domain_spec_block = DOMAIN_SPEC if experiment_type in domain_codes else ""

    # 模板里真实存在的章节优先于实验类型的标准章节：
    # 老师模板多出「误差分析」「思考题」时，要让模型照着模板的栏位写。
    tpl_sections = _sections_from_template_summary(template_summary) or standard_sections

    # 构建系统 prompt（含知识库规范）
    system_prompt = f"""你是一位中国高校工科本科生，正在写自己刚做完的实验的实验报告。
不是"撰写助手"，不追求华丽的排比句，只把做过的事老老实实写清楚。

## 你的任务
1. 根据实验指导手册的 OCR 识别内容，提取关键信息
2. 将信息填入实验报告模板的对应位置
3. 对内容进行适度润色，使其通顺专业，但**绝不编造数据**
4. **最重要：把指导书内容重构为一份"本次实验报告"，而不是照抄指导书**（见下方领域规范）
5. 去除重复内容，保持简洁
6. 语言要像一个学生自己写的：具体、朴实、长短句错落，**不要有 AI 腔**

## 本次实验类型：{experiment_type}
标准章节：{' → '.join(tpl_sections)}
特殊要求：{type_requirements}

{ANTI_AI_STYLE}

## 学术格式规范
{POLISH_GUIDELINES}

{("## 领域专项写作规范（本实验必须照做）\n" + domain_spec_block) if domain_spec_block else ""}

**绝对禁止**：
- 编造任何实验数据！只使用 OCR 文本中明确出现的数据
- 如果某个章节在 OCR 文本中找不到对应内容，标记为 [待补充]
- 不要用"我""我们"，用第三人称（但句子要自然，不要靠"本文""笔者"堆砌）
- 表格数据必须保持原始数值不变
- **不要整段照抄指导书**：各栏必须收窄到本次实际用到的器件/电路/步骤
- **不要出现上方"AI 腔"里列出的任何一个套话句式**"""

    # ── 本课程参考范例（kb/ 知识库，可选）──────────────────────────────
    # 库里没东西、或检索出错时静默跳过——知识库坏了不能拖垮生成。
    if os.getenv("KB_ENABLE", "1").lower() not in ("0", "false", "off", ""):
        try:
            from modules.retrieval import build_context
            _kb_block = build_context(
                ocr_text,
                exp_type=experiment_type,
                k=int(os.getenv("KB_TOP_K", "3")),
                max_chars=int(os.getenv("KB_MAX_CHARS", "1500")),
            )
        except Exception as _e:
            print(f"[retrieval] 跳过知识库检索：{type(_e).__name__}: {_e}")
            _kb_block = ""
        if _kb_block:
            system_prompt += f"""

## 本课程参考范例（只学写法，不抄内容）
以下是本校本课程往届报告的片段，仅用于模仿**文风、段落详略、公式与表格的呈现方式**。
⚠️ 范例中的实验数据、器件型号、数值结果一律不得写入本次报告——本次报告的每个数字
都必须来自上面的 OCR 原文；范例与指导书冲突时，以指导书为准。

{_kb_block}
"""

    # 构建 user prompt
    _req = (requirement or "").strip()
    _req_block = (
        f"\n## 用户的额外要求（必须满足，优先级高于默认写法）\n{_req}\n"
        if _req else ""
    )

    _md = (source_markdown or "").strip()
    _md_block = (
        "\n## 指导手册结构化原文（Markdown，已保留标题/列表/表格层级）\n"
        "⚠️ 下面的【实验目的】【实验仪器】小节请**逐条原样照抄**，"
        "不要改写、不要增删、不要重新概括。这是最权威的原文。\n"
        f"{_md[:9000]}\n"
        if _md else ""
    )

    # JSON 骨架按模板动态生成，模板有几栏就输出几栏
    _skeleton = json.dumps(
        {"sections": {n: "" for n in tpl_sections}, "notes": "补充说明（可空）"},
        ensure_ascii=False, indent=2,
    )

    user_prompt = f"""## 实验类型
{experiment_type}

## 实验报告模板结构
{template_summary}

## 实验指导手册 OCR 识别内容
{ocr_text[:8000]}
{_md_block}
## 识别到的表格数据
{chr(10).join(str(t) for t in ocr_tables) if ocr_tables else '（无表格数据）'}
{_req_block}
【模板章节对齐（硬性要求）】
sections 的键**只能**是这 {len(tpl_sections)} 个：「{'、'.join(tpl_sections)}」
一个都不能多、一个都不能少，顺序也要一致。
- **绝对禁止**输出上面没有的章节；模板没给位置的思考题/设计内容，写进最后一个正文段落里，
  不要自己另开新章节。
- 判断依据只有模板结构，不要依据指导书的章节数。

【表格数据格式】凡涉及实验数据（真值表、测量数据表、误差表等），请把数据写成 Markdown 表格
嵌入对应章节的文本中，格式如下：
| 输入A | 输入B | 输出电压/V | 输出逻辑 |
| --- | --- | --- | --- |
| 0 | 0 | 3.50 | 1 |
| 0 | 1 | 5.04 | 1 |
表格必须以"|"开头和结尾，第一行为表头，第二行为 ----- 分隔行。数据必须来自 OCR 识别内容，严禁编造。

请按以下 JSON 格式返回填入内容（严格 JSON，不要输出 JSON 之外的文字）：
{_skeleton}"""

    api_kwargs = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.3,
        "max_tokens": 8192,
    }

    for attempt in range(MAX_RETRY):
        try:
            response = client.chat.completions.create(**api_kwargs)
            content = response.choices[0].message.content

            m = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', content, re.DOTALL)
            if m:
                content = m.group(1).strip()
            result = _parse_json_robust(content)
            if result is None:
                raise json.JSONDecodeError("无法解析", content, 0)

            secs = result.get("sections", {})
            # ── 文风自检：命中 AI 腔就带着具体修改建议重试一次 ──
            audit = audit_sections(secs)
            if audit and attempt < MAX_RETRY - 1:
                print(f"[filler] 文风自检命中 {sum(len(v) for v in audit.values())} 处 AI 腔，"
                      f"带着修改建议重试…")
                _fix_lines = [f"- 删除/改写「{frag}」（{tag}）" for v in audit.values()
                              for frag, tag, _h in v[:3]]
                nudge = (
                    "\n\n【上一版文风不合格，必须改写】\n"
                    + "\n".join(_fix_lines)
                    + "\n请重写这些句子，不要用套话，直接陈述具体事实与数据。"
                )
                api_kwargs["messages"][1]["content"] = user_prompt + nudge
                time.sleep(0.5)
                continue

            notes = result.get("notes", "")
            _audit_txt = format_audit(audit)
            if _audit_txt:
                notes = (notes + "\n" + _audit_txt).strip()

            return {
                "sections": secs,
                "tables_fill": result.get("tables_fill", []),
                "notes": notes,
                "model_used": model,
                "experiment_type": experiment_type,
                "ai_audit": audit,
                "raw_response": content
            }

        except json.JSONDecodeError:
            print(f"[filler] JSON 解析失败，第 {attempt+1}/{MAX_RETRY} 次重试...")
            time.sleep(1)
        except Exception as e:
            print(f"[filler] API 调用异常: {e}，第 {attempt+1}/{MAX_RETRY} 次重试...")
            last_error = str(e)
            time.sleep(2)

    return {
        "sections": {},
        "tables_fill": [],
        "notes": f"AI 填充失败：{last_error or '请检查 API 配置'}",
        "model_used": model,
        "experiment_type": experiment_type,
        "raw_response": ""
    }


# ══════════ 文风自检：检测"AI 味"套话 ══════════
# 生成后用正则扫一遍，把命中项写进 notes，并在重试时作为强约束回灌。
_AI_TELLS = [
    (r"通过本次实验[,，]?\s*[^。]{0,40}(加深|认识|理解)", "通过本次实验…加深了理解"),
    (r"为后续[^。]{0,20}(奠定|打下)[^。]{0,10}基础", "为后续…奠定基础"),
    (r"不仅[^。]{0,30}而且", "不仅…而且…"),
    (r"综?上?所述|总而言之|值得注意的是|需要指出的是|总体而言|整体而言", "综上所述/值得注意的是"),
    (r"充分(体现|说明|证明|反映)", "充分体现/证明"),
    (r"有效地?(验证|提高|掌握|实现)", "有效验证/提高"),
    (r"全面地?(掌握|理解|分析)", "全面掌握"),
    (r"极大地|显著地?|深入地?", "极大地/显著地/深入地"),
    (r"(掌握|熟悉|了解)[^。]{0,20}[，,]\s*(掌握|熟悉|了解)", "掌握…熟悉…了解…三连排比"),
    (r"(→|—>|->)", "箭头串流程"),
    (r"完整(的)?(流程|过程|链路)", "完整流程"),
    (r"既[^。]{0,20}又", "既…又…"),
    (r"具有一定(的)?(参考|指导|实用)(价值|意义)", "具有一定参考价值"),
    (r"本次实验(以|围绕)[^。]{0,30}(为例|展开)", "本次实验以…为例"),
]

_AUTO_FIX_HINTS = {
    "通过本次实验…加深了理解": "删掉整句，改为具体写出这次得到的结论或数据",
    "为后续…奠定基础": "删掉整句",
    "不仅…而且…": "拆成两句独立的陈述",
    "综上所述/值得注意的是": "删掉该连接词，直接用内容承接",
    "充分体现/证明": "删掉副词“充分”，或整句改为具体事实",
    "有效验证/提高": "删掉“有效”，只留“验证”",
    "全面掌握": "删掉“全面”",
    "极大地/显著地/深入地": "删掉该副词",
    "掌握…熟悉…了解…三连排比": "改成只写本次真正做到的一两件事，不要凑三项",
    "箭头串流程": "把箭头串改写成完整句子，如“先…再…最后…”",
    "完整流程": "删掉“完整流程”，直接说做了什么",
    "既…又…": "拆成两句",
    "具有一定参考价值": "删掉整句",
    "本次实验以…为例": "改为直接陈述实验对象",
}


def detect_ai_tells(text: str) -> list:
    """返回命中的 AI 腔套话列表 [(命中片段, 问题类型, 修改建议)]"""
    hits = []
    seen_tags = set()
    t = text or ""
    for pat, tag in _AI_TELLS:
        # 同一类型只报一次（箭头串、三连排比这类会重复命中）
        if tag in seen_tags:
            continue
        m = re.search(pat, t)
        if m:
            frag = m.group(0)
            if len(frag) > 40:
                frag = frag[:40] + "…"
            hits.append((frag, tag, _AUTO_FIX_HINTS.get(tag, "")))
            seen_tags.add(tag)
            if len(hits) >= 12:
                break
    return hits


def audit_sections(sections: dict) -> dict:
    """逐章节检查 AI 味，返回 {章节: [(片段, 类型, 建议)]}"""
    out = {}
    for name, txt in (sections or {}).items():
        if not txt or not txt.strip():
            continue
        hits = detect_ai_tells(txt)
        if hits:
            out[name] = hits
    return out


def format_audit(audit: dict) -> str:
    """把自检结果格式化成可读的一段话（用于 notes / 日志）"""
    if not audit:
        return ""
    total = sum(len(v) for v in audit.values())
    lines = [f"文风自检：发现 {total} 处 AI 腔表述，建议人工替换"]
    for name, hits in audit.items():
        lines.append(f"· {name}：")
        for frag, tag, hint in hits[:4]:
            lines.append(f"    「{frag}」→ {tag}" + (f"（{hint}）" if hint else ""))
    return "\n".join(lines)


def compare_models(
    ocr_text: str,
    ocr_tables: list,
    template_summary: str,
    models: list = None,
) -> dict:
    """
    多模型对比生成 — 用多个模型生成同一份报告，便于选择最佳结果

    参数:
        models: 模型 ID 列表，默认用 ['DeepSeek-V3', 'Qwen3.5-27B']
    返回:
        {"model_id": {...report...}, ...}
    """
    if models is None:
        models = ["deepseek-ai/DeepSeek-V3", "Qwen/Qwen3.5-27B"]

    results = {}
    for model_id in models:
        print(f"[filler] 正在用 {model_id} 生成...")
        results[model_id] = generate_report(ocr_text, ocr_tables, template_summary, model=model_id)
    return results


def match_section(section_name: str, filled_sections: dict) -> str:
    """模糊匹配章节名称 → 返回对应填充内容"""
    if section_name in filled_sections:
        return filled_sections[section_name]
    for key, value in filled_sections.items():
        if key in section_name or section_name in key:
            return value
    return ""


# ── 便捷：返回可用的模型列表 ──

def list_available_models() -> list:
    """返回 [(model_id, label, desc), ...]"""
    return [(k, v["label"], v["desc"]) for k, v in SILICONFLOW_MODELS.items()]


if __name__ == "__main__":
    print("可用模型:")
    for mid, label, desc in list_available_models():
        print(f"  {mid:<40} {label:<25} {desc}")

    # 测试
    test_ocr = "实验名称：RC电路瞬态过程研究\n实验目的：观察RC电路的充放电过程\n实验器材：电阻箱、电容箱、示波器..."
    exp_type = detect_experiment_type(test_ocr)
    print(f"\n检测到实验类型: {exp_type}")
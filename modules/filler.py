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
from config import LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, MAX_RETRY
from modules.knowledge import EXPERIMENT_TYPES, FORMAT_RULES, POLISH_GUIDELINES

_client_cache = {}

# ─── 硅基流动可用模型列表 ───

SILICONFLOW_MODELS = {
    "deepseek-ai/DeepSeek-V3":       {"label": "DeepSeek V3 ⭐推荐",   "desc": "综合能力最强，中英文俱佳"},
    "deepseek-ai/DeepSeek-V3.2":     {"label": "DeepSeek V3.2",       "desc": "V3 最新迭代，推理更强"},
    "deepseek-ai/DeepSeek-R1":       {"label": "DeepSeek R1",         "desc": "深度推理模型，适合复杂分析"},
    "deepseek-ai/DeepSeek-V4-Pro":   {"label": "DeepSeek V4 Pro",     "desc": "最新旗舰，质量最高"},
    "Qwen/Qwen3.5-27B":              {"label": "Qwen 3.5 27B",        "desc": "阿里通义千问，中文优秀"},
    "Qwen/Qwen3.5-35B-A3B":          {"label": "Qwen 3.5 35B",        "desc": "Qwen 大杯，中文理解强"},
    "zai-org/GLM-5.3":               {"label": "GLM-5.3",             "desc": "智谱最新，中文优异"},
    "zai-org/GLM-5.3-Flash":         {"label": "GLM-5.3 Flash ⚡",    "desc": "智谱轻量版，速度快"},
    "Pro/deepseek-ai/DeepSeek-V3.2": {"label": "DeepSeek V3.2 Pro",   "desc": "Pro 加速版，更高并发"},
    "Qwen/Qwen3-32B":                {"label": "Qwen 3 32B",          "desc": "Qwen3 中杯，平衡性能"},
    "THUDM/GLM-4-32B-0414":          {"label": "ChatGLM 4 32B",       "desc": "清华智谱，老牌中文模型"},
}


def _get_client(model: str = None) -> OpenAI:
    """获取 OpenAI 客户端（按 model 缓存，支持切换不同模型）"""
    global _client_cache
    key = model or LLM_MODEL
    if key not in _client_cache:
        _client_cache[key] = OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL)
    return _client_cache[key]


def detect_experiment_type(ocr_text: str) -> str:
    """
    自动判断实验类型（物理/化学/电路/力学/光学/计算机）
    """
    ocr_lower = ocr_text.lower()
    scores = {}
    keywords = {
        "物理实验": ["物理", "速度", "加速度", "力学", "电磁", "光学", "热学", "波动", "共振", "电阻", "电压", "电流", "电容", "电感"],
        "化学实验": ["化学", "滴定", "溶液", "反应", "浓度", "沉淀", "pH", "mol", "mL", "试剂", "酸碱"],
        "电路/电子实验": ["电路", "电阻", "电压", "示波器", "万用表", "焊接", "波形", "频率", "信号", "接线"],
        "力学实验": ["力学", "牛顿", "质量", "重力", "摩擦", "弹性", "转动", "力矩", "弹簧"],
        "光学实验": ["光学", "透镜", "折射", "反射", "干涉", "衍射", "波长", "偏振", "光栅", "激光"],
        "计算机/编程实验": ["编程", "算法", "代码", "Python", "C++", "数据结构", "编译", "调试", "程序"],
    }
    for exp_type, kws in keywords.items():
        scores[exp_type] = sum(1 for kw in kws if kw in ocr_lower)
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "物理实验"


def generate_report(
    ocr_text: str,
    ocr_tables: list,
    template_summary: str,
    model: str = None,
    experiment_type: str = None,
) -> dict:
    """
    核心方法：根据 OCR 内容和模板结构，调用 AI 生成填入方案

    参数:
        ocr_text: 从指导手册识别出的全文
        ocr_tables: 识别出的表格数据
        template_summary: 模板结构摘要
        model: 指定模型（None 则用 config 默认）
        experiment_type: 实验类型（None 则自动检测）

    返回:
        {
            "sections": {"实验目的": "...", ...},
            "tables_fill": [...],
            "model_used": "deepseek-ai/DeepSeek-V3",
            "notes": "..."
        }
    """
    model = model or LLM_MODEL
    client = _get_client(model)

    # 自动检测实验类型
    if experiment_type is None:
        experiment_type = detect_experiment_type(ocr_text)

    # 获取该实验类型的标准结构
    exp_info = EXPERIMENT_TYPES.get(experiment_type, EXPERIMENT_TYPES["物理实验"])
    standard_sections = exp_info["sections"]
    type_requirements = exp_info["requirements"]

    # 构建系统 prompt（含知识库规范）
    system_prompt = f"""你是一个专业的工科实验报告撰写助手，服务于中国高校学生。

## 你的任务
1. 根据实验指导手册的 OCR 识别内容，提取关键信息
2. 将信息填入实验报告模板的对应位置
3. 对内容进行适度润色，使其通顺专业，但**绝不编造数据**
4. 去除重复内容，保持简洁
5. 严格按高校学术规范组织语言

## 本次实验类型：{experiment_type}
标准章节：{' → '.join(standard_sections)}
特殊要求：{type_requirements}

## 学术格式规范
{POLISH_GUIDELINES}

**绝对禁止**：
- 编造任何实验数据！只使用 OCR 文本中明确出现的数据
- 如果某个章节在 OCR 文本中找不到对应内容，标记为 [待补充]
- 不要用"我""我们"，用第三人称
- 表格数据必须保持原始数值不变"""

    # 构建 user prompt
    user_prompt = f"""## 实验类型
{experiment_type}

## 实验报告模板结构
{template_summary}

## 实验指导手册 OCR 识别内容
{ocr_text[:8000]}

## 识别到的表格数据
{chr(10).join(str(t) for t in ocr_tables) if ocr_tables else '（无表格数据）'}

请按以下 JSON 格式返回填入内容（严格 JSON）：
{{
  "sections": {{
    "实验名称": "",
    "实验目的": "",
    "实验原理": "",
    "实验器材": "",
    "实验步骤": "",
    "实验数据": "",
    "数据处理": "",
    "实验结果": "",
    "误差分析": "",
    "实验结论": ""
  }},
  "tables_fill": [],
  "notes": "补充说明（可空）"
}}"""

    api_kwargs = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.3,
        "max_tokens": 4096,
    }

    for attempt in range(MAX_RETRY):
        try:
            response = client.chat.completions.create(**api_kwargs)
            content = response.choices[0].message.content

            import json, re
            m = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', content, re.DOTALL)
            if m:
                content = m.group(1).strip()
            result = json.loads(content)

            return {
                "sections": result.get("sections", {}),
                "tables_fill": result.get("tables_fill", []),
                "notes": result.get("notes", ""),
                "model_used": model,
                "experiment_type": experiment_type,
                "raw_response": content
            }

        except json.JSONDecodeError:
            print(f"[filler] JSON 解析失败，第 {attempt+1}/{MAX_RETRY} 次重试...")
            time.sleep(1)
        except Exception as e:
            print(f"[filler] API 调用异常: {e}，第 {attempt+1}/{MAX_RETRY} 次重试...")
            time.sleep(2)

    return {
        "sections": {},
        "tables_fill": [],
        "notes": "AI 填充失败，请检查 API 配置",
        "model_used": model,
        "experiment_type": experiment_type,
        "raw_response": ""
    }


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
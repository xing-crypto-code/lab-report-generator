"""
AI 填充模块 · filler.py
职责：调用大模型 API，将 OCR 识别出的实验内容
      智能填入模板对应的位置，润色、去重、补全。
"""
from openai import OpenAI
import time
from config import LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, MAX_RETRY

_client: OpenAI = None


def _get_client() -> OpenAI:
    """延迟初始化 OpenAI 客户端"""
    global _client
    if _client is None:
        _client = OpenAI(
            api_key=LLM_API_KEY,
            base_url=LLM_BASE_URL
        )
    return _client


def generate_report(
    ocr_text: str,
    ocr_tables: list,
    template_summary: str
) -> dict:
    """
    核心方法：根据 OCR 内容和模板结构，调用 AI 生成填入方案

    参数:
        ocr_text: 从指导手册识别出的全文
        ocr_tables: 识别出的表格数据
        template_summary: 模板结构摘要

    返回:
        {
            "sections": {
                "实验目的": "填写内容...",
                "实验步骤": "填写内容...",
                ...
            },
            "tables_fill": [
                {"table_index": 0, "rows": [["数据1", "数据2"], ...]},
                ...
            ],
            "raw_response": "AI 原始返回"
        }
    """
    client = _get_client()

    # 构建 prompt
    system_prompt = """你是一个专业的工科实验报告撰写助手。你的任务是：
1. 根据实验指导手册的 OCR 识别内容，提取关键信息
2. 将信息填入实验报告模板的对应位置
3. 对内容进行适度润色，使其通顺专业，但绝不编造数据
4. 去除重复内容，保持简洁
5. 按实验报告的学术规范组织语言

**重要约束**：
- 不要编造任何实验数据！只使用 OCR 文本中明确出现的数据
- 如果某个章节在 OCR 文本中找不到对应内容，标记为 [待补充]
- 保持学术严谨的语气
- 对于表格数据，保持原始数值不变"""

    user_prompt = f"""## 实验报告模板结构
{template_summary}

## 实验指导手册 OCR 识别内容
{ocr_text[:8000]}  # 截断以保证 token 限制

## 识别到的表格数据
{chr(10).join(str(t) for t in ocr_tables) if ocr_tables else '（无表格数据）'}

请按以下 JSON 格式返回填入内容（严格 JSON，不要加 markdown 代码块标记）：
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
  "notes": "补充说明"
}}

只返回 JSON，不要其他内容。"""

    # 构建 API 调用参数
    api_kwargs = {
        "model": LLM_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.3,
        "max_tokens": 4096,
    }
    # JSON 模式并非所有免费 API 都支持，先不加，靠 prompt 约束
    # 如需开启：api_kwargs["response_format"] = {"type": "json_object"}

    for attempt in range(MAX_RETRY):
        try:
            response = client.chat.completions.create(**api_kwargs)
            content = response.choices[0].message.content

            # 解析 JSON（兼容 markdown 代码块包裹的情况）
            import json, re
            # 尝试提取 ```json ... ``` 中的内容
            m = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', content, re.DOTALL)
            if m:
                content = m.group(1).strip()
            result = json.loads(content)

            return {
                "sections": result.get("sections", {}),
                "tables_fill": result.get("tables_fill", []),
                "notes": result.get("notes", ""),
                "raw_response": content
            }

        except json.JSONDecodeError:
            # JSON 解析失败，重试
            print(f"[filler] JSON 解析失败，第 {attempt+1}/{MAX_RETRY} 次重试...")
            time.sleep(1)
            continue
        except Exception as e:
            print(f"[filler] API 调用异常: {e}，第 {attempt+1}/{MAX_RETRY} 次重试...")
            time.sleep(2)
            continue

    return {"sections": {}, "tables_fill": [], "notes": "AI 填充失败，请检查 API 配置", "raw_response": ""}


def match_section(section_name: str, filled_sections: dict) -> str:
    """
    模糊匹配章节名称 → 返回对应填充内容
    """
    # 精确匹配
    if section_name in filled_sections:
        return filled_sections[section_name]

    # 模糊匹配：包含关系
    for key, value in filled_sections.items():
        if key in section_name or section_name in key:
            return value

    return ""


if __name__ == "__main__":
    # 简单测试
    test_ocr = "实验名称：RC电路瞬态过程研究\n实验目的：1. 观察RC电路的充放电过程\n实验器材：电阻箱、电容箱、示波器..."
    test_template = "## 模板\n- 实验目的\n- 实验器材\n- 实验步骤"
    result = generate_report(test_ocr, [], test_template)
    print(result)
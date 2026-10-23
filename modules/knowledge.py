"""
工科实验知识库 · knowledge.py
=============================
提供标准实验报告结构、常见实验类型模板、学术格式规范等
作为 AI 生成的参考基线，确保输出符合高校实验报告标准
"""

# ========== 常见实验类型 & 标准结构 ==========

EXPERIMENT_TYPES = {
    "物理实验": {
        "sections": [
            "实验名称", "实验目的", "实验原理", "实验器材",
            "实验步骤", "实验数据记录", "数据处理与分析",
            "误差分析", "实验结果", "讨论与结论"
        ],
        "requirements": "数据表格需含表头、单位；公式使用 LaTeX 格式；有效数字保留要规范"
    },
    "化学实验": {
        "sections": [
            "实验名称", "实验目的", "实验原理（反应方程式）",
            "实验试剂与仪器", "实验步骤", "实验数据记录",
            "数据处理", "结果与讨论", "结论"
        ],
        "requirements": "化学方程式要配平；浓度/体积单位规范；现象描述要准确"
    },
    "电路/电子实验": {
        "sections": [
            "实验名称", "实验目的", "实验原理与电路图",
            "实验仪器与器件", "实验步骤与接线图",
            "实验数据（波形/表格）", "数据处理",
            "误差分析", "结果与讨论"
        ],
        "requirements": "电路图要清晰标注参数；波形图要有坐标轴标签；数据表格要有量纲"
    },
    "力学实验": {
        "sections": [
            "实验名称", "实验目的", "实验原理（含公式推导）",
            "实验仪器", "实验步骤", "实验数据",
            "数据处理与图表", "误差分析",
            "结果与讨论"
        ],
        "requirements": "力学参数要有单位；图表要有轴标签和误差棒；有效数字要规范"
    },
    "计算机/编程实验": {
        "sections": [
            "实验名称", "实验目的", "实验环境",
            "算法/原理说明", "实验步骤与代码",
            "实验结果（截图/输出）", "结果分析",
            "实验总结"
        ],
        "requirements": "代码块用等宽字体；截图要清晰；输出结果要如实记录"
    },
    "光学实验": {
        "sections": [
            "实验名称", "实验目的", "实验原理（光路图）",
            "实验仪器", "实验步骤", "实验数据",
            "数据处理", "误差分析",
            "结果与讨论"
        ],
        "requirements": "光路图要标注角度/距离；干涉/衍射图样要清晰"
    },
}

# ========== 学术格式规范 ==========

FORMAT_RULES = {
    "标题": "宋体/黑体，三号（16pt），居中加粗",
    "一级标题": "宋体/黑体，四号（14pt），加粗",
    "正文": "宋体，小四（12pt），1.5倍行距",
    "表格标题": "黑体，五号（10.5pt），表头加粗，表格居中",
    "数据": "Times New Roman，五号，保留适当有效数字",
    "公式": "居中排列，右侧标注编号 (1), (2)...",
    "图表标注": "图下方居中标注图号和图名；表上方居中标注表号和表名",
}

# ========== 润色去重提示 ==========

POLISH_GUIDELINES = """
## 实验报告撰写规范（AI 请严格遵守）

1. **数据真实性**：实验数据必须来自 OCR 识别内容，绝不编造
2. **去重**：同一数据不要重复出现；同一结论不要多次表述
3. **学术语言**：使用客观、严谨的第三人称描述；避免"我""我们"等口语
4. **公式规范**：物理公式使用 LaTeX 语法，化学方程式必须配平
5. **表格规范**：三线表格式（学术标准），表头加粗，数据居中，单位标注在表头
6. **有效数字**：按照测量仪器的精度保留有效数字
7. **误差分析**：即使 OCR 没提到，也要简要分析可能的误差来源（系统误差、偶然误差）
8. **结论凝练**：200字以内，概括实验目的、方法、主要结果和收获
"""

# ========== 标准 Word 模板生成 ==========

def build_standard_template(experiment_type: str = "物理实验") -> bytes:
    """
    生成标准的实验报告 .docx 模板（内置模板，无需用户上传）
    返回 .docx 文件的 bytes
    """
    from docx import Document
    from docx.shared import Pt, Cm, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    import io

    doc = Document()

    # 页面设置
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    # 标题
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("实 验 报 告")
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.name = "黑体"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

    sections = EXPERIMENT_TYPES.get(experiment_type, EXPERIMENT_TYPES["物理实验"])["sections"]

    for i, sec in enumerate(sections):
        # 小节标题
        if i > 0:
            doc.add_paragraph()  # 空行分隔
        h = doc.add_paragraph()
        run = h.add_run(f"{'一二三四五六七八九十'[i] if i < 10 else str(i+1)}、{sec}")
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.name = "黑体"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

        # 占位内容
        p = doc.add_paragraph()
        run = p.add_run("（此区域将由 AI 自动填入实验数据与分析结果）")
        run.font.size = Pt(12)
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.color.rgb = None  # 默认黑色
        from docx.shared import RGBColor
        run.font.color.rgb = RGBColor(150, 150, 150)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


if __name__ == "__main__":
    data = build_standard_template("物理实验")
    with open("templates/physics_standard.docx", "wb") as f:
        f.write(data)
    print("Generated standard template for 物理实验")
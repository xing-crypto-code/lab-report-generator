"""
模板解析模块 · template.py
职责：读取老师发的 .docx 实验报告模板，分析其结构，
      找出需要填入内容的区块，同时保留全部格式信息。
"""
from docx import Document
from docx.shared import Inches, Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pathlib import Path
from typing import Any


def analyze_template(file_path: str) -> dict:
    """
    解析 .docx 模板，返回结构化的模板描述

    返回:
    {
        "paragraphs": [
            {"index": 0, "text": "实验名称：______", "style": "Heading 2", "runs": [...]},
            ...
        ],
        "tables": [
            {"index": 0, "rows": 5, "cols": 3, "headers": [...], "data_area": ...},
            ...
        ],
        "sections_found": ["实验目的", "实验器材", "实验步骤", ...],
        "placeholder_count": 5   # 下划线/空白占位数量
    }
    """
    doc = Document(file_path)
    paragraphs = []
    tables = []
    sections = []

    # 已知的实验报告常见章节标题
    KNOWN_SECTIONS = [
        "实验名称", "课程名称", "实验目的", "实验原理", "实验器材", "实验仪器",
        "实验设备", "实验步骤", "实验内容", "实验过程", "实验数据", "数据记录",
        "数据处理", "实验结果", "实验分析", "误差分析", "实验结论", "实验小结",
        "思考题", "讨论", "心得体会", "注意事项", "实验日期",
        "姓名", "学号", "班级", "同组人员", "指导教师", "教师评阅"
    ]

    # ------ 解析段落 ------
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        style_name = para.style.name if para.style else "Normal"

        para_info = {
            "index": i,
            "text": text,
            "style": style_name,
            "alignment": str(para.alignment) if para.alignment else None,
            "is_heading": "Heading" in style_name or "heading" in style_name,
            "is_empty": len(text) == 0,
            "has_placeholder": "___" in text or "______" in text or "（" in text,
            "runs": [{"text": r.text, "bold": r.bold, "font_name": r.font.name,
                      "font_size": str(r.font.size) if r.font.size else None}
                     for r in para.runs]
        }
        paragraphs.append(para_info)

        # 自动检测章节
        for section in KNOWN_SECTIONS:
            if section in text and section not in sections:
                sections.append(section)

    # ------ 解析表格 ------
    for i, table in enumerate(doc.tables):
        rows = len(table.rows)
        cols = len(table.columns)
        headers = []
        body_rows = []
        row_labels = []  # 每行首个单元格若像章节名则记录

        for j, row in enumerate(table.rows):
            cells = [cell.text.strip() for cell in row.cells]
            if j == 0:
                headers = cells
            else:
                body_rows.append(cells)
            # 检测表格首列/首格的章节标签（如"实验目的"、"课程名称"）
            first = cells[0] if cells else ""
            if first and len(first) <= 22 and first not in row_labels:
                is_label = any(s in first for s in KNOWN_SECTIONS)
                if is_label:
                    row_labels.append(first)
                    if first not in sections:
                        sections.append(first)

        table_info = {
            "index": i,
            "rows": rows,
            "cols": cols,
            "headers": headers,
            "body": body_rows,
            "row_labels": row_labels,
            "placeholder_cells": sum(
                1 for row in table.rows for cell in row.cells
                if "___" in cell.text or "______" in cell.text or cell.text.strip() == ""
            )
        }
        tables.append(table_info)

    # 计placeholder数量
    placeholder_count = sum(1 for p in paragraphs if p["has_placeholder"])
    for t in tables:
        placeholder_count += t["placeholder_cells"]

    return {
        "paragraphs": paragraphs,
        "tables": tables,
        "sections_found": sections,
        "placeholder_count": placeholder_count,
        "total_paragraphs": len(paragraphs),
        "total_tables": len(tables)
    }


def get_template_summary(analysis: dict) -> str:
    """
    生成模板的人类可读摘要，供 AI 理解模板结构
    """
    lines = []
    lines.append(f"## 实验报告模板结构")
    lines.append(f"- 总段落数: {analysis['total_paragraphs']}")
    lines.append(f"- 总表格数: {analysis['total_tables']}")
    lines.append(f"- 识别到的章节: {', '.join(analysis['sections_found'])}")
    lines.append(f"- 待填入占位数: {analysis['placeholder_count']}")
    lines.append("")

    lines.append("### 段落列表（含占位符的优先）")
    for p in analysis["paragraphs"]:
        if p["text"]:  # 只列出有内容的
            marker = "🔲" if p["has_placeholder"] else "  "
            lines.append(f"{marker} [{p['style']}] {p['text'][:80]}{'...' if len(p['text'])>80 else ''}")

    lines.append("")
    lines.append("### 表格列表（每行首列章节名 → 对应填写区）")
    for t in analysis["tables"]:
        lines.append(f"- 表格 {t['index']+1}: {t['rows']}行×{t['cols']}列, 表头={t['headers']}")
        if t.get("row_labels"):
            lines.append(f"  章节行: {'、'.join(t['row_labels'])}")
            lines.append("  → 生成 sections 时请使用这些章节名作为键")

    return "\n".join(lines)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        result = analyze_template(sys.argv[1])
        print(get_template_summary(result))
"""
文档生成模块 · generator.py
职责：将 AI 填充好的内容写入 .docx 模板，保持原格式不变，
      只替换/填入文字内容，避免格式错乱。
      支持输出 .docx 和 .pdf 两种格式。
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy
from pathlib import Path
import shutil

from config import OUTPUT_DIR


def fill_template(
    template_path: str,
    filled_data: dict,
    output_name: str = "实验报告"
) -> str:
    """
    核心方法：将 filled_data 填入模板并保存

    参数:
        template_path: 原始 .docx 模板路径
        filled_data: AI 填充返回的数据 {"sections": {...}, "tables_fill": [...]}
        output_name: 输出文件名（不含扩展名）

    返回:
        输出文件路径
    """
    doc = Document(template_path)

    # ---- 策略：遍历每个段落，匹配章节并替换占位符 ----

    for para in doc.paragraphs:
        original_text = para.text.strip()
        if not original_text:
            continue

        # 查找匹配的章节
        for section_name, content in filled_data.get("sections", {}).items():
            if not content:
                continue

            # 情况1：段落文本正好是章节名（如"实验目的"）
            if original_text == section_name:
                # 在段落后追加内容（新建一个段落，保持格式）
                _insert_content_after(para, content, doc)
                break

            # 情况2：段落包含章节名（如"二、实验目的"）
            elif section_name in original_text:
                _insert_content_after(para, content, doc)
                break

            # 情况3：段落包含占位符（下划线/空白）
            elif "___" in original_text or "______" in original_text:
                # 替换占位符为实际内容
                _replace_placeholder_in_para(para, original_text, content, section_name)
                break

    # ---- 填充表格 ----
    for table_fill in filled_data.get("tables_fill", []):
        table_index = table_fill.get("table_index", 0)
        rows_data = table_fill.get("rows", [])

        if table_index < len(doc.tables):
            table = doc.tables[table_index]
            _fill_table(table, rows_data)

    # ---- 保存 ----
    output_path = OUTPUT_DIR / f"{output_name}.docx"
    doc.save(str(output_path))

    return str(output_path)


def _insert_content_after(para, content: str, doc: Document):
    """
    在指定段落后插入内容段落，继承原段落的部分格式
    """
    # 找到该段落在文档中的位置
    parent = para._element.getparent()
    index = list(parent).index(para._element)

    # 新建段落
    new_para = doc.add_paragraph()
    run = new_para.add_run(content)

    # 继承字体大小（默认小四）
    run.font.size = Pt(12)
    run.font.name = "宋体"

    # 移动新段落到目标段落之后
    parent.remove(new_para._element)
    parent.insert(index + 1, new_para._element)


def _replace_placeholder_in_para(para, original_text: str, content: str, section_name: str):
    """
    替换段落中的占位符（下划线）为实际内容
    """
    # 仅当段落确实与当前 section 相关时才替换
    if section_name not in original_text:
        return

    # 清除原有 runs，写入新内容
    for run in para.runs:
        if "___" in run.text or "______" in run.text:
            run.text = run.text.replace("___", content).replace("______", content)
            return

    # 如果 runs 里没有占位符，尝试整体替换
    new_text = original_text.replace("___", content).replace("______", content)
    if new_text != original_text:
        for run in para.runs:
            run.text = ""
        if para.runs:
            para.runs[0].text = new_text


def _fill_table(table, rows_data: list):
    """
    填充表格数据（从第2行开始填充，保留表头）
    """
    for i, row_data in enumerate(rows_data):
        row_index = i + 1  # 跳过表头行
        if row_index < len(table.rows):
            row = table.rows[row_index]
            for j, cell_text in enumerate(row_data):
                if j < len(row.cells):
                    row.cells[j].text = str(cell_text)


def convert_to_pdf(docx_path: str) -> str:
    """
    .docx → .pdf 转换（Windows 上用 docx2pdf 调用 Word COM）
    """
    try:
        from docx2pdf import convert
        pdf_path = str(Path(docx_path).with_suffix(".pdf"))
        convert(str(docx_path), pdf_path)
        return pdf_path
    except Exception as e:
        # 回退：仅返回 docx
        print(f"[generator] PDF 转换失败 ({e})，仅提供 .docx")
        return docx_path


def generate_report_document(
    template_path: str,
    filled_data: dict,
    output_name: str = "实验报告",
    to_pdf: bool = False
) -> dict:
    """
    一站式生成：填写模板 + 可选转 PDF
    返回 {"docx": "路径", "pdf": "路径或None"}
    """
    docx_path = fill_template(template_path, filled_data, output_name)
    result = {"docx": docx_path, "pdf": None}

    if to_pdf:
        result["pdf"] = convert_to_pdf(docx_path)

    return result


if __name__ == "__main__":
    # 测试
    test_data = {
        "sections": {
            "实验名称": "RC电路瞬态过程研究",
            "实验目的": "1. 观察RC电路的充放电过程\n2. 测定RC电路的时间常数",
            "实验器材": "电阻箱、电容箱、示波器、直流电源",
        },
        "tables_fill": [
            {"table_index": 0, "rows": [
                ["1", "100", "0.01", "10.0"],
                ["2", "200", "0.02", "20.0"],
            ]}
        ]
    }
    # result = fill_template("demo_template.docx", test_data)
    # print(f"生成: {result}")
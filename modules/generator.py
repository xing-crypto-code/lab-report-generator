"""
文档生成模块 · generator.py
===========================
将 AI 填充的内容写入 .docx 模板，保持原格式，只替换文字。
内置三线表格式、学术字体规范（宋体正文/Times New Roman 数据）。
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pathlib import Path

from config import OUTPUT_DIR


def fill_template(template_path: str, filled_data: dict, output_name: str = "实验报告") -> str:
    """
    将 filled_data 填入模板并保存。
    优先匹配已有段落（原地替换内容），没匹配到的追加新段落。
    """
    doc = Document(template_path)
    sections = filled_data.get("sections", {})
    matched_keys = set()

    # ── 遍历已有段落，尝试匹配 ──
    for para in doc.paragraphs:
        original_text = para.text.strip()
        if not original_text:
            continue
        for section_name, content in sections.items():
            if not content or section_name in matched_keys:
                continue

            # 1) 段落刚好是章节名 → 追加内容到段落后面
            if original_text == section_name or original_text.endswith(section_name):
                _append_content(para, doc, content)
                matched_keys.add(section_name)
                break

            # 2) 段落包含章节名
            if section_name in original_text and ("___" in original_text or "______" in original_text):
                _replace_placeholder(para, content)
                matched_keys.add(section_name)
                break

    # ── 未匹配到的 section，追加到文末 ──
    for section_name, content in sections.items():
        if content and section_name not in matched_keys:
            _append_heading_with_content(doc, section_name, content)

    # ── 填充表格 ──
    for table_fill in filled_data.get("tables_fill", []):
        ti = table_fill.get("table_index", 0)
        rows_data = table_fill.get("rows", [])
        if ti < len(doc.tables):
            _fill_table(doc.tables[ti], rows_data)

    # ── 保存 ──
    output_path = OUTPUT_DIR / f"{output_name}.docx"
    doc.save(str(output_path))
    return str(output_path)


def _append_content(after_para, doc, content):
    """在 after_para 段落后插入正文段落，用宋体小四"""
    parent = after_para._element.getparent()
    idx = list(parent).index(after_para._element)
    p = doc.add_paragraph()
    run = p.add_run(content)
    run.font.size = Pt(12)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    parent.remove(p._element)
    parent.insert(idx + 1, p._element)


def _replace_placeholder(para, content):
    """替换段落中的占位符"""
    for run in para.runs:
        if "___" in run.text or "______" in run.text:
            run.text = content
            return
    # fallback
    new_text = para.text.replace("___", content).replace("______", content)
    if new_text != para.text:
        for run in para.runs:
            run.text = ""
        if para.runs:
            para.runs[0].text = new_text


def _append_heading_with_content(doc, section_name, content):
    """追加章节标题 + 正文内容"""
    # 标题
    h = doc.add_paragraph()
    run = h.add_run(section_name)
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "黑体"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    # 正文
    p = doc.add_paragraph()
    run = p.add_run(content)
    run.font.size = Pt(12)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')


def _fill_table(table, rows_data):
    """从表头下一行开始填充数据"""
    for i, row_data in enumerate(rows_data):
        ri = i + 1  # skip header
        if ri < len(table.rows):
            for j, cell_text in enumerate(row_data):
                if j < len(table.rows[ri].cells):
                    table.rows[ri].cells[j].text = str(cell_text)


def convert_to_pdf(docx_path: str) -> str:
    """docx → pdf（Windows 用 Word COM）"""
    try:
        from docx2pdf import convert
        pdf_path = str(Path(docx_path).with_suffix(".pdf"))
        convert(str(docx_path), pdf_path)
        return pdf_path
    except Exception as e:
        print(f"[generator] PDF 转换失败: {e}")
        return docx_path


def generate_report_document(template_path, filled_data, output_name="实验报告", to_pdf=False):
    docx_path = fill_template(template_path, filled_data, output_name)
    result = {"docx": docx_path, "pdf": None}
    if to_pdf:
        result["pdf"] = convert_to_pdf(docx_path)
    return result
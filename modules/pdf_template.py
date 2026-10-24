"""
PDF 模板转换模块 · pdf_template.py
=================================
把用户上传的 PDF 模板（教师线框模板）转成可编辑、可填充的 .docx：
- pdfplumber 提取每一页的表格结构 → 重建为 Word 表格线框
- 提取文本 → 重建为段落（含章节占位符），供 analyze_template / fill_template 使用
这样 PDF 模板就能像 .docx 模板一样走 "识别章节名 → 填入内容" 流程。
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm

# 章节关键词，检测文本行是否为章节标签
_SECTION_KEYWORDS = [
    "实验名称", "课程名称", "实验目的", "实验原理", "实验仪器", "实验器材",
    "实验设备", "实验步骤", "实验内容", "实验过程", "实验数据", "数据记录",
    "数据处理", "实验结果", "实验分析", "误差分析", "实验结论", "实验小结",
    "思考题", "讨论", "心得体会", "注意事项", "班级", "姓名", "学号",
    "同组人员", "指导教师", "教师评阅", "成绩", "实验日期",
]

_BLANK_LABELS = ["___", "＿＿", "______", "＿＿＿＿", "　"]


def _extra_table_cells_to_nodes(pdf_table):
    """把 pdfplumber 的一个表格（list[list[cell]]）转成 python-docx 需要的二维列表，
    空的 cell 视为占位、单元格文本即为原内容。"""
    rows = []
    for row in pdf_table:
        row_nodes = []
        for cell in row:
            text = ""
            if isinstance(cell, dict):      # pdfplumber 新版返回 dict {text, x0...}
                text = (cell.get("text") or "").strip()
            elif isinstance(cell, (list, tuple)):
                text = " ".join(str(c) for c in cell).strip()
            else:
                text = str(cell or "").strip()
            # 单元格内容若看起来是占位（很短的空白/下划线），保留为空框架
            if text in _BLANK_LABELS:
                text = ""
            row_nodes.append(text if text else "")
        rows.append(row_nodes)
    return rows


def _set_table_borders(table):
    """给表格加单线边框（避免依赖 Table Grid 样式，python-docx 新文档没有该样式）"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)


def convert_pdf_to_docx(file_path: str, output_docx: str) -> str:
    """把 PDF 模板转成 .docx，返回输出路径。"""
    import pdfplumber

    pdf = pdfplumber.open(file_path)
    doc = Document()

    # docx 全局默认字体（尽量接近中文模板视觉）
    try:
        style = doc.styles["Normal"]
        style.font.name = "宋体"
        style.font.size = Pt(11)
        from docx.oxml.ns import qn
        style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    except Exception:
        pass

    for pno, page in enumerate(pdf.pages):
        # 1) 每个 PDF 页起一个新 Word 页视角（用分隔符区分，但不强制分页，避免过度空白）
        if pno > 0:
            doc.add_paragraph()  # 简单空段作分隔

        # 2) 该页表格（按顶部位置排序，尽量保持阅读顺序）
        try:
            raw_tables = page.extract_tables()
        except Exception:
            raw_tables = []

        # 3) 表格 → Word 表格线框
        for raw in raw_tables:
            rows = _extra_table_cells_to_nodes(raw)
            if not rows:
                continue
            ncols = max(len(r) for r in rows)
            table = doc.add_table(rows=len(rows), cols=max(ncols, 1))
            _set_table_borders(table)
            for ri, r in enumerate(rows):
                for ci in range(ncols):
                    cell_text = r[ci] if ci < len(r) else ""
                    table.cell(ri, ci).text = cell_text

        # 4) 表格外的文本 → 段落
        try:
            txt = page.extract_text() or ""
        except Exception:
            txt = ""
        for line in txt.splitlines():
            line = line.strip()
            if not line:
                continue
            para = doc.add_paragraph(line)
            # 若该行像是章节标签，套用加粗以利识别
            if any(k in line for k in _SECTION_KEYWORDS):
                for run in para.runs:
                    run.bold = True

    pdf.close()
    out = Path(output_docx)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return str(out)
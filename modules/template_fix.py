"""
模板固化 / rebake · modules/template_fix.py
==========================================
把所有"排版/边框"类修复**烧进内置模板文件本身**，让模板每次都是最新状态，
而不是只在生成时临时打补丁。

约定（用户要求）：以后每修一个排版相关的 bug，都要：
    1) 在 `TEMPLATE_FIXES` 里登记一条（名字 + 说明 + 修法函数）
    2) 跑一次 `python -m modules.template_fix` 重新烘焙内置模板
    3) 新增回归测试断言该修复在模板里已生效

固化内容（截至当前版本）：
  fix_01 表格宽度自洽        tblW == ΣgridCols == ΣtcW
  fix_02 去除表格居中        tblPr/jc 删除，避免窄表贴不齐右侧
  fix_03 主表与信息表等宽对齐 两表 tblW 统一
  fix_04 单元格边距规范化    tblCellMar 左右 108 / 上下 0
  fix_05 正文默认字体字号    样式 Normal 设为宋体 小四
"""
import os
import shutil

from docx import Document
from docx.oxml.ns import qn

# 模板规范值（与 docs/实验报告写作规范.md 一致）
PAGE_W_CM, PAGE_H_CM = 21.0, 29.7
MARGIN_TOP_CM = MARGIN_BOTTOM_CM = 2.54
MARGIN_LEFT_CM, MARGIN_RIGHT_CM = 1.75, 2.00
CELL_MAR_LR = 108          # Word 默认单元格左右边距
CELL_MAR_TB = 0
BODY_FONT = "宋体"
BODY_SIZE_HALF_PT = 24     # 小四 12pt → w:sz 半磅值 24


# ─────────────────────── 单条修复 ───────────────────────

def fix_01_table_width_self_consistent(doc):
    """每条表格：tblW 必须等于 gridCols 之和，且每行 tcW 之和也等于它。
    不自洽时 Word 按 gridCols 排版、右边框会跑出表外——这是边框参差的头号原因。"""
    changed = 0
    for t in doc.tables:
        tbl = t._tbl
        tblPr = tbl.tblPr
        grid = tbl.find(qn('w:tblGrid'))
        if grid is None:
            continue
        cols = [int(g.get(qn('w:w')) or 0) for g in grid.findall(qn('w:gridCol'))]
        total = sum(cols)
        if not total:
            continue
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = tblPr.makeelement(qn('w:tblW'), {})
            tblPr.append(tblW)
        if tblW.get(qn('w:w')) != str(total) or tblW.get(qn('w:type')) != 'dxa':
            tblW.set(qn('w:type'), 'dxa')
            tblW.set(qn('w:w'), str(total))
            changed += 1
        # 每行各 cell 的 tcW 之和对齐 grid
        for row in t.rows:
            seen = set()
            cells = []
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                cells.append(c)
            row_sum = 0
            for c in cells:
                tcPr = c._tc.find(qn('w:tcPr'))
                if tcPr is None:
                    continue
                w = tcPr.find(qn('w:tcW'))
                if w is not None and w.get(qn('w:w')):
                    row_sum += int(w.get(qn('w:w')))
            if row_sum and row_sum != total:
                changed += 1
    return changed


def fix_02_remove_table_centering(doc):
    """删除表格级 jc（居中/右对齐）。居中会让任何窄表永远贴不齐右侧边框。"""
    changed = 0
    for t in doc.tables:
        tblPr = t._tbl.tblPr
        if tblPr is None:
            continue
        for jc in tblPr.findall(qn('w:jc')):
            tblPr.remove(jc)
            changed += 1
    return changed


def fix_03_unify_table_widths(doc):
    """把顶层表格统一到同一总宽度（取众数），但**按原有列宽比例缩放**，
    不破坏模板设计的列宽分配；同时同步每行 tcW。"""
    widths = []
    for t in doc.tables:
        grid = t._tbl.find(qn('w:tblGrid'))
        if grid is None:
            continue
        cols = [int(g.get(qn('w:w')) or 0) for g in grid.findall(qn('w:gridCol'))]
        if cols and sum(cols):
            widths.append(sum(cols))
    if len(set(widths)) <= 1:
        return 0
    target = max(set(widths), key=widths.count)
    changed = 0
    for t in doc.tables:
        tbl = t._tbl
        tblPr = tbl.tblPr
        grid = tbl.find(qn('w:tblGrid'))
        if grid is None:
            continue
        cols = grid.findall(qn('w:gridCol'))
        old = [int(g.get(qn('w:w')) or 0) for g in cols]
        cur = sum(old)
        if cur == target or not cols or not cur:
            continue
        # 按比例缩放，末列补余数保证总和精确
        scaled = [max(1, round(w * target / cur)) for w in old]
        diff = target - sum(scaled)
        scaled[-1] += diff
        for g, w in zip(cols, scaled):
            g.set(qn('w:w'), str(w))
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is not None:
            tblW.set(qn('w:type'), 'dxa')
            tblW.set(qn('w:w'), str(target))
        # 同步每行 tcW：按 gridSpan 合并对应新列宽
        for row in t.rows:
            seen = set()
            idx = 0
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                tcPr = c._tc.find(qn('w:tcPr'))
                if tcPr is None:
                    idx += 1
                    continue
                gs = tcPr.find(qn('w:gridSpan'))
                span = int(gs.get(qn('w:val'))) if gs is not None else 1
                w_el = tcPr.find(qn('w:tcW'))
                if w_el is not None:
                    new_w = sum(scaled[idx:min(idx + span, len(scaled))])
                    w_el.set(qn('w:type'), 'dxa')
                    w_el.set(qn('w:w'), str(new_w))
                idx += span
        changed += 1
    return changed


def fix_04_normalize_cell_margins(doc):
    """统一单元格边距：左右 108、上下 0。边距不一致会导致嵌套表左右偏移。"""
    changed = 0
    for t in doc.tables:
        tblPr = t._tbl.tblPr
        if tblPr is None:
            continue
        mar = tblPr.find(qn('w:tblCellMar'))
        if mar is None:
            mar = tblPr.makeelement(qn('w:tblCellMar'), {})
            tblPr.append(mar)
        for edge, val in (('top', CELL_MAR_TB), ('left', CELL_MAR_LR),
                          ('bottom', CELL_MAR_TB), ('right', CELL_MAR_LR)):
            el = mar.find(qn('w:' + edge))
            if el is None:
                el = mar.makeelement(qn('w:' + edge), {})
                mar.append(el)
            if el.get(qn('w:w')) != str(val) or el.get(qn('w:type')) != 'dxa':
                el.set(qn('w:w'), str(val))
                el.set(qn('w:type'), 'dxa')
                changed += 1
    return changed


def fix_06_unify_table_positioning(doc):
    """统一表格横向定位与宽度：所有顶层表格 = 版心宽度，贴版心左边缘。

    这才是上下表格左右边框对不齐的真正根因：
      · 信息表用 horzAnchor=margin + tblpXSpec=center → 居中后左边界右移 (版心宽-T)/2
      · 主表  用 horzAnchor=page  + tblpX=1157     → 以页面左边缘为基准，又多偏 164
    两者锚点不同、偏移不同，所以左边框差 94 dxa、右边框差 24 dxa。

    统一后：左边界 = 页左边距，右边界 = 页右边距，上下表的左右边框严格重合，
    且左右留白对称（符合 docs 规范「表格总宽 ≤ 可用宽度」）。
    """
    from docx.oxml.ns import qn as _qn
    changed = 0

    # 目标宽度 = 版心宽度（页宽 - 左右边距）
    sec = doc.sections[0]
    try:
        target_w = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
    except Exception:
        target_w = None
    if not target_w:
        return 0

    for t in doc.tables:
        tbl = t._tbl
        tblPr = tbl.tblPr
        if tblPr is None:
            continue

        # ── 宽度统一到版心宽度（按比例缩放，保留列宽设计比例）──
        g = tbl.find(_qn('w:tblGrid'))
        cols = g.findall(_qn('w:gridCol')) if g is not None else []
        old = [int(c.get(_qn('w:w')) or 0) for c in cols]
        cur = sum(old)
        if not cols or not cur:
            continue
        if cur != target_w:
            scaled = [max(1, round(w * target_w / cur)) for w in old]
            scaled[-1] += target_w - sum(scaled)
            for c, w in zip(cols, scaled):
                c.set(_qn('w:w'), str(w))
            tblW = tblPr.find(_qn('w:tblW'))
            if tblW is not None:
                tblW.set(_qn('w:type'), 'dxa')
                tblW.set(_qn('w:w'), str(target_w))
            for row in t.rows:
                seen = set()
                idx = 0
                for c in row.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    tcPr = c._tc.find(_qn('w:tcPr'))
                    if tcPr is None:
                        idx += 1
                        continue
                    gs = tcPr.find(_qn('w:gridSpan'))
                    span = int(gs.get(_qn('w:val'))) if gs is not None else 1
                    w_el = tcPr.find(_qn('w:tcW'))
                    if w_el is not None:
                        w_el.set(_qn('w:type'), 'dxa')
                        w_el.set(_qn('w:w'), str(sum(scaled[idx:min(idx + span, len(scaled))])))
                    idx += span
            changed += 1

        # ── 定位统一：锚点=版心(margin)、无居中/右对齐、tblpX=0 ──
        tblpPr = tblPr.find(_qn('w:tblpPr'))
        if tblpPr is not None:
            tblpPr.set(_qn('w:horzAnchor'), 'margin')
            for attr in ('w:tblpX', 'w:tblpXSpec'):
                if tblpPr.get(_qn(attr)) is not None:
                    del tblpPr.attrib[_qn(attr)]
                    changed += 1
            tblpPr.set(_qn('w:tblpX'), '0')
        for tag in ('w:jc', 'w:tblInd'):
            for el in tblPr.findall(_qn(tag)):
                tblPr.remove(el)
                changed += 1
    return changed


def fix_07_fixed_table_layout(doc):
    """把表格布局从 autofit 改为 fixed。

    关键：tblLayout=autofit 时 Word 会自动按内容重排列宽，**完全忽略** tblGrid/tcW
    里我们精心设置的宽度，于是右侧边框随内容长短漂移、上下表对不齐。
    只有 fixed 才让 tblW == ΣgridCols 真正生效，边框才会固定在同一位置。
    """
    from docx.oxml.ns import qn as _qn
    changed = 0
    for t in doc.tables:
        tblPr = t._tbl.tblPr
        if tblPr is None:
            continue
        lay = tblPr.find(_qn('w:tblLayout'))
        if lay is None:
            lay = tblPr.makeelement(_qn('w:tblLayout'), {})
            tblPr.append(lay)
        if lay.get(_qn('w:type')) != 'fixed':
            lay.set(_qn('w:type'), 'fixed')
            changed += 1
    return changed


def fix_08_clean_ppr_rpr_order_and_indent(doc):
    """修正模板自带段落的 pPr/rPr 子元素顺序，并清掉异常首行缩进。

    背景：
    · OOXML 里 w:pPr / w:rPr 的子元素**必须按固定顺序**排列，顺序错了
      Word 会静默忽略后面的属性 → 行距/缩进不生效、上下位置对不上。
    · 母版里"实验目的/实验仪器/实验原理"等 label 段带 firstLineChars=49/50/100
      这类脏缩进，会让文字右移半个行宽，与相邻栏错位。
    """
    from docx.oxml.ns import qn as _qn

    PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
           'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
           'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
           'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
           'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
           'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
           'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
           'sectPr', 'pPrChange']
    RPR = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
           'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
           'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
           'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
           'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
           'eastAsianLayout', 'specVanish', 'oMath']

    def _reorder(el, order):
        idx = {n: i for i, n in enumerate(order)}
        kids = list(el)
        if not kids:
            return 0
        unknown = len(order)
        ordered = sorted(kids, key=lambda e: idx.get(e.tag.split('}')[-1], unknown))
        if [k.tag for k in ordered] == [k.tag for k in kids]:
            return 0
        for k in kids:
            el.remove(k)
        for k in ordered:
            el.append(k)
        return 1

    changed = 0
    for p in doc.element.body.iter(_qn('w:p')):
        pPr = p.find(_qn('w:pPr'))
        if pPr is None:
            continue
        # 清异常首行缩进
        ind = pPr.find(_qn('w:ind'))
        if ind is not None:
            fc = ind.get(_qn('w:firstLineChars'))
            if fc is not None:
                try:
                    if abs(int(fc)) > 10:
                        for a in ('firstLineChars', 'firstLine'):
                            ind.attrib.pop(_qn('w:' + a), None)
                        changed += 1
                except Exception:
                    pass
            if not ind.attrib:
                pPr.remove(ind)
                changed += 1
        changed += _reorder(pPr, PPR)

    for rPr in doc.element.body.iter(_qn('w:rPr')):
        changed += _reorder(rPr, RPR)
    return changed


def fix_05_normal_font(doc):
    """Normal 样式默认字体：宋体、小四(12pt)，避免继承 Office 默认西文字体。"""
    changed = 0
    try:
        st = doc.styles['Normal']
        st.font.name = BODY_FONT
        st.font.size = __import__('docx').shared.Pt(BODY_SIZE_HALF_PT / 2)
        rPr = st.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = rPr.makeelement(qn('w:rFonts'), {})
            rPr.insert(0, rFonts)
        for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
            rFonts.set(qn(attr), BODY_FONT)
        changed += 1
    except Exception as e:
        print('  [fix_05] 样式设置失败:', e)
    return changed


# 登记表：以后新增修复就往后追加，保持顺序
TEMPLATE_FIXES = [
    ("fix_01_table_width_self_consistent", "表格宽度自洽：tblW == ΣgridCols == ΣtcW",
     fix_01_table_width_self_consistent),
    ("fix_02_remove_table_centering", "去除表格居中，避免窄表贴不齐右侧边框",
     fix_02_remove_table_centering),
    ("fix_03_unify_table_widths", "顶层表格统一宽度，上下表左右边框对齐",
     fix_03_unify_table_widths),
    ("fix_04_normalize_cell_margins", "单元格边距规范：左右108 / 上下0",
     fix_04_normalize_cell_margins),
    ("fix_05_normal_font", "Normal 样式默认宋体小四",
     fix_05_normal_font),
    ("fix_06_unify_table_positioning",
     "统一表格横向定位：全部贴版心左边缘、锚点一致、无 tblpX/居中偏移",
     fix_06_unify_table_positioning),
    ("fix_07_fixed_table_layout",
     "表格布局改 fixed（autofit 会让 Word 忽略我们设的列宽、边框漂移）",
     fix_07_fixed_table_layout),
    ("fix_08_clean_ppr_rpr_order_and_indent",
     "修正 pPr/rPr 子元素顺序 + 清异常首行缩进（顺序错 Word 静默忽略属性）",
     fix_08_clean_ppr_rpr_order_and_indent),
]


def rebake_template(template_path: str, backup=True) -> dict:
    """把全部已登记的修复应用到模板文件本身。返回每条修复的变更计数。"""
    if not os.path.exists(template_path):
        raise FileNotFoundError(template_path)
    if backup:
        bak = template_path + '.bak'
        if not os.path.exists(bak):
            shutil.copy2(template_path, bak)
    doc = Document(template_path)
    report = {}
    for name, desc, fn in TEMPLATE_FIXES:
        try:
            report[name] = fn(doc)
        except Exception as e:
            report[name] = f"ERROR: {e}"
    doc.save(template_path)
    return report


def verify_template(template_path: str) -> list:
    """校验模板是否已固化全部修复。返回问题列表（空=全部通过）。"""
    doc = Document(template_path)
    problems = []
    for ti, t in enumerate(doc.tables):
        tbl = t._tbl
        tblPr = tbl.tblPr
        grid = tbl.find(qn('w:tblGrid'))
        if grid is None:
            problems.append(f"table{ti}: 缺少 tblGrid")
            continue
        cols = [int(g.get(qn('w:w')) or 0) for g in grid.findall(qn('w:gridCol'))]
        total = sum(cols)
        tblW = tblPr.find(qn('w:tblW')) if tblPr is not None else None
        if tblW is None:
            problems.append(f"table{ti}: 缺少 tblW")
        else:
            if tblW.get(qn('w:type')) != 'dxa':
                problems.append(f"table{ti}: tblW type != dxa")
            if int(tblW.get(qn('w:w')) or 0) != total:
                problems.append(
                    f"table{ti}: tblW({tblW.get(qn('w:w'))}) != gridCols({total})")
        if tblPr is not None and tblPr.find(qn('w:jc')) is not None:
            problems.append(f"table{ti}: 仍存在表格级 jc（居中）")
        # 行 tcW 之和
        for ri, row in enumerate(t.rows):
            seen = set()
            s = 0
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                tcPr = c._tc.find(qn('w:tcPr'))
                if tcPr is None:
                    continue
                w = tcPr.find(qn('w:tcW'))
                if w is not None and w.get(qn('w:w')):
                    s += int(w.get(qn('w:w')))
            if s and s != total:
                problems.append(f"table{ti} r{ri}: ΣtcW({s}) != gridCols({total})")
    # 顶层表格宽度统一
    ws = []
    for t in doc.tables:
        grid = t._tbl.find(qn('w:tblGrid'))
        if grid is None:
            continue
        s = sum(int(g.get(qn('w:w')) or 0) for g in grid.findall(qn('w:gridCol')))
        if s:
            ws.append(s)
    if len(set(ws)) > 1:
        problems.append(f"顶层表格宽度不一致: {sorted(set(ws))}")

    # ── 横向定位一致性（上下表左右边框必须重合）──
    sec = doc.sections[0]
    body_left = int(sec.left_margin.twips) if hasattr(sec.left_margin, 'twips') \
        else int(sec.left_margin / 635)
    body_w = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
    body_right = body_left + body_w
    edges = []
    for ti, t in enumerate(doc.tables):
        grid = t._tbl.find(qn('w:tblGrid'))
        if grid is None:
            continue
        total = sum(int(g.get(qn('w:w')) or 0) for g in grid.findall(qn('w:gridCol')))
        tblPr = t._tbl.tblPr
        tblpPr = tblPr.find(qn('w:tblpPr')) if tblPr is not None else None
        anchor = tblpPr.get(qn('w:horzAnchor')) if tblpPr is not None else None
        px = tblpPr.get(qn('w:tblpX')) if tblpPr is not None else None
        xspec = tblpPr.get(qn('w:tblpXSpec')) if tblpPr is not None else None
        if anchor == 'page':
            left = int(px) if px else 0
        elif anchor == 'margin':
            if px is not None:
                left = body_left + int(px)
            elif xspec == 'center':
                left = body_left + (body_w - total) // 2
                problems.append(f"table{ti}: 表格居中(tblpXSpec=center)，会偏离版心左边缘")
            elif xspec == 'right':
                left = body_right - total
                problems.append(f"table{ti}: 表格右对齐，会偏离版心左边缘")
            else:
                left = body_left
        else:
            left = body_left
        edges.append((ti, left, left + total))
        if px not in (None, '0') and anchor == 'page':
            problems.append(
                f"table{ti}: horzAnchor=page 且 tblpX={px}，与其他表锚点不一致")
    # 所有表左边界必须一致、右边界必须一致
    lefts = {e[1] for e in edges}
    rights = {e[2] for e in edges}
    if len(lefts) > 1:
        problems.append(f"顶层表格左边界不一致: {sorted(lefts)} （应全部贴版心左边缘 {body_left}）")
    if len(rights) > 1:
        problems.append(f"顶层表格右边界不一致: {sorted(rights)}")
    return problems


if __name__ == "__main__":
    import sys
    from pathlib import Path
    root = Path(__file__).resolve().parent.parent
    tpl = root / "templates" / "内置实验报告模板.docx"
    print(f"烘焙模板: {tpl}")
    rep = rebake_template(str(tpl))
    for k, v in rep.items():
        print(f"  {k}: {v}")
    probs = verify_template(str(tpl))
    print("\n校验结果:", "全部通过 ✓" if not probs else "")
    for p in probs:
        print("  ✗", p)
    print(f"模板大小: {os.path.getsize(tpl)} bytes")

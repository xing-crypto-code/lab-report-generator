"""
排版黄金标准测试 · tests/test_golden_layout.py
==============================================
把用户已确认满意的排班效果固化为基线，防止后续改动无意破坏。

覆盖（全部来自用户确认的那一版）：
  1. 表格版式：两表等宽、贴版心左右边距、fixed 布局、无居中
  2. 字体：标题宋体四号(14pt)加粗 / 正文宋体小四(12pt)不加粗
  3. 行距：固定值 18pt（exact），字间距正文 -6 / 标题 -4
  4. 编号：统一 `1. 2. 3.`，连续递增，左边缘对齐
  5. 段落：pPr/rPr 子元素顺序合法，无尾部空段
  6. 占位：图片位置渲染为虚线框（保留，不丢）
  7. 结构：主表 7 行，栏名齐全，无模板外多余章节

运行：python tests/test_golden_layout.py
"""
import os
import sys
import re
import warnings

warnings.filterwarnings('ignore')
for _k in ('HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy',
           'ALL_PROXY', 'all_proxy', 'NO_PROXY', 'no_proxy'):
    os.environ.pop(_k, None)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, 'libs'))

from docx import Document                       # noqa: E402
from docx.oxml.ns import qn                     # noqa: E402

TEMPLATE = os.path.join(ROOT, 'templates', '内置实验报告模板.docx')

# ── 黄金标准数值（用户确认版）──
GOLD = {
    'page_w_cm': 21.0, 'page_h_cm': 29.7,
    'margin_left_cm': 1.75, 'margin_right_cm': 2.00,
    'body_width_dxa': 9780,
    'body_left_dxa': 993, 'body_right_dxa': 10773,
    'heading_pt': 14.0, 'body_pt': 12.0,
    'line_pt': 18.0, 'line_rule': 'exact',
    'char_spacing_body': '-6', 'char_spacing_heading': '-4',
    'cell_mar_lr': 108, 'cell_mar_tb': 0,
    'main_table_rows': 7,
}

_fail = []


def check(name, cond, detail=''):
    if cond:
        print(f'  ✓ {name}')
    else:
        print(f'  ✗ {name}   {detail}')
        _fail.append(name)


SAMPLE = {
    '实验目的': '1. 熟悉集成门电路的工作原理和主要参数\n'
                '2. 验证门电路的逻辑功能\n'
                '3. 掌握门电路外型引脚排列及使用注意事项',
    '实验仪器': '1. 数模电实验箱 1 台\n2. 数字万用表 1 只\n3. 74LS00、74LS20 各一片',
    '实验原理': ('1. 逻辑电平判定：输出大于 2.4 V 为逻辑 1，小于 0.4 V 为逻辑 0。\n'
                 '2. 74LS00 二输入与非门：Y＝(A·B)′\n'
                 '3. 74LS86 二输入异或门：Y＝A′B＋AB′'),
    '实验过程': ('一、实验步骤\n'
                 '1. 检查实验箱电源，将芯片插入插座。\n'
                 '2. 按信号流向逐级接线。\n'
                 '3. 测完先断电再拆线。\n'
                 '【此处插入实验接线图】\n'
                 '二、实验数据记录与分析\n'
                 '| A | B | Y |\n| --- | --- | --- |\n| 0 | 0 | 1 |\n| 1 | 1 | 0 |\n'
                 '三、误差分析\n万用表分辨率与接触电阻是主要误差来源。\n'
                 '四、思考题\nTTL 与 CMOS 有什么区别？'),
    '实验小结': '本次用一片 74LS00 完成了门电路功能验证。',
}


def main():
    from modules.generator import fill_template

    out = None
    try:
        out = fill_template(TEMPLATE, {'sections': SAMPLE, 'tables_fill': []},
                            '_golden_check')
        doc = Document(out)
        sec = doc.sections[0]

        print('\n== 1. 页面与版心 ==')
        check(f'页面 {sec.page_width.cm:.1f}x{sec.page_height.cm:.1f}cm',
              abs(sec.page_width.cm - GOLD['page_w_cm']) < 0.05 and
              abs(sec.page_height.cm - GOLD['page_h_cm']) < 0.05)
        check(f'左边距 {sec.left_margin.cm:.2f}cm',
              abs(sec.left_margin.cm - GOLD['margin_left_cm']) < 0.05)
        check(f'右边距 {sec.right_margin.cm:.2f}cm',
              abs(sec.right_margin.cm - GOLD['margin_right_cm']) < 0.05)
        body_w = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
        check(f'版心宽 {body_w} dxa（期望{GOLD["body_width_dxa"]}）',
              body_w == GOLD['body_width_dxa'])

        print('\n== 2. 表格版式（两表边框重合、贴版心）==')
        lefts, rights = set(), set()
        for ti, t in enumerate(doc.tables):
            pr = t._tbl.tblPr
            lay = pr.find(qn('w:tblLayout'))
            check(f'table{ti} 布局 fixed',
                  lay is not None and lay.get(qn('w:type')) == 'fixed')
            check(f'table{ti} 无居中 jc', pr.find(qn('w:jc')) is None)
            grid = t._tbl.find(qn('w:tblGrid'))
            total = sum(int(g.get(qn('w:w')) or 0)
                        for g in grid.findall(qn('w:gridCol')))
            tblW = pr.find(qn('w:tblW'))
            check(f'table{ti} tblW==ΣgridCols ({total})',
                  tblW is not None and int(tblW.get(qn('w:w')) or 0) == total)
            mar = pr.find(qn('w:tblCellMar'))
            if mar is not None:
                l = mar.find(qn('w:left'))
                r = mar.find(qn('w:right'))
                check(f'table{ti} cellMar 左右={GOLD["cell_mar_lr"]}',
                      int(l.get(qn('w:w'))) == GOLD['cell_mar_lr'] and
                      int(r.get(qn('w:w'))) == GOLD['cell_mar_lr'])
            tblp = pr.find(qn('w:tblpPr'))
            px = 0
            if tblp is not None and tblp.get(qn('w:tblpX')):
                px = int(tblp.get(qn('w:tblpX')))
            left = GOLD['body_left_dxa'] + px
            lefts.add(left)
            rights.add(left + total)
        check(f'两表左边界一致 {sorted(lefts)}', len(lefts) == 1)
        check(f'两表右边界一致 {sorted(rights)}', len(rights) == 1)
        check(f'左边界=版心左 {GOLD["body_left_dxa"]}', lefts == {GOLD['body_left_dxa']})
        check(f'右边界=版心右 {GOLD["body_right_dxa"]}', rights == {GOLD['body_right_dxa']})

        print('\n== 3. 字体与行距/字间距（用户确认的紧凑排版）==')
        main_t = doc.tables[1]
        heads, bodies = [], []
        for row in main_t.rows:
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                for p in c.paragraphs:
                    if not (p.text or '').strip():
                        continue
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        sz = r.font.size.pt if r.font.size else None
                        bold = bool(r.font.bold)
                        cs = None
                        rPr = r._element.rPr
                        if rPr is not None:
                            sp = rPr.find(qn('w:spacing'))
                            if sp is not None:
                                cs = sp.get(qn('w:val'))
                        pPr = p._p.find(qn('w:pPr'))
                        line = rule = None
                        if pPr is not None:
                            spc = pPr.find(qn('w:spacing'))
                            if spc is not None:
                                line = int(spc.get(qn('w:line'))) / 20 if spc.get(qn('w:line')) else None
                                rule = spc.get(qn('w:lineRule'))
                        (heads if bold else bodies).append((sz, cs, line, rule))
                        break
        check(f'标题 {GOLD["heading_pt"]}pt 加粗',
              heads and all(h[0] == GOLD['heading_pt'] for h in heads),
              f'{set(h[0] for h in heads)}')
        check(f'标题字间距 {GOLD["char_spacing_heading"]}',
              all(h[1] == GOLD['char_spacing_heading'] for h in heads),
              f'{set(h[1] for h in heads)}')
        check(f'正文 {GOLD["body_pt"]}pt 不加粗',
              bodies and all(b[0] == GOLD['body_pt'] for b in bodies),
              f'{set(b[0] for b in bodies)}')
        check(f'正文字间距 {GOLD["char_spacing_body"]}',
              all(b[1] == GOLD['char_spacing_body'] for b in bodies),
              f'{set(b[1] for b in bodies)}')
        check(f'行距统一 {GOLD["line_pt"]}pt / {GOLD["line_rule"]}',
              all(h[2] == GOLD['line_pt'] and h[3] == GOLD['line_rule'] for h in heads) and
              all(b[2] == GOLD['line_pt'] and b[3] == GOLD['line_rule'] for b in bodies),
              f'heads={set((h[2],h[3]) for h in heads)} bodies={set((b[2],b[3]) for b in bodies)}')

        print('\n== 4. 编号统一 `1. 2. 3.` ==')
        num_pat = re.compile(r'^(?:\d+\s*([\.、．)）])|[（(]\d+[）)]|[①②③④⑤⑥⑦⑧⑨⑩]|[a-zA-Z][\.、)])')
        seps = set()
        for row in main_t.rows:
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                for p in c.paragraphs:
                    m = num_pat.match((p.text or '').strip())
                    if m:
                        seps.add(m.group(1) or 'other')
        check(f'编号分隔符只有 "."（实际 {sorted(seps)}）', seps <= {'.'})

        print('\n== 5. 段落结构合法（无尾部空段 / 顺序正确）==')
        PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
               'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
               'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
               'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
               'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
               'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
               'textDirection', 'textAlignment', 'textboxTightWrap',
               'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
        RPR = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
               'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
               'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color',
               'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight',
               'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl',
               'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']

        def ok_el(el, order):
            tags = [c.tag.split('}')[-1] for c in el]
            idxs = [order.index(t) for t in tags if t in order]
            return idxs == sorted(idxs)

        bad_p = [p for p in doc.element.body.iter(qn('w:pPr')) if not ok_el(p, PPR)]
        bad_r = [r for r in doc.element.body.iter(qn('w:rPr')) if not ok_el(r, RPR)]
        check(f'pPr 顺序合法（错误 {len(bad_p)}）', not bad_p)
        check(f'rPr 顺序合法（错误 {len(bad_r)}）', not bad_r)

        blanks = 0
        for row in main_t.rows:
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                if not (c.text or '').strip():
                    continue
                ps = [x for x in c._tc if x.tag == qn('w:p')]
                if ps and not any((t.text or '').strip()
                                  for t in ps[-1].iter(qn('w:t'))):
                    blanks += 1
        check(f'有内容单元格无尾部空段（{blanks}）', blanks == 0)

        print('\n== 6. 插图虚线占位框（不能丢）==')
        from modules.render import render_docx
        html = render_docx(out)
        check(f'预览含虚线框 {html.count("r-ph")} 处', html.count('r-ph') >= 1)
        check('预览含真值表', "class='r-table'" in html)

        print('\n== 7. 结构：主表 7 行 + 无多余章节 ==')
        check(f'主表 {GOLD["main_table_rows"]} 行（实际{len(main_t.rows)}）',
              len(main_t.rows) == GOLD['main_table_rows'])
        labels = [(main_t.rows[i].cells[0].text or '').strip()
                  for i in range(len(main_t.rows))]
        for want in ('实验目的', '实验仪器', '实验原理', '实验过程', '实验小结'):
            check(f'含「{want}」栏', any(want in lb for lb in labels))
        extra = [p.text.strip() for p in doc.paragraphs
                 if (p.text or '').strip() and (p.text or '').strip() != '成都大学实验报告']
        check(f'无模板外多余顶层段落（{len(extra)}）', not extra, str(extra[:3]))

        print('\n== 8. 实验目的/仪器 照抄指导书（不增删）==')
        text_all = '\n'.join((c.text or '') for row in main_t.rows for c in row.cells)
        for kw in ('数模电实验箱', '数字万用表', '74LS00'):
            check(f'仪器保留原文「{kw}」', kw in text_all)
        for kw in ('熟悉集成门电路', '验证门电路'):
            check(f'目的保留原文「{kw}」', kw in text_all)
    finally:
        if out and os.path.exists(out):
            os.remove(out)

    print('\n' + '=' * 52)
    if _fail:
        print(f'✗ {len(_fail)} 项偏离黄金标准:')
        for f in _fail:
            print('   -', f)
        return 1
    print('✓ 全部符合黄金标准（用户确认的排版）')
    return 0


if __name__ == '__main__':
    sys.exit(main())

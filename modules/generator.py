"""
文档生成模块 · generator.py
===========================
将 AI 填充的内容写入 .docx 模板，保持原格式，只替换文字。
内置三线表格式、学术字体规范（宋体正文/Times New Roman 数据）。
"""
import re as _re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pathlib import Path

from config import OUTPUT_DIR

# ── 排版规范（按用户要求：参照学术论文，行距/字距紧凑）──
# 标题：宋体 四号（14pt）加粗；正文：宋体 小四（12pt）不加粗
HEADING_FONT = "宋体"
HEADING_SIZE_PT = 14      # 四号
BODY_FONT = "宋体"
BODY_SIZE_PT = 12         # 小四

# 正文行距：固定值 18 pt（学术论文常用；此前 1.5 倍≈21.6pt 偏稀）
BODY_LINE_SPACING_EXACT_PT = 18
# 字间距：轻微收紧（-0.3pt＝-6 twips），中文论文排版的常见做法
BODY_CHAR_SPACING_TWIPS = -6
HEADING_CHAR_SPACING_TWIPS = -4

# 段间距：正文段之间不加空（靠首行缩进区分段落，符合学术论文规范）
BODY_SPACE_AFTER = 0
BODY_SPACE_BEFORE = 0
# 标题：段前留 10pt 与上文分开，段后留 4pt 与自己的正文分开
HEADING_SPACE_BEFORE = 10
HEADING_SPACE_AFTER = 4

# 中文序号标题：一、 / （一） / 1、 / # markdown
_HEADING_RE = _re.compile(r"^(?:[一二三四五六七八九十]+[、\.．]|（[一二三四五六七八九十]+）|#{1,6}\s)")
# 占位标记：AI 输出【此处插入实验接线图】这类，系统要画虚线框
_PH_HINT_WORDS = ("此处插入", "此处插", "在此插入", "插入接线图", "插入电路图",
                  "插入波形", "插入照片", "插入数据", "插入记录表", "插入图")
_PH_EXACT = ("【插图】", "[插图]", "【图片】", "[图片]", "【此处附图】",
             "【实验接线图】", "【电路图】", "【波形图】")


def _is_heading_line(text: str) -> bool:
    """判断一行是否是（小）标题"""
    s = (text or "").strip()
    if not s or len(s) > 40:
        return False
    if _HEADING_RE.match(s):
        return True
    return False


# 二级/三级列举项：1. / 1） / （1） / ① —— 这类是正文里的条目，不算标题
_LIST_RE = _re.compile(r"^(?:\d+[\.、．)]|[（(]\d+[）)]|[①②③④⑤⑥⑦⑧⑨⑩]|[a-zA-Z][\.、)])")

# 从各种编号写法里抽出「序号数字」与「分隔符」
_NUM_PREFIX_RE = _re.compile(
    r"^(?:(?P<arabic>\d+)\s*(?P<sep>[\.、．)）])\s*|[（(](?P<par>\d+)[）)]\s*"
    r"|(?P<circ>[①②③④⑤⑥⑦⑧⑨⑩])\s*|(?P<alpha>[a-zA-Z])[\.、)]\s*)"
)
_CIRCLE_MAP = {c: i + 1 for i, c in enumerate("①②③④⑤⑥⑦⑧⑨⑩")}


def _normalize_numbering(items):
    """统一编号格式为 `1.` `2.` `3.`，让左边缘严格对齐。

    用户要求：字间距、行间距、序号都要统一对齐。AI 常混用 `1.` `2、` `（3）`
    `④` `a.` 五种写法，分隔符宽度不同 → 正文起始位置参差不齐。
    这里把同一「编号串」内所有条目重排为 `N. ` 统一格式。
    """
    out = []
    run = []          # 当前连续编号串
    run_is_numbered = None

    def flush():
        nonlocal run_is_numbered
        if not run:
            return
        if run_is_numbered:
            for n, it in enumerate(run, 1):
                body = _strip_num_prefix(it["text"])
                out.append({"kind": "list", "text": f"{n}. {body}"})
        else:
            out.extend(run)
        run.clear()
        run_is_numbered = None

    for it in items:
        if it["kind"] == "list":
            is_num = bool(_NUM_PREFIX_RE.match((it.get("text") or "").strip()))
            if run and run_is_numbered != is_num:
                flush()
            run.append(it)
            run_is_numbered = is_num
        else:
            flush()
            out.append(it)
    flush()
    return out


def _renumber_questions(items):
    """思考题「题目行」编号连续：实证样本里每条题目都写成 `1.`，这是错的。

    识别模式：`N. xxx？`（以问号结尾的编号行）后面紧跟 `答：...`。
    把这类「编号+问号」行统一按出现顺序重排为 1. 2. 3. …，
    而 `答：` 行保持原样（不加编号）。
    """
    # 找出所有「以问号结尾的带编号行」
    idxs = []
    for i, it in enumerate(items):
        if it["kind"] != "list":
            continue
        t = (it.get("text") or "").strip()
        body = _strip_num_prefix(t)
        if body.endswith(("？", "?")):
            idxs.append(i)
    if len(idxs) < 2:
        return items
    n = 0
    out = list(items)
    for i in idxs:
        n += 1
        body = _strip_num_prefix(out[i]["text"])
        out[i] = {"kind": "list", "text": f"{n}. {body}"}
    return out


def _drop_duplicate_headings(items):
    """去掉连续/近邻重复的标题（实证样本里出现过「逻辑电路图 2逻辑电路图 2」）。"""
    out = []
    seen_recent = []
    for it in items:
        if it["kind"] == "heading":
            txt = (it.get("text") or "").strip()
            if txt in seen_recent:
                continue          # 重复标题，丢弃
            seen_recent.append(txt)
            if len(seen_recent) > 8:
                seen_recent.pop(0)
        # 标题完全相同的「同一行重复两遍」也要拆开
        elif it["kind"] in ("para", "list"):
            pass
        out.append(it)
    return out


def _dedupe_repeated_phrase(items):
    """清理「逻辑电路图 2逻辑电路图 2」这类同一短句在一行内重复两遍的残留。"""
    import re as _r
    out = []
    for it in items:
        if "text" not in it:
            out.append(it)
            continue
        t = (it.get("text") or "").strip()
        new = t
        # 情形1：前半 == 后半（允许中间空格/数字差异极小），如
        #   "逻辑电路图 2逻辑电路图 2"
        n = len(t)
        if n >= 8:
            for cut in range(4, n // 2 + 1):
                a, b = t[:cut].strip(), t[cut:].strip()
                # 去掉两边标点/空白后比较
                an = _r.sub(r"[\s\W_]+", "", a)
                bn = _r.sub(r"[\s\W_]+", "", b)
                if an and an == bn:
                    new = a
                    break
        if new != t:
            it = dict(it)
            it["text"] = new
        out.append(it)
    return out


def _strip_num_prefix(text):
    """去掉行首的任意编号前缀（1. / 2、/ （3） / ④ / a.），返回正文"""
    s = (text or "").strip()
    m = _NUM_PREFIX_RE.match(s)
    if m:
        return s[m.end():].strip()
    return s


def _is_list_line(text: str) -> bool:
    s = (text or "").strip()
    return bool(s) and bool(_LIST_RE.match(s))


def _is_placeholder_only_line(text: str) -> bool:
    """整行只是 [待补充] 这类占位，渲染时跳过（不写进正文）"""
    s = (text or "").strip()
    if not s:
        return True
    return _section_is_placeholder_only(s)


def _strip_heading_marks(text: str) -> str:
    """去掉 markdown 井号，保留“一、实验步骤”这类中文序号"""
    return (text or "").strip().lstrip("#").strip()


def _match_placeholder(text: str):
    """若整行是图片/表格占位标记，返回 (label, width_cm, height_cm)；否则 None。
    尺寸按占位内容智能给一个“合理大小”。"""
    s = (text or "").strip()
    if not s or len(s) > 60:
        return None
    if not (any(w in s for w in _PH_HINT_WORDS) or s in _PH_EXACT):
        return None
    # 按用途推出合理框尺寸（cm）
    if any(k in s for k in ("接线图", "连线图", "电路图", "逻辑图", "实验装置")):
        w, h = 12.0, 6.5
    elif any(k in s for k in ("波形", "曲线", "频谱", "示波", "扫频", "李萨如")):
        w, h = 12.0, 6.0
    elif any(k in s for k in ("表格", "数据", "真值表", "记录表")):
        w, h = 12.0, 4.5
    elif any(k in s for k in ("照片", "实物", "仪器")):
        w, h = 10.0, 7.5
    else:
        w, h = 12.0, 6.0
    return s, w, h


def _merge_list_continuations(items):
    """把列表项的"续行"并回列表项。

    模型常写成三行：
        1. 闭环电压放大倍数
        Avf＝Av/(1＋Av·Fv)
        式中 Av 为开环电压放大倍数……
    后两行没有编号，过去会被 autonumber 当成新的列表项，于是编号错成 1./2./3.。
    这里在编号之前先合并：列表项未以句末标点收尾，就把后续无编号段落并进来。
    """
    out = []
    for it in items:
        prev = out[-1] if out else None
        if (prev is not None and prev["kind"] == "list" and it["kind"] == "para"
                and not prev["text"].rstrip().endswith(("。", "；", "！", "？", ";", "!", "?"))):
            body = prev["text"].rstrip()
            tail = it["text"].strip()
            # 短标题后接内容 → 用冒号衔接，读起来像"标题：内容"；
            # 已经有冒号的行不要再补一个，否则会出现"：…："这种双冒号。
            joiner = ("：" if len(body) <= 30 and "：" not in body and ":" not in body
                      else " ")
            prev["text"] = body + joiner + tail
            continue
        out.append(dict(it))
    return out


def _autonumber_paragraphs(items):
    """给散落的无编号正文自动加 `1.` `2.` 序号。

    用户要求：适当的地方要有 1. 2. 这种编号。规则：
    · 只在「一串连续的普通段落」里加号（长度 >=2 才算一串）；
    · 已有编号（1. / （1） / ① / a.）的段落不重复加，且作为一串的结束标志；
    · 单个孤立段落不加（避免给一句话硬编号）；
    · 说明性/过渡性段落（"数据处理："、"说明："、"注："、"其中："等）不加号，
      它们本身是下一段的前言；
    · 标题、表格、占位框前后会打断一串，各自重新从 1 开始。
    """
    out = []
    run_buf = []

    def _is_explanatory(t):
        s = (t or "").strip()
        for lead in ("数据处理", "说明", "注：", "注:", "其中", "注意", "提示",
                     "小结", "综上所述", "因此", "所以", "答：", "答:", "解：",
                     "解:", "分析：", "分析:"):
            if s.startswith(lead):
                return True
        # 以冒号结尾的是「前言句」（如"74LS00 与非门（Y=(A·B)′）："），
        # 它引出的是下面的表格/内容，不该编号。
        return s.endswith("：") or s.endswith(":")

    def flush():
        if not run_buf:
            return
        # 只有「非说明性」段落才参与编号；且需 >=2 条才成串
        numbered = [it for it in run_buf if not _is_explanatory(it["text"])]
        if len(numbered) >= 2:
            i = 0
            for it in run_buf:
                if _is_explanatory(it["text"]):
                    out.append(it)
                else:
                    i += 1
                    out.append({"kind": "list", "text": f"{i}. {it['text']}"})
        else:
            # 串太短（0或1条可编号项）→ 全部保持原样
            out.extend(run_buf)
        run_buf.clear()

    for it in items:
        if it["kind"] == "para":
            run_buf.append(it)
        else:
            flush()
            out.append(it)
    flush()
    return out


def _iter_render_items(content):
    """把一段内容解析成有序的渲染项：
    {"kind": "table",       "rows": [[...]]}
    {"kind": "placeholder", "label":..., "width":.., "height":..}
    {"kind": "heading",     "text":...}
    {"kind": "para",        "text":...}
    """
    parts = (content or "").split("\n")
    blocks = _split_markdown_tables(parts)
    items = []
    for b in blocks:
        if b.get("is_table"):
            items.append({"kind": "table", "rows": b["rows"]})
            continue
        for line in b["lines"]:
            s = line.strip()
            if not s:
                continue
            # 纯占位行（[待补充]）不渲染成正文
            if _is_placeholder_only_line(s):
                continue
            ph = _match_placeholder(s)
            if ph:
                label, w, h = ph
                items.append({"kind": "placeholder", "label": label,
                              "width": w, "height": h})
            elif _is_heading_line(s) and not _is_list_line(s):
                items.append({"kind": "heading", "text": _strip_heading_marks(s)})
            elif _is_list_line(s):
                items.append({"kind": "list", "text": s})
            else:
                items.append({"kind": "para", "text": s})
    items = _merge_list_continuations(items)
    items = _autonumber_paragraphs(items)
    items = _normalize_numbering(items)
    items = _renumber_questions(items)
    items = _drop_duplicate_headings(items)
    items = _dedupe_repeated_phrase(items)
    return items


def _template_section_names(doc, template_summary=""):
    """收集模板里真实存在的章节名（来自表格 label 单元格 + 段落标题）。
    用于判断某个 AI 返回的章节到底该不该追加——模板没有的就不要造出来。"""
    names = set()

    def _add_label(raw):
        """把单元格文本归纳成"章节名"：取行首那段短的标签。

        r4 这类合并格常是「实验过程（可另附页）\\n\\n一、实验步骤\\n…」，
        整段不适合当章节名，要按行拆开、只取像标题的前几行。
        """
        txt = (raw or "").strip()
        if not txt:
            return
        # 先整段短文本直接可用
        if len(txt) <= 24 and "\n" not in txt:
            names.add(txt)
            return
        # 多行 / 长文本：逐行取短标签
        for line in txt.splitlines():
            s = line.strip()
            if not s:
                continue
            if len(s) <= 24:
                names.add(s)
            # 只取前几行，避免把正文句子都当成章节
            if len(names) > 200:
                break

    # 1) 表格单元格里的 label
    for table in doc.tables:
        for row in table.rows:
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                _add_label(c.text)

    # 2) 正文里像标题的段落
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t and len(t) <= 24 and _is_heading_line(t):
            names.add(_strip_heading_marks(t))

    # 3) 模板结构摘要里出现过的章节
    if template_summary:
        for line in str(template_summary).splitlines():
            s = line.strip().strip("-·*0123456789. ")
            if s and len(s) <= 24:
                names.add(s)
    return names


def _normalize_table_layout(doc):
    """对【任意模板】统一表格横向定位与宽度，保证上下表左右边框重合。

    内置模板由 template_fix 预先烘焙好，但用户上传的老师模板不会，
    所以填充时必须再跑一趟同样的规则（幂等，已规范则几乎无改动）：
      · 所有顶层表格锚点=margin、无 tblpX/tblpXSpec/jc/tblInd 偏移
      · 宽度统一为版心宽度，按比例缩放保留原列宽设计
    """
    from docx.oxml.ns import qn as _qn
    try:
        sec = doc.sections[0]
        body_w = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
    except Exception:
        return
    if not body_w:
        return

    tables = doc.tables
    if not tables:
        return

    for t in tables:
        tbl = t._tbl
        tblPr = tbl.tblPr
        if tblPr is None:
            continue

        # 宽度统一 = 版心宽（按比例缩放）
        g = tbl.find(_qn('w:tblGrid'))
        cols = g.findall(_qn('w:gridCol')) if g is not None else []
        old = [int(c.get(_qn('w:w')) or 0) for c in cols]
        cur = sum(old)
        if cols and cur and cur != body_w:
            scaled = [max(1, round(w * body_w / cur)) for w in old]
            scaled[-1] += body_w - sum(scaled)
            for c, w in zip(cols, scaled):
                c.set(_qn('w:w'), str(w))
            tblW = tblPr.find(_qn('w:tblW'))
            if tblW is None:
                tblW = tblPr.makeelement(_qn('w:tblW'), {})
                tblPr.append(tblW)
            tblW.set(_qn('w:type'), 'dxa')
            tblW.set(_qn('w:w'), str(body_w))
            # 同步每行 tcW（按 gridSpan 累加）
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
                        w_el.set(_qn('w:w'),
                                 str(sum(scaled[idx:min(idx + span, len(scaled))])))
                    idx += span

        # 布局必须 fixed：autofit 会让 Word 忽略我们设的列宽，边框随内容漂移
        layout = tblPr.find(_qn('w:tblLayout'))
        if layout is None:
            layout = tblPr.makeelement(_qn('w:tblLayout'), {})
            tblPr.append(layout)
        layout.set(_qn('w:type'), 'fixed')

        # ── 定位统一（强制重建 tblpPr，彻底清掉模板残留的 tblpX 偏移）──
        # 老师模板里表0 常用 tblpXSpec=center、主表用 horzAnchor=page+tblpX=1157，
        # 这些都会让表格整体偏移（实测会出现左偏 -30 / 右偏 -44 dxa 的错位）。
        # 只"改属性"不够稳，直接重建 tblpPr 最彻底。
        tblpPr = tblPr.find(_qn('w:tblpPr'))
        if tblpPr is None:
            tblpPr = tblPr.makeelement(_qn('w:tblpPr'), {})
            # tblpPr 在 tblPr 里应尽量靠前，放在 tblStyle 之后
            kids = list(tblPr)
            idx = 1 if kids and kids[0].tag == _qn('w:tblStyle') else 0
            tblPr.insert(idx, tblpPr)
        # 清空旧属性后仅保留"垂直浮动所需"的最小集合
        for a in list(tblpPr.attrib):
            del tblpPr.attrib[a]
        tblpPr.set(_qn('w:vertAnchor'), 'text')
        tblpPr.set(_qn('w:horzAnchor'), 'margin')
        tblpPr.set(_qn('w:tblpX'), '0')
        # tblpY 保留模板原值（决定上下位置），没有就用 0
        for tag in ('w:jc', 'w:tblInd'):
            for el in tblPr.findall(_qn(tag)):
                tblPr.remove(el)

        # 单元格边距规范（左右108/上下0），保证嵌套表两边等距
        mar = tblPr.find(_qn('w:tblCellMar'))
        if mar is None:
            mar = tblPr.makeelement(_qn('w:tblCellMar'), {})
            tblPr.append(mar)
        for edge, val in (('top', 0), ('left', 108), ('bottom', 0), ('right', 108)):
            el = mar.find(_qn('w:' + edge))
            if el is None:
                el = mar.makeelement(_qn('w:' + edge), {})
                mar.append(el)
            el.set(_qn('w:type'), 'dxa')
            el.set(_qn('w:w'), str(val))


def fill_template(template_path: str, filled_data: dict, output_name: str = "实验报告") -> str:
    """
    将 filled_data 填入模板并保存。

    核心策略（保证内容落在模板线框内）：
    1. 优先扫描模板中的【表格】：找到章节名所在的单元格（如"实验目的"），
       把 AI 内容填入同行的空单元格/数据区。
    2. 再处理段落占位符（___ 或含章节名标题后的段落）。
    3. 仍没匹配到的章节，追加到文末（兜底）。
    """
    doc = Document(template_path)
    sections = filled_data.get("sections", {})
    matched_keys = set()

    # ── 先统一表格版式（对用户上传的老师模板同样生效）──
    try:
        _normalize_table_layout(doc)
    except Exception as _e:
        print("[generator] 表格版式统一失败:", _e)

    # 有效的填充内容（排除空、[待补充]）
    def _usable(k):
        v = sections.get(k) or ""
        return v.strip() and "[待补充]" not in v

    # ── 第1优先：填充模板表格（老师的模板一般是表格线框）──
    matched_keys.update(_fill_via_tables(doc, sections))

    # ── 第2优先：段落占位符 / 章节标题后追加 ──
    for para in doc.paragraphs:
        original_text = para.text.strip()
        if not original_text:
            continue
        for section_name, content in sections.items():
            if not _usable(section_name) or section_name in matched_keys:
                continue
            # 段落本身是章节标题 → 在其后插入正文
            if original_text == section_name or original_text.endswith(section_name):
                _append_content(para, doc, content)
                matched_keys.add(section_name)
                break
            # 段落内含下划线占位符 → 直接替换
            if "_" in original_text and section_name in original_text:
                _replace_placeholder(para, content)
                matched_keys.add(section_name)
                break

    # ── 兜底：仍未匹配的章节，追加到文末 ──
    # 关键：只允许模板中【真实存在】的章节追加，不能凭空造出"实验设计"这种
    # 模板里没有的标题（用户交作业时会被算作多余章节）。
    _allowed = _template_section_names(
        doc, filled_data.get("template_summary", ""))
    for section_name, content in sections.items():
        if not _usable(section_name) or section_name in matched_keys:
            continue
        if section_name not in _allowed:
            print(f"[generator] 跳过模板中不存在的章节：{section_name}")
            continue
        _append_heading_with_content(doc, section_name, content)

    # ── 填充独立数据表（tables_fill 按表头匹配，不硬按 index）──
    _fill_data_tables(doc, filled_data.get("tables_fill", []))

    # ── 收尾清理：pPr 排序 + 去多余空段（否则编辑时按 Delete 会连带删上一行）──
    try:
        _finalize_document(doc)
    except Exception as _e:
        print("[generator] 收尾清理失败:", _e)

    # ── 保存 ──
    output_path = OUTPUT_DIR / f"{output_name}.docx"
    doc.save(str(output_path))
    return str(output_path)


def _finalize_document(doc):
    """收尾清理，修掉"编辑时 Delete 连带删上一行"和"上下位置对不齐"的根因：

    1) 所有段落 w:pPr 子元素按 OOXML 规范顺序重排（顺序错 Word 静默忽略属性）；
    2) 删除每种内容块末尾多余的连续空段落。Word 里单元格/表格后若残留多个空段，
       光标落到空段首按 Delete 会与上一段合并，表现为"把上一排删掉了"。
    """
    from docx.oxml.ns import qn as _qn

    # 1) pPr / rPr 排序（全文 + 所有单元格）
    containers = [doc.element.body]
    for t in doc.tables:
        containers.append(t._tbl)
    for c in containers:
        for ppr in c.iter(_qn('w:pPr')):
            _sort_ppr(ppr)
        for rpr in c.iter(_qn('w:rPr')):
            _sort_rpr(rpr)

    # 1.5) 清掉模板残留的异常首行缩进（firstLineChars 很大 → 文字右移、上下不齐）
    for p in doc.element.body.iter(_qn('w:p')):
        pPr = p.find(_qn('w:pPr'))
        if pPr is None:
            continue
        ind = pPr.find(_qn('w:ind'))
        if ind is None:
            continue
        fc = ind.get(_qn('w:firstLineChars'))
        if fc is not None:
            try:
                if abs(int(fc)) > 10:      # >10 字符的缩进在这类模板里必然是脏数据
                    for a in ('firstLineChars', 'firstLine'):
                        ind.attrib.pop(_qn('w:' + a), None)
            except Exception:
                pass
        # 空的 ind 元素留着没意义，删掉更干净
        if not ind.attrib:
            pPr.remove(ind)
        _sort_ppr(pPr)

    # 1.6) 统一全部段落的行距与字间距
    # 模板原有的 label 格（课程名称/实验目的/实验仪器…）没有动过，缺行距和
    # 字间距，导致 label 列与内容列**上下对不齐**。这里把它们也规范到同一标准。
    for p in doc.element.body.iter(_qn('w:p')):
        has_text = any((t.text or '').strip() for t in p.iter(_qn('w:t')))
        if not has_text:
            continue
        pPr = p.find(_qn('w:pPr'))
        if pPr is None:
            pPr = p.makeelement(_qn('w:pPr'), {})
            p.insert(0, pPr)
        # 判定是否标题：段内首个 run 加粗
        is_heading = False
        for r in p.iter(_qn('w:r')):
            rPr = r.find(_qn('w:rPr'))
            if rPr is not None:
                b = rPr.find(_qn('w:b'))
                if b is not None and b.get(_qn('w:val')) not in ('0', 'false'):
                    is_heading = True
            break
        # 行距：缺失或非 exact 一律补齐为统一固定值
        sp = pPr.find(_qn('w:spacing'))
        need_sp = True
        if sp is not None and sp.get(_qn('w:line')) is not None:
            try:
                need_sp = (abs(int(sp.get(_qn('w:line'))) - int(BODY_LINE_SPACING_EXACT_PT * 20)) > 0
                           or sp.get(_qn('w:lineRule')) != 'exact')
            except Exception:
                need_sp = True
        if need_sp:
            if sp is None:
                sp = pPr.makeelement(_qn('w:spacing'), {})
                pPr.append(sp)
            sp.set(_qn('w:line'), str(int(BODY_LINE_SPACING_EXACT_PT * 20)))
            sp.set(_qn('w:lineRule'), 'exact')
            if sp.get(_qn('w:before')) is None:
                sp.set(_qn('w:before'), '0')
            if sp.get(_qn('w:after')) is None:
                sp.set(_qn('w:after'), '0')
        # 字间距：缺失则补上（标题 -4 / 正文 -6）
        want_cs = str(HEADING_CHAR_SPACING_TWIPS if is_heading
                      else BODY_CHAR_SPACING_TWIPS)
        for r in p.iter(_qn('w:r')):
            if not any((t.text or '').strip() for t in r.iter(_qn('w:t'))):
                continue
            rPr = r.find(_qn('w:rPr'))
            if rPr is None:
                rPr = r.makeelement(_qn('w:rPr'), {})
                r.insert(0, rPr)
            cs = rPr.find(_qn('w:spacing'))
            if cs is None:
                cs = rPr.makeelement(_qn('w:spacing'), {})
                rPr.append(cs)
            if cs.get(_qn('w:val')) is None:
                cs.set(_qn('w:val'), want_cs)
            _sort_rpr(rPr)
        _sort_ppr(pPr)

    # 2) 清理连续空段落（保留最多 1 个）
    def _is_blank(p):
        texts = p.findall('.//' + _qn('w:t'))
        if any((t.text or "").strip() for t in texts):
            return False
        # 含图片/表格的空段不算空白
        if p.find('.//' + _qn('w:drawing')) is not None:
            return False
        if p.find('.//' + _qn('w:tbl')) is not None:
            return False
        return True

    # 正文层
    body = doc.element.body
    prev_blank = False
    for el in list(body):
        if el.tag != _qn('w:p'):
            prev_blank = False
            continue
        if _is_blank(el):
            if prev_blank:
                body.remove(el)
            else:
                prev_blank = True
        else:
            prev_blank = False

    # 每个单元格内
    for t in doc.tables:
        for row in t.rows:
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                tc = c._tc
                kids = list(tc)
                prev_blank = False
                for el in kids:
                    if el.tag != _qn('w:p'):
                        prev_blank = False
                        continue
                    if _is_blank(el):
                        if prev_blank:
                            tc.remove(el)
                        else:
                            prev_blank = True
                    else:
                        prev_blank = False

                # 末尾若还有空段，直接删掉（Word 里该空段按 Delete 会吃掉上一行）
                # 但单元格至少要留一个段落，否则文档结构非法。
                while True:
                    ps = [x for x in tc if x.tag == _qn('w:p')]
                    if len(ps) <= 1:
                        break
                    if not _is_blank(ps[-1]):
                        break
                    tc.remove(ps[-1])

                # 末尾空段带上异常缩进时清零，避免行高被撑开导致上下不齐
                ps = [x for x in tc if x.tag == _qn('w:p')]
                if len(ps) == 1 and _is_blank(ps[0]):
                    pPr = ps[0].find(_qn('w:pPr'))
                    if pPr is not None:
                        ind = pPr.find(_qn('w:ind'))
                        if ind is not None:
                            for a in ('firstLine', 'firstLineChars',
                                      'left', 'leftChars', 'hanging', 'hangingChars'):
                                ind.attrib.pop(_qn('w:' + a), None)
                        _sort_ppr(pPr)


# 预置的章节别名，用于在模板单元格里模糊匹配
_SECTION_ALIASES = {

    "实验名称": ["实验名称", "题目", "实验题目"],
    "课程名称": ["课程名称", "课程"],
    "实验过程": ["实验过程"],
    "实验小结": ["实验小结", "小结"],
    "实验目的": ["实验目的", "目的"],
    "实验原理": ["实验原理", "原理", "原理（含公式推导）", "实验原理与电路图"],
    "实验器材": ["实验器材", "实验仪器", "仪器", "实验仪器与器件", "器材", "设备"],
    "实验试剂与仪器": ["实验试剂", "试剂与仪器", "药品与仪器"],
    "实验步骤": ["实验步骤", "步骤", "实验步骤与接线图", "操作步骤"],
    "实验数据记录": ["实验数据记录", "实验数据", "数据记录", "记录"],
    "实验数据": ["实验数据", "数据记录", "数据（波形/表格）"],
    "数据处理与分析": ["数据处理分析", "数据处理与分析", "数据处理"],
    "数据处理": ["数据处理", "数据处理与分析"],
    "误差分析": ["误差分析"],
    "实验结果": ["实验结果", "结果"],
    "讨论与结论": ["讨论与结论", "结果与讨论"],
    "实验结论": ["实验结论", "结论"],
    "结果与讨论": ["结果与讨论", "讨论"],
    "思考题": ["思考题", "思考"],
    "注意事项": ["注意事项", "注意"],
    "实验日期": ["实验日期", "日期"],
    "姓名": ["姓名"],
    "学号": ["学号", "班级"],
    "班级": ["班级", "专业"],
    "同组人员": ["同组人员", "组员"],
    "指导教师": ["指导教师", "指导老师"],
    "评阅": ["教师评阅", "评阅意见", "评语"],
}


def _section_matches(cell_text: str, section_name: str) -> bool:
    """判断单元格文字是否对应某个章节（含别名匹配）"""
    text = (cell_text or "").strip()
    if not text:
        return False
    aliases = _SECTION_ALIASES.get(section_name, [section_name])
    return any(a and a in text for a in aliases)


def _fill_via_tables(doc, sections: dict) -> set:
    """
    核心：扫描模板所有表格，把 AI 内容填入对应章节的单元格。
    返回已填充的章节名集合。

    对每一行,逐个 label 单元格匹配章节名：
      - 若该小节是"label|空白"结构（右邻为空），填右邻单元格
      - 若一行含多个 label（如"课程名称|空|实验名称|空"），各自向右找值位
    """
    matched = set()
    for table in doc.tables:
        nrows = len(table.rows)
        if nrows == 0:
            continue
        ncols = len(table.columns)
        for ri in range(nrows):
            row = table.rows[ri]
            cells = row.cells
            for ci in range(len(cells)):
                t = (cells[ci].text or "").strip()
                # 只把"短标签文本"当作章节名候选，避免把已填内容格/长文体误判为标签
                if not t or len(t) > 22 or "\n" in t:
                    continue
                for section_name, content in sections.items():
                    if not content or not content.strip():
                        continue
                    # 仅在内容几乎全是占位符时才跳过（避免整段内容含少量[待补充]导致漏填整节）
                    _c = content.strip()
                    if _section_is_placeholder_only(_c):
                        continue
                    if section_name in matched:
                        continue
                    if _section_matches(t, section_name):
                        # 找到章节名单元格，向右找值位填
                        if _fill_row_target(table, ri, ci, content, ncols):
                            matched.add(section_name)
                        break
    return matched


def _fill_row_target(table, ri, ci, content, ncols):
    """把 content 填入章节名单元格 ci 的【右边】值位。"""
    if ri >= len(table.rows):
        return False
    row = table.rows[ri]
    cells = row.cells

    # 情形0：整行是一个单元格（如"实验过程（可另附页）"一整行合并）
    # 直接在该 label 单元格追加/替换内容，绝不波及别的行
    if len(cells) == 1:
        cell = cells[0]
        # 标签文本 = 去掉占位
        _set_or_append_cell(cell, content)
        return True

    # 策略A：同行向右找空单元格（label 在左、值区在右的常见结构）
    for j in range(ci + 1, len(cells)):
        cell = cells[j]
        if _is_empty_cell(cell):
            _set_cell_text(cell, content)
            return True

    # 策略B：同 label 列向下找空行
    for k in range(ri + 1, len(table.rows)):
        nd = table.rows[k]
        if ci < len(nd.cells) and _is_empty_cell(nd.cells[ci]):
            _set_cell_text(nd.cells[ci], content)
            return True

    # 策略C：把内容追加进章节名单元格
    try:
        _append_cell_content(cells[ci], content)
        return True
    except Exception:
        return False


def _set_or_append_cell(cell, content):
    """整行即章节时：保留 label，清掉模板多余空行后写入内容"""
    from docx.oxml.ns import qn as _qn
    tc = cell._tc
    # 移除除第一个段落外的所有空段落（模板占位的空行）
    ps = tc.findall(_qn('w:p'))
    for p in ps[1:]:
        # 空段落（无 run 或无文字）
        texts = p.findall('.//' + _qn('w:t'))
        if not any((t.text or "").strip() for t in texts):
            tc.remove(p)
    # 找到 label 段，在其后追加内容各段
    ps2 = tc.findall(_qn('w:p'))
    anchor = ps2[0] if ps2 else None
    parts = [x for x in content.split("\n") if x.strip()]
    for part in parts:
        p = tc.makeelement(_qn('w:p'), {})
        # create a run
        r = tc.makeelement(_qn('w:r'), {})
        t = tc.makeelement(_qn('w:t'), {})
        t.text = part
        r.append(t)
        p.append(r)
        if anchor is not None:
            anchor.addnext(p)
        else:
            tc.append(p)


def _append_cell_content(cell, content):
    """在章节名单元格写内容（保留 label、清空占位段）"""
    _set_cell_text(cell, content)


def _is_empty_cell(cell):
    """判断单元格是否可写（空内容或只含空行/下划线占位）"""
    t = (cell.text or "").strip()
    if t == "":
        return True
    return "___" in t or "＿" in t or "____" in t


def _section_is_placeholder_only(text):
    """整个章节内容是否等于占位符（此时跳过填充）。
    只有内容去掉空格/标点后拆分出的每个 token 都是占位词（如“[待补充]”
    “待补充”“待定”“…”“None”“略”）才判为纯占位；一段正常内容里夹着
    一处[待补充]不算（会照常填整节）。"""
    import re as _re
    t = (text or "").strip()
    if not t:
        return True
    # 按中英文逗号/换行/分号切分为 token 后再判定，避免“待补充\n待定”拼成新词
    tokens = [x for x in _re.split(r"[\n，,、；;、…。:：\s\[\]()（）]", t)
              if x.strip()]
    if not tokens:
        return True
    placeholder_words = {"待补充", "待定", "补充", "暂缺", "略", "none", "null", "待"}
    return all(tok.strip().lower() in placeholder_words for tok in tokens)


def _set_cell_text(cell, content):
    """把 content 写入单元格。支持 Markdown 表格（转成真正的 Word 嵌套表格）
    和【此处插入接线图】这类图片占位（转成虚线框让用户后续贴图）。
    若单元格含章节名 label 则保留 label、清空其后空占位段，再写内容/表格/占位框。"""
    items = _iter_render_items(content)
    tc = cell._tc
    paragraphs = cell.paragraphs

    first_text = (cell.text or "").strip()
    is_label_cell = (len(first_text) < 20 and "\n" not in first_text
                     and len(paragraphs) > 0 and first_text != "")

    if is_label_cell and paragraphs:
        # 保留 label 段，删除其后空占位段
        _clear_empty_paras(tc)
        # label 段也统一到正文行距/字距，避免整栏一处松一处紧
        # 同时清掉模板可能残留的异常缩进（如 firstLineChars=49 会右移半行）
        _apply_para_format(paragraphs[0], first_line_indent_chars=0)
        _p0 = paragraphs[0]._p
        _pPr0 = _p0.find(qn('w:pPr'))
        if _pPr0 is not None:
            _ind0 = _pPr0.find(qn('w:ind'))
            if _ind0 is not None:
                for _a in ('firstLineChars', 'firstLine',
                           'leftChars', 'left', 'hangingChars', 'hanging'):
                    _ind0.attrib.pop(qn('w:' + _a), None)
            _sort_ppr(_pPr0)
        for r in paragraphs[0].runs:
            _apply_run_style(r, BODY_SIZE_PT, bold=False)
        anchor = paragraphs[0]._element
        for it in items:
            kind = it["kind"]
            if kind == "placeholder":
                _insert_dashed_placeholder_after(tc, anchor, cell, it)
                anchor = _last_child_before(tc, anchor) or anchor
            elif kind == "table":
                t = _insert_nested_table(cell, it["rows"])
                if t is not None:
                    anchor = t._tbl
            elif kind == "heading":
                anchor = _insert_styled_para(tc, anchor, it["text"], heading=True)
            elif kind == "list":
                anchor = _insert_styled_para(tc, anchor, it["text"], list_item=True)
            else:
                anchor = _insert_styled_para(tc, anchor, it["text"], heading=False)
    else:
        # 普通值格：清空首段后按 items 顺序渲染
        if cell.paragraphs:
            first = cell.paragraphs[0]
            for r in list(first.runs):
                try:
                    r._element.getparent().remove(r._element)
                except Exception:
                    pass
            anchor = first._element
            used_first = False
        else:
            p = cell.add_paragraph()
            anchor = p._element
            used_first = False

        for it in items:
            kind = it["kind"]
            if kind == "placeholder":
                _insert_dashed_placeholder_after(tc, anchor, cell, it)
                anchor = _last_child_before(tc, anchor) or anchor
            elif kind == "table":
                t = _insert_nested_table(cell, it["rows"])
                if t is not None:
                    anchor = t._tbl
            elif kind == "heading":
                if not used_first:
                    _force_left(first)
                    _apply_para_format(first, first_line_indent_chars=0)
                    run = first.add_run(it["text"])
                    _apply_run_style(run, HEADING_SIZE_PT, bold=True)
                    used_first = True
                else:
                    anchor = _insert_styled_para(tc, anchor, it["text"], heading=True)
            elif kind == "list":
                if not used_first:
                    _force_left(first)
                    _apply_para_format(first, first_line_indent_chars=0)
                    _apply_hanging_indent(first, 24, 24)
                    run = first.add_run(it["text"])
                    _apply_run_style(run, BODY_SIZE_PT, bold=False)
                    used_first = True
                else:
                    anchor = _insert_styled_para(tc, anchor, it["text"], list_item=True)
            else:
                if not used_first:
                    _force_left(first)
                    _apply_para_format(first, first_line_indent_chars=2)
                    run = first.add_run(it["text"])
                    _apply_run_style(run, BODY_SIZE_PT, bold=False)
                    used_first = True
                else:
                    anchor = _insert_styled_para(tc, anchor, it["text"], heading=False)


def _last_child_before(tc, anchor):
    """返回 anchor 之后最后一个已插入元素的引用（用于继续往后追加）"""
    try:
        children = list(tc)
        idx = children.index(anchor)
        return children[-1] if idx >= 0 and len(children) > idx + 1 else anchor
    except Exception:
        return anchor


def _insert_styled_para(tc, anchor, text, heading=False, list_item=False):
    """在单元格 anchor 后插入一个带正确字体/缩进的段落
    heading=True → 四号加粗；list_item=True → 小四、悬挂缩进；否则小四首行缩进"""
    p = _insert_para_after(tc, anchor, text)
    from docx.text.paragraph import Paragraph
    para = Paragraph(p, None)
    _force_left(para)
    if heading:
        _apply_para_format(para, first_line_indent_chars=0,
                           space_before=HEADING_SPACE_BEFORE,
                           space_after=HEADING_SPACE_AFTER)
        for run in para.runs:
            _apply_run_style(run, HEADING_SIZE_PT, bold=True,
                             char_spacing_twips=HEADING_CHAR_SPACING_TWIPS)
    elif list_item:
        _apply_para_format(para, first_line_indent_chars=0)
        _apply_hanging_indent(para, left_pt=24, hanging_pt=24)
        for run in para.runs:
            _apply_run_style(run, BODY_SIZE_PT, bold=False)
    else:
        _apply_para_format(para, first_line_indent_chars=2)
        for run in para.runs:
            _apply_run_style(run, BODY_SIZE_PT, bold=False)
    return p


def _apply_hanging_indent(para, left_pt=24, hanging_pt=24):
    """列表项悬挂缩进：编号顶到左，文字对齐到 2 字符处"""
    from docx.shared import Pt as _Pt
    fmt = para.paragraph_format
    fmt.left_indent = _Pt(left_pt)
    fmt.first_line_indent = _Pt(-hanging_pt)


def _insert_dashed_placeholder_after(tc, anchor, cell, item):
    """在单元格里、anchor 之后插入虚线占位框（并做上下留空）"""
    from docx.shared import Cm as _Cm
    # 空一行，避免框贴着上一段文字
    from docx.text.paragraph import Paragraph as _Paragraph
    spacer = _insert_styled_para(tc, anchor, "")
    _apply_para_format(_Paragraph(spacer, None), first_line_indent_chars=0,
                       space_before=0, space_after=0)
    tbl = _add_dashed_placeholder(cell, item["label"], item["width"], item["height"])
    # 把表格挪到 anchor 之后（add_table 会加到单元格末尾）
    if tbl is not None and spacer is not None:
        tbl._tbl.getparent().remove(tbl._tbl)
        spacer.addnext(tbl._tbl)
    return tbl


# OOXML 规定的 w:rPr 子元素顺序。顺序错同样会被 Word 静默忽略（字间距/字号不生效）。
_RPR_ORDER = [
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
    'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof',
    'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern',
    'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd',
    'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
    'specVanish', 'oMath',
]


def _sort_rpr(rpr):
    """把 w:rPr 的子元素按 OOXML 规范顺序重排。"""
    if rpr is None:
        return
    idx = {name: i for i, name in enumerate(_RPR_ORDER)}
    children = list(rpr)
    if not children:
        return
    unknown = len(_RPR_ORDER)
    ordered = sorted(children, key=lambda el: idx.get(el.tag.split('}')[-1], unknown))
    if [c.tag for c in ordered] == [c.tag for c in children]:
        return
    for c in children:
        rpr.remove(c)
    for c in ordered:
        rpr.append(c)


def _clear_empty_paras(tc):
    """删除 tc 下除第一个段落外的空占位段"""
    from docx.oxml.ns import qn as _qn
    ps = tc.findall(_qn('w:p'))
    for p in ps[1:]:
        texts = p.findall('.//' + _qn('w:t'))
        if not any((t.text or "").strip() for t in texts):
            tc.remove(p)


# OOXML 规定的 w:pPr 子元素顺序。顺序错了 Word 会**静默忽略**后面那些属性，
# 表现为「行距/缩进不生效、上下位置对不齐」，以及编辑时 Delete 键行为异常
# （删掉上一行）。必须按这个顺序写入。
_PPR_ORDER = [
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
    'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
    'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
    'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
    'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
    'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
    'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
    'sectPr', 'pPrChange',
]


def _sort_ppr(ppr):
    """把 w:pPr 的子元素按 OOXML 规范顺序重排（稳定排序，未知元素放末尾）。"""
    if ppr is None:
        return
    from docx.oxml.ns import qn as _qn
    idx = {name: i for i, name in enumerate(_PPR_ORDER)}
    children = list(ppr)
    if not children:
        return
    unknown_base = len(_PPR_ORDER)

    def key(el):
        name = el.tag.split('}')[-1]
        return idx.get(name, unknown_base)

    ordered = sorted(children, key=key)   # Python 排序稳定，同序保持原有相对位置
    if [c.tag for c in ordered] == [c.tag for c in children]:
        return                            # 已经合法，不动
    for c in children:
        ppr.remove(c)
    for c in ordered:
        ppr.append(c)


def _force_left(paragraph):
    """强行把段落设为左对齐"""
    from docx.oxml.ns import qn as _qn
    pPr = paragraph._p.get_or_add_pPr()
    for jc in pPr.findall(_qn('w:jc')):
        pPr.remove(jc)
    jc = pPr.makeelement(_qn('w:jc'), {})
    jc.set(_qn('w:val'), 'left')
    pPr.append(jc)
    # jc 必须排在 spacing/ind 之后、rPr 之前，统一重排一次
    _sort_ppr(pPr)


# ═══════════════════ 样式写入（宋体 / 四号标题 / 小四正文）═══════════════════

def _apply_run_style(run, size_pt=BODY_SIZE_PT, bold=False, font=BODY_FONT,
                     char_spacing_twips=None):
    """给 run 设置中英文字体、字号、加粗、字间距（收紧）"""
    from docx.shared import Pt as _Pt
    run.font.size = _Pt(size_pt)
    run.font.bold = bold
    run.font.name = font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.insert(0, rFonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rFonts.set(qn(attr), font)
    # 字间距：w:spacing（单位 1/20 pt，负数=收紧）
    if char_spacing_twips is None:
        char_spacing_twips = BODY_CHAR_SPACING_TWIPS
    if char_spacing_twips:
        for old in rPr.findall(qn('w:spacing')):
            rPr.remove(old)
        sp = rPr.makeelement(qn('w:spacing'), {})
        sp.set(qn('w:val'), str(int(char_spacing_twips)))
        rPr.append(sp)
    _sort_rpr(rPr)   # 顺序错会让字间距/字号被 Word 忽略


def _apply_para_format(para, first_line_indent_chars=2, space_before=BODY_SPACE_BEFORE,
                       space_after=BODY_SPACE_AFTER,
                       line_pt=BODY_LINE_SPACING_EXACT_PT,
                       line_rule='exact'):
    """段落格式：首行缩进 2 字符、固定行距（学术论文标准）、段间距。

    line_pt=18 + line_rule='exact' → 固定值 18 磅（中文论文最常用），
    比 1.5 倍行距（≈21.6pt）紧凑得多。
    """
    from docx.shared import Pt as _Pt
    fmt = para.paragraph_format
    fmt.space_before = _Pt(space_before)
    fmt.space_after = _Pt(space_after)
    # 首行缩进 2 字符（小四=12pt，2字符=24pt）
    fmt.first_line_indent = _Pt(24 if first_line_indent_chars else 0)
    # 行距：固定值，用 XML 精确设置
    pPr = para._p.get_or_add_pPr()
    for old in pPr.findall(qn('w:spacing')):
        pPr.remove(old)
    sp = pPr.makeelement(qn('w:spacing'), {})
    sp.set(qn('w:line'), str(int(line_pt * 20)))   # pt → twips
    sp.set(qn('w:lineRule'), line_rule)
    sp.set(qn('w:after'), str(int(space_after * 20)))
    sp.set(qn('w:before'), str(int(space_before * 20)))
    pPr.append(sp)
    _sort_ppr(pPr)   # 保证 pPr 子元素顺序合法，否则 Word 静默忽略


def _add_body_paragraph(container, text, style_src=None):
    """新增一个正文段落：宋体小四、不加粗、首行缩进、左对齐"""
    p = container.add_paragraph()
    _force_left(p)
    _apply_para_format(p)
    run = p.add_run(text)
    _apply_run_style(run, BODY_SIZE_PT, bold=False)
    return p


def _add_heading_paragraph(container, text):
    """新增一个小标题：宋体四号、加粗、不缩进、段前段后留空"""
    p = container.add_paragraph()
    _force_left(p)
    _apply_para_format(p, first_line_indent_chars=0,
                       space_before=HEADING_SPACE_BEFORE, space_after=HEADING_SPACE_AFTER)
    run = p.add_run(text)
    _apply_run_style(run, HEADING_SIZE_PT, bold=True,
                     char_spacing_twips=HEADING_CHAR_SPACING_TWIPS)
    return p


def _add_dashed_placeholder(container, label, width_cm=12.0, height_cm=6.0):
    """插入一个虚线框占位区域：用户可在框内直接贴图/贴表格。

    实现：1×1 的无填充表格，四边 dashed 边框，行高固定 = height_cm，
    单元格内居中放灰字提示（如【此处插入实验接线图】）。
    """
    from docx.shared import Cm as _Cm, Pt as _Pt, RGBColor as _RGB
    from docx.enum.table import WD_TABLE_ALIGNMENT

    tbl = container.add_table(rows=1, cols=1)
    # 左对齐 + 撑满宽度，保证虚线框左右边框与主表对齐（不居中）
    try:
        tbl.alignment = None
    except Exception:
        pass
    tblPr = tbl._tbl.tblPr
    if tblPr is None:
        tblPr = tbl._tbl.makeelement(qn('w:tblPr'), {})
        tbl._tbl.insert(0, tblPr)
    # 撑满父单元格可用宽度（width_cm 仅作没有父容器时的兜底）
    parent_tc = getattr(container, "_tc", None)
    fb = None
    if parent_tc is None and hasattr(container, "sections"):
        fb = _doc_text_width_dxa(container)
    _stretch_table_full_width(tbl, parent_tc, fallback_width=fb)
    # 若容器不是单元格（无父宽度），退回到 width_cm
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is not None and parent_tc is None and fb is None:
        tblW.set(qn('w:type'), 'dxa')
        tblW.set(qn('w:w'), str(int(width_cm * 567)))

    # ── 虚线边框（四边 + 内框全 dashed）──
    borders = tblPr.find(qn('w:tblBorders'))
    if borders is None:
        borders = tblPr.makeelement(qn('w:tblBorders'), {})
        tblPr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        old = borders.find(qn('w:' + edge))
        if old is not None:
            borders.remove(old)
        el = borders.makeelement(qn('w:' + edge), {
            qn('w:val'): 'dashed',
            qn('w:sz'): '8',          # 边框粗细（1/8 pt）
            qn('w:space'): '0',
            qn('w:color'): '808080',  # 灰虚线
        })
        borders.append(el)

    # ── 行高固定 = height_cm ──
    tr = tbl.rows[0]
    trPr = tr._tr.get_or_add_trPr()
    for old in trPr.findall(qn('w:trHeight')):
        trPr.remove(old)
    trH = trPr.makeelement(qn('w:trHeight'), {
        qn('w:val'): str(int(height_cm * 567)),
        qn('w:hRule'): 'atLeast',
    })
    trPr.append(trH)

    # ── 单元格：垂直居中 + 灰字提示 ──
    cell = tbl.cell(0, 0)
    # 垂直居中
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:vAlign')):
        tcPr.remove(old)
    vA = tcPr.makeelement(qn('w:vAlign'), {qn('w:val'): 'center'})
    tcPr.append(vA)
    cell.text = ""
    para = cell.paragraphs[0]
    # 提示文字水平居中，但正文规范要求左对齐 → 占位提示居中更醒目
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _AL
    para.alignment = _AL.CENTER
    _apply_para_format(para, first_line_indent_chars=0, space_before=0, space_after=0)
    run = para.add_run(label if label.startswith("【") else f"【{label}】")
    _apply_run_style(run, 10.5, bold=False)   # 五号灰字
    run.font.color.rgb = _RGB(0x99, 0x99, 0x99)
    return tbl


def _doc_text_width_dxa(doc):
    """读取文档版心宽度（页宽 - 左右页边距），单位 dxa。读不到返回 None。"""
    try:
        sec = doc.sections[0]
        width = sec.page_width - sec.left_margin - sec.right_margin
        return int(width.twips) if hasattr(width, 'twips') else int(width / 635)
    except Exception:
        return None


def _cell_side_margin_dxa(container_tc, default=108):
    """读取父单元格的左右边距（dxa）。优先读 tcMar，其次 tblCellMar，最后用默认 108。"""
    from docx.oxml.ns import qn as _qn
    if container_tc is None:
        return default, default
    left = right = None
    tcPr = container_tc.find(_qn('w:tcPr'))
    if tcPr is not None:
        m = tcPr.find(_qn('w:tcMar'))
        if m is not None:
            for edge, var in (('left', 'l'), ('right', 'r')):
                el = m.find(_qn('w:' + edge))
                if el is not None and el.get(_qn('w:w')) is not None:
                    if edge == 'left':
                        left = int(el.get(_qn('w:w')))
                    else:
                        right = int(el.get(_qn('w:w')))
    return (left if left is not None else default,
            right if right is not None else default)


def _stretch_table_full_width(table, container_tc=None, fallback_width=None):
    """把表格宽度撑满父单元格的【全部内宽】，并用负 tblInd 抵消单元格左边距，
    使表格左右边框与主表外框（及上下各行的 cell 边框）严格对齐。

    三条铁律（之前边框参差的根因）：
    1. tblW 必须严格等于 gridCols 之和，否则 Word 按 gridCols 排版、右边框跑出去；
    2. 绝不能设 jc=center —— 居中会让任何窄表永远贴不齐右边；
    3. 表格放进单元格后默认在左边距（108 dxa）处起始、右侧也留 108，
       所以要么把宽度减 216（右边还是会差一条线），要么宽度用满 + tblInd=-108
       把左边界顶到单元格内框上。本函数用后者，两侧都严丝合缝。
    """
    from docx.oxml.ns import qn as _qn
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = tbl.makeelement(_qn('w:tblPr'), {})
        tbl.insert(0, tblPr)

    left_mar, right_mar = _cell_side_margin_dxa(container_tc)

    # 目标宽度 = 父单元格全宽（用负 tblInd 把左边界顶到内框，右边自然到内框右沿）
    # 注意：宽度取父 tcW 全宽，不是 tcW-左右边距，否则右侧会短一条边距（216 dxa）。
    target = None
    if container_tc is not None:
        tcPr = container_tc.find(_qn('w:tcPr'))
        if tcPr is not None:
            tcW = tcPr.find(_qn('w:tcW'))
            if tcW is not None and tcW.get(_qn('w:w')):
                try:
                    target = int(tcW.get(_qn('w:w')))
                except Exception:
                    target = None
    if target is None and fallback_width:
        target = int(fallback_width)

    grid = tbl.find(_qn('w:tblGrid'))
    cols = grid.findall(_qn('w:gridCol')) if grid is not None else []
    ncols = len(cols) or (len(table.columns) if table.columns else 1)
    if target is None or target <= 0:
        try:
            target = sum(int(c.get(_qn('w:w')) or 0) for c in cols)
        except Exception:
            target = 0
    if not target:
        return

    # 等分列宽，写进 gridCol（末列补齐余数保证总和精确）
    per = target // ncols
    for i, c in enumerate(cols):
        w = per if i < ncols - 1 else target - per * (ncols - 1)
        c.set(_qn('w:w'), str(w))

    # tblW 与 gridCols 之和一致
    tblW = tblPr.find(_qn('w:tblW'))
    if tblW is None:
        tblW = tblPr.makeelement(_qn('w:tblW'), {})
        tblPr.append(tblW)
    tblW.set(_qn('w:type'), 'dxa')
    tblW.set(_qn('w:w'), str(target))

    # 负左缩进，把表格左边界顶到单元格内框上（抵消单元格左边距）
    tblInd = tblPr.find(_qn('w:tblInd'))
    if tblInd is None:
        tblInd = tblPr.makeelement(_qn('w:tblInd'), {})
        tblPr.append(tblInd)
    tblInd.set(_qn('w:type'), 'dxa')
    tblInd.set(_qn('w:w'), str(-left_mar if container_tc is not None else 0))

    # 布局固定，避免 Word 自动重排改变列宽
    layout = tblPr.find(_qn('w:tblLayout'))
    if layout is None:
        layout = tblPr.makeelement(_qn('w:tblLayout'), {})
        tblPr.append(layout)
    layout.set(_qn('w:type'), 'fixed')

    # 去掉居中，改为左对齐（jc 缺失=左）
    for old in tblPr.findall(_qn('w:jc')):
        tblPr.remove(old)

    # 每个单元格 tcW 同步，否则单元格仍按旧宽度显示
    for row in table.rows:
        for ci, c in enumerate(row.cells):
            tcPr = c._tc.find(_qn('w:tcPr'))
            if tcPr is None:
                continue
            w_el = tcPr.find(_qn('w:tcW'))
            if w_el is None:
                w_el = tcPr.makeelement(_qn('w:tcW'), {})
                tcPr.append(w_el)
            w_el.set(_qn('w:type'), 'dxa')
            w_el.set(_qn('w:w'), str(per if ci < ncols - 1 else target - per * (ncols - 1)))


def _add_nested_table(container, rows):
    """在 cell/文档里插入带边框三线表：表头加粗，正文小四宋体"""
    from docx.shared import Pt as _Pt
    if not rows:
        return None
    ncols = max(len(r) for r in rows)
    table = container.add_table(rows=len(rows), cols=ncols)
    # 不再居中：居中会让窄表贴不齐右侧，导致边框参差
    try:
        table.alignment = None
    except Exception:
        pass
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = tbl.makeelement(qn('w:tblPr'), {})
        tbl.insert(0, tblPr)
    borders = tblPr.find(qn('w:tblBorders'))
    if borders is None:
        borders = tblPr.makeelement(qn('w:tblBorders'), {})
        tblPr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = borders.find(qn('w:' + edge))
        if el is None:
            el = borders.makeelement(qn('w:' + edge), {
                qn('w:val'): 'single', qn('w:sz'): '4', qn('w:space'): '0',
                qn('w:color'): '000000'})
            borders.append(el)
    # 撑满父单元格宽度（左右边框与父表对齐）；文档级表格对齐版心
    parent_tc = getattr(container, "_tc", None)
    fb = None
    if parent_tc is None and hasattr(container, "sections"):
        fb = _doc_text_width_dxa(container)
    _stretch_table_full_width(table, parent_tc, fallback_width=fb)
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            val = row[ci] if ci < len(row) else ""
            ccell = table.rows[ri].cells[ci]
            ccell.text = ""
            para = ccell.paragraphs[0]
            from docx.enum.text import WD_ALIGN_PARAGRAPH as _AL
            para.alignment = _AL.CENTER
            run = para.add_run(str(val))
            _apply_run_style(run, BODY_SIZE_PT, bold=(ri == 0))
    return table


def _insert_para_after(tc, anchor, text):
    """在 anchor 后插入一个左对齐的宋体正文段落，返回新段落 element"""
    from docx.oxml.ns import qn as _qn
    p = tc.makeelement(_qn('w:p'), {})
    # 显式左对齐（pPr/jc=left），避免继承单元格的居中样式
    ppe = tc.makeelement(_qn('w:pPr'), {})
    jce = tc.makeelement(_qn('w:jc'), {_qn('w:val'): 'left'})
    ppe.append(jce)
    p.insert(0, ppe)
    r = tc.makeelement(_qn('w:r'), {})
    rpr = tc.makeelement(_qn('w:rPr'), {})
    rf = tc.makeelement(_qn('w:rFonts'), {
        _qn('w:ascii'): BODY_FONT, _qn('w:hAnsi'): BODY_FONT,
        _qn('w:eastAsia'): BODY_FONT, _qn('w:cs'): BODY_FONT})
    sz = tc.makeelement(_qn('w:sz'), {_qn('w:val'): '24'})  # 12pt
    rpr.append(rf); rpr.append(sz); r.append(rpr)
    t = tc.makeelement(_qn('w:t'), {})
    t.set(_qn('xml:space'), 'preserve')
    t.text = text
    r.append(t); p.append(r)
    anchor.addnext(p)
    _sort_ppr(ppe)
    return p


def _insert_nested_table(cell, rows):
    """在 cell 末尾插入带边框的三线表（表头加粗、正文小四宋体）"""
    return _add_nested_table(cell, rows)


def _split_markdown_tables(lines):
    """把文本行拆成 Blocks：普通文本块 和 Markdown 表格块。
    返回 [{ "lines": [...], "is_table": bool, "rows": [[...]] }]"""
    i = 0
    blocks = []
    buf = []
    def flush():
        if buf:
            blocks.append({"lines": list(buf), "is_table": False})
            buf.clear()
    while i < len(lines):
        line = lines[i].strip()
        if _is_md_table_row(line):
            flush()
            header = _parse_md_row(line)
            sep_idx = i + 1
            # 分隔行 |---|---|
            if sep_idx < len(lines) and _is_md_separator(lines[sep_idx].strip()):
                rows = [header]
                j = sep_idx + 1
                while j < len(lines) and _is_md_table_row(lines[j]) and not _is_md_separator(lines[j].strip()):
                    rows.append(_parse_md_row(lines[j]))
                    j += 1
                blocks.append({"lines": lines[i:j], "is_table": True, "rows": rows})
                i = j
                continue
        buf.append(line)
        i += 1
    flush()
    return blocks


def _is_md_table_row(line):
    return line.startswith("|") and line.endswith("|") and "|" in line[1:-1]


def _is_md_separator(line):
    return line.replace("|", "").replace("-", "").replace(" ", "").replace(":", "") == ""


def _parse_md_row(line):
    cells = line.strip().strip("|").split("|")
    return [c.strip() for c in cells]


def _block_text(block):
    if block.get("is_table"):
        return []
    return block["lines"]


def _fill_data_tables(doc, tables_fill):
    """按表头匹配填充数据表，避免硬按 index 串位"""
    for table_fill in tables_fill:
        ti = table_fill.get("table_index", None)
        rows_data = table_fill.get("rows", [])
        if ti is not None and ti < len(doc.tables):
            # 明确给 index 时，跳过表头行填充
            _fill_table_rows(doc.tables[ti], rows_data)


def _fill_table_rows(table, rows_data):
    """从表头下一行开始填，尽量匹配行数"""
    start = 0
    # 若第一行像表头（含单位/符号），则从第2行开始
    if len(table.rows) > 1 and table.rows[0].cells and \
       any('_'.isdigit() or any(ch.isdigit() for ch in c.text) for c in table.rows[0].cells):
        start = 1
    for i, row_data in enumerate(rows_data):
        ri = start + i
        if ri < len(table.rows):
            for j, cell_text in enumerate(row_data):
                if j < len(table.rows[ri].cells):
                    try:
                        table.rows[ri].cells[j].text = str(cell_text)
                    except Exception:
                        pass


def _append_content(after_para, doc, content):
    """在 after_para 段落后插入正文（标题四号加粗 / 正文小四宋体 / 占位虚线框）"""
    parent = after_para._element.getparent()
    idx = list(parent).index(after_para._element)
    inserted = 0
    for it in _iter_render_items(content):
        kind = it["kind"]
        anchor_el = list(parent)[idx + inserted]
        if kind == "placeholder":
            new_el = _add_dashed_placeholder(doc, it["label"], it["width"], it["height"])._tbl
            parent.remove(new_el)
            parent.insert(idx + inserted + 1, new_el)
            inserted += 1
        elif kind == "table":
            new_tbl = _add_nested_table(doc, it["rows"])._tbl
            parent.remove(new_tbl)
            parent.insert(idx + inserted + 1, new_tbl)
            inserted += 1
        else:
            p = doc.add_paragraph()
            run = p.add_run(it["text"])
            if kind == "heading":
                _force_left(p)
                _apply_para_format(p, first_line_indent_chars=0,
                                   space_before=HEADING_SPACE_BEFORE,
                                   space_after=HEADING_SPACE_AFTER)
                _apply_run_style(run, HEADING_SIZE_PT, bold=True,
                                 char_spacing_twips=HEADING_CHAR_SPACING_TWIPS)
            elif kind == "list":
                _force_left(p)
                _apply_para_format(p, first_line_indent_chars=0)
                _apply_hanging_indent(p, 24, 24)
                _apply_run_style(run, BODY_SIZE_PT, bold=False)
            else:
                _force_left(p)
                _apply_para_format(p)
                _apply_run_style(run, BODY_SIZE_PT, bold=False)
            parent.remove(p._element)
            parent.insert(idx + inserted + 1, p._element)
            inserted += 1


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
    """追加章节标题（宋体四号加粗）+ 正文（宋体小四），标题与正文间留空"""
    h = doc.add_paragraph()
    _force_left(h)
    _apply_para_format(h, first_line_indent_chars=0, space_before=HEADING_SPACE_BEFORE,
                       space_after=HEADING_SPACE_AFTER)
    run = h.add_run(section_name)
    _apply_run_style(run, HEADING_SIZE_PT, bold=True,
                     char_spacing_twips=HEADING_CHAR_SPACING_TWIPS)

    for it in _iter_render_items(content):
        kind = it["kind"]
        if kind == "placeholder":
            _add_dashed_placeholder(doc, it["label"], it["width"], it["height"])
        elif kind == "table":
            _add_nested_table(doc, it["rows"])
        elif kind == "heading":
            sub = doc.add_paragraph()
            _force_left(sub)
            _apply_para_format(sub, first_line_indent_chars=0,
                               space_before=HEADING_SPACE_BEFORE,
                               space_after=HEADING_SPACE_AFTER)
            _apply_run_style(sub.add_run(it["text"]), HEADING_SIZE_PT, bold=True,
                             char_spacing_twips=HEADING_CHAR_SPACING_TWIPS)
        elif kind == "list":
            p = doc.add_paragraph()
            _force_left(p)
            _apply_para_format(p, first_line_indent_chars=0)
            _apply_hanging_indent(p, 24, 24)
            _apply_run_style(p.add_run(it["text"]), BODY_SIZE_PT, bold=False)
        else:
            p = doc.add_paragraph()
            _force_left(p)
            _apply_para_format(p)
            _apply_run_style(p.add_run(it["text"]), BODY_SIZE_PT, bold=False)


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
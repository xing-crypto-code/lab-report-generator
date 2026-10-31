"""
实验报告知识库检索 · retrieval.py
==================================
把「优秀范文 / 指导书 / 写作规范」放进 kb/，切块入库；生成时检索最相关的片段，
拼成一小段 context 注入 prompt，让模型照着**本课程真正的好报告**写，而不是照着
通用大模型的默认腔调写。

设计上的三个取舍（都是有意的）：

1. **纯标准库**：docx 用 zipfile + ElementTree 直接读 document.xml，不依赖
   python-docx；这样本模块可以在任何机器上跑，也不会和 libs/ 的安装路径纠缠。
2. **中文用字符二元组做 BM25**，不依赖 jieba。实验报告里「霍尔效应」「示波器」
   「误差传递」这类术语，二元组召回率已经足够，还省掉一个词典依赖。
3. **切块按章节走**：实验报告的章节名（实验目的/原理/步骤/数据处理）是天然边界，
   检索出来的一段话总是"某章节里的完整内容"，比按固定字数硬切好读得多。

用法（在仓库根目录）：

    python -m modules.retrieval stats           # 看库里有什么
    python -m modules.retrieval ingest          # 扫描 kb/ 增量入库
    python -m modules.retrieval ingest 某报告.docx
    python -m modules.retrieval search "霍尔效应 磁场 测量" -k 3

代码里用 build_context() 拿可直接拼进 prompt 的文本块。
"""
from __future__ import annotations

import argparse
import math
import os
import re
import sqlite3
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

# ─── 路径 ────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent.parent
KB_ROOT = Path(os.getenv("KB_DIR") or (ROOT / "kb"))
INDEX_DB = Path(os.getenv("KB_INDEX") or (KB_ROOT / "index.sqlite"))

SUPPORTED = {".docx", ".pdf", ".md", ".txt"}

# ─── 切块 / 检索参数 ─────────────────────────────────────────────────
CHUNK_TARGET = 900      # 单块目标字数（中文按字）
# 丢弃阈值不能设太高：实验报告的「实验器材」「数据记录」小节本身就短，
# 卡在 40 字会把一整节数据表悄悄扔掉。
CHUNK_MIN = 12
K1, B = 1.5, 0.75       # BM25 标准参数

KIND_DIRS = {
    "范文": "exemplar", "范例": "exemplar", "优秀": "exemplar", "高分": "exemplar",
    "指导书": "guide", "指导": "guide", "手册": "guide",
    "规范": "spec", "标准": "spec", "要求": "spec",
    "草稿": "draft", "自撰": "draft",
}
KIND_LABEL = {"exemplar": "范文", "guide": "指导书", "spec": "规范",
              "draft": "本人稿", "unknown": "资料"}
# 范文最值钱：它决定"好报告长什么样"，所以打分时给一点加成。
# 草稿（自己以前的稿子）内容往往更薄，默认不参与检索，避免"原理只写一行"被学走。
KIND_BOOST = {"exemplar": 1.25, "spec": 1.10, "guide": 1.0,
              "draft": 0.55, "unknown": 1.0}

# 章节权重：要学的是"怎么写原理/步骤/数据处理"，不是"实验目的一二三四"这种模板话
_HEAD_HIGH = re.compile(r"(原理|步骤|过程|数据|误差|讨论|分析|结论|小结|思考|设计)")
_HEAD_LOW = re.compile(r"(目的|仪器|器材|名称|环境|日期|成绩)")


def _heading_weight(heading: str) -> float:
    if not heading:
        return 1.0
    if _HEAD_LOW.search(heading):
        return 0.70
    if _HEAD_HIGH.search(heading):
        return 1.15
    return 1.0


def _text_weight(text: str) -> float:
    """纯数据表压一档：格式值得学，但数字被照抄的风险太高。"""
    if not text:
        return 1.0
    ratio = sum(c.isdigit() for c in text) / len(text)
    if ratio > 0.22:
        return 0.75
    if ratio > 0.12:
        return 0.90
    return 1.0

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ─── 实验类型识别（轻量版）───────────────────────────────────────────
# 注：modules/filler.py 里另有一份 detect_experiment_type 用于生成阶段。
# 这里刻意不 import 它——filler 会拉起 openai 客户端，检索模块要保持零依赖。
EXP_KEYWORDS = {
    "物理实验": ["霍尔", "光电效应", "杨氏模量", "示波器", "单摆", "热膨胀",
                 "表面张力", "分光计", "牛顿环", "迈克尔逊", "转动惯量"],
    "电路/电子实验": ["电路", "二极管", "三极管", "运放", "放大电路", "逻辑门",
                   "真值表", "波形", "基尔霍夫", "叠加定理", "戴维南", "触发器"],
    "力学实验": ["力学", "拉伸", "弹性模量", "泊松比", "扭矩", "应力", "应变",
                 "弯曲", "振动", "加速度"],
    "光学实验": ["光学", "干涉", "衍射", "偏振", "透镜焦距", "光栅", "光路"],
    "化学实验": ["化学", "滴定", "溶液", "反应方程式", "吸光度", "分光光度",
                 "浓度", "摩尔", "沉淀"],
    "计算机/编程实验": ["编程", "算法", "代码", "程序", "数据结构", "时间复杂度",
                   "python", "c++", "java", "数据库"],
}

# 章节名（用于判断"这一段是不是标题"，也用于切块边界）
_NUM_PREFIX = re.compile(r"^(?:[一二三四五六七八九十]+\s*[、.．]|\d+\s*[、.．]\s*|[（(]\d+[)）]\s*)")
_SECTION_NAME = re.compile(
    r"^(?:实验名称|实验目的|实验原理|实验内容|实验器材|实验仪器|实验步骤|实验过程|"
    r"实验数据|数据处理|误差分析|实验结果|结果与讨论|结果分析|讨论|结论|小结|"
    r"思考题|注意事项|预习|实验环境|算法说明|原理说明)")
# 指导书里"实验一 常用电子仪器仪表使用"这类大标题，是整本书最有用的切块边界
_EXP_ITEM = re.compile(r"^实验\s*[一二三四五六七八九十\d]+")


def _is_heading(style: str, text: str) -> bool:
    t = text.strip()
    if len(t) > 40:
        return False
    # 电路图/表格里抽出来的 "1  1"、"V cc 2CP 2SD" 这类标注不是标题
    if len(re.findall(r"[\u4e00-\u9fff]", t)) < 2:
        return False
    if re.match(r"(?i)^(heading|标题)\s*\d*$", style or ""):
        return True
    if re.fullmatch(r"\d", style or "") and len(t) <= 20:
        return True      # 中文版 Word 的一级标题 styleId 常常就是 "1"/"2"
    if len(t) <= 30 and _EXP_ITEM.match(t):
        return True
    # 剩下靠内容认：去掉"一、""2."这类编号后，短、以章节名开头、不是整句话。
    # 能匹配到 "实验过程（可另附页）""二、实验数据记录" 这种没套用标题样式的写法。
    core = _NUM_PREFIX.sub("", t).strip()
    return (len(core) <= 16 and bool(_SECTION_NAME.match(core))
            and not core.endswith(("。", "；", "，", ",")))


def guess_exp_type(text: str) -> str | None:
    """按关键词出现次数猜实验类型。

    要求"第一名至少 3 次，且领先第二名 1.6 倍"，否则返回 None。
    这是为了防止写作规范这类通篇举例的文档被强行归到某一类——它一旦带了类型标签，
    检索时就会把真正相关的范例挤掉。
    """
    low = text.lower()
    scored = sorted(((sum(low.count(w.lower()) for w in words), name)
                     for name, words in EXP_KEYWORDS.items()), reverse=True)
    best_score, best = scored[0]
    runner_up = scored[1][0] if len(scored) > 1 else 0
    if best_score < 3 or best_score < runner_up * 1.6:
        return None
    return best


# ─── 分词：CJK 二元组 + 拉丁词 ────────────────────────────────────────
_CJK_RUN = re.compile(r"[\u4e00-\u9fff]+")
_LATIN = re.compile(r"[a-z0-9][a-z0-9._+\-/]*")
_STOP = {"的", "了", "是", "在", "和", "与", "及", "对", "为", "用", "把",
         "我们", "进行", "可以", "以及", "或者", "然后", "其中"}


def tokenize(text: str) -> list[str]:
    toks: list[str] = []
    low = text.lower()
    for run in _CJK_RUN.findall(low):
        if len(run) == 1:
            toks.append(run)
        else:
            toks.extend(run[i:i + 2] for i in range(len(run) - 1))
    toks.extend(_LATIN.findall(low))
    return [t for t in toks if t not in _STOP]


# ─── 解析：docx / pdf / md / txt → [(style, text)] ───────────────────
def _para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(W + "t")).strip()


def _para_style(p) -> str:
    ppr = p.find(W + "pPr")
    if ppr is None:
        return ""
    ps = ppr.find(W + "pStyle")
    return (ps.get(W + "val") or "") if ps is not None else ""


def _blocks_from_docx(path: Path) -> list[tuple[str, str]]:
    with zipfile.ZipFile(path) as z:
        try:
            xml = z.read("word/document.xml")
        except KeyError:
            raise ValueError("不是有效的 .docx（缺 word/document.xml）")
    root = ET.fromstring(xml)
    body = root.find(W + "body")
    if body is None:
        return []

    out: list[tuple[str, str]] = []

    def walk(node):
        for child in node:
            tag = child.tag.split("}")[-1]
            if tag == "p":
                txt = _para_text(child)
                if txt:
                    out.append((_para_style(child), txt))
            elif tag == "tbl":
                for tr in child.findall(W + "tr"):
                    cells = tr.findall(W + "tc")
                    texts = [" ".join(_para_text(p) for p in tc.findall(W + "p")).strip()
                             for tc in cells]
                    filled = [t for t in texts if t]
                    # 数据表：每格都很短（数字/符号）→ 整行保留成 "a | b | c"
                    if len(filled) >= 2 and all(len(t) <= 12 for t in filled):
                        out.append(("", " | ".join(filled)))
                    else:
                        # 表单型模板（中文实验报告常见）：单元格里塞着标题和正段，
                        # 必须逐段展开，否则"实验过程"这类标题会被并进正文里
                        for tc in cells:
                            walk(tc)
            elif tag in ("sdt", "sdtContent"):
                walk(child)

    walk(body)
    return out


def _blocks_from_pdf(path: Path) -> list[tuple[str, str]]:
    try:
        import pymupdf as fitz            # PyMuPDF ≥ 1.24 的新名字
    except ImportError:
        try:
            import fitz                   # 老版本
        except ImportError as exc:
            raise RuntimeError(
                "解析 PDF 需要 pymupdf：pip install pymupdf；"
                "或把 PDF 另存为 .docx / 复制成 .md 再入库"
            ) from exc
    doc = fitz.open(str(path))
    try:
        out = []
        for page in doc:
            for line in page.get_text("text").splitlines():
                if line.strip():
                    out.append(("", line.strip()))
        return out
    finally:
        doc.close()


def _blocks_from_markdown(text: str) -> list[tuple[str, str]]:
    out = []
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", s)
        out.append((f"heading{len(m.group(1))}", m.group(2).strip()) if m else ("", s))
    return out


def load_blocks(path: Path) -> list[tuple[str, str]]:
    suf = path.suffix.lower()
    if suf == ".docx":
        return _blocks_from_docx(path)
    if suf == ".pdf":
        return _blocks_from_pdf(path)
    raw = path.read_bytes()
    for enc in ("utf-8", "gb18030", "utf-16"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("无法识别文本编码（试过 utf-8 / gb18030 / utf-16）")
    text = text.lstrip("\ufeff")
    return _blocks_from_markdown(text) if suf == ".md" else \
        [("", ln.strip()) for ln in text.splitlines() if ln.strip()]


# ─── 切块 ────────────────────────────────────────────────────────────
def _split_long(text: str, target: int) -> list[str]:
    if len(text) <= target:
        return [text]
    parts, buf = [], ""
    for sent in re.split(r"(?<=[。！？；!?;])", text):
        if buf and len(buf) + len(sent) > target:
            parts.append(buf)
            buf = sent
        else:
            buf += sent
    if buf.strip():
        parts.append(buf)
    return parts


def _clean_heading(t: str) -> str:
    """去掉目录页码与多余空格："实验五  时序逻辑电路设计 20" → "实验五 时序逻辑电路设计"。"""
    t = re.sub(r"\s+\d{1,3}$", "", t.strip())
    return re.sub(r"\s{2,}", " ", t)


def chunk_blocks(blocks: list[tuple[str, str]]) -> list[tuple[str, str]]:
    """[(style, text)] → [(heading, chunk_text)]，标题作为切块边界。"""
    chunks: list[tuple[str, str]] = []
    heading, buf = "", ""

    def flush():
        nonlocal buf
        body = buf.strip()
        buf = ""
        if len(body) >= CHUNK_MIN:
            for piece in _split_long(body, CHUNK_TARGET):
                if len(piece.strip()) >= CHUNK_MIN:
                    chunks.append((heading, piece.strip()))

    for style, text in blocks:
        if _is_heading(style, text):
            flush()
            heading = _clean_heading(text)
            continue
        if buf and len(buf) + len(text) > CHUNK_TARGET:
            flush()
        buf += ("\n" if buf else "") + text
    flush()
    return chunks


# ─── 分类 ────────────────────────────────────────────────────────────
def classify(rel: str, text: str) -> tuple[str, str | None]:
    p = Path(rel)
    kind = "unknown"
    for part in p.parts[:-1]:
        if part.strip() in KIND_DIRS:
            kind = KIND_DIRS[part.strip()]
            break
    if kind == "unknown":
        stem = p.stem
        if re.search(r"范文|范例|优秀|高分|模板范例", stem):
            kind = "exemplar"
        elif re.search(r"指导|手册|讲义|实验说明", stem):
            kind = "guide"
        elif re.search(r"规范|标准|要求|评分", stem):
            kind = "spec"
    exp = guess_exp_type(text[:3000] + "\n" + p.stem)
    if kind == "spec":
        exp = None      # 写作规范是全类型通用的，不该被绑定到某一类实验
    return kind, exp


# ─── 存储 ────────────────────────────────────────────────────────────
_SCHEMA = """
CREATE TABLE IF NOT EXISTS docs (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    path     TEXT UNIQUE NOT NULL,
    title    TEXT,
    kind     TEXT,
    exp_type TEXT,
    mtime    REAL,
    size     INTEGER,
    n_chunks INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS chunks (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_id  INTEGER NOT NULL REFERENCES docs(id) ON DELETE CASCADE,
    ord     INTEGER NOT NULL,
    heading TEXT,
    text    TEXT NOT NULL,
    n_chars INTEGER
);
CREATE INDEX IF NOT EXISTS idx_chunks_doc ON chunks(doc_id);
"""


def connect() -> sqlite3.Connection:
    KB_ROOT.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(INDEX_DB))
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    con.executescript(_SCHEMA)
    return con


def _rel(path: Path) -> str:
    path = path.resolve()
    for base in (KB_ROOT, ROOT):
        try:
            return path.relative_to(base.resolve()).as_posix()
        except ValueError:
            continue
    return path.as_posix()


def _collect(paths: list[str] | None) -> list[Path]:
    if paths:
        found: list[Path] = []
        for raw in paths:
            p = Path(raw)
            if not p.is_absolute():
                p = (Path.cwd() / p)
            if p.is_dir():
                found.extend(_iter_files(p))
            elif p.is_file():
                found.append(p)
            else:
                raise FileNotFoundError(f"找不到：{raw}")
        return found
    KB_ROOT.mkdir(parents=True, exist_ok=True)
    return _iter_files(KB_ROOT)


def _iter_files(root: Path) -> list[Path]:
    out = []
    for p in sorted(root.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in SUPPORTED:
            continue
        if p.name.startswith(("~$", ".")) or "index.sqlite" in p.name:
            continue
        # kb/README.md 这类说明文件是给人看的，不是知识
        if p.name.lower() in ("readme.md", "readme.txt", "说明.md"):
            continue
        out.append(p)
    return out


# ─── 入库 ────────────────────────────────────────────────────────────
def ingest(paths: list[str] | None = None, rebuild: bool = False) -> dict:
    stats = {"added": 0, "updated": 0, "removed": 0, "skipped": 0,
             "chunks": 0, "errors": []}
    files = _collect(paths)
    con = connect()
    try:
        with con:
            if rebuild:
                con.execute("DELETE FROM chunks")
                con.execute("DELETE FROM docs")
            existing = {r["path"]: r for r in con.execute("SELECT * FROM docs")}
            seen: set[str] = set()

            for f in files:
                rel = _rel(f)
                seen.add(rel)
                try:
                    st = f.stat()
                except OSError as exc:
                    stats["errors"].append(f"{rel}: {exc}")
                    continue
                old = existing.get(rel)
                if (old and not rebuild and old["size"] == st.st_size
                        and abs(old["mtime"] - st.st_mtime) < 1e-6):
                    stats["skipped"] += 1
                    continue

                try:
                    blocks = load_blocks(f)
                    chunks = chunk_blocks(blocks)
                except Exception as exc:                      # 单份坏文件不该中断整批
                    stats["errors"].append(f"{rel}: {type(exc).__name__}: {exc}")
                    continue
                if not chunks:
                    stats["errors"].append(f"{rel}: 没解析出正文（扫描件需先 OCR）")
                    continue

                full = "\n".join(t for _h, t in blocks)
                kind, exp = classify(rel, full)
                title = next((h for h, _t in chunks if h), f.stem)

                if old:
                    con.execute("DELETE FROM chunks WHERE doc_id = ?", (old["id"],))
                    con.execute(
                        "UPDATE docs SET title=?,kind=?,exp_type=?,mtime=?,size=?,n_chunks=? WHERE id=?",
                        (title, kind, exp, st.st_mtime, st.st_size, len(chunks), old["id"]))
                    doc_id = old["id"]
                    stats["updated"] += 1
                else:
                    cur = con.execute(
                        "INSERT INTO docs(path,title,kind,exp_type,mtime,size,n_chunks) "
                        "VALUES (?,?,?,?,?,?,?)",
                        (rel, title, kind, exp, st.st_mtime, st.st_size, len(chunks)))
                    doc_id = cur.lastrowid
                    stats["added"] += 1

                con.executemany(
                    "INSERT INTO chunks(doc_id,ord,heading,text,n_chars) VALUES (?,?,?,?,?)",
                    [(doc_id, i, h, t, len(t)) for i, (h, t) in enumerate(chunks)])
                stats["chunks"] += len(chunks)

            # 只有"全量扫描 kb/"时才清理磁盘上已删除的文件；
            # 指定单个文件入库时若无条件清理，会把库里其余文档全部删掉。
            if paths is None:
                for rel, row in existing.items():
                    if rel not in seen:
                        con.execute("DELETE FROM docs WHERE id = ?", (row["id"],))
                        stats["removed"] += 1
    finally:
        con.close()
    _CACHE.clear()
    return stats


# ─── 检索 ────────────────────────────────────────────────────────────
_CACHE: dict = {}


def _corpus() -> tuple[list[dict], dict[str, int], float]:
    """加载全部块并建 DF/平均长度；按索引文件 mtime 缓存。"""
    key = INDEX_DB.stat().st_mtime if INDEX_DB.exists() else None
    if _CACHE.get("key") == key and _CACHE.get("items") is not None:
        return _CACHE["items"], _CACHE["df"], _CACHE["avgdl"]

    if not INDEX_DB.exists():
        return [], {}, 0.0

    con = connect()
    try:
        rows = con.execute(
            "SELECT c.id, c.heading, c.text, d.path, d.kind, d.exp_type "
            "FROM chunks c JOIN docs d ON d.id = c.doc_id ORDER BY d.path, c.ord"
        ).fetchall()
    finally:
        con.close()

    items = []
    df: Counter = Counter()
    for r in rows:
        toks = tokenize((r["heading"] or "") + " " + r["text"])
        if not toks:
            continue
        items.append({
            "id": r["id"], "heading": r["heading"] or "", "text": r["text"],
            "path": r["path"], "kind": r["kind"] or "unknown",
            "exp_type": r["exp_type"], "tokens": toks,
        })
        df.update(set(toks))

    avgdl = (sum(len(i["tokens"]) for i in items) / len(items)) if items else 0.0
    _CACHE.update({"key": key, "items": items, "df": dict(df), "avgdl": avgdl})
    return items, _CACHE["df"], avgdl


def _is_cover(text: str) -> bool:
    """报告封面（班/学号/姓名/课程名称那几行表格），注入时没有价值。"""
    return sum(k in text for k in ("学  号", "姓  名", "课程名称", "实验日期", "指导教师")) >= 2


def search(query: str, k: int = 3, exp_type: str | None = None,
           kinds: list[str] | None = None,
           exclude_kinds: set[str] | None = None,
           drop_boilerplate: bool = False) -> list[dict]:
    """BM25 检索。query 通常是整段 OCR 文本，所以按"词是否出现"计分（qt=1），
    避免长文本里高频词把排序带偏。

    drop_boilerplate=True 时剔除"实验目的/实验仪器"这类模板段落：生成时这两栏
    本来就要求照抄指导书原文，范例里的同类段落只会白占上下文预算。
    """
    items, df, avgdl = _corpus()
    if not items or not query.strip():
        return []

    q_terms = set(tokenize(query))
    n = len(items)
    scored: list[tuple[float, dict]] = []

    for it in items:
        if exp_type and it["exp_type"] and it["exp_type"] != exp_type:
            continue
        if kinds and it["kind"] not in kinds:
            continue
        if exclude_kinds and it["kind"] in exclude_kinds:
            continue
        if drop_boilerplate and (_HEAD_LOW.search(it["heading"] or "")
                                 or _is_cover(it["text"])):
            continue
        tf = Counter(it["tokens"])
        dl = len(it["tokens"])
        norm = K1 * (1 - B + B * dl / avgdl) if avgdl else K1
        score = 0.0
        for term in q_terms:
            f = tf.get(term)
            if not f:
                continue
            d = df.get(term, 0)
            idf = math.log(1 + (n - d + 0.5) / (d + 0.5))
            score += idf * (f * (K1 + 1)) / (f + norm)
        if score <= 0:
            continue
        score *= KIND_BOOST.get(it["kind"], 1.0)
        score *= _heading_weight(it["heading"])
        score *= _text_weight(it["text"])
        scored.append((score, it))

    scored.sort(key=lambda x: -x[0])
    results = []
    for score, it in scored[:max(1, k)]:
        results.append({**{x: it[x] for x in ("id", "heading", "text", "path",
                                              "kind", "exp_type")},
                        "score": round(score, 3)})
    return results


def _trim(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    cut = text[:limit]
    m = max(cut.rfind("。"), cut.rfind("；"), cut.rfind("\n"))
    if m > limit * 0.5:
        cut = cut[:m + 1]
    return cut.rstrip() + "……"


def build_context(query: str, exp_type: str | None = None, k: int = 3,
                  max_chars: int = 1500, kinds: list[str] | None = None) -> str:
    """检索并拼成可注入 prompt 的文本块；无命中返回空串（调用方据此跳过）。

    草稿默认不参与：它们通常写得更薄，混进来会拉低"章节该写多深"的示范。
    确认某份草稿质量够好，设 KB_DRAFTS=1 即可放行，或直接把文件移进 范文/。
    """
    if kinds is None:
        exclude = set() if os.getenv("KB_DRAFTS", "").lower() in ("1", "true", "yes") \
            else {"draft"}
    else:
        exclude = set()

    hits = search(query, k=k, exp_type=exp_type, kinds=kinds,
                  exclude_kinds=exclude, drop_boilerplate=True)
    if not hits:
        return ""

    blocks, used = [], 0
    for i, h in enumerate(hits, 1):
        head = f"｜{h['heading']}" if h["heading"] else ""
        block = (f"【参考{i}｜{KIND_LABEL.get(h['kind'], '资料')}："
                 f"{Path(h['path']).stem}{head}】\n{h['text'].strip()}")
        remain = max_chars - used
        if remain < 120:
            break
        block = _trim(block, remain)
        blocks.append(block)
        used += len(block) + 2
    return "\n\n".join(blocks)


def stats() -> dict:
    if not INDEX_DB.exists():
        return {"docs": 0, "chunks": 0, "by_kind": {}, "by_type": {}, "bytes": 0}
    con = connect()
    try:
        docs = con.execute("SELECT COUNT(*) c FROM docs").fetchone()["c"]
        chunks = con.execute("SELECT COUNT(*) c, COALESCE(SUM(n_chars),0) n FROM chunks").fetchone()
        by_kind = {r["kind"]: r["c"] for r in
                   con.execute("SELECT kind, COUNT(*) c FROM docs GROUP BY kind")}
        by_type = {r["exp_type"] or "未识别": r["c"] for r in
                   con.execute("SELECT exp_type, COUNT(*) c FROM docs GROUP BY exp_type")}
    finally:
        con.close()
    return {"docs": docs, "chunks": chunks["c"], "chars": chunks["n"],
            "by_kind": by_kind, "by_type": by_type,
            "bytes": INDEX_DB.stat().st_size}


# ─── CLI ────────────────────────────────────────────────────────────
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        prog="python -m modules.retrieval",
        description="实验报告知识库：入库 / 检索 / 统计")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_in = sub.add_parser("ingest", help="扫描 kb/ 或指定文件，增量入库")
    p_in.add_argument("paths", nargs="*", help="文件或目录（默认扫描 kb/）")
    p_in.add_argument("--rebuild", action="store_true", help="清空索引后重建")

    p_se = sub.add_parser("search", help="检索片段")
    p_se.add_argument("query")
    p_se.add_argument("-k", type=int, default=3)
    p_se.add_argument("--type", dest="exp_type", default=None)
    p_se.add_argument("--full", action="store_true", help="打印完整片段而非摘要")

    p_st = sub.add_parser("stats", help="查看库的规模与分布")
    p_st.add_argument("-v", "--verbose", action="store_true", help="逐文档列出分类结果")

    args = ap.parse_args(argv)

    if args.cmd == "ingest":
        st = ingest(args.paths or None, rebuild=args.rebuild)
        print(f"[入库] 新增 {st['added']}｜更新 {st['updated']}｜"
              f"跳过 {st['skipped']}（未变）｜移除 {st['removed']}｜"
              f"块 {st['chunks']}")
        for e in st["errors"]:
            print(f"  ⚠ {e}", file=sys.stderr)
        return 0 if not st["errors"] else 1

    if args.cmd == "search":
        hits = search(args.query, k=args.k, exp_type=args.exp_type)
        if not hits:
            print("没有命中。先 python -m modules.retrieval ingest 入库？")
            return 1
        for i, h in enumerate(hits, 1):
            body = h["text"] if args.full else h["text"][:120].replace("\n", " ")
            print(f"\n[{i}] score={h['score']}  {KIND_LABEL.get(h['kind'], '资料')}"
                  f"｜{h['exp_type'] or '通用'}｜{h['path']}｜{h['heading'] or '—'}")
            print(f"    {body}")
        return 0

    st = stats()
    print(f"文档 {st['docs']} 份｜块 {st['chunks']} 个｜"
          f"{st.get('chars', 0)} 字｜索引 {st['bytes'] / 1024:.1f} KB")
    print(f"按类型：{st['by_kind'] or '（空）'}")
    print(f"按实验：{st['by_type'] or '（空）'}")
    if args.verbose and st["docs"]:
        con = connect()
        try:
            rows = con.execute("SELECT path, kind, exp_type, n_chunks FROM docs ORDER BY path")
            print("\n逐文档：")
            for r in rows:
                print(f"  {KIND_LABEL.get(r['kind'], r['kind']):<5}｜"
                      f"{r['exp_type'] or '通用':<8}｜{r['n_chunks']:>3} 块｜{r['path']}")
        finally:
            con.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())

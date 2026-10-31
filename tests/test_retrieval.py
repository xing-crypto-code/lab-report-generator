"""
知识库检索回归测试 · tests/test_retrieval.py
============================================
覆盖：分词 / 类型识别 / docx 解析 / 切块 / 入库增量 / BM25 检索 / context 注入规则。

全部用标准库构造样本（连 .docx 都是现拼一个 zip），不依赖 python-docx，
所以在任何环境都能跑。

运行：python tests/test_retrieval.py
"""
import os
import shutil
import sys
import zipfile
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
sys.path.insert(0, os.path.join(ROOT, 'modules'))   # 直接以模块身份导入，绕开包的重量级 __init__

_TMP = os.path.join(ROOT, 'temp', 'kbtest_retrieval')
shutil.rmtree(_TMP, ignore_errors=True)
os.makedirs(_TMP, exist_ok=True)
os.environ['KB_DIR'] = os.path.join(_TMP, 'kb')
os.environ['KB_INDEX'] = os.path.join(_TMP, 'kb', 'index.sqlite')

import retrieval                                  # noqa: E402

_failures = []


def check(name, cond, detail=''):
    status = 'ok  ' if cond else 'FAIL'
    print(f'[{status}] {name}')
    if not cond:
        print(f'       {detail}')
        _failures.append(name)


# ─── 1. 分词 ─────────────────────────────────────────────────────────
toks = retrieval.tokenize('霍尔效应与霍尔电压的测量，B=0.5T')
check('分词：中文出二元组', '霍尔' in toks and '效应' in toks, f'tokens={toks[:12]}')
check('分词：保留拉丁与数字', '0.5t' in toks, f'tokens={toks[-5:]}')
check('分词：无空串与停用词', all(t.strip() for t in toks) and '进行' not in toks)


# ─── 2. 实验类型识别 ──────────────────────────────────────────────────
check('类型：明显物理实验',
      retrieval.guess_exp_type('霍尔效应 霍尔电压 霍尔元件 霍尔传感器 霍尔片') == '物理实验')
check('类型：通篇举例的规范不给标签',
      retrieval.guess_exp_type('实验目的一行；实验原理一行；实验步骤一行；实验数据一行') is None,
      '模糊文本不应被强行归类')


# ─── 3. docx 解析（含表格式模板）──────────────────────────────────────
W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _p(text, style=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def _table(rows):
    trs = ''.join(
        '<w:tr>' + ''.join(f'<w:tc>{_p(c)}</w:tc>' for c in row) + '</w:tr>'
        for row in rows)
    return f'<w:tbl>{trs}</w:tbl>'


def make_docx(path, body_xml):
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{body_xml}</w:body></w:document>')
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('word/document.xml', xml)


docx_path = os.path.join(_TMP, 'kb', '范文', '霍尔效应实验.docx')
make_docx(docx_path,
          _p('霍尔效应实验报告', 'Heading1')
          + _p('一、实验原理')
          + _p('霍尔元件置于螺线管磁场中，磁场方向与元件平面垂直时，载流子受洛伦兹力'
               '发生偏转，在两侧建立霍尔电场。平衡时电场力与洛伦兹力相等，'
               '由此得到霍尔电压与磁感应强度的关系式。')
          + _p('二、实验数据记录')
          + _table([['B/mT', 'U/mV'], ['5.0', '2.14'], ['10.0', '4.29']]))

blocks = retrieval.load_blocks(retrieval.Path(docx_path))
texts = [t for _s, t in blocks]
check('docx：读到正文段落', any('霍尔元件置于' in t for t in texts), f'blocks={texts[:3]}')
check('docx：数据表整行保留', any(t == 'B/mT | U/mV' for t in texts),
      f'tables={[t for t in texts if "|" in t]}')
check('docx：表单型单元格逐段展开（标题没被吃掉）',
      '一、实验原理' in texts and '二、实验数据记录' in texts)


# ─── 4. 切块 ─────────────────────────────────────────────────────────
chunks = retrieval.chunk_blocks(blocks)
check('切块：有内容', len(chunks) >= 2, f'n={len(chunks)}')
check('切块：标题成为块边界', any(h == '一、实验原理' for h, _t in chunks),
      f'headings={[h for h, _t in chunks]}')
long_text = '这是一句用于测试切块的较长正文。' * 80
pieces = retrieval.chunk_blocks([('', long_text)])
check('切块：超长段落被拆开',
      len(pieces) > 1 and all(len(t) <= retrieval.CHUNK_TARGET + 40 for _h, t in pieces),
      f'sizes={[len(t) for _h, t in pieces]}')


# ─── 5. 入库 + 检索 ──────────────────────────────────────────────────
def write_md(rel, body):
    p = os.path.join(os.environ['KB_DIR'], rel)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, 'w', encoding='utf-8') as f:
        f.write(body)


write_md(os.path.join('范文', '门电路实验.md'),
         '# 门电路逻辑功能测试\n\n## 实验原理\n\n真值表把输入组合与输出电平逐一列出，'
         'TTL 门按输出电平高低判定逻辑值，与门、或门、与非门的逻辑关系由真值表给出。\n\n'
         '## 实验步骤\n\n把逻辑开关接到门电路输入端，逐组给出输入组合，'
         '用万用表测量输出电平并记录波形。\n')

st = retrieval.ingest()
check('入库：新增 2 份', st['added'] == 2, str(st))
check('入库：无解析错误', not st['errors'], str(st['errors']))
check('入库：识别出块', st['chunks'] >= 3, str(st))

st2 = retrieval.ingest()
check('入库：未改动文件被跳过', st2['skipped'] == 2 and st2['added'] == 0, str(st2))

hits = retrieval.search('霍尔电压 螺线管 磁场 测量', k=3)
check('检索：物理查询命中物理范文',
      bool(hits) and '霍尔' in hits[0]['path'], f'top={hits[0]["path"] if hits else None}')

hits2 = retrieval.search('霍尔电压 螺线管 磁场 测量', k=3, exp_type='电路/电子实验')
check('检索：类型过滤生效', all('霍尔' not in h['path'] for h in hits2),
      f'paths={[h["path"] for h in hits2]}')

hits3 = retrieval.search('真值表 与门 或门 输出电平', k=2)
check('检索：电路查询命中电路范文',
      bool(hits3) and '门电路' in hits3[0]['path'],
      f'top={hits3[0]["path"] if hits3 else None}')

# 指定单文件入库时，不能把库里其他文档当成"已删除"清掉
single = os.path.join(os.environ['KB_DIR'], '范文', '门电路实验.md')
st3 = retrieval.ingest([single], rebuild=False)
check('入库：单文件入库不误删其他文档',
      st3['removed'] == 0 and retrieval.stats()['docs'] == 2,
      f'st={st3} docs={retrieval.stats()["docs"]}')


# ─── 6. build_context 的注入规则 ─────────────────────────────────────
ctx = retrieval.build_context('霍尔电压 螺线管 磁场', exp_type='物理实验', k=3, max_chars=300)
check('context：非空且带来源标记', ctx.startswith('【参考1｜'), ctx[:40])
check('context：不超过字数上限', len(ctx) <= 300 + 20, f'len={len(ctx)}')
check('context：无命中时返回空串', retrieval.build_context('量子色动力学 胶子禁闭') == '')

# 模板段落（实验目的/仪器）不进 context：那两栏生成时本来就照抄指导书
write_md(os.path.join('范文', '模板段落实验.md'),
         '# 某实验\n\n## 实验目的\n\n1. 了解实验箱的使用方法。2. 掌握基本的测量手段。\n')
retrieval.ingest()
ctx2 = retrieval.build_context('实验箱 使用方法 测量手段', k=3, max_chars=600)
check('context：剔除实验目的等模板段落', '实验目的' not in ctx2, ctx2[:80])

# 草稿默认不参与检索，KB_DRAFTS=1 才放行
write_md(os.path.join('草稿', '旧稿.md'),
         '# 旧稿\n\n## 实验原理\n\n这段文字只出现在草稿里，用于验证默认不注入：'
         '霍尔效应旧稿版本说明。\n')
retrieval.ingest()
default_ctx = retrieval.build_context('霍尔效应旧稿版本说明 草稿', k=3, max_chars=600)
check('context：草稿默认被排除', '旧稿' not in default_ctx, default_ctx[:80])
os.environ['KB_DRAFTS'] = '1'
draft_ctx = retrieval.build_context('霍尔效应旧稿版本说明 草稿', k=3, max_chars=600)
check('context：KB_DRAFTS=1 时纳入草稿', '旧稿' in draft_ctx, draft_ctx[:80])
os.environ.pop('KB_DRAFTS', None)


# ─── 7. 清理 ─────────────────────────────────────────────────────────
shutil.rmtree(_TMP, ignore_errors=True)

print('\n' + ('全部通过' if not _failures else f'{len(_failures)} 项失败：{_failures}'))
sys.exit(1 if _failures else 0)

"""
实验报告库整理器 · scripts/tidy_drill.py
==========================================
把"随手丢的文件"整理成可检索、可复用的报告库。

目录约定（脚本会自动建）：

    drill/
    ├── 00_收件箱/      ← 新文件先丢这里（或直接丢在 drill 根目录）
    ├── 01_范文/        ← 已整理的过关报告，一个实验一个文件夹
    │   └── 负反馈放大器/
    │       ├── 负反馈放大器.docx
    │       └── 负反馈放大器.pdf
    ├── 02_模板/        ← 老师给的空白模板
    ├── 03_指导书/      ← 课程指导书
    ├── 99_备份/        ← 重复文件、旧版本，不删只挪
    ├── index.md        ← 自动生成：有哪些实验、各有哪些文件
    └── 整理记录.md      ← 每次整理的 旧名 → 新位置 对照，可回溯

用法：

    python scripts/tidy_drill.py --root "C:\\Users\\X2373\\Desktop\\drill" --dry-run
    python scripts/tidy_drill.py --root "C:\\Users\\X2373\\Desktop\\drill"

文件名规则：自动去掉「+周曦睿+电信24-1+202310715214(1)」这类学号/姓名/班级/版本尾巴，
只留实验名；同一实验的 docx 与 pdf 放同一文件夹。
"""
from __future__ import annotations

import argparse
import hashlib
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

SUPPORTED = {".docx", ".doc", ".pdf", ".md", ".txt"}

INBOX = "00_收件箱"
REPORTS = "01_范文"
TEMPLATES = "02_模板"
GUIDES = "03_指导书"
BACKUP = "99_备份"

# 名字里要丢掉的片段：学号、班级、版本号、括号序号
DROP_PATTERNS = [
    re.compile(r"^\d{6,}$"),                       # 202310715214
    re.compile(r"^电信\d+-\d+$"),                   # 电信24-1
    re.compile(r"^(?:通信|电科|自动化)\d+-\d+$"),
    re.compile(r"^v\d+(?:\.\d+)*$", re.I),          # v0.1
    re.compile(r"^第?\d+版$"),
    re.compile(r"^\(\d+\)$"),
    re.compile(r"^[（(]\d+[)）]$"),
]

# 常见的实验名错别字 → 规范写法
TYPO_FIX = {"射级跟随器": "射极跟随器", "单极晶体管": "单级晶体管"}


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def experiment_name(stem: str) -> str:
    """从文件名里还原"实验名"。"""
    parts = [p.strip() for p in re.split(r"[+＋]", stem) if p.strip()]
    keep = []
    for p in parts:
        if any(pat.match(p) for pat in DROP_PATTERNS):
            continue
        if re.fullmatch(r"[\u4e00-\u9fff]{2,4}", p) and p not in keep and len(keep) >= 1:
            # 「+周曦睿+」这类 2~4 字的纯中文段，多半是姓名：只在已经拿到实验名后丢弃
            continue
        keep.append(p)
    name = keep[0] if keep else stem
    name = re.sub(r"[（(]\d+[)）]\s*$", "", name).strip()
    name = re.sub(r"\s+", "", name)
    for bad, good in TYPO_FIX.items():
        name = name.replace(bad, good)
    return name or stem


def classify(path: Path) -> str:
    """返回 reports / templates / guides"""
    n = path.name
    if "模板" in n:
        return TEMPLATES
    if re.search(r"指导书|实验指导|手册|讲义", n):
        return GUIDES
    return REPORTS


def unique_path(target: Path, planned: set | None = None) -> Path:
    """避免同名覆盖：既看磁盘上是否已存在，也看本次计划里是否已占用。"""
    planned = planned if planned is not None else set()

    def taken(p: Path) -> bool:
        return p.exists() or p in planned

    if not taken(target):
        return target
    stem, suf = target.stem, target.suffix
    i = 2
    while True:
        cand = target.with_name(f"{stem}_{i}{suf}")
        if not taken(cand):
            return cand
        i += 1


def is_raw(stem: str) -> bool:
    """文件名里带学号/班级的，是原始提交稿，保留但标 _原始。"""
    parts = [p.strip() for p in re.split(r"[+＋]", stem) if p.strip()]
    for p in parts:
        if re.fullmatch(r"\d{6,}", p):
            return True
        if re.fullmatch(r"(?:电信|通信|电科|自动化)\d+-\d+", p):
            return True
    return False


SKIP_NAMES = {"README.md", "index.md", "整理记录.md"}


def collect(root: Path, inbox: Path) -> list[Path]:
    files = []
    for base in (root, inbox):
        if not base.exists():
            continue
        for p in sorted(base.iterdir()):
            if not p.is_file() or p.suffix.lower() not in SUPPORTED:
                continue
            if p.name in SKIP_NAMES or p.name.startswith("~$"):
                continue
            files.append(p)
    return files


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="整理实验报告文件夹")
    ap.add_argument("--root", default=r"C:\Users\X2373\Desktop\drill")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划，不动文件")
    args = ap.parse_args(argv)

    root = Path(args.root)
    if not root.exists():
        print(f"目录不存在：{root}")
        return 1
    dirs = {k: root / k for k in (INBOX, REPORTS, TEMPLATES, GUIDES, BACKUP)}
    if not args.dry_run:
        for d in dirs.values():
            d.mkdir(exist_ok=True)
        (dirs[BACKUP] / "重复").mkdir(exist_ok=True)

    files = collect(root, dirs[INBOX])
    if not files:
        print("没有待整理的文件（根目录和收件箱都是空的）")
        return 0

    seen: dict[str, Path] = {}       # 内容哈希 → 已放置的文件
    planned: set = set()             # 本次计划已占用的目标路径（防同名覆盖）
    moves: list[tuple[Path, Path]] = []
    dups: list[tuple[Path, Path]] = []

    for f in files:
        kind = classify(f)
        digest = sha256_of(f)
        if digest in seen:
            dups.append((f, seen[digest]))
            continue

        name = experiment_name(f.stem)
        raw = is_raw(f.stem)
        if kind == REPORTS:
            target_dir = dirs[REPORTS] / name
            suffix = "_原始" if raw else ""
            target = target_dir / f"{name}{suffix}{f.suffix.lower()}"
        elif kind == TEMPLATES:
            target = dirs[TEMPLATES] / f.name
        else:
            target = dirs[GUIDES] / f.name

        target = unique_path(target, planned)
        planned.add(target)
        seen[digest] = target
        moves.append((f, target))

    for src, dst in moves + dups:
        tag = "重复" if (src, dst) in dups else "移动"
        print(f"[{tag}] {src.name}\n        → {dst.relative_to(root)}")

    if args.dry_run:
        print(f"\n[dry-run] 计划移动 {len(moves)} 个，重复 {len(dups)} 个；加 --dry-run 去掉即执行")
        return 0

    log = [f"# 整理记录 {datetime.now():%Y-%m-%d %H:%M}\n"]
    for src, dst in moves:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src), str(dst))
        log.append(f"- `{src.name}` → `{dst.relative_to(root)}`")
    for src, original in dups:
        dst = unique_path(dirs[BACKUP] / "重复" / src.name)
        shutil.move(str(src), str(dst))
        log.append(f"- `{src.name}` → `{dst.relative_to(root)}`（与 {original.name} 内容相同）")

    with (root / "整理记录.md").open("a", encoding="utf-8") as f:
        f.write("\n".join(log) + "\n\n")

    # ── 生成索引 ──
    rows = []
    for exp_dir in sorted(p for p in dirs[REPORTS].iterdir() if p.is_dir()):
        items = sorted(exp_dir.iterdir())
        docx = [i for i in items if i.suffix.lower() in (".docx", ".doc")]
        pdf = [i for i in items if i.suffix.lower() == ".pdf"]
        size = sum(i.stat().st_size for i in items) / 1024
        latest = max(i.stat().st_mtime for i in items)
        rows.append((exp_dir.name, len(docx), len(pdf), size, latest, items))

    lines = ["# 实验报告库索引", "",
             f"共 {len(rows)} 个实验，更新于 {datetime.now():%Y-%m-%d %H:%M}", "",
             "| 实验 | docx | pdf | 大小 | 最后修改 |", "|---|---|---|---|---|"]
    for name, nd, nf, size, latest, _items in rows:
        lines.append(f"| [{name}]({REPORTS}/{name}/) | {nd} | {nf} | {size:.0f} KB | "
                     f"{datetime.fromtimestamp(latest):%Y-%m-%d} |")
    (root / "index.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"\n完成：移动 {len(moves)} 个，重复归档 {len(dups)} 个，索引已更新")
    return 0


if __name__ == "__main__":
    sys.exit(main())

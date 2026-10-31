"""
公开仓库前的个人信息清理 · scripts/clean_pii.py
================================================
把 kb/ 里所有 .docx 的姓名、学号、教师姓名替换成占位符，并在替换后逐部件复核。

为什么需要它：这个仓库是公开的，而实验报告封面本来就印着姓名学号；
只改正文是不够的——docProps 里的作者、页眉页脚、表格单元格都可能是独立部件。

用法：

    python scripts/clean_pii.py scan                      # 只体检，不改文件
    python scripts/clean_pii.py scrub 周曦睿 邓可 --id 202310715214
    python scripts/clean_pii.py scrub 段黎 --teacher      # → 【教师】

替换规则：姓名 → 【姓名】，--id 的值 → 【学号】，--teacher 的值 → 【教师】。
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
KB = Path(__import__("os").getenv("KB_DIR") or (ROOT / "kb"))

# 体检用的模式：手机号、邮箱、11~12 位连续数字（学号）
PATTERNS = {
    "手机号": re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)"),
    "邮箱": re.compile(r"[\w.+-]+@[\w-]+\.[\w.]{2,}"),
    "学号(11~12位)": re.compile(r"(?<!\d)\d{11,12}(?!\d)"),
}
XML_SUFFIX = (".xml", ".rels")


def _plausible_id(value: str) -> bool:
    """真值表里连着的 000000000000、100000000000 不是学号：要求至少 4 个不同数字。"""
    return len(set(value)) >= 4


def _docx_files() -> list[Path]:
    return sorted(p for p in KB.rglob("*.docx") if not p.name.startswith("~$"))


def _parts(path: Path):
    with zipfile.ZipFile(path) as z:
        return [(i, z.read(i.filename)) for i in z.infolist()]


def _write(path: Path, entries) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in entries:
            z.writestr(info, data)


def _replace_tree(root, old: str, new: str) -> int:
    hits = 0
    for el in root.iter():
        if el.text and old in el.text:
            el.text = el.text.replace(old, new)
            hits += 1
        if el.tail and old in el.tail:
            el.tail = el.tail.replace(old, new)
            hits += 1
    return hits


def scrub(path: Path, pairs: list[tuple[str, str]]) -> int:
    entries = _parts(path)
    hits = 0
    out = []
    for info, data in entries:
        if info.filename.lower().endswith(XML_SUFFIX):
            try:
                root = ET.fromstring(data)
            except ET.ParseError:
                out.append((info, data))
                continue
            n = 0
            for old, new in pairs:
                n += _replace_tree(root, old, new)
            if n:
                data = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
                hits += n
        out.append((info, data))
    if hits:
        _write(path, out)
    return hits


def scan(path: Path, pairs: list[tuple[str, str]]) -> tuple[list[str], int]:
    """返回 (问题列表, 无法审计的二进制部件数)。

    只对 XML 部件做文本扫描：图片按 utf-8 解码是乱码，正则会把随机字节当成邮箱。
    图片里的文字（截图上的水印/姓名）没法在这里查，只能肉眼过一遍。
    """
    issues = []
    binary = 0
    for old, _new in pairs:
        if old in path.name:
            issues.append(f"文件名含「{old}」")
    for info, data in _parts(path):
        name = info.filename
        if not name.lower().endswith(XML_SUFFIX):
            binary += 1
            continue
        blob = data.decode("utf-8", errors="ignore")
        for old, _new in pairs:
            if old in blob:
                issues.append(f"{name} 含「{old}」")
        for label, pat in PATTERNS.items():
            for m in sorted(set(pat.findall(blob))):
                if label.startswith("学号") and not _plausible_id(m):
                    continue
                issues.append(f"{name} 疑似{label}：{m}")
    return issues, binary


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="kb/ 个人信息体检与清理")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_scan = sub.add_parser("scan", help="只体检，不改文件")
    p_scan.add_argument("names", nargs="*", help="已知的姓名（可选）")

    p_scr = sub.add_parser("scrub", help="替换并复核")
    p_scr.add_argument("names", nargs="*", help="要替换的姓名 → 【姓名】")
    p_scr.add_argument("--id", action="append", default=[], help="学号 → 【学号】")
    p_scr.add_argument("--teacher", action="append", default=[], help="教师姓名 → 【教师】")

    args = ap.parse_args(argv)
    files = _docx_files()
    if not files:
        print(f"没找到 docx：{KB}")
        return 1

    if args.cmd == "scan":
        pairs = [(n, "") for n in args.names]
        bad = 0
        media = 0
        for p in files:
            issues, binary = scan(p, pairs)
            media += binary
            if issues:
                bad += 1
                print(f"\n{p.relative_to(KB)}")
                for it in issues:
                    print("   ", it)
        print(f"\n扫描 {len(files)} 份，{bad} 份有需要处理的内容"
              + ("" if bad else "（文本干净）"))
        print(f"另有 {media} 个图片/媒体部件未做文本审计——截图里若出现姓名学号，需肉眼过一遍。")
        return 1 if bad else 0

    pairs = ([(n, "【姓名】") for n in args.names]
             + [(i, "【学号】") for i in args.id]
             + [(t, "【教师】") for t in args.teacher])
    if not pairs:
        print("没有要给替换的词。示例：python scripts/clean_pii.py scrub 张三 --id 20230001")
        return 2

    total = 0
    for p in files:
        n = scrub(p, pairs)
        total += n
        if n:
            print(f"[清理] {p.relative_to(KB)}  替换 {n} 处")

    bad = 0
    for p in files:
        issues, _binary = scan(p, pairs)
        if issues:
            bad += 1
            print(f"\n[残留] {p.relative_to(KB)}")
            for it in issues:
                print("   ", it)

    print(f"\n共替换 {total} 处；复核 {len(files)} 份，仍有问题 {bad} 份")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())

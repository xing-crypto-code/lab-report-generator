"""
重新烘焙内置模板 · scripts/rebake.py
====================================
把 modules/template_fix.py 里登记的全部修复应用到内置模板文件本身。

用法：
    python scripts/rebake.py
    python scripts/rebake.py --force-backup    # 覆盖已有 .bak

修复登记在 modules/template_fix.py 的 TEMPLATE_FIXES 列表里；
以后每修一个排版 bug，就往那里追加一条，再跑本脚本。
"""
import os
import sys
import argparse
import warnings

warnings.filterwarnings('ignore')
for _k in ('HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy',
           'ALL_PROXY', 'all_proxy', 'NO_PROXY', 'no_proxy'):
    os.environ.pop(_k, None)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, 'libs'))

from modules.template_fix import (          # noqa: E402
    rebake_template, verify_template, TEMPLATE_FIXES,
)

TEMPLATE = os.path.join(ROOT, 'templates', '内置实验报告模板.docx')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--force-backup', action='store_true',
                    help='覆盖已有 .bak 备份')
    args = ap.parse_args()

    if not os.path.exists(TEMPLATE):
        print(f'✗ 模板不存在: {TEMPLATE}')
        return 1

    # 只读属性会让 save 失败
    try:
        os.chmod(TEMPLATE, 0o666)
    except Exception:
        pass
    bak = TEMPLATE + '.bak'
    if args.force_backup and os.path.exists(bak):
        os.remove(bak)
        print('已清除旧备份')

    before = os.path.getsize(TEMPLATE)
    print(f'烘焙: {os.path.basename(TEMPLATE)}  ({before} bytes)')
    print(f'已登记修复: {len(TEMPLATE_FIXES)} 条\n')

    report = rebake_template(TEMPLATE)
    for name, desc, _fn in TEMPLATE_FIXES:
        val = report.get(name)
        flag = '✓' if isinstance(val, int) else '✗'
        print(f'  {flag} {name:36s} 改动 {val} 处   （{desc}）')

    problems = verify_template(TEMPLATE)
    print()
    if problems:
        print(f'✗ 校验未通过，仍有 {len(problems)} 个问题:')
        for p in problems:
            print('   -', p)
        return 1
    after = os.path.getsize(TEMPLATE)
    print(f'✓ 校验全部通过     {before} → {after} bytes')
    print(f'备份: {os.path.basename(bak)}（若存在）')
    print('\n下一步：跑回归测试  python tests/test_template.py')
    return 0


if __name__ == '__main__':
    sys.exit(main())

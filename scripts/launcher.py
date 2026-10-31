"""无痛实验报告生成器 · 一键启动
================================
双击 start.bat 即可：自检依赖 → 清理代理环境变量 → 启动 Gradio → 自动打开浏览器。

关键点：
- 依赖装在 libs/（本机 pip --target 安装），脚本优先用 libs/ 里的包（4.x gradio 等）。
- app.py 同款地清掉 HTTP(S)_PROXY / NO_PROXY（其中 [::1] 语法会让 gradio 在
  import 阶段直接崩 Invalid port: ':1]'）。
- 让解释器优先引用本仓库 libs/，避免误用系统里不兼容的旧包。
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
APP = ROOT / "app.py"
LIBS = ROOT / "libs"

# 优先用本仓库 libs/（同 app.py 的一致方式）
if str(LIBS) not in sys.path:
    sys.path.insert(0, str(LIBS))

# 同 app.py：清理代理环境变量（NO_PROXY 里的 [::1] 会让 gradio 崩）
for _k in (
    "HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy",
    "ALL_PROXY", "all_proxy", "NO_PROXY", "no_proxy",
):
    os.environ.pop(_k, None)

SERVER_HOST = "0.0.0.0"
SERVER_PORT = 7860
URL = f"http://127.0.0.1:{SERVER_PORT}"


def _wait_port(port: int, timeout: float = 60.0) -> bool:
    import socket
    deadline = time.time() + timeout
    while time.time() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.4)
            if s.connect_ex(("127.0.0.1", port)) == 0:
                return True
        time.sleep(0.5)
    return False


def check_deps() -> list[str]:
    """返回缺失的关键依赖名（可导入才认为 OK）。"""
    missing = []
    checks = [
        ("gradio", "gradio"),
        ("PyMuPDF", "fitz"),
        ("python-docx", "docx"),
        ("openai", "openai"),
        ("Pillow", "PIL"),
        ("python-dotenv", "dotenv"),
        ("easyocr", "easyocr"),
        ("torch", "torch"),
    ]
    for name, mod in checks:
        try:
            __import__(mod)
        except Exception:
            missing.append(name)
    return missing


def main() -> int:
    print("=" * 56)
    print("  无痛实验报告生成器 · 一键启动")
    print("=" * 56)
    if not APP.exists():
        print(f"[错误] 找不到 {APP}")
        return 1
    if not LIBS.exists():
        print(f"[提示] 未找到 libs/（{LIBS}）")
        print("       首次运行请先执行：python install_deps.py")
        return 1

    missing = check_deps()
    if missing:
        print(f"[提示] 检测到缺失依赖：{', '.join(missing)}")
        print("       尝试用本仓库 libs/ 加载；若失败，请重装依赖：")
        print("       python install_deps.py")
        print()

    if _wait_port(SERVER_PORT, timeout=1.2):
        print(f"[提示] 端口 {SERVER_PORT} 已有一个服务在跑，直接打开浏览器。")
        webbrowser.open(URL)
        return 0

    env = dict(os.environ)
    for _k in (
        "HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy",
        "ALL_PROXY", "all_proxy", "NO_PROXY", "no_proxy",
    ):
        env.pop(_k, None)

    print("[启动] 正在启动服务器…（首次稍慢，请稍候）")
    try:
        proc = subprocess.Popen(
            [sys.executable, str(APP)],
            cwd=str(ROOT),
            env=env,
        )
    except Exception as e:
        print(f"[错误] 启动失败：{e}")
        return 1

    if _wait_port(SERVER_PORT, timeout=90):
        print(f"✅ 服务已就绪：{URL}")
        print("   浏览器将自动打开；如果没弹，手动访问上面的地址。")
        webbrowser.open(URL)
        return 0
    else:
        print("[提示] 90 秒内未检测到端口就绪，请检查上方日志。")
        return 1


if __name__ == "__main__":
    sys.exit(main())
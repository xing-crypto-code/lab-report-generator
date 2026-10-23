"""小工具：在沙箱环境下安全安装依赖"""
import subprocess, sys, os

# 全部使用 8.3 短路径，避免中文编码问题
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
LIBS = os.path.join(WORKSPACE, "libs")
TEMP = os.path.join(WORKSPACE, "temp")

os.makedirs(LIBS, exist_ok=True)
os.makedirs(TEMP, exist_ok=True)

env = dict(os.environ)
env["TMPDIR"] = TEMP
env["TEMP"] = TEMP
env["TMP"] = TEMP
env["PYTHONUNBUFFERED"] = "1"

packages = [
    "gradio",
    "PyMuPDF",
    "pdfplumber",
    "Pillow",
    "python-docx",
    "docx2pdf",
    "reportlab",
    "openai",
    "python-dotenv",
    "lxml",
]

print(f"[install] Workspace: {WORKSPACE}")
print(f"[install] Libs: {LIBS}")
print(f"[install] Temp: {TEMP}")

args = [
    sys.executable, "-m", "pip", "install",
    "--target", LIBS,
    "--no-cache-dir",
    "--no-warn-script-location",
    *packages
]

print(f"[install] Running: {' '.join(args)}")
result = subprocess.run(args, env=env, cwd=TEMP)
sys.exit(result.returncode)
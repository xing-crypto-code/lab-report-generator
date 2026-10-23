"""
无痛实验报告生成器 · 配置中心
"""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ---- 项目路径 ----
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
TEMPLATE_DIR = BASE_DIR / "templates"
LIBS_DIR = BASE_DIR / "libs"

for d in [UPLOAD_DIR, OUTPUT_DIR, TEMPLATE_DIR]:
    d.mkdir(exist_ok=True)

# ---- 把本地 libs 加入 sys.path（免 venv 安装方案） ----
import sys
_lib = str(LIBS_DIR)
if _lib not in sys.path:
    sys.path.insert(0, _lib)

# ---- AI API 配置 ----
# 支持 OpenAI 兼容接口（DeepSeek / Qwen / GPT / Ollama 等）
# 请在 .env 文件或系统环境变量中设置
LLM_API_KEY = os.getenv("LLM_API_KEY", "")
LLM_BASE_URL = os.getenv("LLM_BASE_URL", "https://api.deepseek.com")
LLM_MODEL = os.getenv("LLM_MODEL", "deepseek-chat")

# ---- OCR 配置 ----
# EasyOCR: ch_sim=简体中文, en=英文
OCR_LANG = os.getenv("OCR_LANG", "ch")
OCR_USE_GPU = os.getenv("OCR_USE_GPU", "false").lower() in ("1", "true", "yes")

# ---- 生成配置 ----
MAX_RETRY = 3
TEMP_DIR = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)
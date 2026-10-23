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

for d in [UPLOAD_DIR, OUTPUT_DIR, TEMPLATE_DIR]:
    d.mkdir(exist_ok=True)

# ---- AI API 配置 ----
# 支持 OpenAI 兼容接口（DeepSeek / Qwen / GPT 等）
# 请在 .env 文件或系统环境变量中设置
LLM_API_KEY = os.getenv("LLM_API_KEY", "sk-your-api-key")
LLM_BASE_URL = os.getenv("LLM_BASE_URL", "https://api.deepseek.com")
LLM_MODEL = os.getenv("LLM_MODEL", "deepseek-chat")

# ---- OCR 配置 ----
OCR_LANG = "ch"           # PaddleOCR 语言
OCR_USE_GPU = False       # 是否启用 GPU（默认 CPU）
TABLE_RECOGNITION = True  # 开启表格识别

# ---- 生成配置 ----
MAX_RETRY = 3             # API 调用重试次数
TEMP_DIR = BASE_DIR / "temp"  # 临时文件目录
TEMP_DIR.mkdir(exist_ok=True)
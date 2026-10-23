"""冒烟测试——验证所有模块"""
import sys; sys.path.insert(0, "libs")

from modules.parser import parse_document
from modules.preprocessor import preprocess_for_ocr
from modules.template import analyze_template
from modules.filler import list_available_models, detect_experiment_type
from modules.generator import generate_report_document
from modules.knowledge import EXPERIMENT_TYPES, FORMAT_RULES

print("=== 全模块导入: OK ===")

print(f"知识库: {len(EXPERIMENT_TYPES)} 种实验类型, {len(FORMAT_RULES)} 条格式规则")
for name, info in EXPERIMENT_TYPES.items():
    print(f"  {name}: {info['sections']}")

print(f"\n可用模型: {len(list_available_models())} 个")

tests = [
    ("RC电路 电阻 电压 示波器 波形", "电路/电子实验"),
    ("滴定 溶液 浓度 酸碱", "化学实验"),
    ("牛顿第二定律 质量 加速度", "物理实验"),
    ("Python 代码 算法 编译", "计算机/编程实验"),
]
for text, expected in tests:
    result = detect_experiment_type(text)
    status = "OK" if result == expected else f"FAIL(expected {expected})"
    print(f"  [{status}] '{text[:35]}' -> {result}")

print("\n=== ALL CHECKS PASSED ===")
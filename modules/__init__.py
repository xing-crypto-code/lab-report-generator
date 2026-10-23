# modules 包初始化
from .parser import parse_document
from .template import analyze_template, get_template_summary
from .filler import generate_report, compare_models, list_available_models, detect_experiment_type
from .generator import generate_report_document
from .knowledge import build_standard_template, EXPERIMENT_TYPES
from .preprocessor import preprocess_for_ocr
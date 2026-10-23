"""
文档解析模块 · parser.py
职责：识别用户上传的实验指导手册（PDF / 图片），
      提取其中的文字、表格、实验步骤、数据等信息。

OCR 引擎：EasyOCR（支持 80+ 语言，中文识别优秀，Python 3.14 兼容）
"""
import fitz  # PyMuPDF
from PIL import Image
from docx import Document
import io
from pathlib import Path
from typing import Optional

from config import OCR_LANG, OCR_USE_GPU
from modules.preprocessor import preprocess_for_ocr, detect_document_region

def _detect_cuda() -> bool:
    """
    自动检测当前 torch 是否支持 CUDA（GPU 加速）。
    只有装了 CUDA 版 torch 才返回 True，否则强制 CPU，
    避免 EasyOCR 在无 CUDA 时报错。
    """
    try:
        import torch
        return bool(torch.cuda.is_available())
    except Exception:
        return False


# GPU 可用性自动检测（覆盖 .env 的 OCR_USE_GPU）
_CUDA_OK = _detect_cuda()
if OCR_USE_GPU and not _CUDA_OK:
    print("[parser] 检测到 OCR_USE_GPU=true，但当前 torch 为 CPU 版本，"
          f"已自动降级为 CPU 模式。可尝试安装 CUDA 版 torch 使用 GPU 加速。")

# EasyOCR 延迟加载（首次调用时初始化，避免启动卡顿）
_reader = None

def _get_reader():
    """单例模式获取 EasyOCR Reader 实例"""
    global _reader
    if _reader is None:
        import easyocr
        # EasyOCR 语言代码: ch_sim=简体中文, en=英文
        lang_list = ["ch_sim", "en"] if OCR_LANG == "ch" else ["en"]
        use_gpu = OCR_USE_GPU and _CUDA_OK
        _reader = easyocr.Reader(
            lang_list=lang_list,
            gpu=use_gpu,
            verbose=False
        )
    return _reader


def parse_pdf(file_path: str) -> dict:
    """
    解析 PDF 文件，返回：
    {
        "text": "整段文字内容",
        "tables": [["表头", "列2"], ["数据1", "数据2"]],
        "page_count": 页数
    }
    """
    doc = fitz.open(file_path)
    full_text = []
    all_tables = []
    page_count = len(doc)

    for page_num in range(page_count):
        page = doc[page_num]

        # 1. 提取文字（电子版 PDF）
        text = page.get_text("text")
        if text.strip():
            full_text.append(f"--- 第 {page_num+1} 页 ---\n{text}")

        # 2. 提取图片 → OCR 识别（处理扫描版 PDF / 照片）
        images = page.get_images(full=True)
        for img_index, img in enumerate(images):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            pil_image = Image.open(io.BytesIO(image_bytes))

            ocr_result = ocr_image(pil_image)
            if ocr_result["text"]:
                full_text.append(
                    f"--- 第 {page_num+1} 页 · 图片 {img_index+1} ---\n"
                    f"{ocr_result['text']}"
                )
            all_tables.extend(ocr_result["tables"])

    doc.close()
    return {
        "text": "\n".join(full_text),
        "tables": all_tables,
        "page_count": page_count
    }


def ocr_image(image: Image.Image) -> dict:
    """
    对单张图片做 OCR 识别（含图像预处理增强）
    返回 { "text": "识别文字", "tables": [表格列表] }
    """
    import numpy as np

    # ── 预处理 ──
    processed = preprocess_for_ocr(image)

    # PIL → OpenCV 格式（RGB → BGR）
    img_array = np.array(processed.convert("RGB"))
    img_bgr = img_array[:, :, ::-1].copy()

    reader = _get_reader()
    results = reader.readtext(img_bgr, detail=1)

    texts = []
    tables = []

    for item in results:
        bbox, text, confidence = item
        if confidence > 0.3:
            texts.append(text)

    return {
        "text": "\n".join(texts),
        "tables": tables
    }


def parse_image(file_path: str) -> dict:
    """解析单张图片文件（照片/截图）"""
    image = Image.open(file_path)
    return ocr_image(image)


def parse_docx(file_path: str) -> dict:
    """
    解析 .docx/.doc Word 文档的实验指导手册
    （有些老师发的是 Word 版指导书）
    """
    # .doc（老格式）尝试用 python-docx，失败则报友好错误
    ext = Path(file_path).suffix.lower()
    if ext == ".doc":
        try:
            doc = Document(file_path)
        except Exception:
            raise ValueError(
                "暂不支持 .doc 旧格式 Word 文档。"
                "请用 WPS/Word 打开后「另存为」→ 选择 .docx 格式再上传。"
            )
    else:
        doc = Document(file_path)

    full_text = []
    all_tables = []

    # 提取段落文本
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text:
            full_text.append(text)

    # 提取表格
    for table in doc.tables:
        rows_data = []
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            rows_data.append(cells)
        if rows_data:
            all_tables.append(rows_data)

    return {
        "text": "\n".join(full_text),
        "tables": all_tables,
        "page_count": None  # docx 没有明确页数概念
    }


def parse_document(file_path: str) -> dict:
    """统一入口：自动判断格式，返回统一结构"""
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":
        return parse_pdf(file_path)
    elif ext in {".docx", ".doc"}:
        return parse_docx(file_path)
    elif ext in {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp"}:
        return parse_image(file_path)
    else:
        raise ValueError(f"不支持的文件格式: {ext}")


# ---------- 便捷测试 ----------
if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        result = parse_document(sys.argv[1])
        print(f"页数: {result.get('page_count', 'N/A')}")
        print(f"文字长度: {len(result['text'])} 字符")
        print(f"表格数: {len(result['tables'])}")
        print("--- 前 500 字预览 ---")
        print(result["text"][:500])
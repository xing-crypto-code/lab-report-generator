"""
文档解析模块 · parser.py
职责：识别用户上传的实验指导手册（PDF / 图片），
      提取其中的文字、表格、实验步骤、数据等信息。
"""
import fitz  # PyMuPDF
from PIL import Image
import io
from pathlib import Path
from typing import Optional

from config import OCR_LANG, OCR_USE_GPU, TABLE_RECOGNITION

# PaddleOCR 延迟加载（首次调用时初始化，避免启动卡顿）
_ocr = None

def _get_ocr():
    """单例模式获取 PaddleOCR 实例"""
    global _ocr
    if _ocr is None:
        from paddleocr import PaddleOCR
        _ocr = PaddleOCR(
            lang=OCR_LANG,
            use_gpu=OCR_USE_GPU,
            show_log=False,
            enable_table=TABLE_RECOGNITION
        )
    return _ocr


def parse_pdf(file_path: str) -> dict:
    """
    解析 PDF 文件，返回：
    {
        "text": "整段文字内容",
        "tables": [["表头", "列2"], ["数据1", "数据2"]],
        "images": ["图片base64或路径"],
        "page_count": 页数
    }
    """
    doc = fitz.open(file_path)
    full_text = []
    all_tables = []
    page_count = len(doc)

    for page_num in range(page_count):
        page = doc[page_num]

        # 1. 提取文字
        text = page.get_text("text")
        if text.strip():
            full_text.append(f"--- 第 {page_num+1} 页 ---\n{text}")

        # 2. 提取图片 → OCR 识别（处理扫描版 PDF）
        images = page.get_images(full=True)
        for img_index, img in enumerate(images):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            pil_image = Image.open(io.BytesIO(image_bytes))

            # OCR 识别图片中的文字和表格
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
    对单张图片做 OCR（文字 + 表格识别）
    返回 { "text": "识别文字", "tables": [表格列表] }
    """
    ocr = _get_ocr()
    result = ocr.ocr(image)

    texts = []
    tables = []

    if not result or not result[0]:
        return {"text": "", "tables": []}

    for item in result[0]:
        if item is None:
            continue
        # item 格式: [坐标, (文字, 置信度)]
        if len(item) >= 2:
            text = item[1][0] if isinstance(item[1], (tuple, list)) else str(item[1])
            texts.append(text)

    return {
        "text": "\n".join(texts),
        "tables": tables
    }


def parse_image(file_path: str) -> dict:
    """
    解析单张图片文件（照片/截图）
    """
    image = Image.open(file_path)
    return ocr_image(image)


def parse_document(file_path: str) -> dict:
    """
    统一入口：自动判断 PDF 还是图片，返回统一结构
    """
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":
        return parse_pdf(file_path)
    elif ext in {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp"}:
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
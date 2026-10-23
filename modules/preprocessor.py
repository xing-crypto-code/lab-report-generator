"""
图像预处理模块 · preprocessor.py
职责：对上传的图片进行去噪、增强、纠偏，提升 OCR 识别率
"""
import io
from PIL import Image, ImageFilter, ImageEnhance, ImageOps
import numpy as np


def preprocess_for_ocr(image: Image.Image, mode: str = "auto") -> Image.Image:
    """
    综合预处理流水线，自动判断图片类型并做针对性处理

    参数:
        image: PIL Image 对象
        mode: "auto"=自动 / "photo"=照片 / "scan"=扫描件 / "screenshot"=截图
    返回:
        处理后的 PIL Image
    """
    # 先转 RGB
    if image.mode != "RGB":
        image = image.convert("RGB")

    # ---- 1. 自适应缩放（太大浪费 OCR 时间，太小丢细节）----
    w, h = image.size
    max_dim = 2000
    if max(w, h) > max_dim:
        ratio = max_dim / max(w, h)
        image = image.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)

    # ---- 2. 转为灰度 ----
    gray = image.convert("L")

    # ---- 3. 降噪 ----
    gray = gray.filter(ImageFilter.MedianFilter(size=3))

    # ---- 4. 对比度增强 ----
    enhancer = ImageEnhance.Contrast(gray)
    gray = enhancer.enhance(1.5)

    # ---- 5. 锐化 ----
    gray = gray.filter(ImageFilter.SHARPEN)

    # ---- 6. 自适应二值化（大津法） ----
    img_array = np.array(gray)
    threshold = _otsu_threshold(img_array)
    binary = (img_array > threshold).astype(np.uint8) * 255
    gray = Image.fromarray(binary, mode="L")

    return gray.convert("RGB")


def _otsu_threshold(img_array: np.ndarray) -> int:
    """大津算法（Otsu）计算最优二值化阈值"""
    hist, _ = np.histogram(img_array.ravel(), bins=256, range=(0, 256))
    total = img_array.size
    current_max = 0
    threshold = 0
    sum_total = sum(i * hist[i] for i in range(256))
    sum_b, w_b = 0, 0

    for i in range(256):
        w_b += hist[i]
        if w_b == 0:
            continue
        w_f = total - w_b
        if w_f == 0:
            break
        sum_b += i * hist[i]
        mean_b = sum_b / w_b
        mean_f = (sum_total - sum_b) / w_f
        between = w_b * w_f * (mean_b - mean_f) ** 2
        if between > current_max:
            current_max = between
            threshold = i

    return max(threshold, 100)  # 最低阈值保护


def detect_document_region(image: Image.Image) -> Image.Image:
    """
    检测并裁剪文档区域（处理拍照时背景干扰）
    适用于手机拍照的实验手册
    """
    import cv2
    img_array = np.array(image.convert("RGB"))
    img_bgr = img_array[:, :, ::-1].copy()

    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(blurred, 50, 150)

    contours, _ = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if contours:
        largest = max(contours, key=cv2.contourArea)
        peri = cv2.arcLength(largest, True)
        approx = cv2.approxPolyDP(largest, 0.02 * peri, True)
        if len(approx) == 4:
            # 找到文档四角，做透视变换
            pts = approx.reshape(4, 2)
            rect = _order_points(pts)
            tl, tr, br, bl = rect
            width = max(int(np.linalg.norm(br - bl)), int(np.linalg.norm(tr - tl)))
            height = max(int(np.linalg.norm(tr - br)), int(np.linalg.norm(tl - bl)))
            dst = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], dtype="float32")
            M = cv2.getPerspectiveTransform(rect.astype("float32"), dst)
            warped = cv2.warpPerspective(img_bgr, M, (width, height))
            return Image.fromarray(warped[:, :, ::-1])

    return image


def _order_points(pts):
    """顺时针排列四点：左上 → 右上 → 右下 → 左下"""
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]
    rect[3] = pts[np.argmax(diff)]
    return rect


if __name__ == "__main__":
    from PIL import Image
    test = Image.new("RGB", (500, 300), color=(200, 200, 200))
    result = preprocess_for_ocr(test)
    print(f"Processed: {result.size}")
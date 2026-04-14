"""
OCR 模块（基于 easyocr）

识别 PIL Image 中的文字，返回带坐标的结果列表。
运行测试：python -m capture.ocr <image_path>
"""

import sys
import os
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from PIL import Image
from utils.logger import logger

# 全局 reader 单例，避免重复加载模型（首次加载约需几秒）
_reader = None


def _get_reader():
    global _reader
    if _reader is None:
        import easyocr
        logger.info("加载 easyocr 中文模型（首次加载需要几秒）...")
        _reader = easyocr.Reader(['ch_sim', 'en'], verbose=False)
        logger.info("easyocr 模型加载完成")
    return _reader


class OcrResult:
    """单条 OCR 识别结果"""
    def __init__(self, text: str, x: int, y: int, w: int, h: int, conf: float = 1.0):
        self.text = text
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.conf = conf

    @property
    def center_x(self) -> int:
        return self.x + self.w // 2

    @property
    def center_y(self) -> int:
        return self.y + self.h // 2

    def __repr__(self):
        return f"OcrResult({self.text!r}, x={self.x}, y={self.y}, conf={self.conf:.2f})"


def recognize(img: Image.Image, window_x: int = 0, window_y: int = 0,
              min_conf: float = 0.3, scale: float = 1.0) -> list[OcrResult]:
    """
    对 PIL Image 做 OCR，返回 OcrResult 列表。
    window_x/y: 窗口在屏幕上的逻辑坐标偏移。
    scale: Retina 缩放比例（物理像素/逻辑像素），通常 Retina 屏为 2.0。
           OCR bbox 是物理像素，除以 scale 转为逻辑坐标后再加偏移。
    min_conf: 最低置信度。
    """
    reader = _get_reader()

    # easyocr 接受文件路径或 numpy array，用临时文件传入
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        img.save(tmp_path)
        raw = reader.readtext(tmp_path)
    finally:
        os.unlink(tmp_path)

    results = []
    for bbox, text, conf in raw:
        if conf < min_conf or not text.strip():
            continue

        # bbox 是四个角点 [[x1,y1],[x2,y1],[x2,y2],[x1,y2]]，物理像素
        # 除以 scale 转为逻辑坐标，再加窗口偏移得到屏幕坐标
        xs = [p[0] for p in bbox]
        ys = [p[1] for p in bbox]
        x = int(min(xs) / scale) + window_x
        y = int(min(ys) / scale) + window_y
        w = int((max(xs) - min(xs)) / scale)
        h = int((max(ys) - min(ys)) / scale)

        results.append(OcrResult(text, x, y, w, h, conf))
        logger.info(f"OCR: {text!r}  conf={conf:.2f}  ({x},{y})")

    return results


def find_text(results: list[OcrResult], keyword: str) -> list[OcrResult]:
    """从结果中过滤包含关键词的条目"""
    return [r for r in results if keyword in r.text]


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python -m capture.ocr <image_path>")
        sys.exit(1)

    img = Image.open(sys.argv[1])
    items = recognize(img)
    print(f"\n识别到 {len(items)} 条文字：\n")
    for item in items:
        print(f"  [{item.x:4d},{item.y:4d}] conf={item.conf:.2f}  {item.text}")

"""
Apple Vision OCR 模块

识别 PIL Image 中的文字，返回带坐标的结果列表。
运行测试：python -m capture.ocr <image_path>
"""

import sys
import os
import objc

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from PIL import Image
from utils.logger import logger


class OcrResult:
    """单条 OCR 识别结果"""
    def __init__(self, text: str, x: int, y: int, w: int, h: int):
        self.text = text
        self.x = x      # 屏幕坐标（像素）
        self.y = y
        self.w = w
        self.h = h

    @property
    def center_x(self) -> int:
        return self.x + self.w // 2

    @property
    def center_y(self) -> int:
        return self.y + self.h // 2

    def __repr__(self):
        return f"OcrResult({self.text!r}, x={self.x}, y={self.y})"


def recognize(img: Image.Image, window_x: int = 0, window_y: int = 0) -> list[OcrResult]:
    """
    对 PIL Image 做 OCR，返回 OcrResult 列表。
    window_x/y: 窗口在屏幕上的偏移，用于将图片坐标换算为屏幕坐标。
    """
    try:
        import Vision
        import Quartz
        from Foundation import NSURL
        import objc
    except ImportError:
        raise RuntimeError("请安装：pip install pyobjc-framework-Vision")

    import tempfile, io

    # Vision 需要从文件或 CGImage 读取，先存为临时 PNG
    # 写入 144 DPI，告知 Vision 这是 Retina 2x 截图，否则中文识别会乱码
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp_path = tmp.name
    img.save(tmp_path, dpi=(144, 144))

    try:
        results = _run_vision(tmp_path, img.width, img.height, window_x, window_y)
    finally:
        os.unlink(tmp_path)

    return results


def _run_vision(image_path: str, img_w: int, img_h: int,
                win_x: int, win_y: int) -> list[OcrResult]:
    """调用 Vision VNRecognizeTextRequest"""
    import Vision
    from Foundation import NSURL, NSArray

    url = NSURL.fileURLWithPath_(image_path)
    request_handler = Vision.VNImageRequestHandler.alloc().initWithURL_options_(url, {})

    request = Vision.VNRecognizeTextRequest.alloc().init()
    request.setRecognitionLevel_(1)          # 1 = accurate
    request.setUsesLanguageCorrection_(False)
    # request.setRecognitionLanguages_(NSArray.arrayWithArray_(["zh-Hans", "zh-Hant", "en"]))
    request.setRecognitionLanguages_(["zh-Hans"])
    request.setAutomaticallyDetectsLanguage_(False)

    error_ptr = objc.nil
    request_handler.performRequests_error_(NSArray.arrayWithArray_([request]), None)

    results = []
    observations = request.results()
    if not observations:
        return results

    for obs in observations:
        text = obs.topCandidates_(1)[0].string()
        if not text:
            continue

        # boundingBox: 归一化坐标，原点在左下角
        box = obs.boundingBox()
        x_norm = box.origin.x
        y_norm = box.origin.y
        w_norm = box.size.width
        h_norm = box.size.height

        # 转换为像素坐标（Vision Y轴翻转）
        px = int(x_norm * img_w) + win_x
        py = int((1.0 - y_norm - h_norm) * img_h) + win_y
        pw = int(w_norm * img_w)
        ph = int(h_norm * img_h)

        results.append(OcrResult(text, px, py, pw, ph))
        logger.debug(f"OCR: {text!r}  ({px},{py})")

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
        print(f"  [{item.x:4d},{item.y:4d}] {item.text}")

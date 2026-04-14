"""
钉钉窗口持续截图模块

流程：截图 → 滚动 → 对比底部 → 重复直到内容不再变化
运行：python -m capture.screenshot
"""

import sys
import os
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

try:
    import Quartz
    from AppKit import NSWorkspace
except ImportError:
    print("请先安装：pip install pyobjc-framework-Quartz")
    sys.exit(1)

from PIL import Image, ImageChops
import numpy as np

from capture.find_window import find_dingtalk_windows, find_main_window
from utils.logger import logger

# 截图保存目录
CAPTURE_DIR = os.path.join("output", "tmp", "captures")

# 底部对比区域高度（px），用于判断是否滚到底
COMPARE_STRIP_HEIGHT = 100

# 相似度阈值：底部区域平均像素差小于此值视为"没有新内容"
SIMILARITY_THRESHOLD = 5.0

# 每次滚动后等待渲染的时间（秒）
SCROLL_WAIT = 0.8

# 单次滚动量（负数=向下）
SCROLL_DELTA = -5

# 最多截图张数（防止死循环）
MAX_SCREENSHOTS = 20


def capture_window(window_id: int) -> Image.Image | None:
    """用 Quartz 按 window ID 截图，返回 PIL Image"""
    cg_image = Quartz.CGWindowListCreateImage(
        Quartz.CGRectNull,
        Quartz.kCGWindowListOptionIncludingWindow,
        window_id,
        Quartz.kCGWindowImageBoundsIgnoreFraming,
    )
    if not cg_image:
        return None

    width  = Quartz.CGImageGetWidth(cg_image)
    height = Quartz.CGImageGetHeight(cg_image)
    bpr    = Quartz.CGImageGetBytesPerRow(cg_image)

    data_provider = Quartz.CGImageGetDataProvider(cg_image)
    raw_data = Quartz.CGDataProviderCopyData(data_provider)

    img = Image.frombytes("RGBA", (width, height), bytes(raw_data), "raw", "BGRA", bpr)
    return img.convert("RGB")


def activate_dingtalk(pid: int) -> None:
    """将钉钉窗口带到前台，确保滚动事件发到正确的窗口"""
    running_apps = NSWorkspace.sharedWorkspace().runningApplications()
    for app in running_apps:
        if app.processIdentifier() == pid:
            app.activateWithOptions_(1 << 1)  # NSApplicationActivateIgnoringOtherApps
            time.sleep(0.3)  # 等待窗口激活
            logger.info(f"已激活钉钉窗口 PID={pid}")
            return
    logger.warning(f"未找到 PID={pid} 的应用")


def scroll_window(window: dict, delta: int = SCROLL_DELTA) -> None:
    """激活钉钉窗口后发送滚轮事件，确保事件发到钉钉而非终端"""
    activate_dingtalk(window["pid"])

    cx = window["x"] + window["width"] // 2
    cy = window["y"] + window["height"] // 2

    event = Quartz.CGEventCreateScrollWheelEvent(
        None,
        Quartz.kCGScrollEventUnitLine,
        1,
        delta,
    )
    Quartz.CGEventSetLocation(event, (cx, cy))
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)


def bottom_strip(img: Image.Image) -> np.ndarray:
    """裁剪图片底部条带，转为 numpy 数组"""
    h = img.height
    strip = img.crop((0, h - COMPARE_STRIP_HEIGHT, img.width, h))
    return np.array(strip, dtype=np.float32)


def is_same_content(img_a: Image.Image, img_b: Image.Image) -> bool:
    """比较两张图底部区域，判断内容是否相同"""
    a = bottom_strip(img_a)
    b = bottom_strip(img_b)
    diff = np.abs(a - b).mean()
    logger.info(f"底部区域像素差: {diff:.2f}")
    return diff < SIMILARITY_THRESHOLD


def capture_all(window: dict) -> list[str]:
    """
    持续截图直到内容不再变化，返回所有截图路径列表。
    window: find_window 返回的窗口信息字典
    """
    os.makedirs(CAPTURE_DIR, exist_ok=True)

    paths = []
    prev_img = None

    for i in range(MAX_SCREENSHOTS):
        img = capture_window(window["window_id"])
        if img is None:
            logger.warning("截图失败，窗口可能已关闭")
            break

        path = os.path.join(CAPTURE_DIR, f"capture_{i+1:03d}.png")
        img.save(path)
        paths.append(path)
        logger.info(f"截图 {i+1}: {path}  ({img.width}x{img.height})")

        # 第一张直接滚动继续
        if prev_img is not None and is_same_content(prev_img, img):
            logger.info("底部内容未变化，判断已到底，停止截图")
            break

        prev_img = img
        scroll_window(window, SCROLL_DELTA)
        time.sleep(SCROLL_WAIT)
    else:
        logger.warning(f"已达最大截图数 {MAX_SCREENSHOTS}，强制停止")

    logger.info(f"共截图 {len(paths)} 张，保存在 {CAPTURE_DIR}")
    return paths


if __name__ == "__main__":
    windows = find_dingtalk_windows()
    if not windows:
        print("未找到钉钉窗口，请确认钉钉已启动并已授权屏幕录制权限")
        sys.exit(1)

    main_win = find_main_window(windows)
    print(f"目标窗口: {main_win['name'] or main_win['owner']}  "
          f"({main_win['width']}x{main_win['height']})")
    print(f"截图保存到: {CAPTURE_DIR}\n")

    paths = capture_all(main_win)
    print(f"\n完成，共 {len(paths)} 张截图：")
    for p in paths:
        print(f"  {p}")

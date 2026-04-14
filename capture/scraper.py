"""
钉钉家校本自动抓取 - AI家校本入口版

流程：
  1. 激活钉钉窗口，进入目标群聊
  2. 点击底部"AI家校本"按钮
  3. 点击"全部"tab，查看所有作业
  4. 依次点击每个作业卡片，提取完整内容
  5. 点击"<"返回，处理下一个
  6. 返回 RawMessage 列表供 card_parser 使用

运行测试：python -m capture.scraper [--debug]
"""

import sys
import os
import re
import time
from datetime import date, datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

try:
    import Quartz
    from AppKit import NSWorkspace
except ImportError:
    print("请安装：pip install pyobjc-framework-Quartz")
    sys.exit(1)

from PIL import Image

from capture.find_window import find_dingtalk_windows, find_main_window
from capture.screenshot import capture_window, activate_dingtalk
from capture.ocr import recognize, find_text, OcrResult
from models.models import RawMessage
from utils.logger import logger
import config

# ── 参数 ──────────────────────────────────────────────────────────────────────
CARD_OPEN_WAIT   = 1.5   # 点击后等待渲染（秒）
MAX_HOMEWORK     = 6     # 最多提取作业数量
DEBUG_DIR        = os.path.join("output", "tmp", "captures")

SUBJECT_KEYWORDS = config.SUBJECT_ORDER
RE_CARD_TITLE = re.compile(
    r"(\d{1,2})月(\d{1,2})日\s*(" +
    "|".join(re.escape(s) for s in SUBJECT_KEYWORDS) + r")"
)


# ── 基础操作 ──────────────────────────────────────────────────────────────────

def click(x: int, y: int) -> None:
    pos = (x, y)
    down = Quartz.CGEventCreateMouseEvent(
        None, Quartz.kCGEventLeftMouseDown, pos, Quartz.kCGMouseButtonLeft)
    up = Quartz.CGEventCreateMouseEvent(
        None, Quartz.kCGEventLeftMouseUp, pos, Quartz.kCGMouseButtonLeft)
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, down)
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, up)
    logger.info(f"点击 ({x}, {y})")


def wait_and_capture(window: dict, wait: float = CARD_OPEN_WAIT):
    """等待渲染后截图，只截右侧面板区域，返回 (panel_img, panel_x_logical, scale)"""
    time.sleep(wait)
    img = capture_window(window["window_id"])
    if img is None:
        logger.warning("截图失败")
        return None, 0, 1.0
    # Retina 2x：图片物理宽度是逻辑宽度的2倍
    scale = img.width / window["width"]
    logger.info(f"截图尺寸={img.width}x{img.height}  窗口逻辑={window['width']}x{window['height']}  scale={scale:.1f}")
    # 右侧面板从窗口 37% 处开始（逻辑坐标）
    panel_x_logical = int(window["width"] * 0.37)
    panel_x_physical = int(panel_x_logical * scale)
    panel_img = img.crop((panel_x_physical, 0, img.width, img.height))
    return panel_img, panel_x_logical, scale


def ocr_panel(window: dict, wait: float = CARD_OPEN_WAIT,
              debug_name: str = None, debug: bool = False):
    """截图右侧面板并 OCR，返回屏幕逻辑坐标的 OcrResult 列表"""
    panel_img, panel_x_logical, scale = wait_and_capture(window, wait)
    if panel_img is None:
        return [], 0

    if debug and debug_name:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        panel_img.save(os.path.join(DEBUG_DIR, debug_name))
        logger.info(f"调试截图已保存: {debug_name}")

    # window_x/y 传逻辑坐标偏移，scale 用于将物理像素坐标转为逻辑坐标
    results = recognize(panel_img,
                        window_x=window["x"] + panel_x_logical,
                        window_y=window["y"],
                        scale=scale)
    logger.info(f"OCR [{debug_name or ''}] 识别到 {len(results)} 条：")
    for r in results:
        logger.info(f"  [{r.x:4d},{r.y:4d}] conf={r.conf:.2f}  {r.text!r}")
    return results, panel_x_logical


# ── 步骤函数 ──────────────────────────────────────────────────────────────────

def step1_open_jiaxiaob(window: dict, debug: bool = False) -> bool:
    """点击底部'AI家校本'按钮"""
    logger.info("=== 步骤1：点击 AI家校本 ===")
    activate_dingtalk(window["pid"])

    results, _ = ocr_panel(window, wait=0.5, debug_name="step1_before.png", debug=debug)
    targets = find_text(results, "家校本")
    if not targets:
        logger.error("未找到'AI家校本'按钮，请确认已打开目标群聊")
        return False

    # 取置信度最高的
    target = max(targets, key=lambda r: r.conf)
    logger.info(f"找到'家校本'按钮: ({target.center_x}, {target.center_y})")
    activate_dingtalk(window["pid"])
    click(target.center_x, target.center_y)
    time.sleep(1)
    return True


def step2_click_all_tab(window: dict, debug: bool = False) -> bool:
    """点击'全部'tab"""
    logger.info("=== 步骤2：点击'全部'tab ===")
    results, _ = ocr_panel(window, wait=CARD_OPEN_WAIT,
                           debug_name="step2_list.png", debug=debug)

    targets = find_text(results, "全部")
    if not targets:
        logger.warning("未找到'全部'tab，可能已在全部视图")
        return True

    target = targets[0]
    logger.info(f"点击'全部': ({target.center_x}, {target.center_y})")
    activate_dingtalk(window["pid"])
    click(target.center_x, target.center_y)
    time.sleep(0.5)
    return True


def step3_get_card_list(window: dict, debug: bool = False) -> list[OcrResult]:
    """
    识别作业列表，返回每个作业卡片的标题 OcrResult。
    标题格式："4月13日数学" 或 "数学"（带科目关键词）
    """
    logger.info("=== 步骤3：识别作业列表 ===")
    results, _ = ocr_panel(window, wait=0.8,
                           debug_name="step3_all.png", debug=debug)

    cards = []
    seen_subjects = set()
    for r in results:
        # 匹配 "4月13日数学" 格式
        m = RE_CARD_TITLE.search(r.text)
        if m:
            subject = m.group(3)
            if subject not in seen_subjects:
                seen_subjects.add(subject)
                cards.append(r)
                logger.info(f"  发现作业卡片: {r.text!r}  ({r.center_x},{r.center_y})")

    logger.info(f"共识别到 {len(cards)} 个作业卡片")
    return cards


def step4_extract_one(window: dict, card: OcrResult,
                      debug: bool = False, idx: int = 0) -> str:
    """
    点击一个作业卡片，提取详情页完整内容，点击返回。
    返回提取的文字。
    """
    logger.info(f"=== 步骤4：提取作业 [{card.text}] ===")
    activate_dingtalk(window["pid"])
    click(card.center_x, card.center_y)

    results, _ = ocr_panel(window, wait=CARD_OPEN_WAIT,
                           debug_name=f"step4_detail_{idx}.png", debug=debug)

    # 过滤无关内容
    skip_keywords = ["去补交", "有疑问", "去打印", "优秀作答", "如何被选",
                     "排行榜", "已完成", "预计需", "已截止", "昨日", "首页", "错题本"]
    lines = []
    for r in results:
        if r.conf < 0.4:
            continue
        if any(kw in r.text for kw in skip_keywords):
            logger.debug(f"  跳过: {r.text!r}")
            continue
        lines.append(r.text)

    text = "\n".join(lines)
    logger.info(f"提取内容（前120字）: {text[:120]}")

    # 点击返回按钮 "<"
    back_results, _ = ocr_panel(window, wait=0, debug_name=None)
    back_btn = None
    for r in back_results:
        if r.text.strip() in ("<", "〈") or r.text.startswith("<") or "〈" in r.text or "的练习" in r.text:
            back_btn = r
            break


    if back_btn:
        # 点击行左端 +10px，确保落在 "<" 上而不是整行中心
        bx = back_btn.x + 10
        by = back_btn.center_y
        logger.info(f"点击返回按钮左端: ({bx},{by})  原文={back_btn.text!r}")
        activate_dingtalk(window["pid"])
        click(bx, by)
    else:
        # 返回按钮找不到时，点击面板左上角固定位置
        panel_x = window["x"] + int(window["width"] * 0.37)
        fallback_x = panel_x + 30
        fallback_y = window["y"] + 30
        logger.warning(f"未找到返回按钮，点击左上角固定位置 ({fallback_x},{fallback_y})")
        activate_dingtalk(window["pid"])
        click(fallback_x, fallback_y)

    time.sleep(0.8)
    return text


def step5_close_panel(window: dict) -> None:
    """点击 X 关闭家校本面板"""
    logger.info("=== 步骤5：关闭面板 ===")
    results, _ = ocr_panel(window, wait=0.3)
    for r in results:
        if r.text.strip() == "×" or r.text.strip() == "X":
            activate_dingtalk(window["pid"])
            click(r.center_x, r.center_y)
            logger.info("面板已关闭")
            return
    # 找不到X按钮，点击面板右上角固定位置
    # 找不到X按钮，点击面板左侧外部区域（群聊区域）关闭
    outside_x = window["x"] + int(window["width"] * 0.37) // 2
    outside_y = window["y"] + window["height"] // 2
    logger.warning(f"未找到X按钮，点击面板左侧外部区域 ({outside_x},{outside_y})")
    activate_dingtalk(window["pid"])
    click(outside_x, outside_y)


# ── 主流程 ────────────────────────────────────────────────────────────────────

def scrape(debug: bool = False) -> list[RawMessage]:
    """
    通过 AI家校本 入口抓取今天的作业，返回 RawMessage 列表。
    """
    windows = find_dingtalk_windows()
    if not windows:
        raise RuntimeError("未找到钉钉窗口，请确认钉钉已启动并授权屏幕录制权限")

    window = find_main_window(windows)
    logger.info(f"目标窗口: {window['owner']}  PID={window['pid']}  "
                f"位置=({window['x']},{window['y']})  大小={window['width']}x{window['height']}")

    activate_dingtalk(window["pid"])
    time.sleep(0.5)

    # 步骤1：打开 AI家校本
    if not step1_open_jiaxiaob(window, debug):
        return []

    # 步骤2：点击"全部"tab
    step2_click_all_tab(window, debug)

    # 步骤3：获取作业列表
    cards = step3_get_card_list(window, debug)
    if not cards:
        logger.warning("未识别到任何作业卡片")
        step5_close_panel(window)
        return []

    # 步骤4：逐个提取，最多 MAX_HOMEWORK 个
    messages = []
    for i, card in enumerate(cards[:MAX_HOMEWORK]):
        text = step4_extract_one(window, card, debug=debug, idx=i)
        if text:
            messages.append(RawMessage(
                msg_id=f"jiaxiaob_{i}_{card.text}",
                sender_id="capture",
                sender_name="AI家校本",
                timestamp=datetime.now(),
                msg_type="text",
                text=text,
            ))
            logger.info(f"已收集第 {i+1} 科，当前共 {len(messages)} 科")

    # 步骤5：关闭面板
    step5_close_panel(window)

    logger.info(f"=== 抓取完成，共 {len(messages)} 科作业 ===")
    return messages


if __name__ == "__main__":
    debug_mode = "--debug" in sys.argv
    msgs = scrape(debug=debug_mode)

    if not msgs:
        print("\n未抓取到任何作业内容")
        print("建议加 --debug 参数：python -m capture.scraper --debug")
        sys.exit(0)

    print(f"\n共抓取 {len(msgs)} 科作业：\n")
    for m in msgs:
        print(f"[{m.msg_id}]\n{m.text}\n{'─'*40}")

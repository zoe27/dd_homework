"""
钉钉家校本自动抓取主流程

流程：
  1. 激活钉钉窗口
  2. 滚到底部（最新消息）
  3. 从底部向上扫描，识别家校本卡片
  4. 点开卡片提取作业文字
  5. 遇到昨天或更早的消息时间戳停止
  6. 返回 RawMessage 列表供 card_parser 使用

运行测试：python -m capture.scraper [--debug]
  --debug: 每步截图保存到 output/tmp/captures/
"""

import sys
import os
import re
import time
from datetime import date, datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

try:
    import Quartz
    from AppKit import NSWorkspace
except ImportError:
    print("请安装：pip install pyobjc-framework-Quartz")
    sys.exit(1)

from PIL import Image
import numpy as np

from capture.find_window import find_dingtalk_windows, find_main_window
from capture.screenshot import capture_window, activate_dingtalk
from capture.ocr import recognize, find_text, OcrResult
from models.models import RawMessage
from utils.logger import logger
import config

# ── 参数 ──────────────────────────────────────────────────────────────────────
SCROLL_DOWN_DELTA    = -15   # 第一阶段：快速滚到底，幅度大
SCROLL_UP_DELTA      = 5     # 第二阶段：向上扫描，小幅度避免跳过卡片
SCROLL_WAIT          = 0.6   # 滚动后等待渲染（秒）
CARD_OPEN_WAIT       = 1.8   # 点击卡片后等待面板渲染（秒）
CLOSE_WAIT           = 0.4   # 关闭面板后等待（秒）
COMPARE_HEIGHT       = 120   # 对比条带高度（px）
SIMILARITY_THRESHOLD = 5.0   # 像素差阈值
SAME_COUNT_TO_STOP   = 2     # 连续相同次数判定到顶/底
MAX_SCROLL_STEPS     = 80    # 最大滚动次数

# 消息区域在窗口右侧，左侧列表约占 37% 宽度
MSG_AREA_X_RATIO = 0.37   # 消息区起始 x 比例

DEBUG_DIR = os.path.join("output", "tmp", "captures")

SUBJECT_KEYWORDS = config.SUBJECT_ORDER
# 消息时间戳正则：匹配"昨天"、"04/09"、"4月9日"等

RE_TIMESTAMP_DATE = re.compile(r"(\d{1,2})[/月](\d{1,2})")
#   格式A: "4月10日语文"（带日期）
#   格式B: 只有科目名，如"数学"（不带日期，依赖截止时间判断）
RE_CARD_TITLE_WITH_DATE = re.compile(
    r"(\d{1,2})月(\d{1,2})日\s*(" +
    "|".join(re.escape(s) for s in SUBJECT_KEYWORDS) + r")"
)
RE_CARD_TITLE_SUBJECT_ONLY = re.compile(
    r"^(" + "|".join(re.escape(s) for s in SUBJECT_KEYWORDS) + r")$"
)
# 截止时间：用于从卡片内容中提取日期，如"截止时间：04月13日22:00"
RE_DEADLINE = re.compile(r"截止.*?(\d{1,2})月(\d{1,2})日")


# ── 工具函数 ──────────────────────────────────────────────────────────────────

def scroll(window: dict, delta: int) -> None:
    activate_dingtalk(window["pid"])
    # 滚动发到消息区中心，而不是整个窗口中心
    msg_x_start = window["x"] + int(window["width"] * MSG_AREA_X_RATIO)
    cx = msg_x_start + (window["x"] + window["width"] - msg_x_start) // 2
    cy = window["y"] + window["height"] // 2
    event = Quartz.CGEventCreateScrollWheelEvent(
        None, Quartz.kCGScrollEventUnitLine, 1, delta
    )
    Quartz.CGEventSetLocation(event, (cx, cy))
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)
    logger.debug(f"滚动事件已发送 delta={delta}，消息区中心=({cx},{cy})")


def click(x: int, y: int) -> None:
    pos = (x, y)
    down = Quartz.CGEventCreateMouseEvent(
        None, Quartz.kCGEventLeftMouseDown, pos, Quartz.kCGMouseButtonLeft
    )
    up = Quartz.CGEventCreateMouseEvent(
        None, Quartz.kCGEventLeftMouseUp, pos, Quartz.kCGMouseButtonLeft
    )
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, down)
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, up)
    logger.debug(f"点击事件已发送 ({x},{y})")


def strip_array(img: Image.Image, top: bool = False) -> np.ndarray:
    if top:
        strip = img.crop((0, 0, img.width, COMPARE_HEIGHT))
    else:
        strip = img.crop((0, img.height - COMPARE_HEIGHT, img.width, img.height))
    return np.array(strip, dtype=np.float32)


def is_same(img_a: Image.Image, img_b: Image.Image, top: bool = False) -> bool:
    diff = np.abs(strip_array(img_a, top) - strip_array(img_b, top)).mean()
    zone = "顶部" if top else "底部"
    logger.debug(f"{zone}条带像素差: {diff:.2f}（阈值 {SIMILARITY_THRESHOLD}）")
    return diff < SIMILARITY_THRESHOLD


def save_debug(img: Image.Image, name: str) -> None:
    os.makedirs(DEBUG_DIR, exist_ok=True)
    path = os.path.join(DEBUG_DIR, name)
    img.save(path)
    logger.debug(f"调试截图已保存: {path}")


# ── 终止条件判断 ──────────────────────────────────────────────────────────────

def is_old_timestamp(text: str) -> bool:
    """
    判断文字是否是群聊消息的时间戳且早于今天。
    只匹配独立的时间戳格式（如"昨天"、"04/09"），
    排除卡片内的"截止时间：04月13日"等内容。
    """
    today = date.today()

    # "昨天"是钉钉群聊消息时间戳的固定格式
    if text.strip() == "昨天" or text.strip().startswith("昨天 "):
        logger.debug(f"发现'昨天'时间戳: {text!r}")
        return True

    # 排除含"截止"、"完成"等卡片内容字段，避免误判
    if any(kw in text for kw in ["截止", "完成", "布置", "预计", "已"]):
        return False

    # 匹配独立日期格式 "04/09" 或 "4月9日"
    for m in RE_TIMESTAMP_DATE.finditer(text):
        month, day = int(m.group(1)), int(m.group(2))
        try:
            msg_date = date(today.year, month, day)
            if msg_date < today:
                logger.debug(f"发现早于今天的时间戳 {msg_date}: {text!r}")
                return True
        except ValueError:
            pass

    return False


def should_stop(ocr_results: list[OcrResult]) -> bool:
    """
    扫描当前屏幕的 OCR 结果，判断是否应该停止向上扫描。
    条件：发现昨天或更早的消息时间戳。
    """
    for r in ocr_results:
        if is_old_timestamp(r.text):
            logger.info(f"终止条件触发：发现旧日期文字 → {r.text!r}")
            return True
    return False


# ── 家校本卡片识别 ────────────────────────────────────────────────────────────

def find_card(ocr_results: list[OcrResult]) -> tuple[OcrResult | None, str | None]:
    """
    在 OCR 结果中找今天的家校本卡片。
    规则：
      1. 必须有"家校本"字样
      2. 截止时间是今天（从"截止时间：XX月XX日"提取）
      3. 找到科目名作为点击目标和标识

    返回 (点击目标OcrResult, 科目名)，未找到返回 (None, None)。
    """
    today = date.today()
    all_texts = [r.text for r in ocr_results]

    # 必须有"家校本"
    has_jiaxiaob = any("家校本" in t for t in all_texts)
    if not has_jiaxiaob:
        return None, None

    logger.debug("发现'家校本'关键词，开始识别卡片...")

    # 从截止时间提取日期（最可靠）
    card_date = None
    for t in all_texts:
        dm = RE_DEADLINE.search(t)
        if dm:
            try:
                card_date = date(today.year, int(dm.group(1)), int(dm.group(2)))
                logger.debug(f"截止时间日期: {card_date}  来源: {t!r}")
                break
            except ValueError:
                pass

    # 没有截止时间，尝试从标题"X月X日科目"提取日期
    if card_date is None:
        for r in ocr_results:
            m = RE_CARD_TITLE_WITH_DATE.search(r.text)
            if m:
                try:
                    card_date = date(today.year, int(m.group(1)), int(m.group(2)))
                    logger.debug(f"标题日期: {card_date}  来源: {r.text!r}")
                    break
                except ValueError:
                    pass

    if card_date is None:
        logger.debug("无法确定卡片日期，跳过")
        return None, None

    if card_date < today:
        logger.info(f"卡片日期 {card_date} 早于今天，不点击")
        return None, None

    # 找科目名：优先从标题"X月X日科目"提取，其次找独立科目行
    subject = None
    click_target = None

    for r in ocr_results:
        m = RE_CARD_TITLE_WITH_DATE.search(r.text)
        if m:
            subject = m.group(3)
            click_target = r
            logger.info(f"[格式A] 科目={subject}  日期={card_date}  文字={r.text!r}")
            break

    if subject is None:
        for r in ocr_results:
            m = RE_CARD_TITLE_SUBJECT_ONLY.match(r.text.strip())
            if m:
                subject = m.group(1)
                click_target = r
                logger.info(f"[格式B] 科目={subject}  日期={card_date}  文字={r.text!r}")
                break

    if subject is None:
        logger.debug("有家校本和日期但未识别到科目名")
        return None, None

    return click_target, subject


# ── 面板内容提取 ──────────────────────────────────────────────────────────────

def extract_panel_text(window: dict, debug: bool = False) -> str:
    """截图展开的家校本面板（右半部分），OCR 提取正文"""
    time.sleep(CARD_OPEN_WAIT)
    img = capture_window(window["window_id"])
    if img is None:
        logger.warning("面板截图失败")
        return ""

    if debug:
        save_debug(img, f"panel_{int(time.time())}.png")

    # 裁剪右侧面板区域（右 2/3）
    panel_x = img.width // 3
    panel_img = img.crop((panel_x, 0, img.width, img.height))

    logger.debug(f"面板截图区域: x={panel_x}~{img.width}, 大小={panel_img.width}x{panel_img.height}")

    results = recognize(
        panel_img,
        window_x=window["x"] + panel_x,
        window_y=window["y"],
    )
    logger.info(f"面板 OCR 识别到 {len(results)} 条文字")

    skip_keywords = ["打印", "反馈", "完成时长", "预计需", "已截止", "点此", "↓", "家长"]
    lines = []
    for r in results:
        if any(kw in r.text for kw in skip_keywords):
            logger.debug(f"  跳过: {r.text!r}")
            continue
        logger.debug(f"  保留: {r.text!r}")
        lines.append(r.text)

    text = "\n".join(lines)
    logger.info(f"面板提取文字（前100字）: {text[:100]}")
    return text


def close_panel(window: dict) -> None:
    """点击窗口左侧 1/6 处关闭面板"""
    x = window["x"] + window["width"] // 6
    y = window["y"] + window["height"] // 2
    logger.debug(f"关闭面板，点击坐标: ({x},{y})")
    activate_dingtalk(window["pid"])
    click(x, y)
    time.sleep(CLOSE_WAIT)


# ── 阶段一：滚到底部 ──────────────────────────────────────────────────────────

def scroll_to_bottom(window: dict, debug: bool = False) -> None:
    logger.info("=== 阶段一：滚到底部 ===")
    prev_img = None
    same_count = 0

    for step in range(MAX_SCROLL_STEPS):
        scroll(window, SCROLL_DOWN_DELTA)
        time.sleep(SCROLL_WAIT)
        img = capture_window(window["window_id"])
        if img is None:
            logger.warning("截图失败，跳出")
            break

        if debug:
            save_debug(img, f"down_{step:03d}.png")

        if prev_img is not None and is_same(prev_img, img, top=False):
            same_count += 1
            logger.debug(f"底部相同计数: {same_count}/{SAME_COUNT_TO_STOP}")
            if same_count >= SAME_COUNT_TO_STOP:
                logger.info(f"已到底部（第 {step+1} 步）")
                return
        else:
            same_count = 0

        prev_img = img

    logger.warning("滚到底部超出最大步数，继续执行")


# ── 阶段二：向上扫描 ──────────────────────────────────────────────────────────

def scan_upward(window: dict, debug: bool = False) -> list[RawMessage]:
    logger.info("=== 阶段二：向上扫描作业卡片 ===")
    messages = []
    prev_img = None
    same_count = 0
    seen_subjects: set[str] = set()

    for step in range(MAX_SCROLL_STEPS):
        logger.info(f"--- 扫描步骤 {step+1} ---")

        img = capture_window(window["window_id"])
        if img is None:
            logger.warning("截图失败，跳出")
            break

        if debug:
            save_debug(img, f"up_{step:03d}.png")

        # 到顶判断
        if prev_img is not None and is_same(prev_img, img, top=True):
            same_count += 1
            logger.debug(f"顶部相同计数: {same_count}/{SAME_COUNT_TO_STOP}")
            if same_count >= SAME_COUNT_TO_STOP:
                logger.info("已到顶部，扫描结束")
                break
        else:
            same_count = 0
        prev_img = img

        # OCR，只识别右侧消息区域，打印所有识别到的文字
        msg_x_start = int(img.width * MSG_AREA_X_RATIO)
        msg_img = img.crop((msg_x_start, 0, img.width, img.height))
        if debug:
            save_debug(msg_img, f"msg_{step:03d}.png")
        results = recognize(msg_img,
                            window_x=window["x"] + msg_x_start,
                            window_y=window["y"])
        logger.info(f"OCR 识别到 {len(results)} 条文字：")
        for r in results:
            logger.info(f"  [{r.x:4d},{r.y:4d}] {r.text!r}")

        # 终止条件：发现旧日期时间戳
        if should_stop(results):
            logger.info("发现旧消息时间戳，停止扫描")
            break

        # 识别家校本卡片
        card, subject = find_card(results)
        if card and subject:
            if subject in seen_subjects:
                logger.info(f"科目 [{subject}] 已收集过，跳过")
            else:
                seen_subjects.add(subject)
                logger.info(f"点击卡片坐标=({card.center_x},{card.center_y})  科目={subject}")

                activate_dingtalk(window["pid"])
                click(card.center_x, card.center_y)

                text = extract_panel_text(window, debug=debug)

                if text:
                    messages.append(RawMessage(
                        msg_id=f"capture_{step}_{subject}",
                        sender_id="capture",
                        sender_name="自动抓取",
                        timestamp=datetime.now(),
                        msg_type="text",
                        text=text,
                    ))
                    logger.info(f"已收集科目 [{subject}]，当前共 {len(messages)} 科")
                else:
                    logger.warning(f"科目 [{subject}] 面板提取内容为空")

                close_panel(window)

        # 向上滚动
        scroll(window, SCROLL_UP_DELTA)
        time.sleep(SCROLL_WAIT)

    logger.info(f"=== 扫描完成，共收集 {len(messages)} 科作业 ===")
    return messages


# ── 主入口 ────────────────────────────────────────────────────────────────────

def scrape(debug: bool = False) -> list[RawMessage]:
    """
    自动抓取今天的家校本作业，返回 RawMessage 列表。
    debug=True 时每步截图保存到 output/tmp/captures/。
    """
    windows = find_dingtalk_windows()
    if not windows:
        raise RuntimeError("未找到钉钉窗口，请确认钉钉已启动并授权屏幕录制权限")

    window = find_main_window(windows)
    logger.info(f"目标窗口: {window['owner']}  PID={window['pid']}  "
                f"位置=({window['x']},{window['y']})  大小={window['width']}x{window['height']}")

    if debug:
        logger.info(f"调试模式开启，截图保存到: {DEBUG_DIR}")

    scroll_to_bottom(window, debug=debug)
    messages = scan_upward(window, debug=debug)
    return messages


if __name__ == "__main__":
    debug_mode = "--debug" in sys.argv
    msgs = scrape(debug=debug_mode)

    if not msgs:
        print("\n未抓取到任何作业内容")
        print("建议加 --debug 参数查看截图：python -m capture.scraper --debug")
        sys.exit(0)

    print(f"\n共抓取 {len(msgs)} 科作业：\n")
    for m in msgs:
        print(f"  [{m.msg_id}]\n{m.text}\n")

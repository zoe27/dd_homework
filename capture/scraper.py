"""
钉钉家校本自动抓取 v2 - 工作通知群入口

流程：
  1. OCR 左侧会话列表，找"工作通知:深圳市龙岗区"条目并点击
  2. 右侧群聊消息区滚到底部
  3. 从底部向上扫描，识别今天的作业卡片（"X月X日科目" + "XX老师 发布于"）
  4. 点击卡片进入详情，找"去打印"坐标定位内容区，提取作业文字
  5. 点击"去打印"左侧返回列表，处理下一个
  6. 遇到非今天日期停止，返回 RawMessage 列表

运行：python -m capture.scraper [--debug]
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
import numpy as np

from capture.find_window import find_dingtalk_windows, find_main_window
from capture.screenshot import capture_window, activate_dingtalk
from capture.ocr import recognize, find_text, OcrResult
from models.models import RawMessage
from utils.logger import logger
import config

# ── 参数 ──────────────────────────────────────────────────────────────────────
RENDER_WAIT      = 1.5    # 点击后等待渲染（秒）
SCROLL_WAIT      = 0.6    # 滚动后等待（秒）
MAX_HOMEWORK     = 8      # 最多提取作业数
MAX_SCROLL_STEPS = 40     # 最大滚动次数
SAME_COUNT_STOP  = 2      # 连续相同次数判定到顶/底
DEBUG_DIR        = os.path.join("output", "tmp", "captures")

# 目标会话关键词
TARGET_CONV_KEYWORDS = ["工作通知", "深圳市龙岗区"]

SUBJECT_KEYWORDS = config.SUBJECT_ORDER
RE_CARD_TITLE = re.compile(
    r"(?:(\d{1,2})月)?(\d{1,2})日\s*(" +
    "|".join(re.escape(s) for s in SUBJECT_KEYWORDS) + r")"
)
RE_TEACHER_POST = re.compile(r".{2,6}老师\s*发布于")


# ── 基础操作 ──────────────────────────────────────────────────────────────────

def get_scale(window: dict, img: Image.Image) -> float:
    """计算 Retina 缩放比例"""
    return img.width / window["width"]


def click(x: int, y: int) -> None:
    pos = (x, y)
    down = Quartz.CGEventCreateMouseEvent(
        None, Quartz.kCGEventLeftMouseDown, pos, Quartz.kCGMouseButtonLeft)
    up = Quartz.CGEventCreateMouseEvent(
        None, Quartz.kCGEventLeftMouseUp, pos, Quartz.kCGMouseButtonLeft)
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, down)
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, up)
    logger.info(f"点击 ({x}, {y})")


def scroll_at(x: int, y: int, delta: int) -> None:
    event = Quartz.CGEventCreateScrollWheelEvent(
        None, Quartz.kCGScrollEventUnitLine, 1, delta)
    Quartz.CGEventSetLocation(event, (x, y))
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)


def screenshot(window: dict) -> tuple[Image.Image, float]:
    """截取整个钉钉窗口，返回 (img, scale)"""
    img = capture_window(window["window_id"])
    if img is None:
        return None, 1.0
    scale = get_scale(window, img)
    return img, scale


def ocr_region(img: Image.Image, scale: float,
               win_x: int, win_y: int,
               crop: tuple = None) -> list[OcrResult]:
    """
    对图片（或裁剪区域）做 OCR，坐标转换为屏幕逻辑坐标。
    crop: (x1, y1, x2, y2) 物理像素裁剪区域，None 表示全图。
    """
    if crop:
        x1, y1, x2, y2 = crop
        region = img.crop((x1, y1, x2, y2))
        offset_x = win_x + int(x1 / scale)
        offset_y = win_y + int(y1 / scale)
    else:
        region = img
        offset_x = win_x
        offset_y = win_y
    return recognize(region, window_x=offset_x, window_y=offset_y, scale=scale)


def save_debug(img: Image.Image, name: str) -> None:
    os.makedirs(DEBUG_DIR, exist_ok=True)
    img.save(os.path.join(DEBUG_DIR, name))
    logger.info(f"调试截图: {name}")


def content_same(img_a: Image.Image, img_b: Image.Image,
                 top_ratio: float = 0.2, bottom_ratio: float = 0.15) -> bool:
    """比较两张图中间内容区域是否相同"""
    h = img_a.height
    t = int(h * top_ratio)
    b = int(h * (1 - bottom_ratio))
    a = np.array(img_a.crop((0, t, img_a.width, b)), dtype=np.float32)
    c = np.array(img_b.crop((0, t, img_b.width, b)), dtype=np.float32)
    diff = np.abs(a - c).mean()
    logger.debug(f"内容区像素差: {diff:.2f}")
    return diff < 5.0


# ── 步骤1：找到并点击目标会话 ─────────────────────────────────────────────────

def step1_open_target_conv(window: dict, debug: bool = False) -> bool:
    """
    在左侧会话列表 OCR 找"工作通知:深圳市龙岗区"条目并点击。
    找不到则向下滚动继续找。
    """
    logger.info("=== 步骤1：定位目标会话 ===")
    activate_dingtalk(window["pid"])

    for attempt in range(10):
        time.sleep(0.5)
        img, scale = screenshot(window)
        if img is None:
            continue

        # 只 OCR 左侧会话列表区域（动态：取整图左半部分）
        # 通过找"消息"或会话列表特征确定边界，这里取左侧 40% 作为初始范围
        list_x2 = int(img.width * 0.40)
        results = ocr_region(img, scale,
                             win_x=window["x"], win_y=window["y"],
                             crop=(0, 0, list_x2, img.height))

        if debug:
            save_debug(img.crop((0, 0, list_x2, img.height)),
                       f"step1_list_{attempt:02d}.png")

        logger.info(f"左侧列表 OCR {len(results)} 条")

        # 找包含所有关键词的条目
        for r in results:
            if all(kw in r.text for kw in TARGET_CONV_KEYWORDS):
                logger.info(f"找到目标会话: {r.text!r}  点击 ({r.center_x},{r.center_y})")
                activate_dingtalk(window["pid"])
                click(r.center_x, r.center_y)
                time.sleep(RENDER_WAIT)
                return True

        # 没找到，向下滚动左侧列表
        list_cx = window["x"] + int(window["width"] * 0.20)
        list_cy = window["y"] + window["height"] // 2
        logger.info(f"未找到目标会话，滚动左侧列表（尝试 {attempt+1}）")
        activate_dingtalk(window["pid"])
        scroll_at(list_cx, list_cy, -5)
        time.sleep(SCROLL_WAIT)

    logger.error("多次尝试后仍未找到目标会话")
    return False


# ── 步骤2：消息区滚到底部 ─────────────────────────────────────────────────────

def step2_scroll_to_bottom(window: dict, msg_cx: int, msg_cy: int,
                           debug: bool = False) -> None:
    """向下滚动消息区直到底部"""
    logger.info("=== 步骤2：滚到底部 ===")
    prev_img = None
    same_count = 0

    for step in range(MAX_SCROLL_STEPS):
        activate_dingtalk(window["pid"])
        scroll_at(msg_cx, msg_cy, -15)
        time.sleep(SCROLL_WAIT)
        img, _ = screenshot(window)
        if img is None:
            break
        if debug:
            save_debug(img, f"step2_down_{step:03d}.png")
        if prev_img is not None and content_same(prev_img, img):
            same_count += 1
            if same_count >= SAME_COUNT_STOP:
                logger.info(f"已到底部（第{step+1}步）")
                return
        else:
            same_count = 0
        prev_img = img

    logger.warning("滚到底部超出最大步数")


# ── 步骤3：向上扫描识别作业卡片 ──────────────────────────────────────────────

def find_cards_in_results(results: list[OcrResult]) -> list[OcrResult]:
    """
    从 OCR 结果中找今天的作业卡片标题。
    条件：匹配"X月X日科目"且日期是今天。
    """
    today = date.today()
    cards = []
    seen = set()
    for r in results:
        m = RE_CARD_TITLE.search(r.text)
        if not m:
            continue
        try:
            month = int(m.group(1)) if m.group(1) else date.today().month
            card_date = date(today.year, month, int(m.group(2)))
        except ValueError:
            continue
        subject = m.group(3)
        if card_date == today and subject not in seen:
            seen.add(subject)
            cards.append(r)
            logger.info(f"发现今天作业卡片: {r.text!r}  ({r.center_x},{r.center_y})")
        elif card_date < today:
            logger.info(f"发现非今天卡片 {card_date}，触发终止")
    return cards


def is_old_date(results: list[OcrResult]) -> bool:
    """检测是否出现昨天或更早的消息时间戳"""
    today = date.today()
    for r in results:
        if "昨天" == r.text.strip() or r.text.strip().startswith("昨天 "):
            logger.info(f"发现'昨天'时间戳: {r.text!r}")
            return True
        # 排除含截止/发布等卡片内容
        if any(kw in r.text for kw in ["截止", "发布", "布置", "预计", "已完成"]):
            continue
        for mo in re.finditer(r"(\d{1,2})[/月](\d{1,2})", r.text):
            try:
                d = date(today.year, int(mo.group(1)), int(mo.group(2)))
                if d < today:
                    logger.info(f"发现旧日期 {d}: {r.text!r}")
                    return True
            except ValueError:
                pass
    return False


# ── 步骤4：点开详情，提取作业内容 ────────────────────────────────────────────

def step4_extract_detail(window: dict, card: OcrResult,
                         debug: bool = False, idx: int = 0) -> str:
    """
    点击作业卡片进入详情，提取作业内容，点击返回。
    提取逻辑：找"去打印"坐标，取其上方区域的文字作为作业内容。
    """
    logger.info(f"=== 步骤4：提取详情 [{card.text}] ===")
    activate_dingtalk(window["pid"])
    click(card.center_x, card.center_y)
    time.sleep(RENDER_WAIT)

    img, scale = screenshot(window)
    if img is None:
        return ""
    if debug:
        save_debug(img, f"step4_detail_{idx:02d}.png")

    results = ocr_region(img, scale, win_x=window["x"], win_y=window["y"])
    logger.info(f"详情页 OCR {len(results)} 条")
    for r in results:
        logger.info(f"  [{r.x:4d},{r.y:4d}] conf={r.conf:.2f}  {r.text!r}")

    # 找"去打印"坐标，作为内容区左边界和下边界参考
    print_btn = None
    for r in results:
        if "去打印" in r.text:
            print_btn = r
            logger.info(f"找到'去打印': ({r.x},{r.y})")
            break

    if print_btn is None:
        logger.warning("未找到'去打印'，提取全部文字")
        skip_kw = ["优秀作答", "如何被选", "去补交", "有疑问", "排行榜", "已完成", "预计需", "已截止"]
        lines = [r.text for r in results if r.conf >= 0.4
                 and not any(kw in r.text for kw in skip_kw)]
        return "\n".join(lines)

    # 提取"去打印"上方的内容（y 坐标小于 print_btn.y）
    # 同时过滤掉标题行以外的无关内容
    skip_kw = ["优秀作答", "如何被选", "去补交", "有疑问", "排行榜",
               "已完成", "预计需", "已截止", "今日", "昨日", "周沐曦"]
    lines = []
    for r in results:
        if r.conf < 0.4:
            continue
        if r.y >= print_btn.y:
            continue
        if any(kw in r.text for kw in skip_kw):
            logger.debug(f"  跳过: {r.text!r}")
            continue
        lines.append(r.text)
        logger.debug(f"  保留: {r.text!r}")

    text = "\n".join(lines)
    logger.info(f"提取内容（前120字）: {text[:120]}")

    # 点击"去打印"左侧返回列表
    back_x = print_btn.x - 20
    back_y = print_btn.center_y
    logger.info(f"点击'去打印'左侧返回: ({back_x},{back_y})")
    activate_dingtalk(window["pid"])
    click(back_x, back_y)
    time.sleep(0.8)

    return text


# ── 主流程 ────────────────────────────────────────────────────────────────────

def scrape(debug: bool = False) -> list[RawMessage]:
    """
    通过工作通知群入口抓取今天的作业，返回 RawMessage 列表。
    """
    windows = find_dingtalk_windows()
    if not windows:
        raise RuntimeError("未找到钉钉窗口，请确认钉钉已启动并授权屏幕录制权限")

    window = find_main_window(windows)
    logger.info(f"目标窗口: {window['owner']}  PID={window['pid']}  "
                f"位置=({window['x']},{window['y']})  大小={window['width']}x{window['height']}")

    # 步骤1：找到并点击目标会话
    if not step1_open_target_conv(window, debug):
        return []

    # 消息区中心坐标（用于滚动）：右侧 2/3 区域中心
    # 动态计算：截图后根据实际内容确定，这里先用窗口右侧 2/3 中心
    img, scale = screenshot(window)
    if img is None:
        return []

    # 找消息区边界：用右侧顶部标题"工作通知:深圳市龙岗区"的 x 坐标
    # 这个标题固定在消息区左上角，是最可靠的边界参考
    all_results = ocr_region(img, scale, win_x=window["x"], win_y=window["y"])
    boundary_x = None
    for r in all_results:
        if "工作通知" in r.text and "深圳" in r.text:
            boundary_x = r.x
            logger.info(f"消息区左边界（来自标题栏）: x={boundary_x}  文字={r.text!r}")
            break

    if boundary_x is None:
        # fallback：找最大 x 跳跃点
        xs = sorted(set(r.x for r in all_results))
        gaps = [(xs[i+1] - xs[i], xs[i]) for i in range(len(xs)-1)]
        if gaps:
            gap_val, gap_x = max(gaps)
            boundary_x = gap_x + gap_val // 2
            logger.info(f"消息区左边界（最大x跳跃={gap_val}）: x={boundary_x}")
        else:
            boundary_x = window["x"] + int(window["width"] * 0.37)
            logger.info(f"消息区左边界（默认37%）: x={boundary_x}")

    msg_cx = boundary_x + (window["x"] + window["width"] - boundary_x) // 2
    msg_cy = window["y"] + window["height"] // 2
    logger.info(f"消息区滚动中心: ({msg_cx},{msg_cy})")

    # 步骤2：滚到底部
    step2_scroll_to_bottom(window, msg_cx, msg_cy, debug)

    # 步骤3+4：向上扫描，识别并提取作业
    messages = []
    seen_subjects: set[str] = set()
    prev_img = None
    same_count = 0

    for step in range(MAX_SCROLL_STEPS):
        logger.info(f"--- 扫描步骤 {step+1} ---")
        img, scale = screenshot(window)
        if img is None:
            break
        if debug:
            save_debug(img, f"step3_up_{step:03d}.png")

        # 到顶判断
        if prev_img is not None and content_same(prev_img, img):
            same_count += 1
            if same_count >= SAME_COUNT_STOP:
                logger.info("已到顶部，扫描结束")
                break
        else:
            same_count = 0
        prev_img = img

        # OCR 全图（不裁剪，避免截断消息区左侧内容）
        results = ocr_region(img, scale, win_x=window["x"], win_y=window["y"])
        # 只保留消息区的结果（x > boundary_x）
        results = [r for r in results if r.x >= boundary_x]
        logger.info(f"消息区 OCR {len(results)} 条")
        for r in results:
            logger.info(f"  [{r.x:4d},{r.y:4d}] conf={r.conf:.2f}  {r.text!r}")

        # 终止：发现旧日期
        if is_old_date(results):
            logger.info("发现旧日期，停止扫描")
            break

        # 找今天的作业卡片
        cards = find_cards_in_results(results)
        for card in cards:
            m = RE_CARD_TITLE.search(card.text)
            subject = m.group(3) if m else card.text
            if subject in seen_subjects:
                logger.info(f"科目[{subject}]已收集，跳过")
                continue
            seen_subjects.add(subject)

            text = step4_extract_detail(window, card, debug=debug,
                                        idx=len(messages))
            if text:
                messages.append(RawMessage(
                    msg_id=f"v2_{step}_{subject}",
                    sender_id="capture",
                    sender_name="工作通知群",
                    timestamp=datetime.now(),
                    msg_type="text",
                    text=text,
                ))
                logger.info(f"已收集[{subject}]，共{len(messages)}科")

            if len(messages) >= MAX_HOMEWORK:
                logger.info(f"已达最大数量 {MAX_HOMEWORK}，停止")
                return messages

        # 向上滚动（1/3 内容区高度）
        content_h = window["height"] * 0.65
        delta = max(8, int(content_h / 3 / 10))
        activate_dingtalk(window["pid"])
        scroll_at(msg_cx, msg_cy, delta)
        time.sleep(SCROLL_WAIT)

    logger.info(f"=== 抓取完成，共 {len(messages)} 科作业 ===")
    return messages


if __name__ == "__main__":
    debug_mode = "--debug" in sys.argv
    msgs = scrape(debug=debug_mode)

    if not msgs:
        print("\n未抓取到任何作业内容")
        print("建议加 --debug：python -m capture.scraper --debug")
        sys.exit(0)

    print(f"\n共抓取 {len(msgs)} 科作业：\n")
    for m in msgs:
        print(f"[{m.msg_id}]\n{m.text}\n{'─'*40}")

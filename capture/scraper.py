"""
钉钉家校本自动抓取 - AI家校本入口版

流程：
  1. 激活钉钉窗口
  2. 点击底部"AI家校本"按钮
  3. 点击"全部"tab
  4. 滚动扫描列表，收集所有卡片 card_key（日期+科目+发布人）
  5. 滚回顶部，按顺序逐张点击卡片，提取完整内容（含展开+滚动）
  6. 关闭面板
  7. 返回 RawMessage 列表

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

import numpy as np
from PIL import Image

from capture.find_window import find_dingtalk_windows, find_main_window
from capture.screenshot import capture_window, activate_dingtalk
from capture.ocr import recognize, find_text, OcrResult
from models.models import RawMessage
from utils.logger import logger
import config

# ── 参数 ──────────────────────────────────────────────────────────────────────
CARD_OPEN_WAIT     = 1.5   # 点击卡片后等待渲染（秒）
EXPAND_WAIT        = 0.5   # 点击"∨"展开后等待渲染（秒）
SCROLL_WAIT        = 0.8   # 每次滚动后等待渲染（秒）
MAX_HOMEWORK       = 6     # 最多提取作业数量
MAX_DETAIL_SCROLLS = 5     # 详情页最多滚动次数
CARD_SAFE_MARGIN   = 60    # 点击前卡片距面板顶部/底部的最小安全距离（逻辑像素）
SCROLL_DELTA       = -5    # 向下滚动量
COMPARE_HEIGHT     = 100   # 像素差对比区域高度（物理像素）
SIMILARITY_THRESH  = 5.0   # 底部像素差阈值，低于此值视为内容未变化
DEBUG_DIR          = os.path.join("output", "tmp", "captures")

SUBJECT_KEYWORDS = config.SUBJECT_ORDER
RE_CARD_TITLE = re.compile(
    r"(\d{1,2})月(\d{1,2})日\s*(" +
    "|".join(re.escape(s) for s in SUBJECT_KEYWORDS) + r")"
)
RE_PUBLISHER = re.compile(r"(.+)发布$")

SKIP_KEYWORDS = [
    "去补交", "有疑问", "去打印", "优秀作答", "如何被选",
    "排行榜", "已完成", "预计需", "已截止", "昨日", "首页", "错题本",
    "反馈完成时长", "无需在线提交",
]


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


def scroll_panel(window: dict, delta: int = SCROLL_DELTA) -> None:
    """在右侧面板中心发送滚轮事件"""
    activate_dingtalk(window["pid"])
    panel_x_logical = int(window["width"] * 0.37)
    cx = window["x"] + panel_x_logical + (window["width"] - panel_x_logical) // 2
    cy = window["y"] + window["height"] // 2
    event = Quartz.CGEventCreateScrollWheelEvent(
        None, Quartz.kCGScrollEventUnitLine, 1, delta)
    Quartz.CGEventSetLocation(event, (cx, cy))
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)


def scroll_to_top(window: dict) -> None:
    """滚回列表顶部（多次向上滚动）"""
    for _ in range(15):
        scroll_panel(window, delta=10)
    time.sleep(SCROLL_WAIT)


def ocr_panel(window: dict, wait: float = 0.0,
              debug_name: str = None, debug: bool = False):
    """截图右侧面板并 OCR，返回 (results, panel_x_logical)"""
    if wait > 0:
        time.sleep(wait)
    img = capture_window(window["window_id"])
    if img is None:
        logger.warning("截图失败")
        return [], 0

    scale = img.width / window["width"]
    panel_x_logical = int(window["width"] * 0.37)
    panel_x_physical = int(panel_x_logical * scale)
    panel_img = img.crop((panel_x_physical, 0, img.width, img.height))

    if debug and debug_name:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        panel_img.save(os.path.join(DEBUG_DIR, debug_name))
        logger.info(f"调试截图: {debug_name}")

    results = recognize(panel_img,
                        window_x=window["x"] + panel_x_logical,
                        window_y=window["y"],
                        scale=scale)
    logger.info(f"OCR [{debug_name or ''}] {len(results)} 条")
    for r in results:
        logger.debug(f"  [{r.x:4d},{r.y:4d}] conf={r.conf:.2f}  {r.text!r}")
    return results, panel_x_logical


def capture_panel_img(window: dict) -> tuple[Image.Image | None, float]:
    """截图右侧面板，返回 (panel_img, scale)，用于像素差对比"""
    img = capture_window(window["window_id"])
    if img is None:
        return None, 1.0
    scale = img.width / window["width"]
    panel_x_physical = int(window["width"] * 0.37 * scale)
    panel_img = img.crop((panel_x_physical, 0, img.width, img.height))
    return panel_img, scale


def is_same_bottom(img_a: Image.Image, img_b: Image.Image) -> bool:
    """对比两张面板截图底部区域像素差，判断内容是否不再变化"""
    def strip(img):
        h = img.height
        return np.array(
            img.crop((0, h - COMPARE_HEIGHT, img.width, h)), dtype=np.float32)
    diff = np.abs(strip(img_a) - strip(img_b)).mean()
    logger.info(f"底部像素差: {diff:.2f}")
    return diff < SIMILARITY_THRESH


# ── card_key 相关 ─────────────────────────────────────────────────────────────

def extract_cards_from_ocr(results: list[OcrResult]) -> list[dict]:
    """
    从 OCR 结果中提取完整卡片信息。
    每张卡片需要：标题行（日期+科目）+ 下方紧邻的发布人行。
    只返回 key 完整的卡片（发布人缺失的忽略，等下一轮滚动后完整显示）。

    返回列表，每项：
      { "key": "4月14日数学_朱丹发布", "title": "4月14日数学",
        "publisher": "朱丹", "result": OcrResult(标题行) }
    """
    cards = []
    for i, r in enumerate(results):
        m = RE_CARD_TITLE.search(r.text)
        if not m:
            continue
        title = f"{m.group(1)}月{m.group(2)}日{m.group(3)}"

        # 在标题行下方找发布人（y 坐标更大，且在合理范围内）
        publisher = None
        for j in range(i + 1, min(i + 6, len(results))):
            candidate = results[j]
            # 发布人行应在标题行下方 10~120px 内
            if not (r.y < candidate.y < r.y + 120):
                continue
            pm = RE_PUBLISHER.search(candidate.text)
            if pm:
                publisher = pm.group(1).strip()
                break

        if publisher is None:
            logger.debug(f"  卡片 {title!r} 发布人未识别，跳过（等下一轮）")
            continue

        key = f"{title}_{publisher}发布"
        cards.append({
            "key": key,
            "title": title,
            "publisher": publisher,
            "result": r,
        })
        logger.info(f"  识别卡片: {key!r}  ({r.center_x},{r.center_y})")
    return cards


def find_card_in_ocr(results: list[OcrResult], card_info: dict) -> OcrResult | None:
    """
    在 OCR 结果中按 card_key 定位目标卡片的标题行 OcrResult。
    先匹配标题（日期+科目），再验证下方发布人。
    """
    title = card_info["title"]
    publisher = card_info["publisher"]

    for i, r in enumerate(results):
        if title not in r.text:
            continue
        # 验证下方发布人
        for j in range(i + 1, min(i + 6, len(results))):
            candidate = results[j]
            if not (r.y < candidate.y < r.y + 120):
                continue
            if publisher in candidate.text and "发布" in candidate.text:
                return r
    return None


# ── 步骤函数 ──────────────────────────────────────────────────────────────────

def step1_open_jiaxiaob(window: dict, debug: bool = False) -> bool:
    """点击底部'AI家校本'按钮"""
    logger.info("=== 步骤1：点击 AI家校本 ===")
    activate_dingtalk(window["pid"])
    results, _ = ocr_panel(window, wait=0.5, debug_name="step1.png", debug=debug)
    targets = find_text(results, "家校本")
    if not targets:
        logger.error("未找到'AI家校本'按钮")
        return False
    target = max(targets, key=lambda r: r.conf)
    logger.info(f"找到'家校本': ({target.center_x},{target.center_y})")
    activate_dingtalk(window["pid"])
    click(target.center_x, target.center_y)
    time.sleep(1.0)
    return True


def step2_click_all_tab(window: dict, debug: bool = False) -> None:
    """点击'全部'tab"""
    logger.info("=== 步骤2：点击'全部'tab ===")
    results, _ = ocr_panel(window, wait=CARD_OPEN_WAIT,
                           debug_name="step2.png", debug=debug)
    targets = find_text(results, "全部")
    if not targets:
        logger.warning("未找到'全部'tab，可能已在全部视图")
        return
    activate_dingtalk(window["pid"])
    click(targets[0].center_x, targets[0].center_y)
    time.sleep(0.5)


def step3_scan_card_list(window: dict, debug: bool = False) -> list[dict]:
    """
    滚动扫描作业列表，收集所有完整卡片的 card_key。
    返回有序列表，每项为 card_info dict。
    """
    logger.info("=== 步骤3：滚动扫描作业列表 ===")
    collected: list[dict] = []
    seen_keys: set[str] = set()

    for scroll_round in range(MAX_HOMEWORK + 3):
        debug_name = f"step3_round{scroll_round}.png" if debug else None
        results, _ = ocr_panel(window, wait=SCROLL_WAIT if scroll_round > 0 else 0.8,
                               debug_name=debug_name, debug=debug)
        new_cards = extract_cards_from_ocr(results)

        new_found = 0
        for card in new_cards:
            if card["key"] not in seen_keys:
                seen_keys.add(card["key"])
                collected.append(card)
                new_found += 1
                logger.info(f"  新增卡片 [{len(collected)}]: {card['key']!r}")
                if len(collected) >= MAX_HOMEWORK:
                    logger.info(f"已达 MAX_HOMEWORK={MAX_HOMEWORK}，停止扫描")
                    scroll_to_top(window)
                    return collected

        if scroll_round > 0 and new_found == 0:
            logger.info("本轮无新卡片，扫描完成")
            break

        scroll_panel(window)

    scroll_to_top(window)
    logger.info(f"步骤3完成，共收集 {len(collected)} 张卡片")

    # 打印作业列表
    print(f"\n{'─'*50}")
    print(f"  作业列表（共 {len(collected)} 张）")
    print(f"{'─'*50}")
    for i, c in enumerate(collected, 1):
        print(f"  {i}. {c['key']}")
    print(f"{'─'*50}\n")

    return collected


def _find_card_with_scroll(window: dict, card_info: dict) -> OcrResult | None:
    """
    在当前列表中定位目标卡片。
    找不到时滚回顶部重新查找，仍找不到返回 None。
    """
    # 先在当前视图找
    results, _ = ocr_panel(window)
    r = find_card_in_ocr(results, card_info)
    if r:
        return r

    # 滚回顶部重新找
    logger.info(f"  当前视图未找到 {card_info['key']!r}，滚回顶部重试")
    scroll_to_top(window)
    for _ in range(MAX_HOMEWORK + 2):
        results, _ = ocr_panel(window, wait=SCROLL_WAIT)
        r = find_card_in_ocr(results, card_info)
        if r:
            return r
        scroll_panel(window)

    return None


def _ensure_card_in_safe_zone(window: dict, card_result: OcrResult,
                               card_info: dict) -> OcrResult | None:
    """
    确保卡片 y 坐标在安全区内（距面板顶部/底部各留 CARD_SAFE_MARGIN px）。
    太靠边则微调滚动，重新 OCR 取新坐标。
    """
    panel_top = window["y"]
    panel_bottom = window["y"] + window["height"]
    safe_top = panel_top + CARD_SAFE_MARGIN
    safe_bottom = panel_bottom - CARD_SAFE_MARGIN

    if card_result.center_y < safe_top:
        logger.info("  卡片太靠近顶部，向上微调")
        scroll_panel(window, delta=3)
        time.sleep(0.3)
        results, _ = ocr_panel(window)
        return find_card_in_ocr(results, card_info)

    if card_result.center_y > safe_bottom:
        logger.info("  卡片太靠近底部，向下微调")
        scroll_panel(window, delta=-3)
        time.sleep(0.3)
        results, _ = ocr_panel(window)
        return find_card_in_ocr(results, card_info)

    return card_result


def _find_back_button(results: list[OcrResult]) -> OcrResult | None:
    """
    按优先级匹配返回按钮：
    1. 以 < 或 〈 开头 且 含 '的练习'
    2. 以 < 或 〈 开头
    3. 含 '的练习'
    """
    def starts_with_arrow(text):
        t = text.strip()
        return t.startswith("<") or t.startswith("〈")

    # 优先级1
    for r in results:
        if starts_with_arrow(r.text) and "的练习" in r.text:
            return r
    # 优先级2
    for r in results:
        if starts_with_arrow(r.text):
            return r
    # 优先级3
    for r in results:
        if "的练习" in r.text:
            return r
    return None


def _extract_detail_content(window: dict, idx: int, debug: bool) -> str:
    """
    在详情页提取完整作业内容：
    1. 检查并点击"∨"展开按钮
    2. 在面板内滚动，对比截图像素差，直到内容不再变化
    3. 合并所有轮次 OCR 文字，按行去重
    """
    all_lines: list[str] = []
    seen_lines: set[str] = set()

    def collect_lines(results):
        for r in results:
            if r.conf < 0.4:
                continue
            if any(kw in r.text for kw in SKIP_KEYWORDS):
                continue
            line = r.text.strip()
            if line and line not in seen_lines:
                seen_lines.add(line)
                all_lines.append(line)

    # 首次 OCR
    results, _ = ocr_panel(window, wait=CARD_OPEN_WAIT,
                            debug_name=f"step4_detail_{idx}_0.png", debug=debug)

    # 检查展开按钮（∨ 可能被 OCR 识别为 v/V/∨ 等多种形式）
    expand_targets = [
        r for r in results
        if r.text.strip() in ("∨", "v", "V", "ν") or "展开" in r.text
    ]
    if expand_targets:
        # 取 y 坐标最大的（展开箭头在内容区中间偏下）
        expand_btn = max(expand_targets, key=lambda r: r.y)
        logger.info(f"  发现展开按钮 {expand_btn.text!r}，点击展开")
        activate_dingtalk(window["pid"])
        click(expand_btn.center_x, expand_btn.center_y)
        results, _ = ocr_panel(window, wait=EXPAND_WAIT,
                                debug_name=f"step4_detail_{idx}_expand.png", debug=debug)

    collect_lines(results)

    # 滚动提取剩余内容
    prev_img, _ = capture_panel_img(window)
    for scroll_i in range(MAX_DETAIL_SCROLLS):
        scroll_panel(window)
        time.sleep(SCROLL_WAIT)
        curr_img, _ = capture_panel_img(window)

        if curr_img and prev_img and is_same_bottom(prev_img, curr_img):
            logger.info(f"  详情页内容不再变化，停止滚动（第{scroll_i+1}次）")
            break

        results, _ = ocr_panel(window,
                                debug_name=f"step4_detail_{idx}_scroll{scroll_i+1}.png",
                                debug=debug)
        collect_lines(results)
        prev_img = curr_img

    return "\n".join(all_lines)


def step4_extract_cards(window: dict, card_list: list[dict],
                        debug: bool = False) -> list[RawMessage]:
    """
    按顺序处理每张卡片：定位 → 安全区检查 → 点击 → 提取内容 → 返回
    """
    logger.info("=== 步骤4：逐张提取作业内容 ===")
    messages = []
    panel_x_logical = int(window["width"] * 0.37)

    for i, card_info in enumerate(card_list):
        logger.info(f"--- 处理第 {i+1}/{len(card_list)} 张: {card_info['key']!r} ---")

        # 定位卡片
        card_result = _find_card_with_scroll(window, card_info)
        if card_result is None:
            logger.warning(f"  找不到卡片 {card_info['key']!r}，跳过")
            continue

        # 安全区检查
        card_result = _ensure_card_in_safe_zone(window, card_result, card_info)
        if card_result is None:
            logger.warning(f"  安全区调整后仍找不到 {card_info['key']!r}，跳过")
            continue

        # 点击卡片
        activate_dingtalk(window["pid"])
        click(card_result.center_x, card_result.center_y)

        # 提取详情内容
        text = _extract_detail_content(window, i, debug)
        logger.info(f"  提取内容（前120字）: {text[:120]}")

        if text:
            messages.append(RawMessage(
                msg_id=f"jiaxiaob_{i}_{card_info['key']}",
                sender_id="capture",
                sender_name="AI家校本",
                timestamp=datetime.now(),
                msg_type="text",
                text=text,
            ))

        # 点击返回按钮，同时从返回按钮文字提取学生名
        back_results, _ = ocr_panel(window)
        back_btn = _find_back_button(back_results)
        student_name = ""
        if back_btn:
            # "< 周晓逸的练习" → 提取 "周晓逸"
            m = re.search(r"[<〈]\s*(.+?)的练习", back_btn.text)
            if m:
                student_name = m.group(1).strip()
                logger.info(f"  识别学生: {student_name!r}")
            bx = back_btn.x + 10
            by = back_btn.center_y
            logger.info(f"  点击返回: ({bx},{by})  原文={back_btn.text!r}")
            activate_dingtalk(window["pid"])
            click(bx, by)
        else:
            fallback_x = window["x"] + panel_x_logical + 30
            fallback_y = window["y"] + 30
            logger.warning(f"  未找到返回按钮，点击左上角 ({fallback_x},{fallback_y})")
            activate_dingtalk(window["pid"])
            click(fallback_x, fallback_y)

        if text:
            messages[-1].sender_name = student_name or "AI家校本"

        time.sleep(0.8)

    return messages


def step5_close_panel(window: dict) -> None:
    """关闭家校本面板"""
    logger.info("=== 步骤5：关闭面板 ===")
    results, _ = ocr_panel(window, wait=0.3)
    for r in results:
        if r.text.strip() in ("×", "X"):
            activate_dingtalk(window["pid"])
            click(r.center_x, r.center_y)
            logger.info("面板已关闭（×按钮）")
            return
    # fallback：点击家校本面板左侧外部区域
    outside_x = window["x"] + int(window["width"] * 0.37) // 2
    outside_y = window["y"] + window["height"] // 2
    logger.warning(f"未找到×按钮，点击面板左侧外部 ({outside_x},{outside_y})")
    activate_dingtalk(window["pid"])
    click(outside_x, outside_y)


# ── 主流程 ────────────────────────────────────────────────────────────────────

def scrape(debug: bool = False) -> list[RawMessage]:
    """通过 AI家校本 入口抓取作业，返回 RawMessage 列表"""
    windows = find_dingtalk_windows()
    if not windows:
        raise RuntimeError("未找到钉钉窗口，请确认钉钉已启动并授权屏幕录制权限")

    window = find_main_window(windows)
    logger.info(f"目标窗口: {window['owner']}  PID={window['pid']}  "
                f"位置=({window['x']},{window['y']})  大小={window['width']}x{window['height']}")

    activate_dingtalk(window["pid"])
    time.sleep(0.5)

    if not step1_open_jiaxiaob(window, debug):
        return []

    step2_click_all_tab(window, debug)

    card_list = step3_scan_card_list(window, debug)
    if not card_list:
        logger.warning("未识别到任何作业卡片")
        step5_close_panel(window)
        return []

    messages = step4_extract_cards(window, card_list, debug)

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
        print(f"[{m.sender_name}] [{m.msg_id}]\n{m.text}\n{'─'*40}")

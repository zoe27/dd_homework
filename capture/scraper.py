"""
钉钉作业抓取 v2 - 工作通知群入口

流程：
  1. 左侧会话找「工作通知:深圳市龙岗区」并进入
  2. 消息区滚到底部
  3. 自下而上：以「xx老师 发布于 YYYY-MM-DD」为严格去重主键，「查看详情」为点击目标
  4. 进详情后：在「去打印」上方最后一条作业正文下方轻点一次（促折叠展开），再向下滚合并 OCR
  5. 在「去打印」左侧点一下返回列表，向上滚，继续
  6. 非今天：由 CAPTURE_V2_ON_NON_TODAY 控制停扫或继续找满（仅今日，最多6条）

运行：python -m capture.scraper [--debug]
     CAPTURE_V2_ON_NON_TODAY=continue  # 测试：不因非今天而停止
"""

import sys
import os
import re
import time
import hashlib
from datetime import date, datetime, time as dtime

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
RENDER_WAIT         = 1.5    # 点击后等待渲染（秒）
SCROLL_WAIT         = 0.6    # 滚动后等待（秒）
MAX_HOMEWORK        = int(getattr(config, "CAPTURE_V2_MAX_HOMEWORK", 6))  # 最多提取作业数
MAX_SCROLL_STEPS    = 40     # 最大滚动次数
SAME_COUNT_STOP     = 2      # 连续相同次数判定到顶/底
DEBUG_DIR           = os.path.join("output", "tmp", "captures")
MAX_DETAIL_SCROLLS  = 15     # 详情页向下滑动次数（防止死循环）
COMPARE_HEIGHT      = 100    # 详情页底部条带高度（与像素差比较）
SIMILARITY_THRESH   = 5.0
PAIR_Y_MAX_GAP      = 160    # 「发布于」行与「查看详情」按钮之间的最大 y 距（逻辑像素）
DETAIL_SCROLL_DELTA = -10    # 详情页向下滑动（同消息区向下滑出更多内容）
DETAIL_TAP_WAIT     = 0.5    # 详情内轻点后等界面稳定
# 最后一条作业 OCR 框底边往下一点（折叠热区，须在「去打印」之上）
TAP_AFTER_LAST_BODY_DY = 16

# 目标会话关键词
TARGET_CONV_KEYWORDS = ["工作通知", "深圳市龙岗区"]

# 主键：xx老师 发布于 年月日（与界面一致，用于去重；保留同 key 中较新一条即列表较下优先）
# RE_PUBLISH_FOR_KEY = re.compile(
#     r"(?P<name>.+?老师)\s*发布于\s*"
#     r"(?P<Y>\d{4})[./年\-]?(?P<M>\d{1,2})[./月\-]?(?P<D>\d{1,2})"
# )

RE_PUBLISH_FOR_KEY = re.compile(
    r"(?P<name>[\u4e00-\u9fa5A-Za-z]+老师)[^\w\d]{0,5}发\s*布\s*于\s*"
    r"(?P<Y>\d{4})[./年\-]?(?P<M>\d{1,2})[./月\-]?(?P<D>\d{1,2})"
)
# 详情页标题区：老师 + 发布/布置 行，不作为「最后一行作业正文」
RE_TEACHER_META = re.compile(
    r".+老师\s*(发布于|发布|布置于|布置)"
)

# 详情正文 OCR 时排除的噪声（与 collect_lines_above_print 一致）
DETAIL_SKIP_KW = frozenset(
    (
        "优秀作答", "如何被选", "去补交", "有疑问", "排行榜", "今日", "昨日",
        "无需在线", "去打印", "预计需", "已截止", "反馈完成", "无需在线提交",
    )
)


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


# ── 步骤3：配对准则「老师 发布于 日期」+「查看详情」───────────────────────────

def _tighten_view_detail(s: str) -> str:
    return re.sub(r"\s+", "", (s or ""))


def is_view_detail_button(r: OcrResult) -> bool:
    t = _tighten_view_detail(r.text)
    return t in ("查看详情", "点击详情")


def parse_key_and_date_from_line(text: str) -> tuple[str, date] | None:
    """
    从一行 OCR 解析严格主键「xx老师 发布于 YYYY-MM-DD」与 date。
    """
    m = RE_PUBLISH_FOR_KEY.search((text or "").replace("：", ":"))
    if not m:
        return None
    y, mo, d = int(m.group("Y")), int(m.group("M")), int(m.group("D"))
    try:
        dday = date(y, mo, d)
    except ValueError:
        return None
    name = m.group("name").strip()
    if not name.endswith("老师"):
        return None
    key = f"{name} 发布于 {dday.isoformat()}"
    return (key, dday)


def find_homework_entries(results: list[OcrResult]) -> list[dict]:
    """
    为每个「查看详情」找其正上方、距离最近的含「老师 + 发布于 + 日期」行。
    自屏幕下方优先（y 大在前），同 key 只保留先出现的（时间更新）。
    返回: [{ "key", "card_date", "detail_btn", "teacher_name" }, ...]
    """
    by_key: dict[str, dict] = {}
    details = [r for r in results if is_view_detail_button(r) and r.conf >= 0.3]
    details.sort(key=lambda r: r.y, reverse=True)
    for dbtn in details:
        others = [r for r in results if r is not dbtn and r.conf >= 0.25]
        above = [
            r for r in others
            if r.y < dbtn.y
            and (dbtn.y - r.y) <= PAIR_Y_MAX_GAP
            and "发布于" in r.text
            and "老师" in r.text
        ]
        if not above:
            logger.debug("  无配对发布行，忽略该查看详情")
            continue
        teacher = max(above, key=lambda r: r.y)
        parsed = parse_key_and_date_from_line(teacher.text)
        if not parsed:
            logger.debug(f"  无法解析主键: {teacher.text!r}")
            continue
        key, card_date = parsed
        tname = key.split(" 发布于 ")[0]
        if key not in by_key:
            by_key[key] = {
                "key": key,
                "card_date": card_date,
                "detail_btn": dbtn,
                "teacher_name": tname,
            }
            logger.info(f"  作业条目 key={key!r} 详情按钮=({dbtn.center_x},{dbtn.center_y})")
    return list(by_key.values())


# ── 步骤4：详情内滚动到顶 + 提取 + 返回列表 ─────────────────────────────────

def _crop_msg_from_boundary(img: Image.Image, window: dict, boundary_x: int) -> Image.Image:
    """消息区右半部分，用于对比是否滚到底（逻辑 x >= boundary）。"""
    x0 = int((boundary_x - window["x"]) * img.width / window["width"])
    x0 = max(0, min(x0, img.width - 2))
    return img.crop((x0, 0, img.width, img.height))


def _detail_bottom_array(panel: Image.Image) -> np.ndarray:
    h = min(COMPARE_HEIGHT, panel.height)
    h0 = max(0, panel.height - h)
    return np.array(panel.crop((0, h0, panel.width, panel.height)), dtype=np.float32)


def is_detail_bottom_unchanged(img_a, img_b, window, boundary_x) -> bool:
    a = _detail_bottom_array(_crop_msg_from_boundary(img_a, window, boundary_x))
    b = _detail_bottom_array(_crop_msg_from_boundary(img_b, window, boundary_x))
    if a.shape != b.shape:
        return False
    diff = float(np.abs(a - b).mean())
    logger.info(f"详情区底部像素差: {diff:.2f}")
    return diff < SIMILARITY_THRESH


def ocr_from_boundary(results: list[OcrResult], boundary_x: int) -> list[OcrResult]:
    return [r for r in results if r.x >= boundary_x - 2]


def detail_tap_below_last_body_line(
    window: dict, boundary_x: int, img: Image.Image, scale: float
) -> tuple[Image.Image, float]:
    """
    取「去打印」上方、过滤后的最后一条作业 OCR，在其底边稍下点一下（有折叠会展开，无则多半无影响）。
    无可用行则原样返回。
    """
    rs = ocr_region(img, scale, win_x=window["x"], win_y=window["y"])
    rm = ocr_from_boundary(rs, boundary_x)
    pb = next((r for r in rm if "去打印" in r.text), None)
    boxes: list[OcrResult] = []
    for r in rm:
        if r.conf < 0.3:
            continue
        if pb is not None and r.y >= pb.y:
            continue
        if pb is not None and r.x <= pb.x - 2:
            continue
        if not (r.text or "").strip():
            continue
        if any(kw in r.text for kw in DETAIL_SKIP_KW):
            continue
        if RE_TEACHER_META.search(r.text):
            continue
        boxes.append(r)
    if not boxes:
        return img, scale
    last = max(boxes, key=lambda r: r.y + r.h)
    cx = max(boundary_x + 24, min(last.center_x, window["x"] + window["width"] - 24))
    cy = last.y + last.h + TAP_AFTER_LAST_BODY_DY
    if pb is not None:
        cy = min(cy, pb.y - 14)
    cy = max(cy, last.y + max(10, last.h // 2))
    if pb is not None and cy >= pb.y - 2:
        return img, scale
    logger.info(f"  详情轻点展开区 ({cx},{cy}) 参考行 {last.text[:40]!r}")
    activate_dingtalk(window["pid"])
    click(int(cx), int(cy))
    time.sleep(DETAIL_TAP_WAIT)
    new_img, new_scale = screenshot(window)
    return (img, scale) if new_img is None else (new_img, new_scale)


def collect_lines_above_print(
    results: list[OcrResult], print_btn: OcrResult | None, boundary_x: int
) -> list[str]:
    """只保留消息区内、在「去打印」上方的一行行文本（按 y 排序）。"""
    out = []
    for r in sorted(ocr_from_boundary(results, boundary_x), key=lambda r: (r.y, r.x)):
        if r.conf < 0.3:
            continue
        if print_btn is not None and r.y >= print_btn.y:
            continue
        if print_btn is not None and r.x <= (print_btn.x - 2):
            continue
        if any(kw in r.text for kw in DETAIL_SKIP_KW):
            continue
        if RE_TEACHER_META.search(r.text):
            continue
        out.append(r.text.strip())
    return out


def step4_open_detail_and_extract(
    window: dict,
    boundary_x: int,
    msg_cx: int,
    msg_cy: int,
    detail_btn: OcrResult,
    key: str,
    debug: bool = False,
    idx: int = 0,
) -> str:
    """点击「查看详情」→ 最后一行作业下轻点一次 → 向下滚合并 OCR →「去打印」左侧返回。"""
    logger.info(f"=== 进入详情: {key!r} ===")
    activate_dingtalk(window["pid"])
    click(detail_btn.center_x, detail_btn.center_y)
    time.sleep(RENDER_WAIT)

    img, scale = screenshot(window)
    if img is None:
        return ""

    if debug:
        save_debug(img, f"step4_detail_{idx:02d}.png")

    img, scale = detail_tap_below_last_body_line(window, boundary_x, img, scale)
    if debug:
        save_debug(img, f"step4_detail_{idx:02d}_after_tap.png")

    all_lines: list[str] = []
    seen: set[str] = set()
    prev_img: Image.Image | None = None

    def _one_pass(im: Image.Image, sc: float, tag: str) -> None:
        rs = ocr_region(im, sc, win_x=window["x"], win_y=window["y"])
        rm = ocr_from_boundary(rs, boundary_x)
        pb = next((r for r in rm if "去打印" in r.text), None)
        for line in collect_lines_above_print(rm, pb, boundary_x):
            if line and line not in seen:
                seen.add(line)
                all_lines.append(line)
        if debug and tag:
            save_debug(im, f"step4_detail_{idx:02d}_{tag}.png")

    if img is not None:
        _one_pass(img, scale, "s0")
        prev_img = img

    for s_idx in range(MAX_DETAIL_SCROLLS):
        activate_dingtalk(window["pid"])
        scroll_at(msg_cx, msg_cy, DETAIL_SCROLL_DELTA)
        time.sleep(SCROLL_WAIT)
        curr, s2 = screenshot(window)
        if curr is None:
            break
        _one_pass(curr, s2, f"s{s_idx+1}")
        if prev_img is not None and is_detail_bottom_unchanged(
            curr, prev_img, window, boundary_x
        ):
            logger.info("详情区底部无变化，停止向下滚动")
            break
        prev_img = curr

    text = "\n".join(all_lines)
    logger.info(f"  合并详情（前120字）: {text[:120]!r}")

    # 再截一帧用于点击「去打印」左侧
    img, scale = screenshot(window)
    if img is not None:
        results = ocr_region(img, scale, win_x=window["x"], win_y=window["y"])
        print_btn = next((r for r in ocr_from_boundary(results, boundary_x) if "去打印" in r.text), None)
    else:
        print_btn = None

    if print_btn is None:
        back_x = boundary_x - 20
        back_y = window["y"] + window["height"] // 2
        logger.warning(f"未找到「去打印」，点击边界左侧 ({back_x},{back_y}) 尝试返回")
    else:
        back_x = print_btn.x - 24
        back_y = print_btn.center_y
        logger.info(f"在「去打印」左侧点击返回: ({back_x},{back_y})")

    activate_dingtalk(window["pid"])
    click(int(back_x), int(back_y))
    time.sleep(0.8)
    return text


# ── 主流程 ────────────────────────────────────────────────────────────────────

def _non_today_mode() -> str:
    v = (getattr(config, "CAPTURE_V2_ON_NON_TODAY", "stop") or "stop").lower().strip()
    return "continue" if v in ("continue", "fill", "fill6", "c") else "stop"


def scrape(
    debug: bool = False,
    on_non_today: str | None = None,
) -> list[RawMessage]:
    """
    工作通知 v2：主键=「xx老师 发布于 YYYY-MM-DD」，点「查看详情」进内页，最多 MAX_HOMEWORK 条今日作业。
    on_non_today: None=读 config，stop=早一天即停扫，continue=跳过非今天继续找满(仅今日)条
    """
    mode = (on_non_today or _non_today_mode())
    stop_on_past = mode != "continue"
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
    today = date.today()
    seen_keys: set[str] = set()
    messages: list[RawMessage] = []
    prev_img = None
    same_count = 0
    should_stop = False

    for step in range(MAX_SCROLL_STEPS):
        if should_stop or len(messages) >= MAX_HOMEWORK:
            break
        logger.info(f"--- 列表扫描 {step+1}（非今={mode}，今日最多 {MAX_HOMEWORK}）---")
        img, scale = screenshot(window)
        if img is None:
            break
        if debug:
            save_debug(img, f"step3_up_{step:03d}.png")

        if prev_img is not None and content_same(prev_img, img):
            same_count += 1
            if same_count >= SAME_COUNT_STOP:
                logger.info("已到达消息列表顶部，结束")
                break
        else:
            same_count = 0
        prev_img = img

        # results = ocr_region(img, scale, win_x=window["x"], win_y=window["y"])
        x1 = int((boundary_x - window["x"]) * scale)

        results = ocr_region(
            img,
            scale,
            win_x=window["x"],
            win_y=window["y"],
            crop=(x1, 0, img.width, img.height)
        )
        # results = ocr_region(img, scale, win_x=window["x"], win_y=window["y"],
        #                      crop=(boundary_x, 0, img.width, img.height))
        results = [r for r in results if r.x >= boundary_x]
        logger.info(f"消息区 OCR {len(results)} 条")
        for r in results:
            logger.info(f"  [{r.x:4d},{r.y:4d}] conf={r.conf:.2f}  {r.text!r}")

        entries = find_homework_entries(results)
        # 自下而上：y 大优先；同屏只处理一条，避免点进详情后坐标失效
        entries.sort(key=lambda e: e["detail_btn"].y, reverse=True)

        picked: dict | None = None
        for ent in entries:
            if ent["key"] in seen_keys:
                continue
            cdate = ent["card_date"]
            if not stop_on_past:
                # --all-dates / continue 模式：不过滤日期，直接抓
                picked = ent
                break
            if cdate < today:
                logger.info(
                    f"主键 {ent['key']!r} 发布日早于今天，按 stop 规则结束扫描"
                )
                should_stop = True
                break
            if cdate > today:
                continue
            picked = ent
            break

        if should_stop or len(messages) >= MAX_HOMEWORK:
            break

        if picked is not None:
            seen_keys.add(picked["key"])
            cdate = picked["card_date"]
            text = step4_open_detail_and_extract(
                window,
                boundary_x,
                msg_cx,
                msg_cy,
                picked["detail_btn"],
                picked["key"],
                debug=debug,
                idx=len(messages),
            )
            if text.strip():
                h = hashlib.md5(
                    picked["key"].encode("utf-8")
                ).hexdigest()[:12]
                n = len(messages)
                messages.append(
                    RawMessage(
                        msg_id=f"v2_{n}_{h}",
                        sender_id="capture",
                        sender_name=picked.get("teacher_name", "工作通知群"),
                        timestamp=datetime.combine(
                            cdate, dtime(0, 0, 0)
                        ),
                        msg_type="text",
                        text=text.strip(),
                    )
                )
                logger.info(f"已收第 {len(messages)} 条: {picked['key']!r}")

        if should_stop or len(messages) >= MAX_HOMEWORK:
            break

        # content_h = window["height"] * 0.20
        # delta = max(8, int(content_h / 3 / 10))
        # 动态计算滚动高度为当前卡片的「发布于」到窗口底部的高度
        content_h = window["height"] - picked["detail_btn"].y
        delta = max(8, int(content_h / 3 / 10))
        activate_dingtalk(window["pid"])
        scroll_at(msg_cx, msg_cy, delta)
        time.sleep(SCROLL_WAIT)

    logger.info(
        f"=== 抓取完成，共 {len(messages)} 条；模式={mode}；主键去重 {len(seen_keys)} 个==="
    )
    return messages


if __name__ == "__main__":
    debug_mode = "--debug" in sys.argv
    # --all-dates: 测试用，忽略日期过滤，抓所有作业不管是不是今天
    # all_dates = "--all-dates" in sys.argv
    all_dates = "--all-dates"
    on_non_today = "continue" if all_dates else None
    if all_dates:
        logger.info("--all-dates 模式：忽略日期过滤，抓所有作业")

    msgs = scrape(debug=debug_mode, on_non_today=on_non_today)

    if not msgs:
        print("\n未抓取到任何作业内容")
        print("建议加 --debug：python -m capture.scraper --debug")
        print("测试时加 --all-dates 可忽略日期过滤")
        sys.exit(0)

    print(f"\n共抓取 {len(msgs)} 科作业：\n")
    for m in msgs:
        print(f"[{m.msg_id}]\n{m.text}\n{'─'*40}")

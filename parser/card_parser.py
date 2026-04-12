"""
家校本卡片解析器

输入：群消息文本，例如：
  "4月10日数学\n1订正知能P17,P18\n2完成知能P19,P20"

输出：HomeworkCard 列表

可直接运行测试：
  python -m parser.card_parser
"""

import re
import sys
import os
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from models.models import RawMessage, HomeworkCard
import config


# ── 常量 ──────────────────────────────────────────────────────────────────────

# 科目关键词（按长度降序，避免"道法"被"道"先匹配）
SUBJECTS = sorted(config.SUBJECT_ORDER, key=len, reverse=True)

# 标题正则：匹配 "4月10日数学" / "4月10日 数学" / "04月10日语文"
RE_TITLE = re.compile(
    r"(\d{1,2})月(\d{1,2})日\s*(" + "|".join(re.escape(s) for s in SUBJECTS) + r")"
)

# 作业条目正则：数字编号开头，如 "1订正知能" / "2. 完成练习册" / "①背诵"
RE_ITEM = re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩\d][.、．\s]*(.+)")


# ── 核心解析 ──────────────────────────────────────────────────────────────────

def parse_title(text: str) -> tuple[date | None, str | None]:
    """
    从文本中提取日期和科目。
    返回 (date, subject)，未匹配返回 (None, None)。
    """
    m = RE_TITLE.search(text)
    if not m:
        return None, None

    month = int(m.group(1))
    day = int(m.group(2))
    subject = m.group(3)

    year = date.today().year
    try:
        d = date(year, month, day)
    except ValueError:
        return None, None

    return d, subject


def parse_items(lines: list[str]) -> list[str]:
    """
    从行列表中提取作业条目（跳过标题行和空行）。
    """
    items = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # 跳过标题行（含月日+科目）
        if RE_TITLE.search(line):
            continue
        m = RE_ITEM.match(line)
        if m:
            content = m.group(1).strip()
            if content:
                items.append(content)
        # 没有编号但非空且非标题 → 也作为条目（容错）
        elif len(line) >= 2 and not RE_TITLE.search(line):
            items.append(line)
    return items


def parse_text(text: str, image_paths: list[str] | None = None) -> list[HomeworkCard]:
    """
    解析单条文本消息，返回 HomeworkCard 列表。
    一条消息可能包含多张卡片（多科目）。
    """
    cards = []
    image_paths = image_paths or []

    # 按标题行分割，支持一条消息含多科目
    title_matches = list(RE_TITLE.finditer(text))
    if not title_matches:
        return []

    for i, m in enumerate(title_matches):
        month = int(m.group(1))
        day = int(m.group(2))
        subject = m.group(3)

        try:
            homework_date = date(date.today().year, month, day)
        except ValueError:
            continue

        # 取该标题到下一个标题之间的内容
        start = m.end()
        end = title_matches[i + 1].start() if i + 1 < len(title_matches) else len(text)
        block = text[start:end]

        items = parse_items(block.splitlines())

        # 图片只挂在最后一张卡片上（一次转发的图片与最后一条作业相关）
        card_images = image_paths if i == len(title_matches) - 1 else []

        cards.append(HomeworkCard(
            date=homework_date,
            subject=subject,
            items=items,
            image_paths=card_images,
        ))

    return cards


def parse_messages(messages: list[RawMessage]) -> list[HomeworkCard]:
    """
    批量解析消息列表，返回所有识别到的作业卡片。
    图片消息的路径会关联到前一条（或同一条）文字消息的卡片。
    """
    cards = []
    pending_images: list[str] = []  # 待关联到下一张卡片的图片

    for msg in messages:
        if msg.msg_type == "image" and msg.image_url:
            # 图片消息：先积累，等文字消息来关联
            pending_images.append(msg.image_url)

        elif msg.msg_type in ("text", "mixed") and msg.text:
            new_cards = parse_text(msg.text, image_paths=pending_images)
            if new_cards:
                cards.extend(new_cards)
                pending_images = []  # 图片已关联，清空

    return cards


def sort_cards(cards: list[HomeworkCard]) -> list[HomeworkCard]:
    """
    按 config.SUBJECT_ORDER 排序作业卡片，同科目去重（保留最后一条）。
    """
    # 同科目去重：后来的覆盖前面的（老师可能补充修正）
    seen: dict[str, HomeworkCard] = {}
    for card in cards:
        seen[card.subject] = card

    order = {s: i for i, s in enumerate(config.SUBJECT_ORDER)}
    return sorted(seen.values(), key=lambda c: order.get(c.subject, 999))


# ── CLI 测试入口 ──────────────────────────────────────────────────────────────

def main():
    samples = [
        "4月10日数学\n1订正知能P17,P18\n2完成知能P19,P20",
        "4月10日语文\n1背诵第三课课文\n2完成练习册第5页",
        "4月10日英语\n1听写单词20个\n2朗读课文三遍",
        "今天孩子们表现不错，继续加油！",   # 无效消息
        "4月10日数学\n补充：预习第48页",     # 同科目第二条（覆盖）
    ]

    print(f"\n{'─'*60}")
    print("  家校本卡片解析测试")
    print(f"{'─'*60}\n")

    from datetime import datetime
    messages = []
    for i, text in enumerate(samples):
        messages.append(RawMessage(
            msg_id=str(i),
            sender_id="test",
            sender_name="测试老师",
            timestamp=datetime.now(),
            msg_type="text",
            text=text,
        ))

    cards = parse_messages(messages)
    cards = sort_cards(cards)

    if not cards:
        print("未解析到任何作业卡片")
        return

    for card in cards:
        print(f"[{card.date}] 【{card.subject}】")
        for item in card.items:
            print(f"  - {item}")
        if card.image_paths:
            print(f"  图片: {card.image_paths}")
        print()

    print(f"{'─'*60}")
    print(f"共 {len(cards)} 科作业")


if __name__ == "__main__":
    main()

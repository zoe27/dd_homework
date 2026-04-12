from datetime import datetime, date
from typing import Optional


class RawMessage:
    """Bot 收到的原始群消息"""
    msg_id: str
    sender_id: str
    sender_name: str
    timestamp: datetime
    msg_type: str        # "text" | "image" | "mixed"
    text: str
    image_url: Optional[str]

    def __init__(
        self,
        msg_id: str,
        sender_id: str,
        sender_name: str,
        timestamp: datetime,
        msg_type: str = "text",
        text: str = "",
        image_url: Optional[str] = None,
    ):
        self.msg_id = msg_id
        self.sender_id = sender_id
        self.sender_name = sender_name
        self.timestamp = timestamp
        self.msg_type = msg_type
        self.text = text
        self.image_url = image_url

    def __repr__(self):
        ts = self.timestamp.strftime("%H:%M")
        return f"[{ts}][{self.msg_type}] {self.sender_name}: {self.text[:30]}"


class HomeworkCard:
    """从家校本卡片解析出的单科作业"""
    date: date
    subject: str           # 标准科目名，如 "数学"
    items: list            # 作业条目，如 ["订正知能P17,P18", "完成知能P19,P20"]
    image_paths: list      # 下载后的本地图片路径列表

    def __init__(
        self,
        date: date,
        subject: str,
        items: list[str],
        image_paths: list[str] | None = None,
    ):
        self.date = date
        self.subject = subject
        self.items = items
        self.image_paths = image_paths or []

    def __repr__(self):
        return f"[{self.date}][{self.subject}] {len(self.items)}条作业, {len(self.image_paths)}张图片"


class HomeworkDocument:
    """汇总后的完整作业单，按科目排序"""
    date: date
    cards: list            # list[HomeworkCard]，已按科目顺序排列

    def __init__(self, date: date, cards: list):
        self.date = date
        self.cards = cards

    def is_empty(self) -> bool:
        return len(self.cards) == 0

    def total_items(self) -> int:
        return sum(len(c.items) for c in self.cards)

    def summary(self) -> str:
        parts = [f"共 {len(self.cards)} 科 {self.total_items()} 条作业"]
        for card in self.cards:
            parts.append(f"  【{card.subject}】{len(card.items)}条")
        return "\n".join(parts)

    def __repr__(self):
        return f"HomeworkDocument({self.date}, {len(self.cards)}科)"

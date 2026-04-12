"""
消息内存缓存

负责存储 Bot 监听到的群消息，供触发打印时批量读取。
"""

import sys
import os
from collections import deque
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from models.models import RawMessage
import config


class MessageStore:
    """
    线程安全的消息缓冲队列。
    存储最近 LIMIT 条消息，触发打印后可选择清空。
    """

    def __init__(self, limit: int = None):
        self._limit = limit or config.MESSAGE_STORE_LIMIT
        self._messages: deque[RawMessage] = deque(maxlen=self._limit)

    def add(self, message: RawMessage) -> None:
        self._messages.append(message)

    def get_all(self) -> list[RawMessage]:
        """返回所有缓存消息（按时间升序）"""
        return list(self._messages)

    def clear(self) -> None:
        self._messages.clear()

    def count(self) -> int:
        return len(self._messages)

    def __repr__(self):
        return f"MessageStore({self.count()}/{self._limit} 条)"


# 全局单例，供 handler 和 listener 共享
store = MessageStore()

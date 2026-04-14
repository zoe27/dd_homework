"""
DingTalk Stream SDK 消息监听器

职责：
  - 监听群内所有 @bot 消息（文字 + 图片 + 富文本/转发卡片），存入 store
  - 识别"总结打印"指令，触发完整处理流程

注意：
  DingTalk Stream SDK 的 ChatbotHandler 仅接收 @bot 消息。
  因此用户操作流程为：
    1. 转发家校本卡片时同时 @bot（或将卡片文字粘贴并 @bot 发送）
    2. 也可先转发多张，最后单独发一条 "@bot 总结打印"
       （Bot 会将所有 @bot 消息都存入缓存，触发时批量处理）

使用前提：
  在 open.dingtalk.com 创建企业内部应用，开启机器人能力，
  将 AppKey / AppSecret 填入 .env 文件。
"""

import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import dingtalk_stream
from dingtalk_stream import AckMessage

from utils.logger import logger
from models.models import RawMessage
from bot.store import store
from bot.handler import handle_at_message
import config


class HomeworkBotHandler(dingtalk_stream.ChatbotHandler):
    """
    处理所有 @bot 消息：
      - text     普通文字
      - picture  图片
      - richText 富文本 / 转发的家校本卡片
    """

    async def process(self, callback: dingtalk_stream.CallbackMessage):
        incoming = dingtalk_stream.ChatbotMessage.from_dict(callback.data)
        msg_type = incoming.message_type or "text"

        logger.info(f"[@bot][{msg_type}] {incoming.sender_nick}")

        # ── 提取文字 ──────────────────────────────────────────────────────────
        text_parts = incoming.get_text_list() or []
        full_text = "\n".join(t for t in text_parts if t and t.strip())

        # ── 提取图片 downloadCode ──────────────────────────────────────────────
        image_codes = incoming.get_image_list() or []
        image_urls = []
        for code in image_codes:
            url = self.get_image_download_url(code)
            if url:
                image_urls.append(url)

        # ── 存入缓存 ──────────────────────────────────────────────────────────
        raw = RawMessage(
            msg_id=incoming.message_id or "",
            sender_id=incoming.sender_id or "",
            sender_name=incoming.sender_nick or "",
            timestamp=datetime.fromtimestamp((incoming.create_at or 0) / 1000),
            msg_type="image" if (not full_text and image_urls) else msg_type,
            text=full_text,
            image_url=image_urls[0] if image_urls else None,
        )
        store.add(raw)
        logger.info(f"[缓存] {store}")

        # ── 判断是否为打印指令 ────────────────────────────────────────────────
        def reply(content: str):
            self.reply_text(content, incoming)

        handle_at_message(full_text, reply)
        return AckMessage.STATUS_OK, "ok"


def start():
    """启动 Stream SDK 客户端，阻塞运行"""
    if not config.DINGTALK_APP_KEY or not config.DINGTALK_APP_SECRET:
        logger.error("请在 .env 文件中配置 DINGTALK_APP_KEY 和 DINGTALK_APP_SECRET")
        sys.exit(1)

    credential = dingtalk_stream.Credential(
        config.DINGTALK_APP_KEY,
        config.DINGTALK_APP_SECRET,
    )
    client = dingtalk_stream.DingTalkStreamClient(credential)
    client.register_callback_handler(
        dingtalk_stream.ChatbotMessage.TOPIC,
        HomeworkBotHandler(),
    )

    logger.info("Bot 已启动，等待消息（@bot 触发）...")
    logger.info(f"消息缓存上限：{config.MESSAGE_STORE_LIMIT} 条")
    client.start_forever()

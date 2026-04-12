"""
Bot 指令处理器

负责：
  1. 接收 @bot 消息，识别"总结打印"指令
  2. 从 store 取出缓存消息，解析作业卡片
  3. 下载图片 → 生成 Word → 打印 → 回复确认
"""

import sys
import os
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from utils.logger import logger
from bot.store import store
from parser.card_parser import parse_messages, sort_cards
from models.models import HomeworkDocument
from utils.downloader import download_images

# 触发词：@bot 消息中包含任意一个即触发
TRIGGER_WORDS = ["总结", "打印", "作业"]


def is_trigger(text: str) -> bool:
    """判断 @bot 消息是否为打印指令"""
    return any(w in text for w in TRIGGER_WORDS)


def run_pipeline(reply_func) -> str:
    """
    执行完整流程：解析 → 生成 → 打印
    reply_func: 向群发送消息的回调，签名 reply_func(text: str)
    返回最终回复文本（同时也通过 reply_func 发出）
    """
    from generator.docx_generator import generate
    from printer.printer import print_file

    messages = store.get_all()
    logger.info(f"从缓存取出 {len(messages)} 条消息")

    if not messages:
        msg = "暂无消息记录，请先转发作业卡片再触发打印。"
        reply_func(msg)
        return msg

    # 解析家校本卡片
    cards = parse_messages(messages)
    cards = sort_cards(cards)
    logger.info(f"解析出 {len(cards)} 科作业")

    if not cards:
        msg = "未识别到作业内容，请确认转发的是家校本作业卡片。"
        reply_func(msg)
        return msg

    # 下载卡片中的图片（将 URL 替换为本地路径）
    for card in cards:
        if card.image_paths:
            local_paths = download_images(card.image_paths, msg_id=card.subject)
            card.image_paths = local_paths

    # 取日期（用第一张卡片的日期）
    doc_date = cards[0].date if cards else date.today()
    document = HomeworkDocument(date=doc_date, cards=cards)

    # 生成 Word 文档
    try:
        docx_path = generate(document)
        logger.info(f"文档已生成: {docx_path}")
    except Exception as e:
        logger.error(f"文档生成失败: {e}")
        msg = f"文档生成失败：{e}"
        reply_func(msg)
        return msg

    # 打印
    try:
        print_file(docx_path)
        logger.info("打印任务已发送")
        store.clear()
        msg = f"✓ 已打印\n{document.summary()}"
    except Exception as e:
        logger.error(f"打印失败: {e}")
        msg = f"打印失败：{e}\n文档已保存：{docx_path}"

    reply_func(msg)
    return msg


def handle_at_message(text: str, reply_func) -> None:
    """
    处理 @bot 消息入口。
    text: @bot 之后的文字内容
    reply_func: 向群回复的函数
    """
    text = text.strip()
    logger.info(f"收到指令: {text!r}")

    if is_trigger(text):
        run_pipeline(reply_func)
    else:
        reply_func("收到！发送「总结打印」可汇总并打印作业。")

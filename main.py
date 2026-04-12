import os
import sys
import argparse

os.makedirs("logs", exist_ok=True)
os.makedirs("output", exist_ok=True)

sys.path.insert(0, os.path.dirname(__file__))

from utils.logger import logger


def parse_args():
    parser = argparse.ArgumentParser(description="钉钉作业 Bot")
    parser.add_argument("--mock", action="store_true", help="使用本地测试数据，不连接钉钉")
    parser.add_argument("--input", type=str,
                        default="tests/fixtures/sample_messages.json",
                        help="mock 模式的输入文件（默认 tests/fixtures/sample_messages.json）")
    parser.add_argument("--no-print", action="store_true", help="mock 模式下跳过打印，只生成文档")
    return parser.parse_args()


def run_mock(input_path: str, skip_print: bool = False):
    """加载本地 JSON，直接跑完整流程（解析 → 生成文档 → 打印）"""
    import json
    from datetime import datetime
    from models.models import RawMessage
    from bot.store import store
    from parser.card_parser import parse_messages, sort_cards
    from models.models import HomeworkDocument
    from generator.pdf_generator import generate
    from utils.downloader import download_images

    logger.info(f"Mock 模式，加载: {input_path}")

    with open(input_path, encoding="utf-8") as f:
        raw_data = json.load(f)

    for d in raw_data:
        msg = RawMessage(
            msg_id=d.get("msg_id", ""),
            sender_id=d.get("sender_id", ""),
            sender_name=d.get("sender_name", ""),
            timestamp=datetime.fromisoformat(d["timestamp"]),
            msg_type=d.get("msg_type", "text"),
            text=d.get("text", ""),
            image_url=d.get("image_url"),
        )
        store.add(msg)

    logger.info(f"已加载 {store.count()} 条消息")

    # 解析
    cards = parse_messages(store.get_all())
    cards = sort_cards(cards)

    if not cards:
        logger.warning("未识别到任何作业卡片，请检查输入数据格式")
        return

    logger.info(f"识别结果：{len(cards)} 科")
    for card in cards:
        logger.info(f"  【{card.subject}】{card.items}")

    # 下载图片（mock 数据无真实 URL，跳过失败项）
    for card in cards:
        if card.image_paths:
            card.image_paths = download_images(card.image_paths, card.subject)

    from datetime import date
    doc_date = cards[0].date if cards else date.today()
    document = HomeworkDocument(date=doc_date, cards=cards)

    # 生成文档
    docx_path = generate(document)
    logger.info(f"文档已生成: {docx_path}")

    # 打印
    if skip_print:
        logger.info("--no-print 已设置，跳过打印")
        print(f"\n文档路径: {docx_path}")
        return

    from printer.printer import print_file
    try:
        print_file(docx_path)
        logger.info("打印任务已提交")
    except Exception as e:
        logger.warning(f"打印失败（{e}），文档已保存: {docx_path}")
        print(f"\n文档路径: {docx_path}")


def main():
    args = parse_args()

    if args.mock:
        logger.info("启动模式: Mock")
        run_mock(args.input, skip_print=args.no_print)
        return

    # 真实模式
    logger.info("启动模式: DingTalk Stream SDK")

    import config
    if not config.DINGTALK_APP_KEY or not config.DINGTALK_APP_SECRET:
        logger.error("未配置 DINGTALK_APP_KEY / DINGTALK_APP_SECRET")
        logger.error("请创建 .env 文件，参考 .env.example")
        sys.exit(1)

    from bot.listener import start
    start()


if __name__ == "__main__":
    main()

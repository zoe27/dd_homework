"""
PDF 文档生成器（基于 reportlab，无需 Word/WPS）

将 HomeworkDocument 直接生成 .pdf 文件，支持中文 + 图片嵌入。

可直接运行测试：
  python -m generator.pdf_generator
"""

import os
import sys
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, HRFlowable
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

from models.models import HomeworkDocument, HomeworkCard
from utils.logger import logger
import config

# ── 字体注册 ──────────────────────────────────────────────────────────────────

# macOS 系统中文字体候选（按优先级）
_FONT_CANDIDATES = [
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/Supplemental/Songti.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
]

_FONT_NAME = "Chinese"
_FONT_BOLD_NAME = "Chinese-Bold"


def _register_fonts():
    """注册中文字体，找到第一个可用的即止"""
    for path in _FONT_CANDIDATES:
        if not os.path.exists(path):
            continue
        try:
            pdfmetrics.registerFont(TTFont(_FONT_NAME, path, subfontIndex=0))
            # TTC 通常 index=0 是 Regular，尝试 index=1 作为 Bold
            try:
                pdfmetrics.registerFont(TTFont(_FONT_BOLD_NAME, path, subfontIndex=1))
            except Exception:
                pdfmetrics.registerFont(TTFont(_FONT_BOLD_NAME, path, subfontIndex=0))
            logger.debug(f"已注册字体: {path}")
            return True
        except Exception as e:
            logger.warning(f"字体注册失败 {path}: {e}")
    return False


_fonts_ok = _register_fonts()
_FONT = _FONT_NAME if _fonts_ok else "Helvetica"
_FONT_BOLD = _FONT_BOLD_NAME if _fonts_ok else "Helvetica-Bold"

# ── 样式 ──────────────────────────────────────────────────────────────────────

def _make_styles():
    title_style = ParagraphStyle(
        "Title",
        fontName=_FONT_BOLD,
        fontSize=22,
        leading=30,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#222222"),
    )
    subtitle_style = ParagraphStyle(
        "Subtitle",
        fontName=_FONT,
        fontSize=13,
        leading=20,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#666666"),
    )
    subject_style = ParagraphStyle(
        "Subject",
        fontName=_FONT_BOLD,
        fontSize=14,
        leading=22,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#1a1a1a"),
        spaceBefore=6,
    )
    item_style = ParagraphStyle(
        "Item",
        fontName=_FONT,
        fontSize=12,
        leading=20,
        alignment=TA_LEFT,
        leftIndent=14,
        textColor=colors.HexColor("#333333"),
    )
    return title_style, subtitle_style, subject_style, item_style


# ── 生成逻辑 ──────────────────────────────────────────────────────────────────

def generate(document: HomeworkDocument) -> str:
    """生成 PDF，返回文件路径"""
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)

    date_str = document.date.strftime("%Y%m%d")
    filename = f"今日作业_{date_str}.pdf"
    filepath = os.path.join(config.OUTPUT_DIR, filename)

    doc = SimpleDocTemplate(
        filepath,
        pagesize=A4,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        leftMargin=2.5 * cm,
        rightMargin=2.5 * cm,
    )

    title_style, subtitle_style, subject_style, item_style = _make_styles()
    story = []

    # 标题
    story.append(Paragraph("今日作业", title_style))
    story.append(Spacer(1, 4))

    # 日期副标题
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    weekday = weekdays[document.date.weekday()]
    date_label = (
        f"{document.date.year}年{document.date.month}月"
        f"{document.date.day}日（{weekday}）"
    )
    story.append(Paragraph(date_label, subtitle_style))
    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cccccc")))
    story.append(Spacer(1, 10))

    # 各科作业
    page_width = A4[0] - 5 * cm  # 可用宽度
    for card in document.cards:
        story.append(Paragraph(f"【{card.subject}】", subject_style))

        for i, item in enumerate(card.items, 1):
            story.append(Paragraph(f"{i}.  {item}", item_style))

        # 图片
        for img_path in card.image_paths:
            if not os.path.exists(img_path):
                logger.warning(f"图片不存在，跳过: {img_path}")
                continue
            try:
                img = RLImage(img_path, width=page_width)
                # 等比缩放高度
                from PIL import Image as PILImage
                with PILImage.open(img_path) as pil_img:
                    w, h = pil_img.size
                img.drawHeight = page_width * h / w
                story.append(Spacer(1, 6))
                story.append(img)
            except Exception as e:
                logger.warning(f"图片插入失败: {img_path} — {e}")

        story.append(Spacer(1, 10))

    doc.build(story)
    logger.info(f"PDF 已保存: {filepath}")
    return filepath


# ── CLI 测试入口 ──────────────────────────────────────────────────────────────

def main():
    from models.models import HomeworkCard, HomeworkDocument

    today = date.today()
    cards = [
        HomeworkCard(date=today, subject="数学",
                     items=["订正知能P17,P18", "完成知能P19,P20"]),
        HomeworkCard(date=today, subject="语文",
                     items=["背诵第三课课文", "完成练习册第5页"]),
        HomeworkCard(date=today, subject="英语",
                     items=["听写单词20个", "朗读课文三遍"]),
    ]
    doc = HomeworkDocument(date=today, cards=cards)
    path = generate(doc)
    print(f"\nPDF 已生成: {path}")


if __name__ == "__main__":
    main()

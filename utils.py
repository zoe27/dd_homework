import requests
import os
import win32api
import win32print
from docx import Document
from datetime import datetime
from config import AI_API_KEY, AI_API_URL, AI_MODEL, AUTO_PRINT

def ai_summary(content):
    """免费AI整理作业（智谱GLM-4 Flash）"""
    headers = {
        "Authorization": f"Bearer {AI_API_KEY}",
        "Content-Type": "application/json"
    }

    prompt = """
    你是作业整理助手，只提取今日作业，格式整洁、适合打印，严格按以下格式输出：
    【今日作业】
    语文：xxx
    数学：xxx
    英语：xxx
    其他科目：xxx
    不要多余文字、表情、解释
    """

    data = {
        "model": AI_MODEL,
        "messages": [{"role": "user", "content": f"{prompt}\n作业内容：{content}"}]
    }

    try:
        resp = requests.post(AI_API_URL, json=data, headers=headers, timeout=10)
        resp.raise_for_status()
        result = resp.json()
        return result["choices"][0]["message"]["content"].strip()
    except Exception as e:
        return f"AI调用失败：{str(e)}，请检查KEY或网络"

def create_word(content):
    """生成Word文档"""
    doc = Document()
    doc.add_paragraph(content)
    today = datetime.now().strftime("%Y-%m-%d")
    filename = f"今日作业_{today}.docx"
    doc.save(filename)
    return filename

def auto_print(filename):
    """Windows自动打印"""
    if not AUTO_PRINT:
        return
    try:
        printer_name = win32print.GetDefaultPrinter()
        win32api.ShellExecute(0, "print", filename, f'/d:"{printer_name}"', ".", 0)
        print("✅ 已发送打印任务到默认打印机")
    except Exception as e:
        print(f"❌ 打印失败：{str(e)}，请检查打印机连接")
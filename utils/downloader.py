"""
图片下载工具

从钉钉 URL 下载图片到本地 output/tmp/ 目录。
"""

import os
import sys
import urllib.request

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import config
from utils.logger import logger

TMP_DIR = os.path.join(config.OUTPUT_DIR, "tmp")


def download_image(url: str, msg_id: str, index: int = 0) -> str | None:
    """
    下载图片，返回本地路径。失败返回 None。
    """
    os.makedirs(TMP_DIR, exist_ok=True)
    ext = "jpg"
    if "." in url.split("?")[0].split("/")[-1]:
        ext = url.split("?")[0].split("/")[-1].split(".")[-1]

    filename = f"{msg_id}_{index}.{ext}"
    local_path = os.path.join(TMP_DIR, filename)

    if os.path.exists(local_path):
        return local_path

    try:
        urllib.request.urlretrieve(url, local_path)
        logger.debug(f"图片下载成功: {filename}")
        return local_path
    except Exception as e:
        logger.warning(f"图片下载失败: {url} — {e}")
        return None


def download_images(urls: list[str], msg_id: str) -> list[str]:
    """批量下载，返回成功下载的本地路径列表"""
    paths = []
    for i, url in enumerate(urls):
        path = download_image(url, msg_id, i)
        if path:
            paths.append(path)
    return paths

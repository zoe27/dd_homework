"""
跨平台打印模块

macOS/Linux: lp 命令（直接打印 PDF）
Windows:     win32api.ShellExecute
"""

import sys
import os
import platform
import subprocess

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from utils.logger import logger


def print_file(filepath: str, printer: str | None = None) -> None:
    """
    将文件发送到打印机。
    filepath: 要打印的文件路径（.pdf）
    printer:  指定打印机名称，None 则使用系统默认打印机
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"文件不存在: {filepath}")

    system = platform.system()

    if system == "Darwin":
        _print_macos(filepath, printer)
    elif system == "Windows":
        _print_windows(filepath, printer)
    else:
        _print_linux(filepath, printer)


def _lp_print(filepath: str, printer: str | None) -> None:
    """用 lp 命令打印（macOS / Linux 通用）"""
    cmd = ["lp"]
    if printer:
        cmd += ["-d", printer]
    cmd.append(filepath)

    logger.info(f"打印命令: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"lp 命令失败: {result.stderr.strip()}")
    logger.info(f"打印任务已提交: {result.stdout.strip()}")


def _print_macos(filepath: str, printer: str | None) -> None:
    _lp_print(filepath, printer)


def _print_linux(filepath: str, printer: str | None) -> None:
    _lp_print(filepath, printer)


def _print_windows(filepath: str, printer: str | None) -> None:
    try:
        import win32api
        import win32print
    except ImportError:
        raise RuntimeError("Windows 打印需要安装 pywin32：pip install pywin32")

    if printer:
        win32print.SetDefaultPrinter(printer)

    win32api.ShellExecute(0, "print", filepath, None, ".", 0)
    logger.info(f"打印任务已提交（Windows）: {filepath}")


def list_printers() -> list[str]:
    """列出可用打印机（macOS/Linux）"""
    try:
        result = subprocess.run(["lpstat", "-a"], capture_output=True, text=True)
        lines = result.stdout.strip().splitlines()
        return [line.split()[0] for line in lines if line]
    except Exception:
        return []

"""
钉钉窗口探测

验证是否能找到钉钉主窗口并获取坐标。
运行：python -m capture.find_window
"""

import sys
import subprocess


def find_dingtalk_windows():
    """
    用 Quartz CGWindowListCopyWindowInfo 枚举所有窗口，
    过滤出钉钉相关窗口，返回窗口信息列表。
    """
    try:
        import Quartz
    except ImportError:
        print("需要安装 pyobjc-framework-Quartz：pip install pyobjc-framework-Quartz")
        sys.exit(1)

    window_list = Quartz.CGWindowListCopyWindowInfo(
        Quartz.kCGWindowListOptionOnScreenOnly | Quartz.kCGWindowListExcludeDesktopElements,
        Quartz.kCGNullWindowID,
    )

    results = []
    for w in window_list:
        owner = w.get("kCGWindowOwnerName", "")
        name  = w.get("kCGWindowName", "")
        print(f"窗口: {owner} / {name}")
        if "钉钉" in owner or "DingTalk" in owner.lower() or "dingtalk" in owner.lower():
            bounds = w.get("kCGWindowBounds", {})
            results.append({
                "owner": owner,
                "name": name,
                "pid": w.get("kCGWindowOwnerPID"),
                "window_id": w.get("kCGWindowNumber"),
                "x": int(bounds.get("X", 0)),
                "y": int(bounds.get("Y", 0)),
                "width": int(bounds.get("Width", 0)),
                "height": int(bounds.get("Height", 0)),
            })
    return results


def find_main_window(windows: list) -> dict | None:
    """从窗口列表中挑出主聊天窗口（面积最大的）"""
    if not windows:
        return None
    return max(windows, key=lambda w: w["width"] * w["height"])


if __name__ == "__main__":
    print("正在枚举屏幕窗口...\n")
    windows = find_dingtalk_windows()

    if not windows:
        print("未找到钉钉窗口，请确认：")
        print("  1. 钉钉客户端已启动")
        print("  2. 本程序已获得「屏幕录制」权限（系统设置 → 隐私与安全性 → 屏幕录制）")
        sys.exit(1)

    print(f"找到 {len(windows)} 个钉钉相关窗口：\n")
    for i, w in enumerate(windows):
        print(f"  [{i}] {w['owner']} / {w['name'] or '(无标题)'}")
        print(f"       PID={w['pid']}  WindowID={w['window_id']}")
        print(f"       位置=({w['x']}, {w['y']})  大小={w['width']}x{w['height']}")
        print()

    main = find_main_window(windows)
    print(f"主窗口（面积最大）: {main['name'] or main['owner']}")
    print(f"  坐标: ({main['x']}, {main['y']})  大小: {main['width']}x{main['height']}")

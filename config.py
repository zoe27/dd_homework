import os
from dotenv import load_dotenv

load_dotenv()

# ── 钉钉应用配置 ──────────────────────────────────────────────────────────────
# 在 open.dingtalk.com 创建企业内部应用后获取
DINGTALK_APP_KEY    = os.getenv("DINGTALK_APP_KEY", "")
DINGTALK_APP_SECRET = os.getenv("DINGTALK_APP_SECRET", "")

# 监听的目标群 ID（在钉钉群设置中查看）
TARGET_GROUP_ID = os.getenv("TARGET_GROUP_ID", "")

# ── 作业解析 ──────────────────────────────────────────────────────────────────
# 科目顺序（Word 文档按此顺序排列科目）
SUBJECT_ORDER = ["语文", "数学", "英语", "科学", "道法", "体育", "美术", "音乐"]

# ── 消息缓存 ──────────────────────────────────────────────────────────────────
# 最多缓存多少条群消息（触发打印后自动清空）
MESSAGE_STORE_LIMIT = int(os.getenv("MESSAGE_STORE_LIMIT", "100"))

# ── 输出 ──────────────────────────────────────────────────────────────────────
OUTPUT_DIR = os.getenv("OUTPUT_DIR", "./output")

# ── 打印 ──────────────────────────────────────────────────────────────────────
# 指定打印机名称，留空则使用系统默认打印机
PRINTER_NAME = os.getenv("PRINTER_NAME", "")

# ── 日志 ──────────────────────────────────────────────────────────────────────
LOG_DIR   = "./logs"
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")

# ── 截图 v2 工作通知抓取 ────────────────────────────────────────────────────────
# stop：遇到发布日期早于今天则停止向上扫描；continue：跳过非今天卡片继续找，直至凑满 MAX 或到顶
CAPTURE_V2_ON_NON_TODAY = os.getenv("CAPTURE_V2_ON_NON_TODAY", "stop")
CAPTURE_V2_MAX_HOMEWORK = int(os.getenv("CAPTURE_V2_MAX_HOMEWORK", "6"))

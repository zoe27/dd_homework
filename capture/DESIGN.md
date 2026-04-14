# 钉钉客户端自动抓取方案设计

## 背景

钉钉 Stream SDK Bot 无法接收家校本专属群的消息，因此改为直接操控本机钉钉客户端，
通过截图 + OCR 的方式自动抓取作业内容。

---

## 模块结构

```
capture/
  find_window.py    找到钉钉主窗口，获取坐标和 window ID
  screenshot.py     按 window ID 截图；激活窗口
  ocr.py            easyocr 中文 OCR，返回带屏幕逻辑坐标的结果列表
  scraper.py        主流程（AI家校本入口版）
  scraper_scroll.py 旧版（群聊滚动扫描，已废弃，保留备用）
```

---

## 当前方案：AI家校本入口

### 流程概览

```
激活钉钉窗口，进入目标群聊
  ↓
步骤1：OCR 识别底部导航栏，找到"AI家校本"按钮并点击
  ↓
步骤2：等待面板渲染，点击"全部"tab（查看所有作业，而非只看未完成）
  ↓
步骤3：OCR 识别作业列表，提取所有"X月X日科目"格式的卡片标题和坐标
  ↓
步骤4：循环处理（最多6个）：
  点击作业卡片
    ↓
  等待 1.5s 渲染
    ↓
  截图右侧面板 → OCR 提取完整作业内容
    ↓
  点击返回按钮"<"（点击行左端 +10px，精准落在箭头上）
    ↓
  等待 0.8s，处理下一个
  ↓
步骤5：关闭家校本面板
  优先：OCR 找到"×"按钮点击
  fallback：点击面板左侧外部区域（群聊区域）关闭
  ↓
返回 RawMessage 列表 → card_parser → 生成PDF → 打印
```

---

## 坐标系说明

这是整个方案最关键的细节，涉及三层坐标转换：

```
截图（物理像素，Retina 2x = 逻辑尺寸 × 2）
  ↓ 裁剪右侧面板（物理像素偏移）
  ↓ 传入 easyocr，返回 bbox（物理像素）
  ↓ ÷ scale（2.0）→ 逻辑像素
  ↓ + panel_x_logical（面板裁剪偏移，逻辑）
  ↓ + window["x/y"]（窗口在屏幕上的位置，逻辑）
  = OcrResult.x/y（屏幕逻辑坐标）✓
  → 直接传给 CGEventCreateMouseEvent 点击 ✓
```

关键：`window["width/height"]` 是逻辑尺寸，截图是物理尺寸，`scale = 截图宽 / 窗口逻辑宽`。

---

## 各步骤技术细节

### 激活窗口
`NSWorkspace.sharedWorkspace().runningApplications()` 找到钉钉进程，
`app.activateWithOptions_(NSApplicationActivateIgnoringOtherApps)` 带到前台。
每次点击前都激活，确保事件发到钉钉。

### 截图
`Quartz.CGWindowListCreateImage` 按 window ID 截图，不要求窗口在最前面。
只截右侧面板区域（窗口宽度 37% 处往右），减少 OCR 处理量。

### OCR
`easyocr`，语言 `['ch_sim', 'en']`，置信度阈值 0.3。
首次加载模型约需几秒，之后复用单例。

### 点击
`CGEventCreateMouseEvent`，mouseDown + mouseUp 模拟单击。
坐标已经过 scale 转换，直接使用 OcrResult.center_x/y 或 x+offset。

### 返回按钮
匹配文字以 `<` 或 `〈` 开头的 OCR 结果（通常是"< 周沐曦的练习"整行）。
点击 `r.x + 10`（行左端），而非 `center_x`，确保落在箭头上。

### 关闭面板
优先 OCR 找"×"点击；找不到则点击面板左侧外部区域（`window_x + panel_x/2`），
利用钉钉"点击外部关闭浮层"的交互特性。

---

## 参数配置

| 参数 | 值 | 说明 |
|---|---|---|
| CARD_OPEN_WAIT | 1.5s | 点击卡片后等待面板渲染 |
| MAX_HOMEWORK | 6 | 最多提取作业数量 |
| panel_x_ratio | 37% | 右侧面板起始位置（窗口宽度比例） |
| min_conf | 0.3 | OCR 最低置信度 |

---

## 与现有代码的衔接

`scraper.scrape()` 返回 `list[RawMessage]`，与 `card_parser.parse_messages()` 接口完全兼容。

```python
from capture.scraper import scrape
from parser.card_parser import parse_messages, sort_cards

messages = scrape()
cards = parse_messages(messages)
cards = sort_cards(cards)
# → 生成PDF → 打印，流程不变
```

---

## 风险与限制

- 需授权「屏幕录制」和「辅助功能」权限
- 钉钉需处于目标群聊界面，底部能看到"AI家校本"按钮
- 钉钉界面更新后坐标比例或文字可能变化，需重新调整
- 最小化状态下无法截图
- easyocr 首次加载模型需要几秒，后续复用

# 钉钉客户端自动抓取方案设计

## 背景

钉钉 Stream SDK Bot 无法接收家校本专属群的消息，因此改为直接操控本机钉钉客户端，
通过截图 + OCR 的方式自动抓取作业内容。

---

## 模块结构

```
capture/
  find_window.py   找到钉钉主窗口，获取坐标和 window ID
  screenshot.py    按 window ID 截图，支持持续截图+滚动
  ocr.py           Apple Vision OCR，返回文字内容和边界框坐标
  scraper.py       主流程，串联所有模块，返回作业文字列表
```

---

## 完整流程

```
激活钉钉窗口（带到前台）
  ↓
进入群聊消息列表
  ↓
【第一阶段：滚到底部】
连续向下滚动，直到连续 2 次底部截图相同，确保在最新消息处
  ↓
【第二阶段：从底部向上扫描】
┌─────────────────────────────────────┐
│  截图当前屏幕                        │
│  ↓                                  │
│  OCR 识别截图文字                    │
│  ↓                                  │
│  扫描识别结果：                      │
│    有"家校本" + "X月X日" + 科目？    │
│    ├── 有，且日期是今天              │
│    │     → 记录卡片中心坐标          │
│    │       点击卡片                  │
│    │       等待 1.5s（渲染完成）     │
│    │       截图展开后的面板          │
│    │       OCR 识别面板内容          │
│    │       提取标题（日期+科目）     │
│    │       提取正文（作业文字）      │
│    │       存入 cards 列表           │
│    │       点击左侧群聊区域关闭面板  │
│    │       等待 0.3s                 │
│    ├── 有，但日期不是今天            │
│    │     → 停止扫描，今天内容已全部收集
│    └── 无 → 跳过                    │
│  ↓                                  │
│  向上滚动（delta=+3，约3行）         │
│  等待 0.5s                          │
│  ↓                                  │
│  连续 2 次顶部截图相同？             │
│    ├── 是 → 退出循环（已到顶）       │
│    └── 否 → 继续                    │
└─────────────────────────────────────┘
  ↓
返回 cards 列表 → 喂给 card_parser → 生成PDF → 打印
```

**从底部向上扫描的优势：**
- 不需要从头遍历历史消息，效率高
- 同一科目若老师发了多次，越靠下越新，第一个遇到的就是最新版本，天然去重
- 遇到非今天的卡片立即终止，边界条件清晰

---

## 各步骤技术细节

### 1. 激活窗口
- 用 `NSWorkspace.sharedWorkspace().runningApplications()` 找到钉钉进程
- `app.activateWithOptions_(NSApplicationActivateIgnoringOtherApps)`
- 等待 0.3s 确保窗口在前台

### 2. 截图
- 用 `Quartz.CGWindowListCreateImage` 按 window ID 截图
- 不要求窗口在最前面，被遮挡也能截到（最小化除外）
- 截图返回 PIL Image 对象

### 3. OCR（Apple Vision）
- 使用 macOS 原生 Vision framework（`pyobjc-framework-Vision`）
- `VNRecognizeTextRequest` 识别文字，支持中英文混合
- 每个识别结果包含：文字内容 + `boundingBox`（归一化坐标 0~1）
- 坐标换算：`screen_x = box.x * img_width + window_x`
- Vision 的坐标系原点在左下角，需要翻转 Y 轴：`screen_y = window_y + (1 - box.y - box.height) * img_height`

### 4. 识别作业卡片
在 OCR 结果中查找满足以下条件的文字块：
- 包含"家校本"
- 同一区域附近有"X月X日"格式的日期
- 同一区域附近有科目关键词（语文/数学/英语等）

点击坐标取匹配文字块的 boundingBox 中心点。

### 5. 展开面板后的内容提取
点击卡片后面板在右侧弹出，结构：
- 标题行："4月10日语文"
- 副标题："预计需X分钟  XX.XX 已截止"（可忽略）
- 正文：作业描述文字（需要提取）
- 可能有图片（字帖等，跳过）
- 底部可能有"↓"箭头，说明面板内还有更多内容需要在面板内滚动

面板内滚动：如果检测到"↓"箭头或底部有截断，对面板右半区域继续滚动截图，直到内容完整。

内容完整判断：OCR 结果中不再出现"↓"，或连续两次面板截图底部相同。

### 6. 关闭面板
点击左侧群聊区域（窗口 x 坐标的 1/4 处，y 坐标取窗口中间），稳定可靠。

### 7. 到底判断
- 每次滚动后截图，裁剪底部 100px 条带
- 计算与上一张的像素均值差
- 连续 2 次差值 < 5.0，判定已到底，退出

---

## 参数配置

| 参数 | 值 | 说明 |
|---|---|---|
| SCROLL_DOWN_DELTA | -10 | 第一阶段向下滚到底，每次滚动行数 |
| SCROLL_UP_DELTA | +3 | 第二阶段向上扫描，每次滚动行数 |
| SCROLL_WAIT | 0.5s | 滚动后等待渲染 |
| CARD_OPEN_WAIT | 1.5s | 点击卡片后等待面板渲染 |
| CLOSE_WAIT | 0.3s | 关闭面板后等待 |
| COMPARE_STRIP_HEIGHT | 100px | 顶部/底部对比区域高度 |
| SIMILARITY_THRESHOLD | 5.0 | 像素差阈值，低于此值视为相同 |
| SAME_COUNT_TO_STOP | 2 | 连续相同次数才判定到顶/底 |
| MAX_SCREENSHOTS | 30 | 最大截图数，防止死循环 |

---

## 与现有代码的衔接

`scraper.py` 最终返回 `list[RawMessage]`，与现有 `card_parser.parse_messages()` 接口完全兼容，
不需要修改下游任何代码。

```python
# 调用方式
from capture.scraper import scrape
from parser.card_parser import parse_messages, sort_cards

messages = scrape()           # 自动抓取
cards = parse_messages(messages)
cards = sort_cards(cards)
# → 后续生成PDF、打印，流程不变
```

---

## 风险与限制

- 钉钉界面更新后卡片布局可能变化，需要重新调整识别逻辑
- 家校本卡片标题是图片渲染（绿色背景白字），Apple Vision 对高对比度图片识别率高，预计没问题
- 最小化状态下无法截图，需要钉钉窗口处于正常显示状态
- 需要授权「屏幕录制」和「辅助功能」权限

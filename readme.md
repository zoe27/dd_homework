# 钉钉作业 Bot

从家校本作业卡片自动汇总作业并打印。

## 功能

- 监听个人钉钉群中转发的家校本作业卡片
- 识别日期、科目、作业条目及图片
- `@bot 总结打印` 触发：生成 Word 文档并发送打印任务
- 支持多科目、图片嵌入，按科目顺序排版

## 使用流程

```
家校本（看作业）→ 转发卡片到个人群 → @bot 总结打印 → 打印机输出
```

1. 在钉钉家校本查看老师发布的作业卡片
2. 将卡片转发到你与 Bot 的个人群
3. 发送 `@bot 总结打印`
4. Bot 解析所有卡片，生成 Word 文档并打印
5. 群内收到回复：`✓ 已打印，共 N 科 M 条作业`

## 环境要求

- Python 3.11+
- macOS（打印使用 `lp` 命令）；Windows 需额外安装 `pywin32`

## 安装

```bash
pip install -r requirements.txt
```

## 配置

```bash
cp .env.example .env
```

编辑 `.env`：

```
DINGTALK_APP_KEY=your_app_key
DINGTALK_APP_SECRET=your_app_secret
```

### 获取 AppKey / AppSecret

1. 登录 [open.dingtalk.com](https://open.dingtalk.com)
2. 应用开发 → 企业内部应用 → 新建应用
3. 添加能力：**机器人**
4. 在应用凭证页面获取 AppKey 和 AppSecret
5. 将 Bot 添加到你的个人群

## 启动

```bash
# 连接真实钉钉
python main.py

# Mock 模式（不需要钉钉，用本地测试数据）
python main.py --mock --no-print   # 只生成文档
python main.py --mock              # 生成文档并打印

# 使用自定义测试数据
python main.py --mock --input path/to/messages.json
```

## 项目结构

```
dd_homework/
├── main.py                  # 入口
├── config.py                # 配置项
├── .env.example             # 配置模板
├── requirements.txt
│
├── bot/
│   ├── listener.py          # Stream SDK，接收 @bot 消息
│   ├── handler.py           # 指令识别，串联全流程
│   └── store.py             # 消息内存缓存
│
├── parser/
│   └── card_parser.py       # 家校本卡片解析
│
├── generator/
│   └── docx_generator.py    # Word 文档生成
│
├── printer/
│   └── printer.py           # 系统打印
│
├── utils/
│   ├── logger.py
│   └── downloader.py        # 钉钉图片下载
│
├── models/
│   └── models.py            # 数据模型
│
└── tests/
    └── fixtures/
        └── sample_messages.json   # Mock 测试数据
```

## 测试数据格式

`tests/fixtures/sample_messages.json` 中每条消息的格式：

```json
{
  "msg_id": "msg001",
  "sender_id": "teacher_wang",
  "sender_name": "王老师",
  "timestamp": "2026-04-12T15:30:00",
  "msg_type": "text",
  "text": "4月12日语文\n1背诵第三课课文\n2完成练习册第5页",
  "image_url": null
}
```

家校本卡片文本格式：标题为 `M月D日科目`，内容为数字编号列表。

## 配置说明

| 配置项 | 说明 | 默认值 |
|---|---|---|
| `DINGTALK_APP_KEY` | 钉钉应用 AppKey | 必填 |
| `DINGTALK_APP_SECRET` | 钉钉应用 AppSecret | 必填 |
| `SUBJECT_ORDER` | 科目排列顺序 | 语数英科道体美音 |
| `MESSAGE_STORE_LIMIT` | 消息缓存条数上限 | 100 |
| `OUTPUT_DIR` | Word 文档保存目录 | `./output` |
| `PRINTER_NAME` | 指定打印机，留空用默认 | 空 |
| `LOG_LEVEL` | 日志级别 | INFO |

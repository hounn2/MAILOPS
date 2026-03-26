# Outlook邮件助手

一个强大的Outlook自动化工具，能够根据配置规则自动甄别邮件并执行相应操作（回复、转发、移动等）。

## 功能特性

- **智能规则匹配**：支持多种条件组合（主题、发件人、内容、附件等）
- **自动回复**：根据模板自动回复邮件
- **自动转发**：将邮件转发给指定人员
- **自动归档**：将邮件移动到指定文件夹
- **灵活配置**：通过JSON配置文件定义规则
- **安全运行**：支持试运行模式（Dry Run）

## 安装要求

### 系统要求
- Windows 操作系统
- Microsoft Outlook 已安装并配置
- Python 3.7+

### 安装依赖
```bash
pip install -r requirements.txt
```

## 快速开始

### 1. 基础配置
编辑 `config.json` 文件，配置你的邮件规则：

```json
{
  "rules": [
    {
      "id": "rule_001",
      "name": "示例规则",
      "enabled": true,
      "conditions": {
        "match_all": true,
        "items": [
          {
            "field": "subject",
            "operator": "contains",
            "value": ["关键字1", "关键字2"]
          }
        ]
      },
      "actions": [
        {
          "type": "move",
          "target": "指定文件夹"
        }
      ]
    }
  ]
}
```

### 2. 试运行
在正式运行前，先使用试运行模式测试配置：

```bash
python outlook_assistant.py --dry-run --once
```

### 3. 正式运行
```bash
# 运行一次
python outlook_assistant.py --once

# 持续运行（每60秒检查一次）
python outlook_assistant.py
```

## 配置详解

### 规则条件（Conditions）

支持的条件字段：
- `subject` - 邮件主题
- `body` - 邮件正文
- `sender` - 发件人邮箱
- `sender_domain` - 发件人域名
- `recipients` - 收件人列表
- `has_attachments` - 是否有附件
- `body_length` - 正文长度
- `received_time` - 接收时间

支持的操作符：
- `equals` / `not_equals` - 等于/不等于
- `contains` / `not_contains` - 包含/不包含（支持列表）
- `starts_with` / `ends_with` - 以...开头/结尾
- `in` / `not_in` - 在列表中/不在列表中
- `regex` - 正则表达式匹配
- `greater_than` / `less_than` - 大于/小于
- `between` - 在范围内

### 动作（Actions）

#### 1. 回复邮件（reply）
```json
{
  "type": "reply",
  "template": "template_name",
  "include_original": false
}
```

#### 2. 转发邮件（forward）
```json
{
  "type": "forward",
  "to": ["manager@company.com"],
  "subject_prefix": "[转发] ",
  "additional_body": "请查看附件"
}
```

#### 3. 移动邮件（move）
```json
{
  "type": "move",
  "target": "目标文件夹名称"
}
```

#### 4. 标记已读（mark_as_read）
```json
{
  "type": "mark_as_read"
}
```

### 模板（Templates）

在配置中定义回复模板：

```json
{
  "templates": {
    "auto_reply": {
      "subject": "RE: {original_subject}",
      "body": "您好，\n\n感谢您的来信。\n\n此邮件为自动回复。\n\n{sender}"
    }
  }
}
```

支持的模板变量：
- `{original_subject}` - 原始主题
- `{sender}` - 发件人
- `{received_time}` - 接收时间
- `{return_date}` - 返回日期
- `{backup_contact}` - 备用联系人

### 设置（Settings）

```json
{
  "settings": {
    "check_interval": 60,           // 检查间隔（秒）
    "process_unread_only": true,    // 只处理未读邮件
    "max_emails_per_batch": 50,     // 每批处理最大邮件数
    "log_level": "INFO",           // 日志级别
    "dry_run": false,              // 试运行模式
    "mark_as_read_after_process": true  // 处理后标记已读
  }
}
```

## 命令行参数

```
python outlook_assistant.py [选项]

选项:
  -c, --config PATH     配置文件路径 (默认: config.json)
  --once               只运行一次
  --dry-run            试运行模式（不实际执行操作）
  -h, --help           显示帮助信息
```

## 示例场景

### 场景1：垃圾邮件自动处理
```json
{
  "id": "spam_filter",
  "name": "垃圾邮件过滤",
  "enabled": true,
  "conditions": {
    "match_all": false,
    "items": [
      {
        "field": "subject",
        "operator": "contains",
        "value": ["促销", "广告", "优惠"]
      },
      {
        "field": "sender",
        "operator": "not_in",
        "value": ["whitelist@company.com"]
      }
    ]
  },
  "actions": [
    {
      "type": "move",
      "target": "已删除邮件"
    }
  ]
}
```

### 场景2：紧急邮件转发
```json
{
  "id": "urgent_forward",
  "name": "紧急邮件转发",
  "enabled": true,
  "conditions": {
    "match_all": true,
    "items": [
      {
        "field": "subject",
        "operator": "contains",
        "value": ["紧急", "urgent"]
      }
    ]
  },
  "actions": [
    {
      "type": "forward",
      "to": ["manager@company.com"],
      "subject_prefix": "[紧急] "
    },
    {
      "type": "move",
      "target": "紧急邮件"
    }
  ]
}
```

### 场景3：客户咨询自动回复
```json
{
  "id": "customer_auto_reply",
  "name": "客户咨询自动回复",
  "enabled": true,
  "conditions": {
    "match_all": true,
    "items": [
      {
        "field": "subject",
        "operator": "contains",
        "value": ["咨询", "问题"]
      },
      {
        "field": "sender_domain",
        "operator": "not_in",
        "value": ["@company.com"]
      }
    ]
  },
  "actions": [
    {
      "type": "reply",
      "template": "customer_auto_reply"
    },
    {
      "type": "move",
      "target": "客户咨询"
    }
  ]
}
```

## 日志

程序会同时输出到控制台和日志文件 `outlook_assistant.log`，便于排查问题。

## 注意事项

1. **备份重要邮件**：首次使用前请备份重要邮件
2. **先试运行**：正式运行前务必使用 `--dry-run` 测试
3. **Outlook必须运行**：程序需要Outlook正在运行才能工作
4. **权限**：确保程序有访问Outlook的权限

## 故障排除

### 问题：无法连接Outlook
**解决**：确保Outlook已启动并登录

### 问题：找不到文件夹
**解决**：检查文件夹名称是否与Outlook中完全一致（包括大小写）

### 问题：中文乱码
**解决**：确保配置文件保存为UTF-8编码

## 许可证

MIT License

## 作者

Your Name

## 更新日志

### v1.0.0 (2026-03-26)
- 初始版本发布
- 支持规则匹配、自动回复、转发、移动功能
- 支持试运行模式

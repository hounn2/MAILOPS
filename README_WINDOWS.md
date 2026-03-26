# Outlook邮件助手 - Windows版本使用说明

## 版本选择

本项目提供两个版本：

### 1. Windows COM版本（推荐在Windows上使用）
- **文件**: `outlook_assistant_win.py`
- **特点**: 直接操作本地Outlook客户端，反应更快
- **要求**: Windows系统 + Microsoft Outlook

### 2. 跨平台Graph API版本
- **文件**: `outlook_assistant.py`
- **特点**: 使用Microsoft 365 API，支持Windows/macOS/Linux
- **要求**: Microsoft 365账户 + Azure AD应用注册

---

## Windows版本快速开始

### 环境要求
- Windows 7/8/10/11
- Microsoft Outlook 2013/2016/2019/2021/365
- Python 3.7+

### 安装步骤

#### 1. 安装Python
如果尚未安装，请从 https://python.org 下载并安装Python 3.7或更高版本。

#### 2. 安装依赖
```bash
pip install pywin32
```

#### 3. 配置文件
编辑 `config.json`，配置你的邮件规则。

已有示例规则，可根据需要修改。

#### 4. 运行程序

**试运行（测试配置）**
```bash
python outlook_assistant_win.py --dry-run --once
```

**正式运行（只运行一次）**
```bash
python outlook_assistant_win.py --once
```

**正式运行（持续监控）**
```bash
python outlook_assistant_win.py
```

程序会每60秒检查一次新邮件（可在配置中修改）。

---

## 命令行参数

```bash
python outlook_assistant_win.py [选项]

选项:
  -c, --config PATH    配置文件路径 (默认: config.json)
  --once               只运行一次
  --dry-run            试运行模式（不实际执行操作）
  -h, --help           显示帮助信息
```

---

## 配置文件说明

### config.json 结构

```json
{
  "rules": [
    {
      "id": "规则ID",
      "name": "规则名称",
      "enabled": true,
      "conditions": {
        "match_all": true,
        "items": [
          {
            "field": "subject",
            "operator": "contains",
            "value": ["关键词1", "关键词2"]
          }
        ]
      },
      "actions": [
        {
          "type": "move",
          "target": "目标文件夹"
        }
      ]
    }
  ],
  "templates": {
    "template_name": {
      "subject": "主题模板",
      "body": "正文模板"
    }
  },
  "settings": {
    "check_interval": 60,
    "process_unread_only": true,
    "max_emails_per_batch": 50,
    "dry_run": false
  }
}
```

### 示例规则

#### 1. 垃圾邮件过滤
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

#### 2. 自动回复
```json
{
  "id": "auto_reply",
  "name": "自动回复",
  "enabled": true,
  "conditions": {
    "match_all": true,
    "items": [
      {
        "field": "subject",
        "operator": "contains",
        "value": ["咨询"]
      }
    ]
  },
  "actions": [
    {
      "type": "reply",
      "template": "auto_reply_customer",
      "include_original": false
    }
  ]
}
```

#### 3. 邮件转发
```json
{
  "id": "forward_urgent",
  "name": "紧急邮件转发",
  "enabled": true,
  "conditions": {
    "match_all": true,
    "items": [
      {
        "field": "subject",
        "operator": "contains",
        "value": ["紧急"]
      }
    ]
  },
  "actions": [
    {
      "type": "forward",
      "to": ["manager@company.com"],
      "subject_prefix": "[转发] "
    }
  ]
}
```

---

## 故障排除

### 问题1: 无法连接到Outlook
**错误信息**: `连接Outlook失败`

**解决方案**:
1. 确保Microsoft Outlook已启动
2. 确保Outlook已登录邮箱账户
3. 以管理员身份运行命令提示符
4. 检查Outlook是否被其他程序占用

### 问题2: pywin32安装失败
**错误信息**: `Could not find a version that satisfies the requirement pywin32`

**解决方案**:
```bash
# 使用管理员权限运行
pip install --upgrade pip
pip install pywin32

# 或者使用conda
conda install pywin32
```

### 问题3: 找不到文件夹
**错误信息**: `目标文件夹未找到`

**解决方案**:
1. 检查Outlook中是否存在该文件夹
2. 注意文件夹名称大小写
3. 如果是中文文件夹，确保编码为UTF-8

### 问题4: 中文乱码
**解决方案**:
1. 确保配置文件保存为UTF-8编码
2. 使用支持UTF-8的文本编辑器（如VS Code、Notepad++）

### 问题5: 权限被拒绝
**解决方案**:
1. 以管理员身份运行命令提示符
2. 检查Outlook的安全设置
3. 确保防病毒软件未阻止脚本

---

## 日志查看

程序会生成日志文件 `outlook_assistant.log`，可用于排查问题：

```bash
type outlook_assistant.log
```

或在资源管理器中双击打开。

---

## 定时运行（可选）

### 使用Windows任务计划程序

1. 打开"任务计划程序"
2. 创建基本任务
3. 设置触发器（如"登录时"或"每隔1小时"）
4. 设置操作为"启动程序"
5. 程序: `python.exe`
6. 参数: `outlook_assistant_win.py --once`
7. 起始于: 项目所在目录

---

## 文件说明

```
MAILOPS/
├── outlook_assistant_win.py  # Windows版本主程序 ⭐
├── outlook_assistant.py      # 跨平台Graph API版本
├── actions.py               # Windows COM接口操作
├── rules.py                 # 规则引擎
├── config.json              # 配置文件
├── requirements.txt         # 依赖列表
├── README.md               # 主文档
├── README_WINDOWS.md       # Windows使用说明（本文件）
└── outlook_assistant.log   # 运行日志
```

---

## 安全建议

1. **备份邮件**: 首次使用前备份重要邮件
2. **试运行**: 正式使用前务必使用 `--dry-run` 测试
3. **测试规则**: 先用简单规则测试，确认无误后再使用复杂规则
4. **定期检查日志**: 关注 `outlook_assistant.log` 文件

---

## 技术支持

如有问题：
1. 查看日志文件 `outlook_assistant.log`
2. 使用 `--dry-run` 模式测试
3. 检查配置文件格式

---

**版本**: 1.0.0  
**更新日期**: 2026-03-26

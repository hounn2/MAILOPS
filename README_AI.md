# AI智能回复功能使用说明（基于本地LMStudio）

## 🎯 功能概述

AI智能回复功能可以将邮件内容发送给本地部署的大语言模型（通过LMStudio），AI根据配置的知识库文件生成智能回复。

**核心优势：**
- 🤖 使用本地AI模型，数据不上传云端，保护隐私
- 📚 支持自定义知识库，AI基于知识库回答问题
- 🎨 可自定义AI角色和回复风格
- ⚡ 无需网络依赖（除了邮件发送）

## 📋 前置要求

### 1. 安装LMStudio

1. 下载LMStudio: https://lmstudio.ai/
2. 安装并启动LMStudio
3. 下载模型（推荐）：
   - **中文场景**: chatglm3-6b, Qwen-7B-Chat
   - **通用场景**: Llama-2-7B, Mistral-7B
   - **轻量级**: Phi-2, TinyLlama

### 2. 启动LMStudio服务器

1. 打开LMStudio
2. 在左侧菜单点击 "Local Server"
3. 选择已下载的模型
4. 点击 "Start Server"
5. 确认服务器运行在 `http://localhost:1234`

## 📁 文件说明

```
MAILOPS/
├── ai_engine.py               # ⭐ AI引擎模块
├── knowledge_base.py          # ⭐ 知识库管理模块
├── config_with_ai.json        # ⭐ AI功能配置文件示例
├── knowledge_base/            # 知识库目录（需要创建）
│   ├── product_manual.txt
│   ├── faq.md
│   └── ...
└── outlook_assistant_win_standalone.py  # 已集成AI功能
```

## 🚀 快速开始

### 1. 准备知识库

创建 `knowledge_base` 目录，放入你的知识库文件：

```bash
mkdir knowledge_base
```

支持格式：
- **.txt** - 纯文本文件
- **.md** - Markdown文件
- **.pdf** - PDF文件（需要安装PyPDF2: `pip install PyPDF2`）
- **.docx** - Word文档（需要安装python-docx: `pip install python-docx`）

**示例：** `knowledge_base/product_info.txt`

```
# 产品信息

## 产品A介绍
产品A是我们公司的旗舰产品，具有以下特点：
- 高性能处理器
- 长续航电池
- 轻薄设计

## 价格信息
- 基础版：¥2999
- 专业版：¥4999
- 企业版：¥9999

## 技术支持
电话：400-xxx-xxxx
邮箱：support@company.com
```

### 2. 配置AI功能

复制 `config_with_ai.json` 为 `config.json`：

```bash
cp config_with_ai.json config.json
```

编辑 `config.json`：

```json
{
  "lmstudio": {
    "enabled": true,              // 启用AI功能
    "base_url": "http://localhost:1234",
    "model": null,                // null表示使用默认模型
    "timeout": 60,
    "system_prompt": "你是一个专业的客服助手..."  // 自定义AI角色
  },
  "knowledge_base": {
    "enabled": true,              // 启用知识库
    "path": "knowledge_base",     // 知识库路径
    "search_top_k": 3             // 每次查询最多使用3篇相关文档
  }
}
```

### 3. 启用AI回复规则

在 `config.json` 中找到 `rule_006`，启用它：

```json
{
  "id": "rule_006",
  "name": "AI智能回复",
  "enabled": true,    // 改为true启用
  "conditions": {
    "match_all": true,
    "items": [
      {
        "field": "subject",
        "operator": "contains",
        "value": ["咨询", "问题", "技术支持"]
      }
    ]
  },
  "actions": [
    {
      "type": "ai_reply",
      "use_knowledge_base": true,
      "include_original": false,
      "temperature": 0.7
    }
  ]
}
```

### 4. 运行程序

```bash
# 测试LMStudio连接
python ai_engine.py

# 运行邮件助手
python outlook_assistant_win_standalone.py --once
```

## 🔧 配置详解

### LMStudio配置（lmstudio）

| 字段 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| enabled | bool | 否 | false | 是否启用AI功能 |
| base_url | string | 否 | http://localhost:1234 | LMStudio服务器地址 |
| model | string | 否 | null | 指定模型名称，null使用默认 |
| timeout | int | 否 | 60 | API请求超时时间（秒） |
| system_prompt | string | 否 | 见下文 | AI系统提示词，定义角色和行为 |

**默认system_prompt：**
```
你是一个专业的客服助手。请根据客户邮件内容和提供的知识库信息，生成礼貌、专业、准确的回复。

要求：
1. 使用中文回复
2. 语气友好、专业
3. 基于知识库信息回答，如果不确定就诚实说明
4. 不要编造信息
5. 适当使用换行和列表提高可读性
```

### 知识库配置（knowledge_base）

| 字段 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| enabled | bool | 否 | false | 是否启用知识库 |
| path | string | 否 | knowledge_base | 知识库文件路径 |
| search_top_k | int | 否 | 3 | 每次查询最多使用几篇相关文档 |

### AI回复动作参数（ai_reply）

```json
{
  "type": "ai_reply",
  "use_knowledge_base": true,    // 是否搜索知识库
  "include_original": false,     // 是否包含原始邮件
  "subject": "RE: {original_subject}",  // 回复主题模板
  "temperature": 0.7             // AI创造性参数（0-1）
}
```

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| type | string | 是 | - | 固定值"ai_reply" |
| use_knowledge_base | bool | 否 | true | 是否从知识库搜索相关内容 |
| include_original | bool | 否 | false | 是否在回复中包含原始邮件 |
| subject | string | 否 | RE: {original_subject} | 回复主题模板 |
| temperature | float | 否 | 0.7 | AI温度参数，越高越有创造性 |

## 🧠 工作原理

### 1. 邮件处理流程

```
收到邮件
    ↓
匹配AI回复规则
    ↓
从知识库搜索相关内容
    ↓
构建AI提示词（系统提示 + 知识库内容 + 邮件内容）
    ↓
调用LMStudio API生成回复
    ↓
发送AI生成的回复
```

### 2. 知识库搜索

系统使用文本相似度算法从知识库中找出与邮件最相关的内容：
- 切分长文档为片段
- 计算相似度排序
- 取最相关的前N篇

### 3. AI提示词构建

```
[系统提示词]
你是一个专业的客服助手...

[知识库上下文]
[文档1]
产品A是我们公司的旗舰产品...

[文档2]
基础版：¥2999，专业版：¥4999...

[客户邮件]
主题：产品咨询
正文：
你好，我想了解一下产品A的价格是多少？

请根据以上信息生成回复：
```

## 📝 知识库编写建议

### 1. 文件组织

```
knowledge_base/
├── 01_product_info.txt      # 产品信息（编号前缀方便排序）
├── 02_pricing.md            # 价格信息
├── 03_shipping.txt          # 发货物流
├── 04_refund_policy.md      # 退款政策
└── 05_tech_support.txt      # 技术支持
```

### 2. 内容格式

**推荐格式：**
```markdown
# 标题

## 问题1
详细回答...

## 问题2
详细回答...

联系方式：
- 电话：400-xxx-xxxx
- 邮箱：support@company.com
```

**优化技巧：**
- 使用清晰的标题结构
- 重要的信息放在前面
- 每个段落不要太长（500字以内）
- 包含常见问题的关键词

### 3. 示例知识库

**product_manual.txt:**
```
# 智能手表X1用户手册

## 产品规格
- 屏幕：1.5英寸AMOLED
- 续航：14天
- 防水：5ATM
- 重量：45g

## 常见问题

Q: 如何开机？
A: 长按右侧按钮3秒直到出现logo。

Q: 如何连接手机？
A: 
1. 下载App"HealthMonitor"
2. 打开手机蓝牙
3. 在App中添加设备
4. 选择"Watch X1"

Q: 充电需要多久？
A: 充满约需2小时，快充15分钟可使用1天。

## 售后服务
保修期：2年
客服电话：400-xxx-xxxx
服务时间：周一至周日 9:00-21:00
```

## 🎨 自定义AI角色

### 示例1：技术支持风格

```json
{
  "system_prompt": "你是技术支持专家。回答要简洁、准确，必要时提供步骤说明。如果问题超出知识库范围，建议用户提交工单。"
}
```

### 示例2：销售顾问风格

```json
{
  "system_prompt": "你是热情的销售顾问。回答要突出产品优势，适当使用营销话术，最后引导客户下单或咨询。"
}
```

### 示例3：严谨的法务风格

```json
{
  "system_prompt": "你是法务专员。回答必须基于提供的政策文件，对于不确定的问题明确告知需要进一步确认，避免给出法律建议。"
}
```

## 🧪 测试AI功能

### 1. 测试LMStudio连接

```bash
python ai_engine.py
```

如果连接成功，会显示：
```
LMStudio连接成功！
...
AI生成的回复：
您好！感谢您的咨询。关于产品价格...
```

### 2. 测试知识库

```bash
python knowledge_base.py
```

会显示知识库统计和搜索结果示例。

### 3. 试运行模式

```bash
python outlook_assistant_win_standalone.py --dry-run --once
```

查看日志确认AI回复生成是否正常，但不会实际发送邮件。

## 🔍 故障排除

### 问题1: 无法连接到LMStudio

**错误**: `无法连接到LMStudio服务器`

**解决**:
1. 确认LMStudio已启动
2. 确认在LMStudio中点击了"Start Server"
3. 检查端口号是否正确（默认1234）
4. 检查防火墙设置

### 问题2: AI生成回复很慢

**原因**: 模型太大或电脑配置不足

**解决**:
1. 换用更小的模型（如Phi-2 2.7B）
2. 增加timeout时间：`"timeout": 120`
3. 使用GPU加速（如果显卡支持）

### 问题3: AI回复质量差

**优化方法**:
1. 完善知识库内容
2. 调整system_prompt明确角色和要求
3. 降低temperature（如0.5）让回答更稳定
4. 确保知识库文档包含相关关键词

### 问题4: 知识库内容未生效

**检查**:
1. 确认`knowledge_base.enabled`为true
2. 确认知识库路径正确
3. 查看日志中是否显示"知识库初始化成功"
4. 检查文件格式是否支持

### 问题5: 回复包含编造信息

**解决**:
1. 在system_prompt中强调"不要编造信息"
2. 完善知识库，覆盖更多问题
3. 对于无法回答的问题，AI会建议联系人工客服

## 📊 性能优化

### 模型选择建议

| 模型 | 显存需求 | 生成速度 | 中文能力 | 适用场景 |
|------|----------|----------|----------|----------|
| Phi-2 | 4GB | 快 | 一般 | 简单问答 |
| ChatGLM3-6B | 8GB | 中等 | 优秀 | 中文客服 |
| Qwen-7B | 10GB | 中等 | 优秀 | 通用场景 |
| Llama-2-7B | 10GB | 中等 | 良好 | 英文场景 |
| Mistral-7B | 10GB | 快 | 良好 | 平衡选择 |

### 知识库优化

- 总文档数控制在50个以内
- 每个文档大小不超过100KB
- 定期清理过期内容
- 使用关键词优化搜索命中率

## 🔄 进阶配置

### 多知识库支持

可以创建多个知识库目录，在不同规则中使用：

```json
{
  "rules": [
    {
      "id": "rule_tech",
      "actions": [
        {
          "type": "ai_reply",
          "knowledge_base_path": "kb_technical"  // 技术支持知识库
        }
      ]
    },
    {
      "id": "rule_sales",
      "actions": [
        {
          "type": "ai_reply",
          "knowledge_base_path": "kb_sales"  // 销售知识库
        }
      ]
    }
  ]
}
```

### 分层处理

先使用规则匹配，未匹配再用AI：

```json
{
  "rules": [
    {
      "id": "rule_qa",
      "name": "精确问答",
      "actions": [{"type": "qa_reply"}]  // 先尝试精确匹配
    },
    {
      "id": "rule_ai",
      "name": "AI智能回复",
      "conditions": {
        // 可以添加更多条件
      },
      "actions": [{"type": "ai_reply"}]  // 再用AI处理
    }
  ]
}
```

## 📞 技术支持

如有问题：
1. 查看日志文件 `outlook_assistant.log`
2. 确认LMStudio版本为最新
3. 检查模型是否正确加载
4. 测试知识库搜索功能

## 📄 更新日志

### v2.0.0 (2026-03-26)
- 新增AI智能回复功能
- 支持本地LMStudio部署
- 知识库文档检索
- 可自定义AI角色和行为
- 支持多种文档格式

---

**现在你可以拥有一个真正智能的邮件助手了！** 🤖✨

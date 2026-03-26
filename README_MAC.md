# Outlook邮件助手 - macOS版本使用说明

## 📋 功能特点

✅ **跨平台支持** - 使用Microsoft Graph API，支持macOS和Linux  
✅ **无需Windows Outlook** - 直接使用Microsoft 365/Outlook.com邮箱  
✅ **完整功能** - 支持自动回复、转发、移动、标记已读等操作  
✅ **规则引擎** - 强大的规则匹配系统  
✅ **Web界面** - 可视化配置规则（可选）  

## 🚀 快速开始

### 方法一：直接运行（推荐）

#### 1. 安装依赖

```bash
# 使用Homebrew安装Python（如果尚未安装）
brew install python

# 安装项目依赖
pip3 install -r requirements_mac.txt
```

#### 2. 首次配置

```bash
python3 outlook_assistant_mac.py --setup
```

按照提示完成：
1. 在Azure Portal注册应用并获取Client ID
2. 使用Microsoft账户登录授权

#### 3. 运行程序

```bash
# 运行一次
python3 outlook_assistant_mac.py --once

# 持续运行（每60秒检查一次）
python3 outlook_assistant_mac.py

# 试运行模式（不实际执行操作）
python3 outlook_assistant_mac.py --dry-run
```

### 方法二：使用安装脚本

```bash
chmod +x install_mac.sh
./install_mac.sh
```

## 📦 打包成独立App

如果你想打包成独立的 `.app` 应用，可以按以下步骤操作：

### 1. 安装 py2app

```bash
pip3 install py2app
```

### 2. 打包应用

```bash
python3 setup_mac.py py2app
```

### 3. 使用打包后的应用

打包完成后，在 `dist/` 目录下会生成 `OutlookAssistant.app`：

```bash
# 运行应用
open dist/OutlookAssistant.app

# 或者将应用移动到应用程序目录
mv dist/OutlookAssistant.app /Applications/
```

## 🔧 Azure AD 配置步骤

### 1. 注册Azure应用

1. 访问 https://portal.azure.com
2. 登录你的Microsoft账户
3. 搜索并进入 "Azure Active Directory"
4. 点击左侧菜单 "App registrations"
5. 点击 "New registration"

### 2. 配置应用

- **Name**: Outlook邮件助手（或其他你喜欢的名字）
- **Supported account types**: Accounts in any organizational directory and personal Microsoft accounts
- **Redirect URI**: 
  - Platform: Public client (mobile & desktop)
  - URI: `http://localhost:8080`

### 3. 添加API权限

1. 进入应用详情页
2. 点击 "API permissions"
3. 点击 "Add a permission"
4. 选择 "Microsoft Graph"
5. 选择 "Delegated permissions"
6. 添加以下权限：
   - `Mail.Read` - 读取邮件
   - `Mail.ReadWrite` - 读写邮件
   - `Mail.Send` - 发送邮件
   - `User.Read` - 读取用户信息
7. 点击 "Grant admin consent"（如果是个人账户则不需要）

### 4. 获取Client ID

1. 在应用详情页的 "Overview" 页面
2. 复制 **Application (client) ID**
3. 在程序设置中输入此ID

## 📁 文件说明

```
MAILOPS/
├── outlook_assistant_mac.py    # ⭐ macOS主程序
├── requirements_mac.txt        # macOS依赖
├── setup_mac.py                # 打包配置
├── install_mac.sh              # 安装脚本
├── config.json                 # 配置文件（自动创建）
└── token_cache.json           # Token缓存（自动创建）
```

## ⚙️ 配置文件

首次运行后会自动生成 `config.json`：

```json
{
  "azure_ad": {
    "client_id": "your-client-id",
    "tenant_id": "common",
    "token_cache_path": "token_cache.json"
  },
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
    "auto_reply": {
      "subject": "RE: {original_subject}",
      "body": "感谢您的来信..."
    }
  },
  "settings": {
    "check_interval": 60,
    "process_unread_only": true,
    "max_emails_per_batch": 50
  }
}
```

## 🔍 故障排除

### 问题1: 缺少msal模块
**错误**: `ModuleNotFoundError: No module named 'msal'`

**解决**:
```bash
pip3 install msal requests python-dateutil
```

### 问题2: 认证失败
**错误**: `认证失败` 或无法获取token

**解决**:
1. 检查Client ID是否正确
2. 确保Azure AD中已添加正确的API权限
3. 删除 `token_cache.json` 重新认证
4. 检查网络连接

### 问题3: 无法访问邮箱
**错误**: `403 Forbidden` 或 `Access denied`

**解决**:
1. 确保使用的是Microsoft 365或Outlook.com账户
2. 检查Azure AD中的API权限是否已授予
3. 如果是工作账户，可能需要管理员授权

### 问题4: Python版本过低
**错误**: `SyntaxError` 或其他Python错误

**解决**:
```bash
# 检查Python版本
python3 --version

# 如果低于3.7，请升级
brew install python@3.11
brew link python@3.11
```

### 问题5: 打包后无法运行
**错误**: App打开后立即关闭

**解决**:
1. 检查是否包含了所有依赖
2. 尝试在终端运行查看错误:
   ```bash
   /Applications/OutlookAssistant.app/Contents/MacOS/OutlookAssistant
   ```
3. 确保 `config.json` 在应用包内或用户目录下

## 📝 命令行参数

```bash
python3 outlook_assistant_mac.py [选项]

选项:
  --setup              首次设置向导
  --once               只运行一次
  --dry-run            试运行模式（不实际执行操作）
  -h, --help           显示帮助信息
```

## 🌐 与Web界面配合使用

macOS版本也可以配合Web控制界面使用：

```bash
# 安装Web依赖
pip3 install flask flask-cors

# 运行Web服务
python3 web_app.py
```

然后在浏览器中打开 http://localhost:5000

**注意**: Web版本在macOS上使用Graph API，配置与命令行版本相同。

## 🔐 安全性

1. **Token安全**: 访问token保存在本地 `token_cache.json`，不会上传到任何服务器
2. **Client ID**: 不要在代码中硬编码Client ID，使用配置文件
3. **权限最小化**: 只申请必要的API权限
4. **定期清理**: 定期清理日志文件和token缓存

## 📊 日志查看

```bash
# 实时查看日志
tail -f outlook_assistant.log

# 查看最后100行
tail -n 100 outlook_assistant.log
```

## 🔄 自动运行（可选）

### 使用launchd设置定时任务

创建 `~/Library/LaunchAgents/com.outlook.assistant.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.outlook.assistant</string>
    <key>ProgramArguments</key>
    <array>
        <string>/usr/local/bin/python3</string>
        <string>/path/to/outlook_assistant_mac.py</string>
        <string>--once</string>
    </array>
    <key>StartInterval</key>
    <integer>300</integer>
    <key>StandardOutPath</key>
    <string>/path/to/outlook_assistant.log</string>
    <key>StandardErrorPath</key>
    <string>/path/to/outlook_assistant_error.log</string>
</dict>
</plist>
```

加载配置：
```bash
launchctl load ~/Library/LaunchAgents/com.outlook.assistant.plist
```

## 📞 技术支持

如有问题：
1. 查看日志文件 `outlook_assistant.log`
2. 检查Azure AD配置
3. 确保Microsoft账户正常

## 📄 许可证

MIT License

## 🎉 版本历史

### v1.0.0 (2026-03-26)
- 初始macOS版本发布
- 支持Microsoft Graph API
- 完整的邮件自动处理功能
- Web界面支持

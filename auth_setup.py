#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Microsoft Graph API认证工具
用于首次登录和获取访问令牌

使用方法:
    python auth_setup.py

说明:
    1. 运行此脚本会打开浏览器或显示设备代码
    2. 使用Microsoft账户登录
    3. 授权应用访问邮件
    4. 成功后token会被保存，主程序可以直接使用
"""

import os
import sys
import json
import logging

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def load_config():
    """加载配置"""
    config_path = "config.json"
    if not os.path.exists(config_path):
        logger.error(f"配置文件不存在: {config_path}")
        return None

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"加载配置失败: {e}")
        return None


def main():
    """主函数"""
    print("=" * 60)
    print("Outlook邮件助手 - Microsoft Graph API认证")
    print("=" * 60)
    print()

    # 加载配置
    config = load_config()
    if not config:
        print("错误: 无法加载配置文件")
        sys.exit(1)

    # 获取Azure AD配置
    azure_config = config.get("azure_ad", {})
    client_id = azure_config.get("client_id")

    if not client_id:
        print("错误: 配置文件中缺少 Azure AD Client ID")
        print()
        print("请按以下步骤配置:")
        print("1. 访问 https://portal.azure.com")
        print(
            "2. 注册新应用: Azure Active Directory -> App registrations -> New registration"
        )
        print("3. 设置应用名称和重定向URI (http://localhost:8080)")
        print("4. 在API权限中添加: Microsoft Graph -> Delegated permissions:")
        print("   - Mail.Read")
        print("   - Mail.ReadWrite")
        print("   - Mail.Send")
        print("   - User.Read")
        print("5. 复制Application (client) ID到配置文件")
        print()
        print("配置文件格式:")
        print(
            json.dumps(
                {
                    "azure_ad": {
                        "client_id": "your-client-id-here",
                        "tenant_id": "common",
                    }
                },
                indent=2,
            )
        )
        sys.exit(1)

    print(f"Client ID: {client_id}")
    print()

    try:
        # 导入认证模块
        from graph_auth import GraphAuth

        # 创建认证实例
        auth = GraphAuth(client_id=client_id)

        # 执行认证
        print("开始认证流程...")
        print()

        if auth.authenticate():
            print()
            print("=" * 60)
            print("认证成功!")
            print("=" * 60)
            print()
            print("Token已保存，现在可以运行主程序:")
            print("  python outlook_assistant.py")
            print()
            return 0
        else:
            print()
            print("=" * 60)
            print("认证失败!")
            print("=" * 60)
            return 1

    except ImportError as e:
        print(f"错误: 缺少必要的依赖 - {e}")
        print()
        print("请安装依赖:")
        print("  pip install -r requirements.txt")
        return 1

    except Exception as e:
        print(f"错误: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

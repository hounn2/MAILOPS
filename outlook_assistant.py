#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook邮件助手主程序 - Graph API版本
支持跨平台（Windows/macOS/Linux）

使用方法:
    python outlook_assistant.py [--config path/to/config.json] [--once]

参数:
    --config: 配置文件路径 (默认: config.json)
    --once:   只运行一次，不进入循环模式
    --dry-run: 试运行模式，不实际执行操作

首次使用:
    1. 在Azure Portal注册应用获取Client ID
    2. 运行: python auth_setup.py 完成认证
    3. 运行: python outlook_assistant.py
"""

import os
import sys
import json
import time
import logging
import argparse
from datetime import datetime
from typing import Dict, Any

# 导入自定义模块
from rules import RuleEngine, TemplateEngine
from graph_auth import GraphAuth
from graph_actions import GraphOutlookActions, GraphActionExecutor


# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("outlook_assistant.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)


class OutlookAssistant:
    """Outlook邮件助手主类 - Graph API版本"""

    def __init__(self, config_path: str, dry_run: bool = False):
        self.config_path = config_path
        self.dry_run = dry_run
        self.config = self._load_config()
        self.rule_engine = None
        self.template_engine = None
        self.auth = None
        self.outlook_actions = None
        self.action_executor = None
        self._init_engines()

    def _load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            logger.info(f"成功加载配置文件: {self.config_path}")
            return config
        except Exception as e:
            logger.error(f"加载配置文件失败: {e}")
            raise

    def _init_engines(self):
        """初始化引擎"""
        try:
            # 初始化模板引擎
            templates = self.config.get("templates", {})
            self.template_engine = TemplateEngine(templates)

            # 初始化规则引擎
            rules = self.config.get("rules", [])
            self.rule_engine = RuleEngine(rules)

            # 获取Azure AD配置
            azure_config = self.config.get("azure_ad", {})
            client_id = azure_config.get("client_id")

            if not client_id:
                logger.error("配置文件缺少 Azure AD Client ID")
                logger.error("请先在 Azure Portal 注册应用并配置 client_id")
                logger.error("然后运行: python auth_setup.py 完成认证")
                raise ValueError("缺少 Azure AD Client ID")

            token_cache_path = azure_config.get("token_cache_path", "token_cache.json")

            # 初始化Graph认证
            self.auth = GraphAuth(
                client_id=client_id, token_cache_path=token_cache_path
            )

            # 测试认证
            logger.info("正在验证Graph API认证...")
            if not self.auth.authenticate():
                raise Exception("Graph API认证失败")
            logger.info("Graph API认证成功")

            # 初始化Outlook操作
            settings = self.config.get("settings", {})
            dry_run = self.dry_run or settings.get("dry_run", False)
            self.outlook_actions = GraphOutlookActions(auth=self.auth, dry_run=dry_run)

            # 初始化动作执行器
            self.action_executor = GraphActionExecutor(
                self.outlook_actions, self.template_engine
            )

            logger.info("所有引擎初始化成功")

        except Exception as e:
            logger.error(f"初始化引擎失败: {e}")
            raise

    def reload_config(self):
        """重新加载配置"""
        logger.info("重新加载配置...")
        self.config = self._load_config()
        self._init_engines()
        logger.info("配置重新加载完成")

    def process_emails(self) -> Dict[str, Any]:
        """
        处理邮件

        Returns:
            处理结果统计
        """
        settings = self.config.get("settings", {})
        unread_only = settings.get("process_unread_only", True)
        excluded_folders = settings.get("excluded_folders", [])
        max_emails = settings.get("max_emails_per_batch", 50)
        mark_as_read = settings.get("mark_as_read_after_process", True)

        stats = {
            "processed": 0,
            "matched": 0,
            "actions_executed": 0,
            "errors": 0,
            "start_time": datetime.now(),
        }

        try:
            # 获取邮件
            emails = self.outlook_actions.get_inbox_emails(
                unread_only=unread_only,
                excluded_folders=excluded_folders,
                max_emails=max_emails,
            )

            logger.info(f"开始处理 {len(emails)} 封邮件...")

            for email_item in emails:
                try:
                    # 提取邮件数据
                    email_data = self.outlook_actions.get_email_data(email_item)

                    if not email_data:
                        logger.warning("无法提取邮件数据，跳过")
                        continue

                    stats["processed"] += 1

                    # 匹配规则
                    matched_rules = self.rule_engine.match_email(email_data)

                    if matched_rules:
                        stats["matched"] += 1
                        logger.info(
                            f"邮件 '{email_data.get('subject')}' 匹配了 {len(matched_rules)} 条规则"
                        )

                        # 执行动作
                        for rule in matched_rules:
                            actions = rule.get("actions", [])
                            if actions:
                                result = self.action_executor.execute_actions(
                                    email_item, actions, email_data
                                )
                                stats["actions_executed"] += result["success"]
                    else:
                        logger.debug(
                            f"邮件 '{email_data.get('subject')}' 未匹配任何规则"
                        )

                    # 标记为已读
                    if mark_as_read and not email_data.get("is_read", False):
                        self.outlook_actions.mark_as_read(email_item)

                except Exception as e:
                    logger.error(f"处理单封邮件时出错: {e}")
                    stats["errors"] += 1
                    continue

            stats["end_time"] = datetime.now()
            stats["duration"] = (
                stats["end_time"] - stats["start_time"]
            ).total_seconds()

            logger.info(
                f"处理完成: 处理 {stats['processed']} 封, "
                f"匹配 {stats['matched']} 封, "
                f"执行 {stats['actions_executed']} 个动作, "
                f"错误 {stats['errors']} 个"
            )

            return stats

        except Exception as e:
            logger.error(f"处理邮件过程中出错: {e}")
            stats["errors"] += 1
            return stats

    def run(self, once: bool = False):
        """
        运行助手

        Args:
            once: 是否只运行一次
        """
        settings = self.config.get("settings", {})
        interval = settings.get("check_interval", 60)

        logger.info("Outlook邮件助手启动 (Graph API版本)")
        logger.info(f"运行模式: {'单次' if once else '循环'}")
        logger.info(f"检查间隔: {interval}秒")
        logger.info(f"试运行模式: {self.dry_run or settings.get('dry_run', False)}")

        try:
            while True:
                # 处理邮件
                stats = self.process_emails()

                # 如果只运行一次
                if once:
                    logger.info("单次运行完成")
                    break

                # 等待下一次检查
                logger.info(f"等待 {interval} 秒后再次检查...")
                time.sleep(interval)

        except KeyboardInterrupt:
            logger.info("用户中断，程序退出")
        except Exception as e:
            logger.error(f"运行出错: {e}")
            raise


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="Outlook邮件助手 - Microsoft Graph API版本\n支持跨平台（Windows/macOS/Linux）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
    # 首次使用 - 先完成认证
    python auth_setup.py
    
    # 使用默认配置运行
    python outlook_assistant.py
    
    # 只运行一次
    python outlook_assistant.py --once
    
    # 试运行模式
    python outlook_assistant.py --dry-run
    
    # 使用自定义配置
    python outlook_assistant.py -c myconfig.json

环境要求:
    - Python 3.7+
    - Microsoft 365账户
    - Azure AD应用注册
        """,
    )

    parser.add_argument(
        "-c", "--config", default="config.json", help="配置文件路径 (默认: config.json)"
    )

    parser.add_argument("--once", action="store_true", help="只运行一次，不进入循环")

    parser.add_argument(
        "--dry-run", action="store_true", help="试运行模式，不实际执行操作"
    )

    args = parser.parse_args()

    # 检查配置文件是否存在
    if not os.path.exists(args.config):
        logger.error(f"配置文件不存在: {args.config}")
        print(f"\n错误: 配置文件不存在: {args.config}")
        print("\n请创建配置文件 config.json，包含以下内容:")
        print(
            json.dumps(
                {
                    "azure_ad": {
                        "client_id": "your-azure-app-client-id",
                        "tenant_id": "common",
                        "token_cache_path": "token_cache.json",
                    },
                    "rules": [],
                    "templates": {},
                    "settings": {
                        "check_interval": 60,
                        "process_unread_only": True,
                        "max_emails_per_batch": 50,
                    },
                },
                indent=2,
            )
        )
        sys.exit(1)

    try:
        # 创建并运行助手
        assistant = OutlookAssistant(args.config, dry_run=args.dry_run)
        assistant.run(once=args.once)
    except Exception as e:
        logger.error(f"程序异常退出: {e}")
        print(f"\n程序异常: {e}")
        print("\n如果需要帮助，请查看 README.md 或运行: python auth_setup.py")
        sys.exit(1)


if __name__ == "__main__":
    main()

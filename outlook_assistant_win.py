#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook邮件助手主程序 - Windows COM接口版本
使用pywin32直接操作本地Outlook客户端

使用方法:
    python outlook_assistant_win.py [--config path/to/config.json] [--once]

参数:
    --config: 配置文件路径 (默认: config.json)
    --once:   只运行一次，不进入循环模式
    --dry-run: 试运行模式，不实际执行操作

环境要求:
    - Windows操作系统
    - Microsoft Outlook已安装并配置
    - Python 3.7+
    - pywin32库
"""

import os
import sys
import json
import time
import logging
import argparse
from datetime import datetime
from typing import Dict, Any

# 确保在Windows上运行
if sys.platform != "win32":
    print("错误: 此版本仅支持Windows操作系统")
    print("如需跨平台版本，请使用 outlook_assistant.py (Graph API版本)")
    sys.exit(1)

# 导入自定义模块
from rules import RuleEngine, TemplateEngine
from actions import OutlookActions, ActionExecutor


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


class OutlookAssistantWindows:
    """Outlook邮件助手主类 - Windows COM版本"""

    def __init__(self, config_path: str, dry_run: bool = False):
        self.config_path = config_path
        self.dry_run = dry_run
        self.config = self._load_config()
        self.rule_engine = None
        self.template_engine = None
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

            # 初始化Outlook操作（COM接口）
            settings = self.config.get("settings", {})
            dry_run = self.dry_run or settings.get("dry_run", False)
            logger.info("正在连接到Outlook...")
            self.outlook_actions = OutlookActions(dry_run=dry_run)

            # 初始化动作执行器
            self.action_executor = ActionExecutor(
                self.outlook_actions, self.template_engine
            )

            logger.info("所有引擎初始化成功")

        except Exception as e:
            logger.error(f"初始化引擎失败: {e}")
            logger.error("请确保Microsoft Outlook已运行并正确配置")
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

        logger.info("Outlook邮件助手启动 (Windows COM版本)")
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
        description="Outlook邮件助手 - Windows COM接口版本",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
    python outlook_assistant_win.py                    # 使用默认配置运行
    python outlook_assistant_win.py --once            # 只运行一次
    python outlook_assistant_win.py --dry-run         # 试运行模式
    python outlook_assistant_win.py -c myconfig.json  # 使用自定义配置

环境要求:
    - Windows操作系统
    - Microsoft Outlook已安装并运行
    - Python 3.7+
    - pywin32库 (pip install pywin32)

注意:
    此版本仅适用于Windows。如需跨平台版本，请使用 outlook_assistant.py
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
        sys.exit(1)

    try:
        # 创建并运行助手
        assistant = OutlookAssistantWindows(args.config, dry_run=args.dry_run)
        assistant.run(once=args.once)
    except Exception as e:
        logger.error(f"程序异常退出: {e}")
        print(f"\n程序异常: {e}")
        print("\n请确保:")
        print("  1. Microsoft Outlook已安装并运行")
        print("  2. pywin32已安装: pip install pywin32")
        print("  3. 配置文件格式正确")
        sys.exit(1)


if __name__ == "__main__":
    main()

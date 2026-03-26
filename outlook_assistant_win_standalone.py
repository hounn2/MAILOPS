#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook邮件助手 - Windows单文件版本
所有功能集成在一个文件中，无需导入其他模块

使用方法:
    python outlook_assistant_win_standalone.py [--config path/to/config.json] [--once]

环境要求:
    - Windows操作系统
    - Microsoft Outlook已安装并配置
    - Python 3.7+
    - pywin32库
"""

import os
import sys
import json
import re
import time
import logging
import argparse
from datetime import datetime
from typing import List, Dict, Any, Optional

# 检查Windows平台
if sys.platform != "win32":
    print("错误: 此版本仅支持Windows操作系统")
    sys.exit(1)

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

# ============================================================================
# 规则引擎（内嵌）
# ============================================================================


class RuleEngine:
    """规则引擎主类"""

    def __init__(self, rules: List[Dict[str, Any]]):
        self.rules = rules
        self._compile_rules()

    def _compile_rules(self):
        """编译规则，准备匹配"""
        self.compiled_rules = []
        for rule in self.rules:
            if not rule.get("enabled", True):
                continue
            compiled_rule = {
                "id": rule["id"],
                "name": rule["name"],
                "conditions": self._compile_conditions(rule["conditions"]),
                "actions": rule.get("actions", []),
            }
            self.compiled_rules.append(compiled_rule)

    def _compile_conditions(self, conditions: Dict[str, Any]) -> Dict[str, Any]:
        """编译条件配置"""
        return {
            "match_all": conditions.get("match_all", True),
            "items": [
                self._compile_condition_item(item)
                for item in conditions.get("items", [])
            ],
        }

    def _compile_condition_item(self, item: Dict[str, Any]) -> Dict[str, Any]:
        """编译单个条件项"""
        field = item["field"]
        operator = item["operator"]
        value = item["value"]

        if operator in ["in", "not_in"] and isinstance(value, list):
            value = set(value)

        if operator == "contains" and isinstance(value, list):
            patterns = [re.compile(re.escape(v), re.IGNORECASE) for v in value]
            value = patterns

        return {"field": field, "operator": operator, "value": value}

    def match_email(self, email_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """匹配邮件数据，返回匹配的规则列表"""
        matched_rules = []

        for rule in self.compiled_rules:
            if self._evaluate_conditions(rule["conditions"], email_data):
                matched_rules.append(rule)
                logger.info(
                    f"邮件 '{email_data.get('subject', 'Unknown')}' 匹配规则: {rule['name']}"
                )

        return matched_rules

    def _evaluate_conditions(
        self, conditions: Dict[str, Any], email_data: Dict[str, Any]
    ) -> bool:
        """评估条件是否满足"""
        match_all = conditions["match_all"]
        items = conditions["items"]

        if not items:
            return True

        results = [self._evaluate_condition_item(item, email_data) for item in items]

        if match_all:
            return all(results)
        else:
            return any(results)

    def _evaluate_condition_item(
        self, item: Dict[str, Any], email_data: Dict[str, Any]
    ) -> bool:
        """评估单个条件项"""
        field = item["field"]
        operator = item["operator"]
        value = item["value"]

        email_value = self._get_email_field(email_data, field)

        if email_value is None:
            return False

        return self._apply_operator(email_value, operator, value)

    def _get_email_field(self, email_data: Dict[str, Any], field: str) -> Any:
        """获取邮件字段值"""
        if field == "sender_domain":
            sender = email_data.get("sender", "")
            if "@" in sender:
                return sender[sender.index("@") :]
            return ""
        elif field == "body_length":
            body = email_data.get("body", "")
            return len(body) if body else 0
        elif field == "has_attachments":
            attachments = email_data.get("attachments", [])
            return len(attachments) > 0
        elif field == "received_time":
            return email_data.get("received_time")
        else:
            return email_data.get(field)

    def _apply_operator(self, email_value: Any, operator: str, rule_value: Any) -> bool:
        """应用操作符进行比较"""
        try:
            if operator == "equals":
                return str(email_value).lower() == str(rule_value).lower()
            elif operator == "not_equals":
                return str(email_value).lower() != str(rule_value).lower()
            elif operator == "contains":
                email_str = str(email_value).lower()
                if isinstance(rule_value, list):
                    return any(pattern.search(email_str) for pattern in rule_value)
                return str(rule_value).lower() in email_str
            elif operator == "not_contains":
                email_str = str(email_value).lower()
                if isinstance(rule_value, list):
                    return not any(pattern.search(email_str) for pattern in rule_value)
                return str(rule_value).lower() not in email_str
            elif operator == "starts_with":
                return str(email_value).lower().startswith(str(rule_value).lower())
            elif operator == "ends_with":
                return str(email_value).lower().endswith(str(rule_value).lower())
            elif operator == "in":
                return str(email_value).lower() in {str(v).lower() for v in rule_value}
            elif operator == "not_in":
                return str(email_value).lower() not in {
                    str(v).lower() for v in rule_value
                }
            elif operator == "regex":
                if isinstance(rule_value, str):
                    pattern = re.compile(rule_value, re.IGNORECASE)
                    return bool(pattern.search(str(email_value)))
                return False
            elif operator == "greater_than":
                return float(email_value) > float(rule_value)
            elif operator == "less_than":
                return float(email_value) < float(rule_value)
            elif operator == "between":
                if isinstance(rule_value, (list, tuple)) and len(rule_value) == 2:
                    val = float(email_value)
                    return float(rule_value[0]) <= val <= float(rule_value[1])
                return False
            else:
                logger.warning(f"未知操作符: {operator}")
                return False
        except Exception as e:
            logger.error(f"评估操作符时出错: {operator}, error: {e}")
            return False


class TemplateEngine:
    """模板引擎，用于处理回复模板"""

    def __init__(self, templates: Dict[str, Any]):
        self.templates = templates

    def render(self, template_name: str, context: Dict[str, Any]) -> Dict[str, str]:
        """渲染模板"""
        template = self.templates.get(template_name)
        if not template:
            logger.error(f"模板未找到: {template_name}")
            return {"subject": "", "body": ""}

        subject = self._render_template(template.get("subject", ""), context)
        body = self._render_template(template.get("body", ""), context)

        return {"subject": subject, "body": body}

    def _render_template(self, template_str: str, context: Dict[str, Any]) -> str:
        """使用简单字符串替换渲染模板"""
        result = template_str
        for key, value in context.items():
            placeholder = f"{{{key}}}"
            result = result.replace(placeholder, str(value))
        return result


# ============================================================================
# Outlook操作（内嵌）
# ============================================================================


class OutlookActions:
    """Outlook邮件操作类"""

    def __init__(self, dry_run: bool = False):
        self.dry_run = dry_run
        self.outlook = None
        self.namespace = None
        self._connect()

    def _connect(self):
        """连接到Outlook应用"""
        try:
            import win32com.client
            import pythoncom

            # 初始化COM线程（解决多线程或Flask中的COM初始化问题）
            try:
                pythoncom.CoInitialize()
            except:
                pass  # 可能已经初始化过

            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            logger.info("成功连接到Outlook")
        except Exception as e:
            logger.error(f"连接Outlook失败: {e}")
            raise

    def reply_email(
        self, email_item, subject: str, body: str, include_original: bool = False
    ) -> bool:
        """回复邮件"""
        try:
            if self.dry_run:
                logger.info(f"[DRY RUN] 将回复邮件: {email_item.Subject}")
                logger.info(f"[DRY RUN] 回复主题: {subject}")
                logger.info(f"[DRY RUN] 回复内容: {body[:100]}...")
                return True

            reply = email_item.Reply()
            reply.Subject = subject

            if include_original:
                reply.Body = body + "\n\n--- 原始邮件 ---\n" + email_item.Body
            else:
                reply.Body = body

            reply.Send()
            logger.info(f"成功回复邮件: {email_item.Subject}")
            return True

        except Exception as e:
            logger.error(f"回复邮件失败: {e}")
            return False

    def forward_email(
        self,
        email_item,
        to: List[str],
        subject_prefix: str = "",
        additional_body: str = "",
    ) -> bool:
        """转发邮件"""
        try:
            if self.dry_run:
                logger.info(f"[DRY RUN] 将转发邮件: {email_item.Subject}")
                logger.info(f"[DRY RUN] 转发到: {', '.join(to)}")
                return True

            forward = email_item.Forward()
            recipients_str = "; ".join(to)
            forward.To = recipients_str

            if subject_prefix:
                forward.Subject = subject_prefix + email_item.Subject

            if additional_body:
                forward.Body = additional_body + "\n\n" + forward.Body

            forward.Send()
            logger.info(f"成功转发邮件 '{email_item.Subject}' 到: {recipients_str}")
            return True

        except Exception as e:
            logger.error(f"转发邮件失败: {e}")
            return False

    def move_email(self, email_item, target_folder: str) -> bool:
        """移动邮件到指定文件夹"""
        try:
            if self.dry_run:
                logger.info(
                    f"[DRY RUN] 将移动邮件 '{email_item.Subject}' 到文件夹: {target_folder}"
                )
                return True

            store = email_item.Parent.Store
            folder = self._find_folder(store, target_folder)

            if folder:
                email_item.Move(folder)
                logger.info(
                    f"成功移动邮件 '{email_item.Subject}' 到文件夹: {target_folder}"
                )
                return True
            else:
                logger.warning(f"目标文件夹未找到: {target_folder}")
                return False

        except Exception as e:
            logger.error(f"移动邮件失败: {e}")
            return False

    def _find_folder(self, store, folder_name: str):
        """在存储中查找文件夹"""
        try:
            inbox = store.GetDefaultFolder(6)
            folder = self._search_folder_recursive(inbox, folder_name)
            if folder:
                return folder

            root_folder = store.GetRootFolder()
            folder = self._search_folder_recursive(root_folder, folder_name)

            return folder

        except Exception as e:
            logger.error(f"查找文件夹失败: {e}")
            return None

    def _search_folder_recursive(self, parent_folder, target_name: str):
        """递归搜索文件夹"""
        try:
            if parent_folder.Name == target_name:
                return parent_folder

            for folder in parent_folder.Folders:
                result = self._search_folder_recursive(folder, target_name)
                if result:
                    return result

            return None

        except Exception as e:
            logger.error(f"搜索文件夹出错: {e}")
            return None

    def mark_as_read(self, email_item) -> bool:
        """标记邮件为已读"""
        try:
            if self.dry_run:
                logger.info(f"[DRY RUN] 将标记邮件为已读: {email_item.Subject}")
                return True

            email_item.UnRead = False
            logger.info(f"标记邮件为已读: {email_item.Subject}")
            return True

        except Exception as e:
            logger.error(f"标记邮件已读失败: {e}")
            return False

    def get_email_data(self, email_item) -> Dict[str, Any]:
        """从Outlook邮件对象提取数据"""
        try:
            # 安全地获取发件人信息
            sender = ""
            try:
                if email_item.Sender:
                    # 尝试获取EmailAddress，如果不存在则尝试其他属性
                    if hasattr(email_item.Sender, "EmailAddress"):
                        sender = email_item.Sender.EmailAddress
                    elif hasattr(email_item.Sender, "Address"):
                        sender = email_item.Sender.Address
                    elif hasattr(email_item.Sender, "Name"):
                        sender = email_item.Sender.Name
            except Exception as e:
                logger.debug(f"获取发件人信息失败: {e}")
                sender = ""

            # 安全地获取收件人信息
            recipients = []
            try:
                for recipient in email_item.Recipients:
                    try:
                        addr = (
                            recipient.Address
                            if hasattr(recipient, "Address")
                            else str(recipient)
                        )
                        recipients.append(addr)
                    except:
                        pass
            except Exception as e:
                logger.debug(f"获取收件人信息失败: {e}")

            # 安全地获取附件信息
            attachments = []
            try:
                for attachment in email_item.Attachments:
                    try:
                        filename = (
                            attachment.FileName
                            if hasattr(attachment, "FileName")
                            else "unknown"
                        )
                        size = attachment.Size if hasattr(attachment, "Size") else 0
                        attachments.append({"filename": filename, "size": size})
                    except:
                        pass
            except Exception as e:
                logger.debug(f"获取附件信息失败: {e}")

            return {
                "entry_id": getattr(email_item, "EntryID", ""),
                "subject": getattr(email_item, "Subject", ""),
                "body": getattr(email_item, "Body", ""),
                "html_body": getattr(email_item, "HTMLBody", ""),
                "sender": sender,
                "recipients": recipients,
                "received_time": getattr(email_item, "ReceivedTime", None),
                "sent_time": getattr(email_item, "SentOn", None),
                "is_read": not getattr(email_item, "UnRead", False),
                "importance": getattr(email_item, "Importance", 1),
                "attachments": attachments,
                "conversation_id": getattr(email_item, "ConversationID", ""),
                "message_class": getattr(email_item, "MessageClass", ""),
            }

        except Exception as e:
            logger.error(f"提取邮件数据失败: {e}")
            return {}

    def get_inbox_emails(
        self, unread_only: bool = True, excluded_folders=None, max_emails: int = 50
    ) -> List[Any]:
        """获取收件箱中的邮件"""
        if excluded_folders is None:
            excluded_folders = []

        emails = []

        try:
            if not self.namespace:
                logger.error("未连接到Outlook")
                return []

            for store in self.namespace.Stores:
                try:
                    inbox = store.GetDefaultFolder(6)

                    if inbox.Name in excluded_folders:
                        continue

                    items = inbox.Items
                    items.Sort("[ReceivedTime]", True)

                    for item in items:
                        if len(emails) >= max_emails:
                            break

                        if item.MessageClass not in ["IPM.Note", "IPM.Post"]:
                            continue

                        if unread_only and not item.UnRead:
                            continue

                        emails.append(item)

                except Exception as e:
                    logger.error(f"获取账户邮件失败: {e}")
                    continue

            logger.info(f"获取到 {len(emails)} 封邮件")
            return emails

        except Exception as e:
            logger.error(f"获取收件箱邮件失败: {e}")
            return []


class ActionExecutor:
    """操作执行器，负责执行匹配规则的动作"""

    def __init__(self, outlook_actions: OutlookActions, template_engine):
        self.outlook = outlook_actions
        self.template_engine = template_engine
        self.action_handlers = {
            "reply": self._handle_reply,
            "forward": self._handle_forward,
            "move": self._handle_move,
            "mark_as_read": self._handle_mark_as_read,
        }

    def execute_actions(
        self, email_item, actions: List[Dict[str, Any]], email_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """执行一系列动作"""
        results = {"total": len(actions), "success": 0, "failed": 0, "actions": []}

        for action in actions:
            action_type = action.get("type", "")
            if not action_type:
                logger.warning("动作缺少类型")
                results["failed"] += 1
                continue

            handler = self.action_handlers.get(action_type)

            if handler:
                try:
                    success = handler(email_item, action, email_data)
                    if success:
                        results["success"] += 1
                    else:
                        results["failed"] += 1
                    results["actions"].append({"type": action_type, "success": success})
                except Exception as e:
                    logger.error(f"执行动作 {action_type} 失败: {e}")
                    results["failed"] += 1
                    results["actions"].append(
                        {"type": action_type, "success": False, "error": str(e)}
                    )
            else:
                logger.warning(f"未知的动作类型: {action_type}")
                results["failed"] += 1

        return results

    def _handle_reply(
        self, email_item, action: Dict[str, Any], email_data: Dict[str, Any]
    ) -> bool:
        """处理回复动作"""
        template_name = action.get("template")
        include_original = action.get("include_original", False)

        if template_name:
            context = {
                "original_subject": email_data.get("subject", ""),
                "sender": email_data.get("sender", ""),
                "received_time": str(email_data.get("received_time", "")),
                "return_date": "2026-01-01",
                "backup_contact": "backup@company.com",
            }

            template_result = self.template_engine.render(template_name, context)
            subject = template_result["subject"]
            body = template_result["body"]
        else:
            subject = action.get("subject", f"RE: {email_data.get('subject', '')}")
            body = action.get("body", "")

        return self.outlook.reply_email(email_item, subject, body, include_original)

    def _handle_forward(
        self, email_item, action: Dict[str, Any], email_data: Dict[str, Any]
    ) -> bool:
        """处理转发动作"""
        to = action.get("to", [])
        subject_prefix = action.get("subject_prefix", "")
        additional_body = action.get("additional_body", "")

        if not to:
            logger.error("转发动作缺少收件人")
            return False

        return self.outlook.forward_email(
            email_item, to, subject_prefix, additional_body
        )

    def _handle_move(
        self, email_item, action: Dict[str, Any], email_data: Dict[str, Any]
    ) -> bool:
        """处理移动动作"""
        target_folder = action.get("target")

        if not target_folder:
            logger.error("移动动作缺少目标文件夹")
            return False

        return self.outlook.move_email(email_item, target_folder)

    def _handle_mark_as_read(
        self, email_item, action: Dict[str, Any], email_data: Dict[str, Any]
    ) -> bool:
        """处理标记已读动作"""
        return self.outlook.mark_as_read(email_item)


# ============================================================================
# 主程序
# ============================================================================


class OutlookAssistantWindows:
    """Outlook邮件助手主类 - Windows版本"""

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
            templates = self.config.get("templates", {})
            self.template_engine = TemplateEngine(templates)

            rules = self.config.get("rules", [])
            self.rule_engine = RuleEngine(rules)

            settings = self.config.get("settings", {})
            dry_run = self.dry_run or settings.get("dry_run", False)
            self.outlook_actions = OutlookActions(dry_run=dry_run)

            self.action_executor = ActionExecutor(
                self.outlook_actions, self.template_engine
            )

            logger.info("所有引擎初始化成功")

        except Exception as e:
            logger.error(f"初始化引擎失败: {e}")
            raise

    def process_emails(self) -> Dict[str, Any]:
        """处理邮件"""
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
            emails = self.outlook_actions.get_inbox_emails(
                unread_only=unread_only,
                excluded_folders=excluded_folders,
                max_emails=max_emails,
            )

            logger.info(f"开始处理 {len(emails)} 封邮件...")

            for email_item in emails:
                try:
                    email_data = self.outlook_actions.get_email_data(email_item)

                    if not email_data:
                        logger.warning("无法提取邮件数据，跳过")
                        continue

                    stats["processed"] += 1

                    matched_rules = self.rule_engine.match_email(email_data)

                    if matched_rules:
                        stats["matched"] += 1
                        logger.info(
                            f"邮件 '{email_data.get('subject')}' 匹配了 {len(matched_rules)} 条规则"
                        )

                        for rule in matched_rules:
                            actions = rule.get("actions", [])
                            if actions:
                                result = self.action_executor.execute_actions(
                                    email_item, actions, email_data
                                )
                                stats["actions_executed"] += result["success"]

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
        """运行助手"""
        settings = self.config.get("settings", {})
        interval = settings.get("check_interval", 60)

        logger.info("Outlook邮件助手启动 (Windows版本)")
        logger.info(f"运行模式: {'单次' if once else '循环'}")
        logger.info(f"检查间隔: {interval}秒")
        logger.info(f"试运行模式: {self.dry_run or settings.get('dry_run', False)}")

        try:
            while True:
                stats = self.process_emails()

                if once:
                    logger.info("单次运行完成")
                    break

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
        description="Outlook邮件助手 - Windows单文件版本",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
    python outlook_assistant_win_standalone.py --once
    python outlook_assistant_win_standalone.py --dry-run
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

    if not os.path.exists(args.config):
        logger.error(f"配置文件不存在: {args.config}")
        print(f"错误: 配置文件不存在: {args.config}")
        sys.exit(1)

    try:
        assistant = OutlookAssistantWindows(args.config, dry_run=args.dry_run)
        assistant.run(once=args.once)
    except Exception as e:
        logger.error(f"程序异常退出: {e}")
        print(f"程序异常: {e}")
        print("请确保Microsoft Outlook已运行")
        sys.exit(1)


if __name__ == "__main__":
    main()

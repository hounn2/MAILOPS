"""
邮件操作模块
实现回复、转发、移动等操作
"""

import logging
import win32com.client
from typing import List, Dict, Any, Optional
from datetime import datetime

logger = logging.getLogger(__name__)


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
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            logger.info("成功连接到Outlook")
        except Exception as e:
            logger.error(f"连接Outlook失败: {e}")
            raise

    def reply_email(
        self, email_item, subject: str, body: str, include_original: bool = False
    ) -> bool:
        """
        回复邮件

        Args:
            email_item: Outlook邮件对象
            subject: 回复主题
            body: 回复内容
            include_original: 是否包含原始邮件内容

        Returns:
            是否成功
        """
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
        """
        转发邮件

        Args:
            email_item: Outlook邮件对象
            to: 收件人列表
            subject_prefix: 主题前缀
            additional_body: 附加内容

        Returns:
            是否成功
        """
        try:
            if self.dry_run:
                logger.info(f"[DRY RUN] 将转发邮件: {email_item.Subject}")
                logger.info(f"[DRY RUN] 转发到: {', '.join(to)}")
                return True

            forward = email_item.Forward()

            # 设置收件人
            recipients_str = "; ".join(to)
            forward.To = recipients_str

            # 修改主题
            if subject_prefix:
                forward.Subject = subject_prefix + email_item.Subject

            # 添加附加内容
            if additional_body:
                forward.Body = additional_body + "\n\n" + forward.Body

            forward.Send()
            logger.info(f"成功转发邮件 '{email_item.Subject}' 到: {recipients_str}")
            return True

        except Exception as e:
            logger.error(f"转发邮件失败: {e}")
            return False

    def move_email(self, email_item, target_folder: str) -> bool:
        """
        移动邮件到指定文件夹

        Args:
            email_item: Outlook邮件对象
            target_folder: 目标文件夹名称

        Returns:
            是否成功
        """
        try:
            if self.dry_run:
                logger.info(
                    f"[DRY RUN] 将移动邮件 '{email_item.Subject}' 到文件夹: {target_folder}"
                )
                return True

            # 获取邮件所在账户的收件箱
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
            # 首先在收件箱下查找
            inbox = store.GetDefaultFolder(6)  # 6 = olFolderInbox
            folder = self._search_folder_recursive(inbox, folder_name)
            if folder:
                return folder

            # 然后在整个存储中查找
            root_folder = store.GetRootFolder()
            folder = self._search_folder_recursive(root_folder, folder_name)

            return folder

        except Exception as e:
            logger.error(f"查找文件夹失败: {e}")
            return None

    def _search_folder_recursive(self, parent_folder, target_name: str):
        """递归搜索文件夹"""
        try:
            # 检查当前文件夹
            if parent_folder.Name == target_name:
                return parent_folder

            # 递归搜索子文件夹
            for folder in parent_folder.Folders:
                result = self._search_folder_recursive(folder, target_name)
                if result:
                    return result

            return None

        except Exception as e:
            logger.error(f"搜索文件夹出错: {e}")
            return None

    def mark_as_read(self, email_item) -> bool:
        """
        标记邮件为已读

        Args:
            email_item: Outlook邮件对象

        Returns:
            是否成功
        """
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
        """
        从Outlook邮件对象提取数据

        Args:
            email_item: Outlook邮件对象

        Returns:
            邮件数据字典
        """
        try:
            # 获取发件人
            sender = ""
            if email_item.Sender:
                sender = email_item.Sender.EmailAddress

            # 获取收件人
            recipients = []
            for recipient in email_item.Recipients:
                recipients.append(recipient.Address)

            # 获取附件信息
            attachments = []
            for attachment in email_item.Attachments:
                attachments.append(
                    {"filename": attachment.FileName, "size": attachment.Size}
                )

            return {
                "entry_id": email_item.EntryID,
                "subject": email_item.Subject,
                "body": email_item.Body,
                "html_body": getattr(email_item, "HTMLBody", ""),
                "sender": sender,
                "recipients": recipients,
                "received_time": email_item.ReceivedTime,
                "sent_time": email_item.SentOn,
                "is_read": not email_item.UnRead,
                "importance": email_item.Importance,
                "attachments": attachments,
                "conversation_id": getattr(email_item, "ConversationID", ""),
                "message_class": email_item.MessageClass,
            }

        except Exception as e:
            logger.error(f"提取邮件数据失败: {e}")
            return {}

    def get_inbox_emails(
        self,
        unread_only: bool = True,
        excluded_folders=None,
        max_emails: int = 50,
    ):
        """
        获取收件箱中的邮件

        Args:
            unread_only: 是否只获取未读邮件
            excluded_folders: 排除的文件夹列表
            max_emails: 最大邮件数量

        Returns:
            邮件对象列表
        """
        if excluded_folders is None:
            excluded_folders = []

        emails = []

        try:
            # 检查是否已连接
            if not self.namespace:
                logger.error("未连接到Outlook")
                return []

            # 遍历所有账户
            for store in self.namespace.Stores:
                try:
                    inbox = store.GetDefaultFolder(6)  # 6 = olFolderInbox

                    # 检查是否在排除列表中
                    if inbox.Name in excluded_folders:
                        continue

                    # 获取邮件
                    items = inbox.Items
                    items.Sort("[ReceivedTime]", True)  # 按接收时间降序

                    for item in items:
                        if len(emails) >= max_emails:
                            break

                        # 检查是否为邮件项
                        if item.MessageClass not in ["IPM.Note", "IPM.Post"]:
                            continue

                        # 如果只获取未读邮件
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
        """
        执行一系列动作

        Args:
            email_item: Outlook邮件对象
            actions: 动作列表
            email_data: 邮件数据

        Returns:
            执行结果统计
        """
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
                "return_date": "2026-01-01",  # 可从配置获取
                "backup_contact": "backup@company.com",  # 可从配置获取
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

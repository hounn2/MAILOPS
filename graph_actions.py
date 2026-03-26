"""
Microsoft Graph API邮件操作模块
使用REST API替代Windows COM接口，支持跨平台
"""

import re
import base64
import logging
import requests
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from urllib.parse import quote
from graph_auth import GraphAuth

logger = logging.getLogger(__name__)


class GraphOutlookActions:
    """基于Microsoft Graph API的Outlook操作类"""

    GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

    def __init__(self, auth: GraphAuth, dry_run: bool = False):
        """
        初始化

        Args:
            auth: GraphAuth认证实例
            dry_run: 试运行模式
        """
        self.auth = auth
        self.dry_run = dry_run
        self.session = requests.Session()

    def _make_request(
        self, method: str, endpoint: str, data: Dict = None, params: Dict = None
    ) -> Optional[Dict]:
        """
        发送HTTP请求

        Args:
            method: HTTP方法
            endpoint: API端点
            data: 请求体数据
            params: URL参数

        Returns:
            响应数据或None
        """
        try:
            headers = self.auth.get_headers()
            url = f"{self.GRAPH_API_BASE}{endpoint}"

            if method.upper() == "GET":
                response = self.session.get(url, headers=headers, params=params)
            elif method.upper() == "POST":
                response = self.session.post(
                    url, headers=headers, json=data, params=params
                )
            elif method.upper() == "PATCH":
                response = self.session.patch(url, headers=headers, json=data)
            elif method.upper() == "DELETE":
                response = self.session.delete(url, headers=headers)
            else:
                logger.error(f"不支持的HTTP方法: {method}")
                return None

            if response.status_code in [200, 201, 202, 204]:
                if response.status_code == 204:
                    return {}
                return response.json()
            else:
                logger.error(f"请求失败: {response.status_code} - {response.text}")
                return None

        except Exception as e:
            logger.error(f"请求出错: {e}")
            return None

    def get_inbox_emails(
        self,
        unread_only: bool = True,
        excluded_folders: list = None,
        max_emails: int = 50,
    ) -> List[Dict[str, Any]]:
        """
        获取收件箱邮件

        Args:
            unread_only: 只获取未读邮件
            excluded_folders: 排除的文件夹
            max_emails: 最大数量

        Returns:
            邮件列表
        """
        if excluded_folders is None:
            excluded_folders = []

        try:
            # 构建过滤条件
            filter_conditions = []
            if unread_only:
                filter_conditions.append("isRead eq false")

            params = {
                "$top": max_emails,
                "$orderby": "receivedDateTime desc",
                "$select": "id,subject,bodyPreview,from,toRecipients,receivedDateTime,isRead,hasAttachments,conversationId,importance",
            }

            if filter_conditions:
                params["$filter"] = " and ".join(filter_conditions)

            # 获取收件箱邮件
            endpoint = "/me/mailFolders/inbox/messages"
            response = self._make_request("GET", endpoint, params=params)

            if response and "value" in response:
                emails = response["value"]
                logger.info(f"获取到 {len(emails)} 封邮件")
                return emails
            else:
                logger.warning("未获取到邮件")
                return []

        except Exception as e:
            logger.error(f"获取邮件失败: {e}")
            return []

    def get_email_data(self, email_item: Dict[str, Any]) -> Dict[str, Any]:
        """
        转换Graph API邮件格式为内部格式

        Args:
            email_item: Graph API邮件对象

        Returns:
            内部邮件数据格式
        """
        try:
            # 获取完整邮件内容
            email_id = email_item.get("id")
            if email_id:
                endpoint = f"/me/messages/{email_id}"
                full_email = self._make_request("GET", endpoint)
                if full_email:
                    email_item = full_email

            # 解析发件人
            sender = ""
            from_data = email_item.get("from", {})
            if from_data and "emailAddress" in from_data:
                sender = from_data["emailAddress"].get("address", "")

            # 解析收件人
            recipients = []
            to_recipients = email_item.get("toRecipients", [])
            for recipient in to_recipients:
                if "emailAddress" in recipient:
                    recipients.append(recipient["emailAddress"].get("address", ""))

            # 解析附件
            attachments = []
            if email_item.get("hasAttachments"):
                email_id = email_item.get("id")
                attachments = self._get_attachments(email_id)

            # 解析接收时间
            received_time = None
            received_str = email_item.get("receivedDateTime")
            if received_str:
                try:
                    received_time = datetime.fromisoformat(
                        received_str.replace("Z", "+00:00")
                    )
                except:
                    pass

            # 解析正文
            body_content = ""
            body_data = email_item.get("body", {})
            if body_data:
                body_content = body_data.get("content", "")

            return {
                "entry_id": email_item.get("id", ""),
                "subject": email_item.get("subject", ""),
                "body": body_content,
                "body_preview": email_item.get("bodyPreview", ""),
                "sender": sender,
                "recipients": recipients,
                "received_time": received_time,
                "is_read": email_item.get("isRead", False),
                "importance": email_item.get("importance", "normal"),
                "has_attachments": email_item.get("hasAttachments", False),
                "conversation_id": email_item.get("conversationId", ""),
                "internet_message_id": email_item.get("internetMessageId", ""),
            }

        except Exception as e:
            logger.error(f"转换邮件数据失败: {e}")
            return {}

    def _get_attachments(self, email_id: str) -> List[Dict[str, Any]]:
        """获取邮件附件列表"""
        try:
            endpoint = f"/me/messages/{email_id}/attachments"
            response = self._make_request("GET", endpoint)

            if response and "value" in response:
                attachments = []
                for att in response["value"]:
                    attachments.append(
                        {
                            "id": att.get("id"),
                            "filename": att.get("name", ""),
                            "content_type": att.get("contentType", ""),
                            "size": att.get("size", 0),
                        }
                    )
                return attachments
            return []

        except Exception as e:
            logger.error(f"获取附件失败: {e}")
            return []

    def reply_email(
        self,
        email_item: Dict[str, Any],
        subject: str,
        body: str,
        include_original: bool = False,
    ) -> bool:
        """
        回复邮件

        Args:
            email_item: 邮件数据
            subject: 回复主题
            body: 回复内容
            include_original: 是否包含原始邮件

        Returns:
            是否成功
        """
        try:
            email_id = email_item.get("entry_id") or email_item.get("id")
            if not email_id:
                logger.error("邮件ID为空")
                return False

            if self.dry_run:
                logger.info(
                    f"[DRY RUN] 将回复邮件: {email_item.get('subject', 'Unknown')}"
                )
                logger.info(f"[DRY RUN] 回复主题: {subject}")
                logger.info(f"[DRY RUN] 回复内容: {body[:100]}...")
                return True

            # 构建回复内容
            if include_original:
                original_body = email_item.get("body", "")
                full_body = f"{body}\n\n--- 原始邮件 ---\n{original_body}"
            else:
                full_body = body

            # 发送回复
            endpoint = f"/me/messages/{email_id}/reply"
            data = {"comment": full_body}

            # Graph API的reply不直接支持修改主题，我们使用createReply然后发送
            endpoint = f"/me/messages/{email_id}/createReply"
            reply_draft = self._make_request("POST", endpoint)

            if reply_draft:
                # 更新回复内容
                reply_id = reply_draft.get("id")
                update_endpoint = f"/me/messages/{reply_id}"
                update_data = {"body": {"contentType": "text", "content": full_body}}

                if subject:
                    update_data["subject"] = subject

                self._make_request("PATCH", update_endpoint, update_data)

                # 发送邮件
                send_endpoint = f"/me/messages/{reply_id}/send"
                self._make_request("POST", send_endpoint)

                logger.info(f"成功回复邮件: {email_item.get('subject', 'Unknown')}")
                return True
            else:
                logger.error("创建回复草稿失败")
                return False

        except Exception as e:
            logger.error(f"回复邮件失败: {e}")
            return False

    def forward_email(
        self,
        email_item: Dict[str, Any],
        to: List[str],
        subject_prefix: str = "",
        additional_body: str = "",
    ) -> bool:
        """
        转发邮件

        Args:
            email_item: 邮件数据
            to: 收件人列表
            subject_prefix: 主题前缀
            additional_body: 附加内容

        Returns:
            是否成功
        """
        try:
            email_id = email_item.get("entry_id") or email_item.get("id")
            if not email_id:
                logger.error("邮件ID为空")
                return False

            if not to:
                logger.error("转发收件人为空")
                return False

            if self.dry_run:
                logger.info(
                    f"[DRY RUN] 将转发邮件: {email_item.get('subject', 'Unknown')}"
                )
                logger.info(f"[DRY RUN] 转发到: {', '.join(to)}")
                return True

            # 构建收件人列表
            recipients = [{"emailAddress": {"address": addr}} for addr in to]

            # 创建转发草稿
            endpoint = f"/me/messages/{email_id}/createForward"
            forward_draft = self._make_request("POST", endpoint)

            if forward_draft:
                forward_id = forward_draft.get("id")

                # 更新收件人和内容
                update_endpoint = f"/me/messages/{forward_id}"
                update_data = {"toRecipients": recipients}

                # 修改主题
                if subject_prefix:
                    original_subject = forward_draft.get("subject", "")
                    update_data["subject"] = subject_prefix + original_subject

                # 添加附加内容
                if additional_body:
                    current_body = forward_draft.get("body", {})
                    new_content = (
                        additional_body + "\n\n" + current_body.get("content", "")
                    )
                    update_data["body"] = {
                        "contentType": "text",
                        "content": new_content,
                    }

                self._make_request("PATCH", update_endpoint, update_data)

                # 发送邮件
                send_endpoint = f"/me/messages/{forward_id}/send"
                self._make_request("POST", send_endpoint)

                logger.info(
                    f"成功转发邮件 '{email_item.get('subject', 'Unknown')}' 到: {', '.join(to)}"
                )
                return True
            else:
                logger.error("创建转发草稿失败")
                return False

        except Exception as e:
            logger.error(f"转发邮件失败: {e}")
            return False

    def move_email(self, email_item: Dict[str, Any], target_folder: str) -> bool:
        """
        移动邮件到指定文件夹

        Args:
            email_item: 邮件数据
            target_folder: 目标文件夹名称

        Returns:
            是否成功
        """
        try:
            email_id = email_item.get("entry_id") or email_item.get("id")
            if not email_id:
                logger.error("邮件ID为空")
                return False

            if self.dry_run:
                logger.info(
                    f"[DRY RUN] 将移动邮件 '{email_item.get('subject', 'Unknown')}' 到文件夹: {target_folder}"
                )
                return True

            # 查找目标文件夹
            folder_id = self._find_folder_id(target_folder)

            if not folder_id:
                # 如果文件夹不存在，尝试创建
                folder_id = self._create_folder(target_folder)
                if not folder_id:
                    logger.error(f"目标文件夹未找到且无法创建: {target_folder}")
                    return False

            # 移动邮件
            endpoint = f"/me/messages/{email_id}/move"
            data = {"destinationId": folder_id}

            result = self._make_request("POST", endpoint, data)

            if result:
                logger.info(
                    f"成功移动邮件 '{email_item.get('subject', 'Unknown')}' 到文件夹: {target_folder}"
                )
                return True
            else:
                logger.error("移动邮件失败")
                return False

        except Exception as e:
            logger.error(f"移动邮件失败: {e}")
            return False

    def _find_folder_id(self, folder_name: str) -> Optional[str]:
        """根据名称查找文件夹ID"""
        try:
            # 获取所有邮件文件夹
            endpoint = "/me/mailFolders"
            params = {"$select": "id,displayName", "$top": 100}

            response = self._make_request("GET", endpoint, params=params)

            if response and "value" in response:
                for folder in response["value"]:
                    if folder.get("displayName", "").lower() == folder_name.lower():
                        return folder.get("id")

                # 递归搜索子文件夹
                for folder in response["value"]:
                    folder_id = folder.get("id")
                    child_id = self._find_child_folder(folder_id, folder_name)
                    if child_id:
                        return child_id

            return None

        except Exception as e:
            logger.error(f"查找文件夹失败: {e}")
            return None

    def _find_child_folder(self, parent_id: str, folder_name: str) -> Optional[str]:
        """递归查找子文件夹"""
        try:
            endpoint = f"/me/mailFolders/{parent_id}/childFolders"
            params = {"$select": "id,displayName", "$top": 100}

            response = self._make_request("GET", endpoint, params=params)

            if response and "value" in response:
                for folder in response["value"]:
                    if folder.get("displayName", "").lower() == folder_name.lower():
                        return folder.get("id")

                    # 递归搜索
                    child_id = self._find_child_folder(folder.get("id"), folder_name)
                    if child_id:
                        return child_id

            return None

        except Exception as e:
            logger.error(f"查找子文件夹失败: {e}")
            return None

    def _create_folder(
        self, folder_name: str, parent_id: str = "inbox"
    ) -> Optional[str]:
        """创建新文件夹"""
        try:
            # 默认在收件箱下创建
            endpoint = f"/me/mailFolders/{parent_id}/childFolders"
            data = {"displayName": folder_name}

            response = self._make_request("POST", endpoint, data)

            if response and "id" in response:
                logger.info(f"成功创建文件夹: {folder_name}")
                return response["id"]

            return None

        except Exception as e:
            logger.error(f"创建文件夹失败: {e}")
            return None

    def mark_as_read(self, email_item: Dict[str, Any]) -> bool:
        """
        标记邮件为已读

        Args:
            email_item: 邮件数据

        Returns:
            是否成功
        """
        try:
            email_id = email_item.get("entry_id") or email_item.get("id")
            if not email_id:
                logger.error("邮件ID为空")
                return False

            if self.dry_run:
                logger.info(
                    f"[DRY RUN] 将标记邮件为已读: {email_item.get('subject', 'Unknown')}"
                )
                return True

            endpoint = f"/me/messages/{email_id}"
            data = {"isRead": True}

            result = self._make_request("PATCH", endpoint, data)

            if result:
                logger.info(f"标记邮件为已读: {email_item.get('subject', 'Unknown')}")
                return True
            else:
                logger.error("标记已读失败")
                return False

        except Exception as e:
            logger.error(f"标记已读失败: {e}")
            return False


class GraphActionExecutor:
    """Graph API动作执行器"""

    def __init__(self, outlook_actions: GraphOutlookActions, template_engine):
        self.outlook = outlook_actions
        self.template_engine = template_engine
        self.action_handlers = {
            "reply": self._handle_reply,
            "forward": self._handle_forward,
            "move": self._handle_move,
            "mark_as_read": self._handle_mark_as_read,
        }

    def execute_actions(
        self,
        email_item: Dict[str, Any],
        actions: List[Dict[str, Any]],
        email_data: Dict[str, Any],
    ) -> Dict[str, Any]:
        """
        执行一系列动作

        Args:
            email_item: 原始邮件数据（Graph API格式）
            actions: 动作列表
            email_data: 转换后的邮件数据

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
        self,
        email_item: Dict[str, Any],
        action: Dict[str, Any],
        email_data: Dict[str, Any],
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
        self,
        email_item: Dict[str, Any],
        action: Dict[str, Any],
        email_data: Dict[str, Any],
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
        self,
        email_item: Dict[str, Any],
        action: Dict[str, Any],
        email_data: Dict[str, Any],
    ) -> bool:
        """处理移动动作"""
        target_folder = action.get("target")

        if not target_folder:
            logger.error("移动动作缺少目标文件夹")
            return False

        return self.outlook.move_email(email_item, target_folder)

    def _handle_mark_as_read(
        self,
        email_item: Dict[str, Any],
        action: Dict[str, Any],
        email_data: Dict[str, Any],
    ) -> bool:
        """处理标记已读动作"""
        return self.outlook.mark_as_read(email_item)

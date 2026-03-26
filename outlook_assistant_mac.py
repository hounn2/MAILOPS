#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook邮件助手 - macOS独立版本
使用Microsoft Graph API，无需Windows Outlook客户端

使用方法:
    1. 首次运行: python outlook_assistant_mac.py --setup
    2. 正常运行: python outlook_assistant_mac.py

环境要求:
    - macOS 10.14+
    - Python 3.7+
    - Microsoft 365账户
"""

import os
import sys
import json
import time
import re
import logging
import argparse
from datetime import datetime
from typing import List, Dict, Any, Optional

# 确保不在Windows上运行
if sys.platform == 'win32':
    print("错误: 此版本仅支持macOS和Linux系统")
    print("如需Windows版本，请使用 outlook_assistant_win_standalone.py")
    sys.exit(1)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('outlook_assistant.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

CONFIG_FILE = 'config.json'
TOKEN_CACHE = 'token_cache.json'

# ============================================================================
# Microsoft Graph API 认证
# ============================================================================

class GraphAuth:
    """Microsoft Graph API认证管理器"""
    
    AUTHORITY = "https://login.microsoftonline.com/common"
    SCOPE = ["Mail.Read", "Mail.ReadWrite", "Mail.Send", "User.Read"]
    
    def __init__(self, client_id: str, token_cache_path: str = TOKEN_CACHE):
        self.client_id = client_id
        self.token_cache_path = token_cache_path
        self.access_token = None
        self.app = None
        
    def _load_or_create_cache(self):
        """加载或创建token缓存"""
        try:
            from msal import SerializableTokenCache
            cache = SerializableTokenCache()
            
            if os.path.exists(self.token_cache_path):
                with open(self.token_cache_path, "r") as f:
                    cache.deserialize(f.read())
            
            return cache
        except ImportError:
            return None
    
    def _save_cache(self, cache):
        """保存token缓存"""
        if cache and hasattr(cache, 'has_state_changed') and cache.has_state_changed:
            try:
                with open(self.token_cache_path, "w") as f:
                    f.write(cache.serialize())
            except Exception as e:
                logger.error(f"保存token缓存失败: {e}")
    
    def authenticate(self) -> bool:
        """执行认证流程"""
        try:
            import msal
            
            # 加载或创建token缓存
            cache = self._load_or_create_cache()
            
            self.app = msal.PublicClientApplication(
                client_id=self.client_id,
                authority=self.AUTHORITY,
                token_cache=cache
            )
            
            # 尝试获取已有账户
            accounts = self.app.get_accounts()
            if accounts:
                # 尝试静默获取token
                result = self.app.acquire_token_silent(self.SCOPE, account=accounts[0])
                if result:
                    self.access_token = result["access_token"]
                    logger.info("从缓存成功获取token")
                    self._save_cache(cache)
                    return True
            
            # 使用设备代码流
            return self._device_code_flow(cache)
            
        except ImportError:
            logger.error("缺少msal库，请运行: pip install msal")
            return False
        except Exception as e:
            logger.error(f"认证失败: {e}")
            return False
    
    def _device_code_flow(self, cache) -> bool:
        """使用设备代码流进行认证"""
        try:
            # 获取设备代码
            flow = self.app.initiate_device_flow(scopes=self.SCOPE)
            
            if "user_code" not in flow:
                logger.error("无法创建设备代码流")
                return False
            
            # 显示给用户
            print("\n" + "=" * 60)
            print("需要进行Microsoft账户认证")
            print("=" * 60)
            print(f"1. 在浏览器中打开: {flow['verification_uri']}")
            print(f"2. 输入代码: {flow['user_code']}")
            print("=" * 60 + "\n")
            
            # 等待用户完成认证
            result = self.app.acquire_token_by_device_flow(flow)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("设备代码流认证成功")
                self._save_cache(cache)
                return True
            else:
                error = result.get("error_description", "未知错误")
                logger.error(f"认证失败: {error}")
                return False
                
        except Exception as e:
            logger.error(f"设备代码流认证失败: {e}")
            return False
    
    def get_headers(self) -> Dict[str, str]:
        """获取HTTP请求头"""
        if not self.access_token:
            if not self.authenticate():
                raise Exception("无法获取有效的access token")
        
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }


# ============================================================================
# 规则引擎（与Windows版本相同）
# ============================================================================

class RuleEngine:
    """规则引擎主类"""
    
    def __init__(self, rules: List[Dict[str, Any]]):
        self.rules = rules
        self._compile_rules()
    
    def _compile_rules(self):
        """编译规则"""
        self.compiled_rules = []
        for rule in self.rules:
            if not rule.get('enabled', True):
                continue
            compiled_rule = {
                'id': rule['id'],
                'name': rule['name'],
                'conditions': self._compile_conditions(rule['conditions']),
                'actions': rule.get('actions', [])
            }
            self.compiled_rules.append(compiled_rule)
    
    def _compile_conditions(self, conditions: Dict[str, Any]) -> Dict[str, Any]:
        """编译条件配置"""
        return {
            'match_all': conditions.get('match_all', True),
            'items': [self._compile_condition_item(item) for item in conditions.get('items', [])]
        }
    
    def _compile_condition_item(self, item: Dict[str, Any]) -> Dict[str, Any]:
        """编译单个条件项"""
        field = item['field']
        operator = item['operator']
        value = item['value']
        
        if operator in ['in', 'not_in'] and isinstance(value, list):
            value = set(value)
        
        if operator == 'contains' and isinstance(value, list):
            patterns = [re.compile(re.escape(v), re.IGNORECASE) for v in value]
            value = patterns
        
        return {'field': field, 'operator': operator, 'value': value}
    
    def match_email(self, email_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """匹配邮件数据"""
        matched_rules = []
        
        for rule in self.compiled_rules:
            if self._evaluate_conditions(rule['conditions'], email_data):
                matched_rules.append(rule)
                logger.info(f"邮件 '{email_data.get('subject', 'Unknown')}' 匹配规则: {rule['name']}")
        
        return matched_rules
    
    def _evaluate_conditions(self, conditions: Dict[str, Any], email_data: Dict[str, Any]) -> bool:
        """评估条件是否满足"""
        match_all = conditions['match_all']
        items = conditions['items']
        
        if not items:
            return True
        
        results = [self._evaluate_condition_item(item, email_data) for item in items]
        
        if match_all:
            return all(results)
        else:
            return any(results)
    
    def _evaluate_condition_item(self, item: Dict[str, Any], email_data: Dict[str, Any]) -> bool:
        """评估单个条件项"""
        field = item['field']
        operator = item['operator']
        value = item['value']
        
        email_value = self._get_email_field(email_data, field)
        
        if email_value is None:
            return False
        
        return self._apply_operator(email_value, operator, value)
    
    def _get_email_field(self, email_data: Dict[str, Any], field: str) -> Any:
        """获取邮件字段值"""
        if field == 'sender_domain':
            sender = email_data.get('sender', '')
            if '@' in sender:
                return sender[sender.index('@'):]
            return ''
        elif field == 'body_length':
            body = email_data.get('body', '')
            return len(body) if body else 0
        elif field == 'has_attachments':
            attachments = email_data.get('attachments', [])
            return len(attachments) > 0
        else:
            return email_data.get(field)
    
    def _apply_operator(self, email_value: Any, operator: str, rule_value: Any) -> bool:
        """应用操作符进行比较"""
        try:
            if operator == 'equals':
                return str(email_value).lower() == str(rule_value).lower()
            elif operator == 'not_equals':
                return str(email_value).lower() != str(rule_value).lower()
            elif operator == 'contains':
                email_str = str(email_value).lower()
                if isinstance(rule_value, list):
                    return any(pattern.search(email_str) for pattern in rule_value)
                return str(rule_value).lower() in email_str
            elif operator == 'not_contains':
                email_str = str(email_value).lower()
                if isinstance(rule_value, list):
                    return not any(pattern.search(email_str) for pattern in rule_value)
                return str(rule_value).lower() not in email_str
            elif operator == 'starts_with':
                return str(email_value).lower().startswith(str(rule_value).lower())
            elif operator == 'ends_with':
                return str(email_value).lower().endswith(str(rule_value).lower())
            elif operator == 'in':
                return str(email_value).lower() in {str(v).lower() for v in rule_value}
            elif operator == 'not_in':
                return str(email_value).lower() not in {str(v).lower() for v in rule_value}
            elif operator == 'regex':
                if isinstance(rule_value, str):
                    pattern = re.compile(rule_value, re.IGNORECASE)
                    return bool(pattern.search(str(email_value)))
                return False
            elif operator == 'greater_than':
                return float(email_value) > float(rule_value)
            elif operator == 'less_than':
                return float(email_value) < float(rule_value)
            else:
                logger.warning(f"未知操作符: {operator}")
                return False
        except Exception as e:
            logger.error(f"评估操作符时出错: {operator}, error: {e}")
            return False


class TemplateEngine:
    """模板引擎"""
    
    def __init__(self, templates: Dict[str, Any]):
        self.templates = templates
    
    def render(self, template_name: str, context: Dict[str, Any]) -> Dict[str, str]:
        """渲染模板"""
        template = self.templates.get(template_name)
        if not template:
            logger.error(f"模板未找到: {template_name}")
            return {'subject': '', 'body': ''}
        
        subject = self._render_template(template.get('subject', ''), context)
        body = self._render_template(template.get('body', ''), context)
        
        return {'subject': subject, 'body': body}
    
    def _render_template(self, template_str: str, context: Dict[str, Any]) -> str:
        """使用简单字符串替换渲染模板"""
        result = template_str
        for key, value in context.items():
            placeholder = f'{{{key}}}'
            result = result.replace(placeholder, str(value))
        return result


# ============================================================================
# Microsoft Graph API 邮件操作
# ============================================================================

import requests

class GraphOutlookActions:
    """基于Microsoft Graph API的Outlook操作类"""
    
    GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
    
    def __init__(self, auth: GraphAuth, dry_run: bool = False):
        self.auth = auth
        self.dry_run = dry_run
        self.session = requests.Session()
    
    def _make_request(self, method: str, endpoint: str, 
                     data: Dict = None, params: Dict = None) -> Optional[Dict]:
        """发送HTTP请求"""
        try:
            headers = self.auth.get_headers()
            url = f"{self.GRAPH_API_BASE}{endpoint}"
            
            if method.upper() == "GET":
                response = self.session.get(url, headers=headers, params=params)
            elif method.upper() == "POST":
                response = self.session.post(url, headers=headers, json=data, params=params)
            elif method.upper() == "PATCH":
                response = self.session.patch(url, headers=headers, json=data)
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
    
    def get_inbox_emails(self, unread_only: bool = True, 
                        excluded_folders: list = None,
                        max_emails: int = 50) -> List[Dict[str, Any]]:
        """获取收件箱邮件"""
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
                "$select": "id,subject,bodyPreview,from,toRecipients,receivedDateTime,isRead,hasAttachments,conversationId,importance"
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
        """转换Graph API邮件格式为内部格式"""
        try:
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
                    received_time = datetime.fromisoformat(received_str.replace("Z", "+00:00"))
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
                "internet_message_id": email_item.get("internetMessageId", "")
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
                    attachments.append({
                        "id": att.get("id"),
                        "filename": att.get("name", ""),
                        "content_type": att.get("contentType", ""),
                        "size": att.get("size", 0)
                    })
                return attachments
            return []
            
        except Exception as e:
            logger.error(f"获取附件失败: {e}")
            return []
    
    def reply_email(self, email_item: Dict[str, Any], subject: str, 
                   body: str, include_original: bool = False) -> bool:
        """回复邮件"""
        try:
            email_id = email_item.get("entry_id") or email_item.get("id")
            if not email_id:
                logger.error("邮件ID为空")
                return False
            
            if self.dry_run:
                logger.info(f"[DRY RUN] 将回复邮件: {email_item.get('subject', 'Unknown')}")
                return True
            
            # 构建回复内容
            if include_original:
                original_body = email_item.get("body", "")
                full_body = f"{body}\n\n--- 原始邮件 ---\n{original_body}"
            else:
                full_body = body
            
            # 创建回复草稿
            endpoint = f"/me/messages/{email_id}/createReply"
            reply_draft = self._make_request("POST", endpoint)
            
            if reply_draft:
                # 更新回复内容
                reply_id = reply_draft.get("id")
                update_endpoint = f"/me/messages/{reply_id}"
                update_data = {
                    "body": {
                        "contentType": "text",
                        "content": full_body
                    }
                }
                
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
    
    def forward_email(self, email_item: Dict[str, Any], to: List[str],
                     subject_prefix: str = "", additional_body: str = "") -> bool:
        """转发邮件"""
        try:
            email_id = email_item.get("entry_id") or email_item.get("id")
            if not email_id:
                logger.error("邮件ID为空")
                return False
            
            if not to:
                logger.error("转发收件人为空")
                return False
            
            if self.dry_run:
                logger.info(f"[DRY RUN] 将转发邮件: {email_item.get('subject', 'Unknown')}")
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
                update_data = {
                    "toRecipients": recipients
                }
                
                # 修改主题
                if subject_prefix:
                    original_subject = forward_draft.get("subject", "")
                    update_data["subject"] = subject_prefix + original_subject
                
                # 添加附加内容
                if additional_body:
                    current_body = forward_draft.get("body", {})
                    new_content = additional_body + "\n\n" + current_body.get("content", "")
                    update_data["body"] = {
                        "contentType": "text",
                        "content": new_content
                    }
                
                self._make_request("PATCH", update_endpoint, update_data)
                
                # 发送邮件
                send_endpoint = f"/me/messages/{forward_id}/send"
                self._make_request("POST", send_endpoint)
                
                logger.info(f"成功转发邮件 '{email_item.get('subject', 'Unknown')}' 到: {', '.join(to)}")
                return True
            else:
                logger.error("创建转发草稿失败")
                return False
                
        except Exception as e:
            logger.error(f"转发邮件失败: {e}")
            return False
    
    def move_email(self, email_item: Dict[str, Any], target_folder: str) -> bool:
        """移动邮件到指定文件夹"""
        try:
            email_id = email_item.get("entry_id") or email_item.get("id")
            if not email_id:
                logger.error("邮件ID为空")
                return False
            
            if self.dry_run:
                logger.info(f"[DRY RUN] 将移动邮件 '{email_item.get('subject', 'Unknown')}' 到文件夹: {target_folder}")
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
            data = {
                "destinationId": folder_id
            }
            
            result = self._make_request("POST", endpoint, data)
            
            if result:
                logger.info(f"成功移动邮件 '{email_item.get('subject', 'Unknown')}' 到文件夹: {target_folder}")
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
            params = {
                "$select": "id,displayName",
                "$top": 100
            }
            
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
            params = {
                "$select": "id,displayName",
                "$top": 100
            }
            
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
    
    def _create_folder(self, folder_name: str, parent_id: str = "inbox") -> Optional[str]:
        """创建新文件夹"""
        try:
            # 默认在收件箱下创建
            endpoint = f"/me/mailFolders/{parent_id}/childFolders"
            data = {
                "displayName": folder_name
            }
            
            response = self._make_request("POST", endpoint, data)
            
            if response and "id" in response:
                logger.info(f"成功创建文件夹: {folder_name}")
                return response["id"]
            
            return None
            
        except Exception as e:
            logger.error(f"创建文件夹失败: {e}")
            return None
    
    def mark_as_read(self, email_item: Dict[str, Any]) -> bool:
        """标记邮件为已读"""
        try:
            email_id = email_item.get("entry_id") or email_item.get("id")
            if not email_id:
                logger.error("邮件ID为空")
                return False
            
            if self.dry_run:
                logger.info(f"[DRY RUN] 将标记邮件为已读: {email_item.get('subject', 'Unknown')}")
                return True
            
            endpoint = f"/me/messages/{email_id}"
            data = {
                "isRead": True
            }
            
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


# ============================================================================
# 主程序
# ============================================================================

def load_config():
    """加载配置"""
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config
    except Exception as e:
        logger.error(f"加载配置失败: {e}")
        return {
            'rules': [],
            'templates': {},
            'settings': {
                'check_interval': 60,
                'process_unread_only': True,
                'max_emails_per_batch': 50,
                'dry_run': False
            }
        }

def setup():
    """首次设置向导"""
    print("\n" + "=" * 60)
    print("Outlook邮件助手 - macOS版本 - 首次设置")
    print("=" * 60)
    print()
    
    # 检查Azure AD配置
    config = load_config()
    azure_config = config.get('azure_ad', {})
    
    if not azure_config.get('client_id'):
        print("⚠️  需要配置 Azure AD Client ID")
        print()
        print("请按以下步骤配置:")
        print("1. 访问 https://portal.azure.com")
        print("2. 注册新应用: Azure Active Directory → App registrations → New registration")
        print("3. 设置应用名称和重定向URI (http://localhost:8080)")
        print("4. 在API权限中添加: Microsoft Graph → Delegated permissions:")
        print("   - Mail.Read")
        print("   - Mail.ReadWrite")
        print("   - Mail.Send")
        print("   - User.Read")
        print("5. 复制 Application (client) ID")
        print()
        client_id = input("请输入 Azure AD Client ID: ").strip()
        
        if client_id:
            config['azure_ad'] = {
                'client_id': client_id,
                'tenant_id': 'common',
                'token_cache_path': TOKEN_CACHE
            }
            
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            print("✅ Client ID 已保存")
        else:
            print("❌ 未输入 Client ID，退出设置")
            return False
    
    print()
    print("🔐 接下来需要进行Microsoft账户认证...")
    print()
    
    # 进行认证
    auth = GraphAuth(azure_config.get('client_id'))
    if auth.authenticate():
        print()
        print("✅ 认证成功！")
        print()
        print("设置完成！现在可以运行: python outlook_assistant_mac.py")
        return True
    else:
        print()
        print("❌ 认证失败，请检查Client ID和网络连接")
        return False

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='Outlook邮件助手 - macOS版本')
    parser.add_argument('--setup', action='store_true', help='首次设置')
    parser.add_argument('--once', action='store_true', help='只运行一次')
    parser.add_argument('--dry-run', action='store_true', help='试运行模式')
    args = parser.parse_args()
    
    if args.setup:
        if setup():
            sys.exit(0)
        else:
            sys.exit(1)
    
    # 加载配置
    config = load_config()
    azure_config = config.get('azure_ad', {})
    
    if not azure_config.get('client_id'):
        print("❌ 未配置 Azure AD Client ID")
        print("请先运行: python outlook_assistant_mac.py --setup")
        sys.exit(1)
    
    # 初始化认证
    auth = GraphAuth(azure_config.get('client_id'))
    if not auth.authenticate():
        print("❌ 认证失败")
        sys.exit(1)
    
    # 初始化Outlook操作
    dry_run = args.dry_run or config.get('settings', {}).get('dry_run', False)
    outlook = GraphOutlookActions(auth, dry_run=dry_run)
    
    # 初始化规则引擎
    rule_engine = RuleEngine(config.get('rules', []))
    template_engine = TemplateEngine(config.get('templates', {}))
    
    # 获取设置
    settings = config.get('settings', {})
    interval = settings.get('check_interval', 60)
    unread_only = settings.get('process_unread_only', True)
    max_emails = settings.get('max_emails_per_batch', 50)
    
    print(f"\n{'=' * 60}")
    print("Outlook邮件助手 - macOS版本")
    print(f"{'=' * 60}")
    print(f"运行模式: {'单次' if args.once else '循环'}")
    print(f"试运行: {'是' if dry_run else '否'}")
    print(f"检查间隔: {interval}秒")
    print(f"{'=' * 60}\n")
    
    try:
        while True:
            # 获取邮件
            emails = outlook.get_inbox_emails(
                unread_only=unread_only,
                max_emails=max_emails
            )
            
            logger.info(f"开始处理 {len(emails)} 封邮件...")
            
            processed = 0
            matched = 0
            actions_executed = 0
            
            for email_item in emails:
                try:
                    email_data = outlook.get_email_data(email_item)
                    
                    if not email_data:
                        continue
                    
                    processed += 1
                    
                    # 匹配规则
                    matched_rules = rule_engine.match_email(email_data)
                    
                    if matched_rules:
                        matched += 1
                        logger.info(f"邮件 '{email_data.get('subject')}' 匹配了 {len(matched_rules)} 条规则")
                        
                        for rule in matched_rules:
                            actions = rule.get('actions', [])
                            for action in actions:
                                action_type = action.get('type')
                                
                                if action_type == 'reply':
                                    template_name = action.get('template')
                                    if template_name:
                                        context = {
                                            'original_subject': email_data.get('subject', ''),
                                            'sender': email_data.get('sender', ''),
                                            'received_time': str(email_data.get('received_time', ''))
                                        }
                                        template_result = template_engine.render(template_name, context)
                                        if outlook.reply_email(email_item, template_result['subject'], 
                                                              template_result['body'], 
                                                              action.get('include_original', False)):
                                            actions_executed += 1
                                
                                elif action_type == 'forward':
                                    to = action.get('to', [])
                                    if outlook.forward_email(email_item, to, 
                                                            action.get('subject_prefix', ''),
                                                            action.get('additional_body', '')):
                                        actions_executed += 1
                                
                                elif action_type == 'move':
                                    target = action.get('target')
                                    if outlook.move_email(email_item, target):
                                        actions_executed += 1
                                
                                elif action_type == 'mark_as_read':
                                    if outlook.mark_as_read(email_item):
                                        actions_executed += 1
                    
                except Exception as e:
                    logger.error(f"处理邮件时出错: {e}")
                    continue
            
            logger.info(f"处理完成: 处理 {processed} 封, 匹配 {matched} 封, 执行 {actions_executed} 个动作")
            
            if args.once:
                break
            
            logger.info(f"等待 {interval} 秒后再次检查...")
            time.sleep(interval)
            
    except KeyboardInterrupt:
        print("\n用户中断，程序退出")

if __name__ == '__main__':
    main()

"""
邮件规则引擎模块
负责解析和匹配邮件规则
"""

import re
import logging
from typing import List, Dict, Any, Callable
from datetime import datetime

logger = logging.getLogger(__name__)


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

        # 将值转换为集合以提高查找效率
        if operator in ["in", "not_in"] and isinstance(value, list):
            value = set(value)

        # 为contains操作符编译正则表达式
        if operator == "contains" and isinstance(value, list):
            patterns = [re.compile(re.escape(v), re.IGNORECASE) for v in value]
            value = patterns

        return {"field": field, "operator": operator, "value": value}

    def match_email(self, email_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        匹配邮件数据，返回匹配的规则列表

        Args:
            email_data: 邮件数据字典，包含subject, body, sender, etc.

        Returns:
            匹配的规则列表，包含对应的actions
        """
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

        # 获取邮件字段值
        email_value = self._get_email_field(email_data, field)

        if email_value is None:
            return False

        # 执行操作符比较
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
                    # 使用正则表达式匹配
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
        """
        渲染模板

        Args:
            template_name: 模板名称
            context: 上下文变量

        Returns:
            包含subject和body的字典
        """
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

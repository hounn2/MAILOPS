"""
AI引擎模块 - 支持本地LMStudio
调用本地大语言模型生成智能回复
"""

import json
import logging
import requests
from typing import Dict, Any, Optional, List
from datetime import datetime

logger = logging.getLogger(__name__)


class LMStudioEngine:
    """LMStudio AI引擎"""

    def __init__(
        self,
        base_url: str = "http://localhost:1234",
        model: str = None,
        timeout: int = 60,
    ):
        """
        初始化LMStudio引擎

        Args:
            base_url: LMStudio服务器地址
            model: 模型名称（None则使用默认模型）
            timeout: 请求超时时间（秒）
        """
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.timeout = timeout
        self.session = requests.Session()

        # 测试连接
        if not self._test_connection():
            logger.warning("无法连接到LMStudio服务器，请确保LMStudio已启动")

    def _test_connection(self) -> bool:
        """测试与LMStudio的连接"""
        try:
            response = self.session.get(f"{self.base_url}/v1/models", timeout=5)
            if response.status_code == 200:
                models = response.json()
                logger.info(
                    f"成功连接到LMStudio，可用模型: {len(models.get('data', []))}个"
                )
                return True
            return False
        except Exception as e:
            logger.error(f"连接LMStudio失败: {e}")
            return False

    def chat_completion(
        self,
        messages: List[Dict[str, str]],
        temperature: float = 0.7,
        max_tokens: int = 2000,
    ) -> Optional[str]:
        """
        调用聊天补全API

        Args:
            messages: 消息列表，格式 [{"role": "user", "content": "..."}]
            temperature: 温度参数（创造性）
            max_tokens: 最大生成token数

        Returns:
            AI生成的回复文本
        """
        try:
            payload = {
                "messages": messages,
                "temperature": temperature,
                "max_tokens": max_tokens,
                "stream": False,
            }

            if self.model:
                payload["model"] = self.model

            logger.info(f"正在调用LMStudio生成回复...")

            response = self.session.post(
                f"{self.base_url}/v1/chat/completions",
                json=payload,
                timeout=self.timeout,
            )

            if response.status_code == 200:
                result = response.json()
                if "choices" in result and len(result["choices"]) > 0:
                    content = result["choices"][0].get("message", {}).get("content", "")
                    logger.info("AI回复生成成功")
                    return content.strip()
                else:
                    logger.error("API返回格式异常")
                    return None
            else:
                logger.error(f"API调用失败: {response.status_code} - {response.text}")
                return None

        except requests.exceptions.Timeout:
            logger.error("AI请求超时，请检查LMStudio是否正常运行")
            return None
        except Exception as e:
            logger.error(f"AI调用异常: {e}")
            return None

    def generate_email_reply(
        self, email_content: str, knowledge_context: str = "", system_prompt: str = None
    ) -> Optional[str]:
        """
        生成邮件回复

        Args:
            email_content: 邮件内容
            knowledge_context: 知识库上下文
            system_prompt: 系统提示词

        Returns:
            生成的回复内容
        """
        if system_prompt is None:
            system_prompt = """你是一个专业的客服助手。请根据客户邮件内容和提供的知识库信息，生成礼貌、专业、准确的回复。

重要提示：
- 系统为你提供了多篇相关文档（标记为[文档1]、[文档2]等），请仔细阅读所有文档
- 综合所有文档的信息来回答，不要只看第一篇文档
- 如果不同文档中有互补信息，请整合在一起回答
- 如果文档中没有相关信息，请诚实说明"需要进一步确认"

要求：
1. 使用中文回复
2. 语气友好、专业
3. 基于知识库信息回答，不要编造信息
4. 适当使用换行和列表提高可读性
5. 结尾可以询问是否还有其他问题"""

        # 构建提示词
        user_prompt = f"客户邮件内容：\n{email_content}\n"

        if knowledge_context:
            user_prompt += f"\n\n参考知识库信息：\n{knowledge_context}\n"

        user_prompt += "\n\n请根据以上信息生成回复："

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]

        return self.chat_completion(messages)

    def is_available(self) -> bool:
        """检查AI服务是否可用"""
        return self._test_connection()


class AIReplyEngine:
    """AI回复引擎 - 结合知识库和AI生成回复"""

    def __init__(self, lmstudio_config: Dict[str, Any], knowledge_base=None):
        """
        初始化AI回复引擎

        Args:
            lmstudio_config: LMStudio配置
            knowledge_base: 知识库管理器实例
        """
        self.lmstudio = LMStudioEngine(
            base_url=lmstudio_config.get("base_url", "http://localhost:1234"),
            model=lmstudio_config.get("model"),
            timeout=lmstudio_config.get("timeout", 60),
        )
        self.knowledge_base = knowledge_base
        self.system_prompt = lmstudio_config.get("system_prompt")

    def generate_reply(
        self, email_data: Dict[str, Any], search_knowledge: bool = True
    ) -> Optional[str]:
        """
        生成邮件回复

        Args:
            email_data: 邮件数据
            search_knowledge: 是否搜索知识库

        Returns:
            生成的回复内容
        """
        # 构建邮件内容
        email_content = f"主题: {email_data.get('subject', '')}\n\n"
        email_content += f"正文:\n{email_data.get('body', '')}"

        # 获取知识库上下文
        knowledge_context = ""
        logger.info(
            f"知识库搜索: search_knowledge={search_knowledge}, knowledge_base={self.knowledge_base is not None}"
        )

        if search_knowledge and self.knowledge_base:
            # 从知识库搜索相关内容
            search_query = (
                email_data.get("subject", "") + " " + email_data.get("body", "")[:200]
            )
            logger.info(f"正在搜索知识库，查询: {search_query[:100]}...")

            relevant_docs = self.knowledge_base.search_relevant(search_query, top_k=3)

            if relevant_docs:
                # 构建知识库上下文，明确标记每个文档
                knowledge_parts = []
                for i, doc in enumerate(relevant_docs):
                    doc_header = (
                        f"=== 文档{i + 1} (相似度: {doc.get('score', 0):.1%}) ==="
                    )
                    doc_content = doc["content"][:1000]
                    knowledge_parts.append(f"{doc_header}\n{doc_content}")

                knowledge_context = "\n\n".join(knowledge_parts)

                logger.info(f"从知识库找到 {len(relevant_docs)} 篇相关文档")
                logger.info(
                    f"文档相似度: "
                    + ", ".join(
                        [
                            f"文档{i + 1}:{doc.get('score', 0):.1%}"
                            for i, doc in enumerate(relevant_docs)
                        ]
                    )
                )
                logger.debug(f"知识库内容预览:\n{knowledge_context[:800]}...")
            else:
                logger.warning("未从知识库找到相关文档")
        else:
            if not search_knowledge:
                logger.info("知识库搜索被禁用 (search_knowledge=False)")
            if not self.knowledge_base:
                logger.warning("知识库未初始化 (knowledge_base=None)")

        # 调用AI生成回复
        return self.lmstudio.generate_email_reply(
            email_content=email_content,
            knowledge_context=knowledge_context,
            system_prompt=self.system_prompt,
        )

    def is_ready(self) -> bool:
        """检查引擎是否就绪"""
        return self.lmstudio.is_available()


# 测试代码
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
    )

    # 测试配置
    config = {"base_url": "http://localhost:1234", "timeout": 60}

    engine = LMStudioEngine(**config)

    if engine.is_available():
        print("\nLMStudio连接成功！")
        print("=" * 60)

        # 测试生成回复
        test_email = """
        主题：产品咨询

        你好，我想了解一下你们的产品价格是多少？
        还有企业版包含哪些功能？
        """

        print("\n测试邮件内容：")
        print(test_email)
        print("\n" + "=" * 60)

        reply = engine.generate_email_reply(test_email)

        if reply:
            print("\nAI生成的回复：")
            print(reply)
        else:
            print("\n生成回复失败")
    else:
        print("\n无法连接到LMStudio，请确保：")
        print("1. LMStudio已启动")
        print("2. 服务器地址正确（默认: http://localhost:1234）")
        print("3. 至少加载了一个模型")

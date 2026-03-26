"""
智能问答引擎模块
支持从问答库中模糊匹配问题并返回答案
"""

import json
import re
import logging
from typing import List, Dict, Any, Optional, Tuple
from difflib import SequenceMatcher

logger = logging.getLogger(__name__)


class QAEngine:
    """问答引擎，用于智能匹配问答库"""

    def __init__(self, qa_database_path: str = "qa_database.json"):
        self.qa_database_path = qa_database_path
        self.qa_pairs = []
        self.settings = {}
        self._load_database()

    def _load_database(self):
        """加载问答库"""
        try:
            with open(self.qa_database_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.qa_pairs = data.get("qa_pairs", [])
                self.settings = data.get("settings", {})
            logger.info(f"成功加载问答库，共 {len(self.qa_pairs)} 条问答")
        except Exception as e:
            logger.error(f"加载问答库失败: {e}")
            self.qa_pairs = []
            self.settings = {
                "default_similarity_threshold": 0.6,
                "max_matches": 1,
                "include_unmatched_notice": True,
                "unmatched_notice": "您好！您的问题我已收到，我会尽快为您处理。",
            }

    def reload_database(self):
        """重新加载问答库"""
        self._load_database()

    def _preprocess_text(self, text: str) -> str:
        """预处理文本：去除标点、空格，转为小写"""
        # 去除所有标点符号
        text = re.sub(r"[^\w\s]", "", text)
        # 去除多余空格
        text = " ".join(text.split())
        # 转为小写
        return text.lower().strip()

    def _calculate_similarity(self, text1: str, text2: str) -> float:
        """
        计算两段文本的相似度
        使用多种算法综合评分
        """
        # 预处理
        t1 = self._preprocess_text(text1)
        t2 = self._preprocess_text(text2)

        if not t1 or not t2:
            return 0.0

        # 1. SequenceMatcher相似度（基于最长公共子序列）
        seq_sim = SequenceMatcher(None, t1, t2).ratio()

        # 2. Jaccard相似度（基于词集合）
        words1 = set(t1.split())
        words2 = set(t2.split())

        if not words1 or not words2:
            jaccard_sim = 0.0
        else:
            intersection = words1 & words2
            union = words1 | words2
            jaccard_sim = len(intersection) / len(union) if union else 0.0

        # 3. 包含关系检查（如果文本1包含文本2或反之，提高相似度）
        contain_bonus = 0.0
        if t1 in t2 or t2 in t1:
            contain_bonus = 0.1

        # 4. 关键词匹配（计算匹配的关键词比例）
        keyword_sim = 0.0
        if words1 and words2:
            matched_words = words1 & words2
            keyword_sim = len(matched_words) / max(len(words1), len(words2))

        # 综合评分（加权平均）
        # SequenceMatcher: 40%, Jaccard: 30%, Keyword: 20%, Contain: 10%
        final_sim = (
            seq_sim * 0.4 + jaccard_sim * 0.3 + keyword_sim * 0.2 + contain_bonus
        )

        return min(final_sim, 1.0)  # 确保不超过1.0

    def find_best_answer(
        self, query: str, similarity_threshold: float = None
    ) -> Optional[Dict[str, Any]]:
        """
        查找最佳匹配的答案

        Args:
            query: 用户问题/邮件内容
            similarity_threshold: 相似度阈值，None则使用默认值

        Returns:
            包含答案和匹配信息的字典，或None
        """
        if not self.qa_pairs:
            logger.warning("问答库为空")
            return None

        if similarity_threshold is None:
            similarity_threshold = self.settings.get(
                "default_similarity_threshold", 0.6
            )

        best_match = None
        best_score = 0.0

        # 遍历所有问答对
        for qa in self.qa_pairs:
            questions = qa.get("questions", [])
            threshold = qa.get("similarity_threshold", similarity_threshold)

            # 与每个标准问题计算相似度，取最高值
            max_sim = 0.0
            best_question = ""

            for std_q in questions:
                sim = self._calculate_similarity(query, std_q)
                if sim > max_sim:
                    max_sim = sim
                    best_question = std_q

            # 如果超过阈值，记录为候选
            if max_sim >= threshold and max_sim > best_score:
                best_score = max_sim
                best_match = {
                    "qa_id": qa.get("id", ""),
                    "matched_question": best_question,
                    "similarity": max_sim,
                    "answer": qa.get("answer", ""),
                    "threshold": threshold,
                }

        if best_match:
            logger.info(
                f"找到匹配: '{best_match['matched_question']}' "
                f"(相似度: {best_match['similarity']:.2f})"
            )
            return best_match
        else:
            logger.info(f"未找到匹配问题 (查询: '{query[:50]}...')")
            return None

    def find_multiple_answers(
        self, query: str, top_n: int = 3, similarity_threshold: float = None
    ) -> List[Dict[str, Any]]:
        """
        查找多个匹配的答案（用于推荐相关问题）

        Args:
            query: 用户问题/邮件内容
            top_n: 返回的最大结果数
            similarity_threshold: 相似度阈值

        Returns:
            匹配结果列表
        """
        if not self.qa_pairs:
            return []

        if similarity_threshold is None:
            similarity_threshold = self.settings.get(
                "default_similarity_threshold", 0.6
            )

        matches = []

        for qa in self.qa_pairs:
            questions = qa.get("questions", [])
            threshold = qa.get("similarity_threshold", similarity_threshold)

            max_sim = 0.0
            best_question = ""

            for std_q in questions:
                sim = self._calculate_similarity(query, std_q)
                if sim > max_sim:
                    max_sim = sim
                    best_question = std_q

            if max_sim >= threshold:
                matches.append(
                    {
                        "qa_id": qa.get("id", ""),
                        "matched_question": best_question,
                        "similarity": max_sim,
                        "answer": qa.get("answer", ""),
                        "threshold": threshold,
                    }
                )

        # 按相似度排序，取前N个
        matches.sort(key=lambda x: x["similarity"], reverse=True)
        return matches[:top_n]

    def get_answer_or_fallback(
        self, query: str, similarity_threshold: float = None
    ) -> str:
        """
        获取答案，如果没有匹配则返回默认提示

        Args:
            query: 用户问题/邮件内容
            similarity_threshold: 相似度阈值

        Returns:
            答案字符串
        """
        match = self.find_best_answer(query, similarity_threshold)

        if match:
            return match["answer"]
        else:
            # 返回默认提示
            if self.settings.get("include_unmatched_notice", True):
                return self.settings.get(
                    "unmatched_notice", "您好！您的问题我已收到，我会尽快为您处理。"
                )
            else:
                return ""

    def test_match(self, query: str) -> None:
        """
        测试匹配功能，打印详细信息（用于调试）
        """
        print(f"\n查询: {query}")
        print("=" * 60)

        matches = self.find_multiple_answers(query, top_n=5, similarity_threshold=0.3)

        if matches:
            print(f"找到 {len(matches)} 个匹配:\n")
            for i, match in enumerate(matches, 1):
                print(f"{i}. 匹配问题: {match['matched_question']}")
                print(f"   相似度: {match['similarity']:.2%}")
                print(f"   阈值: {match['threshold']}")
                print(f"   答案预览: {match['answer'][:100]}...")
                print()
        else:
            print("未找到匹配")
            print(f"\n默认回复:\n{self.get_answer_or_fallback(query)}")


# 测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
    )

    # 创建引擎并测试
    engine = QAEngine("qa_database.json")

    # 测试查询
    test_queries = [
        "我想知道产品多少钱",
        "什么时候能收到货",
        "怎么申请退款",
        "系统报错了怎么办",
        "忘记了密码怎么找回",
        "周末你们上班吗",
    ]

    for query in test_queries:
        engine.test_match(query)
        print("\n" + "=" * 60 + "\n")

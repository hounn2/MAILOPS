"""
知识库管理模块
支持加载和搜索各种格式的知识库文件
"""

import os
import re
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path
from difflib import SequenceMatcher

logger = logging.getLogger(__name__)


class KnowledgeBase:
    """知识库管理器"""

    def __init__(self, kb_path: str = "knowledge_base"):
        """
        初始化知识库

        Args:
            kb_path: 知识库文件/目录路径
        """
        self.kb_path = kb_path
        self.documents = []
        self.metadata = {}
        self._load_knowledge_base()

    def _load_knowledge_base(self):
        """加载知识库"""
        if not os.path.exists(self.kb_path):
            logger.warning(f"知识库路径不存在: {self.kb_path}")
            return

        if os.path.isdir(self.kb_path):
            # 加载目录中的所有文件
            self._load_directory(self.kb_path)
        else:
            # 加载单个文件
            self._load_file(self.kb_path)

        logger.info(f"知识库加载完成，共 {len(self.documents)} 个文档片段")

    def _load_directory(self, directory: str):
        """加载目录中的所有文件"""
        supported_extensions = [".txt", ".md", ".pdf", ".docx", ".doc"]

        for root, dirs, files in os.walk(directory):
            for filename in files:
                file_path = os.path.join(root, filename)
                ext = os.path.splitext(filename)[1].lower()

                if ext in supported_extensions:
                    try:
                        self._load_file(file_path)
                    except Exception as e:
                        logger.error(f"加载文件失败 {file_path}: {e}")

    def _load_file(self, file_path: str):
        """加载单个文件"""
        ext = os.path.splitext(file_path)[1].lower()

        if ext == ".pdf":
            self._load_pdf(file_path)
        elif ext in [".docx", ".doc"]:
            self._load_word(file_path)
        elif ext in [".txt", ".md", ""]:
            self._load_text(file_path)
        else:
            logger.warning(f"不支持的文件格式: {file_path}")

    def _load_text(self, file_path: str):
        """加载文本文件"""
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()

            # 将长文本切分成片段
            chunks = self._split_text(content, chunk_size=1000, overlap=100)

            for i, chunk in enumerate(chunks):
                self.documents.append(
                    {
                        "id": f"{file_path}#{i}",
                        "source": file_path,
                        "content": chunk,
                        "type": "text",
                    }
                )

            logger.info(f"加载文本文件: {file_path} ({len(chunks)} 个片段)")

        except Exception as e:
            logger.error(f"加载文本文件失败 {file_path}: {e}")

    def _load_pdf(self, file_path: str):
        """加载PDF文件"""
        try:
            # 尝试导入PyPDF2
            import PyPDF2

            with open(file_path, "rb") as f:
                pdf_reader = PyPDF2.PdfReader(f)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"

            # 切分成片段
            chunks = self._split_text(text, chunk_size=1000, overlap=100)

            for i, chunk in enumerate(chunks):
                self.documents.append(
                    {
                        "id": f"{file_path}#{i}",
                        "source": file_path,
                        "content": chunk,
                        "type": "pdf",
                    }
                )

            logger.info(f"加载PDF文件: {file_path} ({len(chunks)} 个片段)")

        except ImportError:
            logger.warning(f"未安装PyPDF2，无法加载PDF文件。请运行: pip install PyPDF2")
        except Exception as e:
            logger.error(f"加载PDF文件失败 {file_path}: {e}")

    def _load_word(self, file_path: str):
        """加载Word文档"""
        try:
            # 尝试导入python-docx
            import docx

            doc = docx.Document(file_path)
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])

            # 切分成片段
            chunks = self._split_text(text, chunk_size=1000, overlap=100)

            for i, chunk in enumerate(chunks):
                self.documents.append(
                    {
                        "id": f"{file_path}#{i}",
                        "source": file_path,
                        "content": chunk,
                        "type": "word",
                    }
                )

            logger.info(f"加载Word文件: {file_path} ({len(chunks)} 个片段)")

        except ImportError:
            logger.warning(
                f"未安装python-docx，无法加载Word文件。请运行: pip install python-docx"
            )
        except Exception as e:
            logger.error(f"加载Word文件失败 {file_path}: {e}")

    def _split_text(
        self, text: str, chunk_size: int = 1000, overlap: int = 100
    ) -> List[str]:
        """
        将长文本切分成片段

        Args:
            text: 原始文本
            chunk_size: 每个片段的大小
            overlap: 片段间的重叠大小

        Returns:
            文本片段列表
        """
        # 按段落分割
        paragraphs = text.split("\n")
        chunks = []
        current_chunk = ""

        for paragraph in paragraphs:
            paragraph = paragraph.strip()
            if not paragraph:
                continue

            if len(current_chunk) + len(paragraph) + 1 <= chunk_size:
                current_chunk += paragraph + "\n"
            else:
                if current_chunk:
                    chunks.append(current_chunk.strip())
                current_chunk = paragraph + "\n"

        if current_chunk:
            chunks.append(current_chunk.strip())

        # 如果没有切分（文本较短），直接返回
        if not chunks and text:
            chunks = [text]

        return chunks if chunks else [text]

    def _calculate_similarity(self, text1: str, text2: str) -> float:
        """
        计算两段文本的相似度
        """
        # 预处理
        t1 = self._preprocess(text1)
        t2 = self._preprocess(text2)

        if not t1 or not t2:
            return 0.0

        # SequenceMatcher相似度
        return SequenceMatcher(None, t1, t2).ratio()

    def _preprocess(self, text: str) -> str:
        """文本预处理"""
        # 去除多余空格
        text = " ".join(text.split())
        # 转为小写
        return text.lower()

    def search_relevant(
        self, query: str, top_k: int = 3, min_score: float = 0.01
    ) -> List[Dict[str, Any]]:
        """
        搜索相关的知识库文档

        Args:
            query: 查询文本
            top_k: 返回的最相关文档数量
            min_score: 最低相似度阈值（默认0.01，几乎不限制）

        Returns:
            相关文档列表
        """
        if not self.documents:
            logger.warning("知识库为空，无法搜索")
            return []

        logger.info(
            f"开始搜索知识库，查询: '{query[:100]}...', 总文档数: {len(self.documents)}"
        )

        # 计算每个文档与查询的相似度
        scored_docs = []
        for doc in self.documents:
            similarity = self._calculate_similarity(query, doc["content"])
            if similarity > 0:  # 只要有点相似就记录
                scored_docs.append({**doc, "score": similarity})
                logger.debug(f"文档相似度: {similarity:.2%} - {doc['id'][:50]}...")

        if not scored_docs:
            logger.warning("所有文档相似度都为0，可能查询和文档完全不相关")
            return []

        # 按相似度排序
        scored_docs.sort(key=lambda x: x["score"], reverse=True)

        logger.info(f"相似度最高的5个文档:")
        for i, doc in enumerate(scored_docs[:5]):
            logger.info(f"  {i + 1}. {doc['id'][:60]}... - 相似度: {doc['score']:.2%}")

        # 返回前K个，过滤掉相似度过低的
        relevant = [doc for doc in scored_docs[:top_k] if doc["score"] > min_score]

        if relevant:
            logger.info(f"找到 {len(relevant)} 篇相关文档（阈值>{min_score:.2%}）")
        else:
            logger.warning(
                f"未找到相似度>{min_score:.2%}的文档，但最高相似度为: {scored_docs[0]['score']:.2%}"
            )

        return relevant

    def reload(self):
        """重新加载知识库"""
        self.documents = []
        self._load_knowledge_base()

    def get_stats(self) -> Dict[str, Any]:
        """获取知识库统计信息"""
        return {
            "total_documents": len(self.documents),
            "kb_path": self.kb_path,
            "sources": list(set(doc["source"] for doc in self.documents)),
        }


# 测试代码
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
    )

    # 创建测试知识库目录和文件
    test_kb_dir = "test_knowledge_base"
    os.makedirs(test_kb_dir, exist_ok=True)

    # 创建测试文件
    test_content = """# 产品使用手册

## 产品介绍

我们的产品是业界领先的解决方案，具有以下特点：
- 高性能
- 易用性
- 安全可靠

## 常见问题

Q: 如何安装产品？
A: 请参考安装指南，步骤如下：
1. 下载安装包
2. 运行安装程序
3. 配置环境变量

Q: 产品价格是多少？
A: 基础版免费，专业版¥999/年，企业版¥2999/年。

## 技术支持

如有问题，请联系：
- 邮箱: support@company.com
- 电话: 400-xxx-xxxx
"""

    test_file = os.path.join(test_kb_dir, "product_manual.txt")
    with open(test_file, "w", encoding="utf-8") as f:
        f.write(test_content)

    # 加载知识库
    kb = KnowledgeBase(test_kb_dir)

    # 显示统计信息
    stats = kb.get_stats()
    print(f"\n知识库统计:")
    print(f"  文档片段数: {stats['total_documents']}")
    print(f"  来源文件: {stats['sources']}")

    # 测试搜索
    test_queries = ["产品怎么安装", "多少钱", "技术支持联系方式"]

    print("\n" + "=" * 60)
    for query in test_queries:
        print(f"\n查询: {query}")
        results = kb.search_relevant(query, top_k=2)

        if results:
            print(f"找到 {len(results)} 个相关片段:")
            for i, doc in enumerate(results, 1):
                print(f"\n  片段 {i} (相似度: {doc['score']:.2%}):")
                print(f"  {doc['content'][:200]}...")
        else:
            print("未找到相关内容")

        print("\n" + "-" * 60)

    # 清理测试文件
    import shutil

    shutil.rmtree(test_kb_dir)
    print("\n测试完成")

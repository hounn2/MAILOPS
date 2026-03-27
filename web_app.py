#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook邮件助手 - Web控制界面
包含规则配置和执行结果记录

使用方法:
    1. 安装依赖: pip install -r requirements_web.txt
    2. 运行: python web_app.py
    3. 在浏览器中打开: http://localhost:5000

功能:
    - 可视化规则配置（增删改查）
    - 实时执行邮件处理
    - 执行结果记录和查看
    - 日志查询和导出
"""

import os
import sys
import json
import sqlite3
import threading
import time
import logging
from datetime import datetime
from typing import Dict, Any, List
from pathlib import Path

import requests
from flask import Flask, render_template, jsonify, request, send_file
from flask_cors import CORS

# 确保在Windows上运行
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

app = Flask(__name__)
CORS(app)

# 全局变量
CONFIG_FILE = "config.json"
DB_FILE = "execution_logs.db"
processing_thread = None
is_processing = False

# 自动执行相关
auto_execution_thread = None
auto_execution_running = False
auto_execution_stop_event = None

# ============================================================================
# 数据库操作
# ============================================================================


class LogDatabase:
    """执行日志数据库"""

    def __init__(self, db_file: str = DB_FILE):
        self.db_file = db_file
        self._init_db()

    def _init_db(self):
        """初始化数据库"""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        # 创建执行记录表
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS execution_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                execution_time TEXT NOT NULL,
                total_emails INTEGER DEFAULT 0,
                matched_emails INTEGER DEFAULT 0,
                actions_executed INTEGER DEFAULT 0,
                errors INTEGER DEFAULT 0,
                duration REAL DEFAULT 0,
                status TEXT DEFAULT 'success',
                message TEXT
            )
        """)

        # 创建邮件处理详情表
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS email_details (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                log_id INTEGER,
                email_subject TEXT,
                sender TEXT,
                matched_rules TEXT,
                actions_taken TEXT,
                status TEXT,
                processed_time TEXT,
                FOREIGN KEY (log_id) REFERENCES execution_logs(id)
            )
        """)

        conn.commit()
        conn.close()
        logger.info("数据库初始化完成")

    def add_execution_log(self, stats: Dict[str, Any]):
        """添加执行记录"""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        cursor.execute(
            """
            INSERT INTO execution_logs 
            (execution_time, total_emails, matched_emails, actions_executed, errors, duration, status, message)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
            (
                datetime.now().isoformat(),
                stats.get("processed", 0),
                stats.get("matched", 0),
                stats.get("actions_executed", 0),
                stats.get("errors", 0),
                stats.get("duration", 0),
                stats.get("status", "success"),
                stats.get("message", ""),
            ),
        )

        log_id = cursor.lastrowid
        conn.commit()
        conn.close()

        return log_id

    def add_email_detail(self, log_id: int, detail: Dict[str, Any]):
        """添加邮件处理详情"""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        cursor.execute(
            """
            INSERT INTO email_details 
            (log_id, email_subject, sender, matched_rules, actions_taken, status, processed_time)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
            (
                log_id,
                detail.get("subject", ""),
                detail.get("sender", ""),
                json.dumps(detail.get("matched_rules", []), ensure_ascii=False),
                json.dumps(detail.get("actions_taken", []), ensure_ascii=False),
                detail.get("status", "processed"),
                datetime.now().isoformat(),
            ),
        )

        conn.commit()
        conn.close()

    def get_execution_logs(self, limit: int = 100, offset: int = 0) -> List[Dict]:
        """获取执行记录"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT * FROM execution_logs 
            ORDER BY execution_time DESC 
            LIMIT ? OFFSET ?
        """,
            (limit, offset),
        )

        rows = cursor.fetchall()
        logs = [dict(row) for row in rows]

        conn.close()
        return logs

    def get_email_details(self, log_id: int) -> List[Dict]:
        """获取某次执行的邮件详情"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT * FROM email_details 
            WHERE log_id = ?
            ORDER BY processed_time DESC
        """,
            (log_id,),
        )

        rows = cursor.fetchall()
        details = [dict(row) for row in rows]

        conn.close()
        return details

    def get_statistics(self) -> Dict[str, Any]:
        """获取统计信息"""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        # 总执行次数
        cursor.execute("SELECT COUNT(*) FROM execution_logs")
        total_executions = cursor.fetchone()[0]

        # 总处理邮件数
        cursor.execute("SELECT SUM(total_emails) FROM execution_logs")
        total_emails = cursor.fetchone()[0] or 0

        # 总匹配邮件数
        cursor.execute("SELECT SUM(matched_emails) FROM execution_logs")
        total_matched = cursor.fetchone()[0] or 0

        # 总执行动作数
        cursor.execute("SELECT SUM(actions_executed) FROM execution_logs")
        total_actions = cursor.fetchone()[0] or 0

        # 今日统计
        today = datetime.now().strftime("%Y-%m-%d")
        cursor.execute("""
            SELECT SUM(total_emails), SUM(matched_emails), SUM(actions_executed)
            FROM execution_logs
            WHERE date(execution_time) = date('now')
        """)
        today_stats = cursor.fetchone()

        conn.close()

        return {
            "total_executions": total_executions,
            "total_emails": total_emails,
            "total_matched": total_matched,
            "total_actions": total_actions,
            "today_emails": today_stats[0] or 0,
            "today_matched": today_stats[1] or 0,
            "today_actions": today_stats[2] or 0,
        }


# 初始化数据库
db = LogDatabase()

# ============================================================================
# 配置管理
# ============================================================================


def load_config() -> Dict[str, Any]:
    """加载配置文件"""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"加载配置失败: {e}")
        return {
            "rules": [],
            "templates": {},
            "settings": {
                "check_interval": 60,
                "process_unread_only": True,
                "max_emails_per_batch": 50,
                "dry_run": False,
            },
        }


def save_config(config: Dict[str, Any]) -> bool:
    """保存配置文件"""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        logger.error(f"保存配置失败: {e}")
        return False


# ============================================================================
# 邮件处理核心（内嵌）
# ============================================================================


class OutlookAssistantCore:
    """邮件处理核心类"""

    def __init__(self, config: Dict[str, Any], dry_run: bool = False):
        self.config = config
        self.dry_run = dry_run
        self.db = LogDatabase()

    def process_once(self) -> Dict[str, Any]:
        """执行一次邮件处理"""
        import pythoncom

        try:
            # 初始化COM（解决后台线程中的COM初始化问题）
            try:
                pythoncom.CoInitialize()
            except:
                pass  # 可能已经初始化过

            # 导入standalone版本的代码
            import importlib.util

            spec = importlib.util.spec_from_file_location(
                "standalone", "outlook_assistant_win_standalone.py"
            )
            if spec is None or spec.loader is None:
                raise Exception("无法加载outlook_assistant_win_standalone.py模块")
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)

            # 创建助手实例
            assistant = module.OutlookAssistantWindows(
                CONFIG_FILE, dry_run=self.dry_run
            )
            stats = assistant.process_emails()

            # 添加状态信息
            stats["status"] = "success"
            stats["message"] = "执行成功"

            # 记录到数据库，获取log_id
            log_id = self.db.add_execution_log(stats)

            # 记录每封邮件的详情
            if log_id:
                email_details = stats.get("email_details", [])
                for detail in email_details:
                    self.db.add_email_detail(log_id, detail)

            return stats

        except Exception as e:
            logger.error(f"邮件处理失败: {e}")
            error_stats = {
                "status": "error",
                "message": str(e),
                "processed": 0,
                "matched": 0,
                "actions_executed": 0,
                "errors": 1,
                "duration": 0,
            }
            self.db.add_execution_log(error_stats)
            return error_stats
        finally:
            # 释放COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass


# ============================================================================
# Web路由
# ============================================================================


@app.route("/")
def index():
    """主页"""
    return render_template("index.html")


@app.route("/api/config", methods=["GET"])
def get_config():
    """获取配置"""
    config = load_config()
    return jsonify(config)


@app.route("/api/config", methods=["POST"])
def update_config():
    """更新配置"""
    try:
        config = request.json
        if save_config(config):
            return jsonify({"success": True, "message": "配置保存成功"})
        else:
            return jsonify({"success": False, "message": "配置保存失败"}), 500
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/rules", methods=["GET"])
def get_rules():
    """获取所有规则"""
    config = load_config()
    return jsonify(config.get("rules", []))


@app.route("/api/rules", methods=["POST"])
def add_rule():
    """添加规则"""
    try:
        config = load_config()
        new_rule = request.json

        # 生成规则ID
        import uuid

        new_rule["id"] = f"rule_{uuid.uuid4().hex[:8]}"

        config["rules"].append(new_rule)

        if save_config(config):
            return jsonify({"success": True, "rule": new_rule})
        else:
            return jsonify({"success": False, "message": "保存失败"}), 500
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/rules/<rule_id>", methods=["PUT"])
def update_rule(rule_id):
    """更新规则"""
    try:
        config = load_config()
        updated_rule = request.json

        for i, rule in enumerate(config["rules"]):
            if rule.get("id") == rule_id:
                config["rules"][i] = updated_rule
                break

        if save_config(config):
            return jsonify({"success": True})
        else:
            return jsonify({"success": False, "message": "保存失败"}), 500
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/rules/<rule_id>", methods=["DELETE"])
def delete_rule(rule_id):
    """删除规则"""
    try:
        config = load_config()
        config["rules"] = [r for r in config["rules"] if r.get("id") != rule_id]

        if save_config(config):
            return jsonify({"success": True})
        else:
            return jsonify({"success": False, "message": "保存失败"}), 500
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/execute", methods=["POST"])
def execute_now():
    """立即执行邮件处理"""
    global is_processing

    if is_processing:
        return jsonify({"success": False, "message": "正在执行中，请稍候"}), 429

    try:
        data = request.json or {}
        dry_run = data.get("dry_run", False)

        config = load_config()
        assistant = OutlookAssistantCore(config, dry_run=dry_run)

        # 在后台线程中执行
        def run_processing():
            global is_processing
            is_processing = True
            try:
                result = assistant.process_once()
                logger.info(f"执行完成: {result}")
            finally:
                is_processing = False

        thread = threading.Thread(target=run_processing)
        thread.start()

        return jsonify(
            {
                "success": True,
                "message": "已开始执行" + (" (试运行模式)" if dry_run else ""),
            }
        )

    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/logs", methods=["GET"])
def get_logs():
    """获取执行日志"""
    try:
        limit = request.args.get("limit", 50, type=int)
        offset = request.args.get("offset", 0, type=int)

        logs = db.get_execution_logs(limit=limit, offset=offset)
        return jsonify({"success": True, "logs": logs})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/logs/<log_id>/details", methods=["GET"])
def get_log_details(log_id):
    """获取某次执行的详情"""
    try:
        details = db.get_email_details(log_id)
        return jsonify({"success": True, "details": details})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/statistics", methods=["GET"])
def get_statistics():
    """获取统计信息"""
    try:
        stats = db.get_statistics()
        return jsonify({"success": True, "statistics": stats})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/status", methods=["GET"])
def get_status():
    """获取当前状态"""
    return jsonify(
        {
            "is_processing": is_processing,
            "config_file_exists": os.path.exists(CONFIG_FILE),
        }
    )


@app.route("/api/debug/db", methods=["GET"])
def debug_db():
    """调试API - 查看数据库内容"""
    try:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # 获取最近的5条执行记录
        cursor.execute("SELECT * FROM execution_logs ORDER BY id DESC LIMIT 5")
        logs = [dict(row) for row in cursor.fetchall()]

        # 获取所有邮件详情
        cursor.execute("SELECT * FROM email_details ORDER BY id DESC LIMIT 10")
        details = [dict(row) for row in cursor.fetchall()]

        # 获取表结构
        cursor.execute("PRAGMA table_info(email_details)")
        schema = [dict(row) for row in cursor.fetchall()]

        conn.close()

        return jsonify(
            {
                "success": True,
                "recent_logs": logs,
                "recent_details": details,
                "schema": schema,
            }
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


# ============================================================================
# AI功能相关API
# ============================================================================


@app.route("/api/ai/status", methods=["GET"])
def get_ai_status():
    """获取AI引擎状态"""
    try:
        # 获取查询参数中的自定义配置
        custom_base_url = request.args.get("base_url")

        config = load_config()
        lmstudio_config = config.get("lmstudio", {})
        kb_config = config.get("knowledge_base", {})

        # 使用查询参数或配置文件中的地址
        base_url = custom_base_url or lmstudio_config.get(
            "base_url", "http://localhost:1234"
        )

        # 测试LMStudio连接
        ai_ready = False
        ai_error = None
        model_count = 0
        try:
            response = requests.get(base_url.rstrip("/") + "/v1/models", timeout=5)
            if response.status_code == 200:
                ai_ready = True
                data = response.json()
                model_count = len(data.get("data", []))
            else:
                ai_error = f"HTTP {response.status_code}"
        except Exception as e:
            if "Connection" in str(type(e)):
                ai_error = "无法连接到LMStudio服务器，请确保LMStudio已启动"
            elif "Timeout" in str(type(e)):
                ai_error = "连接超时"
            else:
                ai_error = str(e)

        # 检查知识库
        kb_ready = False
        kb_stats = {"total_documents": 0, "sources": []}
        if kb_config.get("enabled", False):
            kb_path = kb_config.get("path", "knowledge_base")
            if os.path.exists(kb_path):
                try:
                    from knowledge_base import KnowledgeBase

                    kb = KnowledgeBase(kb_path)
                    kb_stats = kb.get_stats()
                    kb_ready = True
                except Exception as e:
                    kb_stats["error"] = str(e)

        return jsonify(
            {
                "success": True,
                "ai": {
                    "enabled": lmstudio_config.get("enabled", False),
                    "ready": ai_ready,
                    "error": ai_error,
                    "config": lmstudio_config,
                    "model_count": model_count if ai_ready else 0,
                },
                "knowledge_base": {
                    "enabled": kb_config.get("enabled", False),
                    "ready": kb_ready,
                    "stats": kb_stats,
                    "config": kb_config,
                },
            }
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/ai/test", methods=["POST"])
def test_ai():
    """测试AI生成功能"""
    try:
        data = request.json
        test_email = data.get("email_content", "")
        use_kb = data.get("use_knowledge_base", True)

        config = load_config()
        lmstudio_config = config.get("lmstudio", {})
        kb_config = config.get("knowledge_base", {})

        if not lmstudio_config.get("enabled", False):
            return jsonify({"success": False, "message": "AI功能未启用"}), 400

        # 初始化AI引擎
        from ai_engine import AIReplyEngine

        knowledge_base = None
        logger.info(
            f"API测试 - 知识库配置: use_kb={use_kb}, kb_enabled={kb_config.get('enabled', False)}"
        )

        if use_kb and kb_config.get("enabled", False):
            from knowledge_base import KnowledgeBase

            kb_path = kb_config.get("path", "knowledge_base")
            kb_path_full = os.path.abspath(kb_path)
            logger.info(f"API测试 - 尝试加载知识库: {kb_path_full}")

            if os.path.exists(kb_path):
                try:
                    knowledge_base = KnowledgeBase(kb_path)
                    stats = knowledge_base.get_stats()
                    logger.info(
                        f"API测试 - 知识库加载成功，文档数: {stats.get('total_documents', 0)}"
                    )
                except Exception as e:
                    logger.error(f"API测试 - 知识库加载失败: {e}")
            else:
                logger.error(f"API测试 - 知识库路径不存在: {kb_path_full}")

        ai_engine = AIReplyEngine(lmstudio_config, knowledge_base)

        if not ai_engine.is_ready():
            return jsonify(
                {
                    "success": False,
                    "message": "AI引擎未就绪，请检查LMStudio是否正常运行",
                }
            ), 500

        # 模拟邮件数据
        email_data = {
            "subject": "测试邮件",
            "body": test_email,
            "sender": "test@example.com",
        }

        # 生成回复
        reply = ai_engine.generate_reply(email_data, search_knowledge=use_kb)

        if reply:
            return jsonify({"success": True, "reply": reply})
        else:
            return jsonify({"success": False, "message": "生成回复失败"}), 500

    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/kb/files", methods=["GET"])
def get_kb_files():
    """获取知识库文件列表"""
    try:
        config = load_config()
        kb_config = config.get("knowledge_base", {})
        kb_path = kb_config.get("path", "knowledge_base")

        files = []
        if os.path.exists(kb_path):
            for root, dirs, filenames in os.walk(kb_path):
                for filename in filenames:
                    file_path = os.path.join(root, filename)
                    rel_path = os.path.relpath(file_path, kb_path)
                    files.append(
                        {
                            "name": filename,
                            "path": rel_path,
                            "size": os.path.getsize(file_path),
                            "modified": datetime.fromtimestamp(
                                os.path.getmtime(file_path)
                            ).isoformat(),
                        }
                    )

        return jsonify({"success": True, "files": files})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/kb/upload", methods=["POST"])
def upload_kb_file():
    """上传知识库文件"""
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "没有文件"}), 400

        file = request.files["file"]
        if not file or not file.filename:
            return jsonify({"success": False, "message": "文件名为空"}), 400

        # 获取文件名并安全检查
        filename = str(file.filename)

        # 检查文件类型
        allowed_extensions = {".txt", ".md", ".pdf", ".docx", ".doc"}
        ext = os.path.splitext(filename)[1].lower()
        if ext not in allowed_extensions:
            return jsonify(
                {"success": False, "message": f"不支持的文件类型: {ext}"}
            ), 400

        # 确保知识库目录存在
        config = load_config()
        kb_config = config.get("knowledge_base", {})
        kb_path = kb_config.get("path", "knowledge_base")
        os.makedirs(kb_path, exist_ok=True)

        # 保存文件
        file_path = os.path.join(kb_path, filename)
        file.save(file_path)

        return jsonify({"success": True, "message": f"文件上传成功: {filename}"})

    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/kb/delete/<path:filename>", methods=["DELETE"])
def delete_kb_file(filename):
    """删除知识库文件"""
    try:
        if not filename:
            return jsonify({"success": False, "message": "文件名不能为空"}), 400

        config = load_config()
        kb_config = config.get("knowledge_base", {})
        kb_path = str(kb_config.get("path", "knowledge_base"))
        file_path = os.path.join(kb_path, filename)

        # 安全检查：确保文件在知识库目录内
        real_path = os.path.realpath(file_path)
        real_kb_path = os.path.realpath(kb_path)
        if not real_path.startswith(real_kb_path):
            return jsonify({"success": False, "message": "无效的文件路径"}), 400

        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({"success": True, "message": f"文件已删除: {filename}"})
        else:
            return jsonify({"success": False, "message": "文件不存在"}), 404

    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


# ============================================================================
# 自动执行相关API
# ============================================================================


@app.route("/api/auto_execution", methods=["GET"])
def get_auto_execution_status():
    """获取自动执行状态"""
    try:
        config = load_config()
        settings = config.get("settings", {})

        return jsonify(
            {
                "success": True,
                "enabled": settings.get("auto_execution", False),
                "interval": settings.get("check_interval", 60),
                "running": auto_execution_running,
            }
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/auto_execution", methods=["POST"])
def toggle_auto_execution():
    """开启/关闭自动执行"""
    global auto_execution_running, auto_execution_thread, auto_execution_stop_event

    try:
        data = request.json or {}
        enabled = data.get("enabled", False)

        config = load_config()
        config["settings"]["auto_execution"] = enabled
        save_config(config)

        if enabled:
            if not auto_execution_running:
                # 启动自动执行线程
                auto_execution_stop_event = threading.Event()
                auto_execution_thread = threading.Thread(
                    target=auto_execution_worker,
                    args=(auto_execution_stop_event,),
                    daemon=True,
                )
                auto_execution_thread.start()
                auto_execution_running = True
                logger.info("自动执行已启动")
                return jsonify({"success": True, "message": "自动执行已启动"})
        else:
            if auto_execution_running and auto_execution_stop_event:
                # 停止自动执行
                auto_execution_stop_event.set()
                auto_execution_running = False
                logger.info("自动执行已停止")
                return jsonify({"success": True, "message": "自动执行已停止"})

        return jsonify({"success": True, "message": "设置已保存"})
    except Exception as e:
        logger.error(f"切换自动执行状态失败: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


def auto_execution_worker(stop_event):
    """自动执行工作线程"""
    logger.info("自动执行工作线程已启动")

    while not stop_event.is_set():
        try:
            config = load_config()
            settings = config.get("settings", {})

            # 检查是否启用了自动执行
            if not settings.get("auto_execution", False):
                time.sleep(5)
                continue

            interval = settings.get("check_interval", 60)
            dry_run = settings.get("dry_run", False)

            # 执行邮件处理
            assistant = OutlookAssistantCore(config, dry_run=dry_run)
            result = assistant.process_once()

            if result.get("status") == "success":
                logger.info(f"自动执行完成: 处理{result.get('processed', 0)}封邮件")
            else:
                logger.error(f"自动执行失败: {result.get('message', '未知错误')}")

            # 等待间隔时间（可被提前终止）
            for _ in range(interval):
                if stop_event.is_set():
                    break
                time.sleep(1)

        except Exception as e:
            logger.error(f"自动执行异常: {e}")
            time.sleep(10)  # 出错后等待10秒再试

    logger.info("自动执行工作线程已停止")


# ============================================================================
# 主程序
# ============================================================================

if __name__ == "__main__":
    print("=" * 60)
    print("Outlook邮件助手 - Web控制界面")
    print("=" * 60)
    print()
    print("请在浏览器中打开: http://localhost:5000")
    print()
    print("按 Ctrl+C 停止服务")
    print("=" * 60)

    # 确保templates和static目录存在
    Path("templates").mkdir(exist_ok=True)
    Path("static/css").mkdir(parents=True, exist_ok=True)
    Path("static/js").mkdir(parents=True, exist_ok=True)

    # 检查是否需要自动启动
    try:
        config = load_config()
        settings = config.get("settings", {})
        if settings.get("auto_execution", False):
            # 启动自动执行
            auto_execution_stop_event = threading.Event()
            auto_execution_thread = threading.Thread(
                target=auto_execution_worker,
                args=(auto_execution_stop_event,),
                daemon=True,
            )
            auto_execution_thread.start()
            auto_execution_running = True
            print("\n✅ 自动执行已启动")
    except Exception as e:
        print(f"\n⚠️ 自动执行启动失败: {e}")

    app.run(host="0.0.0.0", port=5000, debug=False)

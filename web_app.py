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
import logging
from datetime import datetime
from typing import Dict, Any, List
from pathlib import Path

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

            # 记录到数据库
            self.db.add_execution_log(stats)

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

    app.run(host="0.0.0.0", port=5000, debug=False)

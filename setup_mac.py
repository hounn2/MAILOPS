#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
macOS App Bundle 打包配置
使用 py2app 打包成独立的 .app 应用

使用方法:
    1. 安装 py2app: pip install py2app
    2. 运行: python setup.py py2app
    3. 在 dist/ 目录下生成 OutlookAssistant.app
"""

from setuptools import setup

APP = ["outlook_assistant_mac.py"]
DATA_FILES = [
    "config.json",
    "templates",
    "static",
]

OPTIONS = {
    "argv_emulation": True,
    "packages": ["msal", "requests", "dateutil"],
    "includes": [
        "json",
        "re",
        "time",
        "logging",
        "argparse",
        "datetime",
        "sqlite3",
        "threading",
        "pathlib",
        "urllib.parse",
    ],
    "excludes": ["tkinter", "matplotlib", "numpy", "pandas"],
    # "iconfile": "icon.icns",  # 如果有图标文件的话，取消注释这行
    "plist": {
        "CFBundleName": "Outlook邮件助手",
        "CFBundleDisplayName": "Outlook邮件助手",
        "CFBundleGetInfoString": "Outlook邮件助手 - macOS版本",
        "CFBundleIdentifier": "com.outlook.assistant",
        "CFBundleVersion": "1.0.0",
        "CFBundleShortVersionString": "1.0.0",
        "NSHumanReadableCopyright": "Copyright 2026",
        "LSMinimumSystemVersion": "10.14",
    },
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={"py2app": OPTIONS},
    setup_requires=["py2app"],
)

#!/bin/bash
# macOS 安装脚本
# 自动安装依赖并配置环境

echo "=========================================="
echo "Outlook邮件助手 - macOS安装脚本"
echo "=========================================="
echo

# 检查Python版本
echo "🔍 检查Python版本..."
python3 --version
if [ $? -ne 0 ]; then
    echo "❌ 未找到Python3，请先安装Python 3.7+"
    echo "可以从 https://www.python.org/downloads/mac-osx/ 下载"
    exit 1
fi
echo "✅ Python3已安装"
echo

# 创建虚拟环境（可选）
read -p "是否创建虚拟环境? (y/n) " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo "📦 创建虚拟环境..."
    python3 -m venv venv
    source venv/bin/activate
    echo "✅ 虚拟环境已创建并激活"
    echo
fi

# 安装依赖
echo "📦 安装依赖..."
pip install -r requirements_mac.txt
if [ $? -ne 0 ]; then
    echo "❌ 依赖安装失败"
    exit 1
fi
echo "✅ 依赖安装完成"
echo

# 首次设置
echo "⚙️  开始首次设置..."
python3 outlook_assistant_mac.py --setup
if [ $? -ne 0 ]; then
    echo "❌ 设置失败"
    exit 1
fi

echo
echo "=========================================="
echo "安装完成！"
echo "=========================================="
echo
echo "使用方法:"
echo "  python3 outlook_assistant_mac.py --once    # 运行一次"
echo "  python3 outlook_assistant_mac.py           # 持续运行"
echo "  python3 outlook_assistant_mac.py --dry-run # 试运行模式"
echo

if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo "注意: 虚拟环境已激活，退出时请运行: deactivate"
fi

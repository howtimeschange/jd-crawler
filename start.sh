#!/bin/bash

# 京东商品爬虫 - 一键启动脚本

set -e

echo ""
echo "🛒 京东商品爬虫工具"
echo "=============================="

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 未找到 Python3，请先安装 Python"
    exit 1
fi

echo "✅ Python3: $(python3 --version)"

# 检查并安装依赖
echo ""
echo "📦 检查依赖..."

install_deps() {
    pip3 install --quiet playwright openpyxl 2>/dev/null || \
    pip3 install --user --quiet playwright openpyxl 2>/dev/null || \
    python3 -m pip install --quiet playwright openpyxl 2>/dev/null || true
}

if ! python3 -c "from playwright.sync_api import sync_playwright" 2>/dev/null; then
    echo "   安装 playwright..."
    install_deps
fi

if ! python3 -c "import openpyxl" 2>/dev/null; then
    echo "   安装 openpyxl..."
    install_deps
fi

echo "   ✅ 依赖就绪"

# 运行爬虫（交互模式）
echo ""
echo "🚀 启动爬虫..."
echo ""

python3 jd_crawler.py

echo ""
echo "👋 完成！"

#!/bin/bash

# 京东商品爬虫 - 一键启动脚本

echo ""
echo "🛒 京东商品爬虫工具"
echo "=============================="

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 未找到 Python3，请先安装 Python"
    exit 1
fi

echo "✅ Python3: $(python3 --version)"

# 直接运行Python检查，不尝试安装
echo ""
echo "📦 检查依赖..."

# 用python3直接检查
python3 -c "from playwright.sync_api import sync_playwright" 2>/dev/null
if [ $? -eq 0 ]; then
    echo "   ✅ playwright: OK"
else
    echo "   ❌ playwright: 未安装"
    echo "   请运行: pip3 install playwright"
fi

python3 -c "import openpyxl" 2>/dev/null
if [ $? -eq 0 ]; then
    echo "   ✅ openpyxl: OK"
else
    echo "   ❌ openpyxl: 未安装"
    echo "   请运行: pip3 install openpyxl"
fi

# 运行爬虫
echo ""
echo "🚀 启动爬虫..."
echo ""

cd "$(dirname "$0")"
python3 jd_crawler.py

echo ""
echo "👋 完成！"

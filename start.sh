#!/bin/bash

# 京东商品爬虫 - 一键启动脚本

set -e

echo "🛒 京东商品爬虫启动器"
echo "=============================="

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 未找到 Python3，请先安装 Python"
    exit 1
fi

echo "✅ Python3 已安装: $(python3 --version)"

# 创建虚拟环境（可选）
VENV_DIR=".venv"
if [ ! -d "$VENV_DIR" ]; then
    echo "📦 创建虚拟环境..."
    python3 -m venv "$VENV_DIR"
fi

# 激活虚拟环境
source "$VENV_DIR/bin/activate"

# 安装依赖
echo "📦 安装依赖..."
pip install --quiet playwright openpyxl 2>/dev/null || pip install --user --quiet playwright openpyxl 2>/dev/null || true

# 安装chromium
echo "🌐 安装浏览器..."
python3 -c "from playwright.sync_api import sync_playwright" 2>/dev/null && {
    playwright install chromium 2>/dev/null || true
}

# 获取URL
if [ -n "$1" ]; then
    SHOP_URL="$1"
else
    echo ""
    echo "📝 请输入店铺商品列表链接:"
    echo "   例如: https://mall.jd.com/advance_search-2863474-1000462158-1000462158-0-0-0-1-1-60.html"
    read -p "🔗 > " SHOP_URL
fi

if [ -z "$SHOP_URL" ]; then
    echo "❌ URL不能为空"
    exit 1
fi

echo ""
echo "🚀 开始爬取..."
echo "   店铺: $SHOP_URL"
echo ""

# 运行爬虫
python3 jd_crawler.py "$SHOP_URL"

echo ""
echo "✅ 完成！"
echo "   查看输出: ls -la output/"

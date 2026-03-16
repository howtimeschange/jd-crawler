#!/bin/bash
# 京东商品爬虫 - 一键启动脚本
# 支持 Mac 和 Windows

set -e

echo ""
echo "🛒 京东商品爬虫工具"
echo "=============================="

# 检测系统
SYSTEM=$(uname -s)
if [ "$SYSTEM" = "Darwin" ]; then
    echo "✅ 系统: macOS"
elif [ "$SYSTEM" = "MINGW"* ] || [ "$SYSTEM" = "MSYS"* ]; then
    echo "✅ 系统: Windows"
else
    echo "✅ 系统: Linux"
fi

# 检测Python
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
else
    echo "❌ 未找到 Python，请先安装 Python"
    exit 1
fi

echo "✅ Python: $($PYTHON_CMD --version)"

# 检查依赖
echo ""
echo "📦 检查依赖..."

check_module() {
    $PYTHON_CMD -c "import $1" 2>/dev/null
}

if check_module "playwright"; then
    echo "   ✅ playwright: OK"
else
    echo "   ❌ playwright: 未安装"
    echo "   正在安装..."
    $PYTHON_CMD -m pip install --user playwright 2>/dev/null || \
    $PYTHON_CMD -m pip install playwright 2>/dev/null || true
fi

if check_module "openpyxl"; then
    echo "   ✅ openpyxl: OK"
else
    echo "   ❌ openpyxl: 未安装"
    echo "   正在安装..."
    $PYTHON_CMD -m pip install --user openpyxl 2>/dev/null || \
    $PYTHON_CMD -m pip install openpyxl 2>/dev/null || true
fi

# 启动爬虫
echo ""
echo "🚀 启动爬虫..."
echo ""

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

$PYTHON_CMD jd_crawler.py

echo ""
echo "👋 完成！"

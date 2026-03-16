#!/usr/bin/env python3
"""
京东商品爬虫工具 - 交互式CLI版本
功能：爬取京东店铺商品信息，支持自动翻页、Excel导出
特点：友好的交互界面，无需记住URL格式
"""

import sys
import os
import json
import re
import time
import subprocess
from pathlib import Path
from datetime import datetime

# 尝试导入playwright
try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("❌ 请先安装 playwright: pip install playwright")
    print("   或运行: pip install -r requirements.txt")
    sys.exit(1)

# 颜色定义
class Colors:
    GREEN = '\033[92m'
    BLUE = '\033[94m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    CYAN = '\033[96m'
    BOLD = '\033[1m'
    END = '\033[0m'

def print_banner():
    """打印欢迎界面"""
    print()
    print(f"{Colors.CYAN}{'='*60}{Colors.END}")
    print(f"{Colors.CYAN}  🛒  京东商品爬虫工具 v1.1{Colors.END}")
    print(f"{Colors.CYAN}{'='*60}{Colors.END}")
    print()

def print_menu():
    """打印菜单"""
    print(f"{Colors.BOLD}📋 功能菜单:{Colors.END}")
    print(f"  {Colors.GREEN}1.{Colors.END} 爬取京东店铺商品")
    print(f"  {Colors.GREEN}2.{Colors.END} 查看已有数据")
    print(f"  {Colors.GREEN}3.{Colors.END} 导出Excel表格")
    print(f"  {Colors.GREEN}0.{Colors.END} 退出")
    print()

def get_shop_url() -> str:
    """获取店铺URL - 友好交互"""
    print(f"{Colors.BOLD}📝 步骤1: 获取店铺商品列表链接{Colors.END}")
    print()
    print("请按以下步骤操作:")
    print(f"  {Colors.YELLOW}1.{Colors.END} 打开京东店铺首页")
    print(f"       例如: https://mall.jd.com/index-1000462158.html")
    print()
    print(f"  {Colors.YELLOW}2.{Colors.END} 点击店铺页面的 {Colors.CYAN}\"所有商品\"{Colors.END} 或 {Colors.CYAN}\"商品列表\"{Colors.END}")
    print()
    print(f"  {Colors.YELLOW}3.{Colors.END} 复制浏览器地址栏的完整URL")
    print()
    print(f"  {Colors.YELLOW}4.{Colors.END} 粘贴到下方")
    print()
    
    url = input(f"{Colors.GREEN}🔗 请粘贴店铺商品列表URL:{Colors.END} ").strip()
    
    if not url:
        print(f"{Colors.RED}❌ URL不能为空{Colors.END}")
        return ""
    
    if 'jd.com' not in url:
        print(f"{Colors.RED}❌ 请输入正确的京东链接{Colors.END}")
        return ""
    
    return url

# 输出目录
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
DEFAULT_OUTPUT_JSON = OUTPUT_DIR / "jd_products.json"
DEFAULT_OUTPUT_XLSX = OUTPUT_DIR / "jd_products.xlsx"


def click_next_page(page) -> bool:
    """点击下一页按钮"""
    try:
        # 多种可能的下一页选择器
        selectors = [
            '#jTopPage > a.next',  # 下一页链接
            '.next',  # class为next的链接
            'a[class*="next"]',  # 包含next的链接
            'a:has-text("下一页")',  # 文字包含下一页
            '#jTopPage a:last-child',  # 分页最后一个链接
            '[class*="pagination"] a:last-child',  # 分页容器最后链接
        ]
        
        for selector in selectors:
            try:
                next_btn = page.query_selector(selector)
                if next_btn and next_btn.is_visible():
                    # 检查是否禁用
                    if 'disabled' in (next_btn.get_attribute('class') or ''):
                        return False  # 最后一页
                    next_btn.click()
                    return True
            except:
                continue
        
        return False
    except Exception as e:
        return False


def crawl_shop(url: str) -> list:
    """爬取店铺商品"""
    print(f"\n{Colors.BOLD}🚀 开始爬取...{Colors.END}")
    print(f"   URL: {url[:50]}...")
    print()
    
    products = []
    max_pages = 6  # 最多爬取6页
    
    with sync_playwright() as p:
        # 启动浏览器（有头模式）
        browser = p.chromium.launch(headless=False, args=['--disable-blink-features=AutomationControlled'])
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
        )
        page = context.new_page()
        
        # 打开第一页
        print(f"   🌐 打开第一页...")
        page.goto(url, wait_until="networkidle", timeout=60000)
        
        for page_num in range(1, max_pages + 1):
            print(f"\n{Colors.CYAN}{'='*50}{Colors.END}")
            print(f"📄 正在处理第 {page_num}/{max_pages} 页")
            
            # 等待价格加载
            print(f"   ⏳ 等待价格加载 (15秒)...")
            page.wait_for_timeout(15000)
            
            # 滚动触发更多价格加载
            for _ in range(3):
                page.evaluate("window.scrollBy(0, 500)")
                page.wait_for_timeout(2000)
            
            # 提取数据
            print(f"   🔍 提取数据...")
            
            page_products = page.evaluate("""
                () => {
                    const products = [];
                    const items = document.querySelectorAll('.jSubObject');
                    
                    items.forEach(item => {
                        try {
                            const link = item.querySelector('a[href*="item.jd.com"]');
                            let sku = '';
                            if (link) {
                                const match = link.href.match(/item\\.jd\\.com\\/(\\d+)/);
                                if (match) sku = match[1];
                            }
                            
                            const nameEl = item.querySelector('.jGoodsInfo a');
                            let name = nameEl ? nameEl.innerText.trim() : '';
                            if (name && name.length > 60) name = name.substring(0, 60);
                            
                            let price = '';
                            const priceEl = item.querySelector('.jdNum');
                            if (priceEl) price = priceEl.innerText.trim() || '';
                            
                            if (!price) price = '待获取';
                            
                            if (sku && name) {
                                products.push({
                                    sku_id: sku,
                                    name: name,
                                    price: price,
                                    url: 'https://item.jd.com/' + sku + '.html'
                                });
                            }
                        } catch(e) {}
                    });
                    return products;
                }
            """)
            
            products.extend(page_products)
            
            has_price = sum(1 for p in page_products if p['price'] != '待获取')
            print(f"   ✅ 获取 {len(page_products)} 个商品, 有价格: {has_price}")
            
            # 保存进度
            with open(DEFAULT_OUTPUT_JSON, 'w', encoding='utf-8') as f:
                json.dump(products, f, ensure_ascii=False, indent=2)
            
            # 尝试点击下一页
            if page_num < max_pages:
                print(f"   👉 点击下一页...")
                success = click_next_page(page)
                if not success:
                    print(f"   ⚠️ 无法找到下一页按钮，可能已到最后一页")
                    break
                page.wait_for_timeout(3000)
        
        browser.close()
    
    return products


def export_excel(products: list):
    """导出Excel"""
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment
    except ImportError:
        print(f"{Colors.RED}❌ 请先安装 openpyxl: pip install openpyxl{Colors.END}")
        return False
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "商品列表"
    
    # 表头
    headers = ["SKU ID", "商品名称", "列表价", "商品链接"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    # 数据
    for row_idx, p in enumerate(products, 2):
        ws.cell(row=row_idx, column=1, value=p.get('sku_id', ''))
        ws.cell(row=row_idx, column=2, value=p.get('name', ''))
        ws.cell(row=row_idx, column=3, value=p.get('price', ''))
        ws.cell(row=row_idx, column=4, value=p.get('url', ''))
    
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 55
    
    wb.save(DEFAULT_OUTPUT_XLSX)
    return True


def show_existing_data():
    """查看已有数据"""
    if not DEFAULT_OUTPUT_JSON.exists():
        print(f"{Colors.YELLOW}⚠️ 暂无数据，请先爬取{Colors.END}")
        return
    
    with open(DEFAULT_OUTPUT_JSON, 'r', encoding='utf-8') as f:
        products = json.load(f)
    
    print(f"\n{Colors.BOLD}📊 已有数据:{Colors.END}")
    print(f"   商品数量: {len(products)}")
    
    has_price = sum(1 for p in products if p.get('price') != '待获取')
    print(f"   有价格: {has_price}")
    print(f"   待获取: {len(products) - has_price}")
    print()
    
    # 显示前5个
    print(f"{Colors.BOLD}前5个商品:{Colors.END}")
    for i, p in enumerate(products[:5], 1):
        print(f"  {i}. {p.get('name', '')[:40]}...")
        print(f"     💰 {p.get('price', '待获取')}")
    print()


def main():
    """主函数"""
    print_banner()
    
    while True:
        print_menu()
        choice = input(f"{Colors.GREEN}👉 请选择 (0-3): {Colors.END}").strip()
        print()
        
        if choice == '1':
            # 爬取
            url = get_shop_url()
            if not url:
                continue
            
            products = crawl_shop(url)
            
            # 导出Excel
            if products:
                export_excel(products)
                
                print()
                print(f"{Colors.GREEN}{'='*60}{Colors.END}")
                print(f"🎉 爬取完成!")
                print(f"{Colors.GREEN}{'='*60}{Colors.END}")
                print(f"   📦 商品总数: {len(products)}")
                has_price = sum(1 for p in products if p.get('price') != '待获取')
                print(f"   💰 有价格: {has_price}")
                print(f"   📁 输出文件:")
                print(f"      JSON: {DEFAULT_OUTPUT_JSON}")
                print(f"      Excel: {DEFAULT_OUTPUT_XLSX}")
                print()
        
        elif choice == '2':
            # 查看数据
            show_existing_data()
        
        elif choice == '3':
            # 导出Excel
            if DEFAULT_OUTPUT_JSON.exists():
                with open(DEFAULT_OUTPUT_JSON, 'r', encoding='utf-8') as f:
                    products = json.load(f)
                if export_excel(products):
                    print(f"{Colors.GREEN}✅ Excel导出成功: {DEFAULT_OUTPUT_XLSX}{Colors.END}")
            else:
                print(f"{Colors.YELLOW}⚠️ 暂无数据，请先爬取{Colors.END}")
            print()
        
        elif choice == '0':
            print(f"{Colors.CYAN}👋 再见！{Colors.END}")
            break
        
        else:
            print(f"{Colors.RED}❌ 无效选择，请重试{Colors.END}")
            print()


if __name__ == "__main__":
    main()

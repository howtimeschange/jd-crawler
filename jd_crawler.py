#!/usr/bin/env python3
"""
京东商品爬虫工具 - 交互式CLI版本 v1.2
功能：爬取京东店铺商品信息，支持自动翻页、Excel导出
特点：友好的交互界面，支持跨平台
"""

import sys
import os
import json
import re
import time
import subprocess
import platform
from pathlib import Path
from datetime import datetime

# 尝试导入playwright
try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("❌ 请先安装 playwright: pip install playwright")
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

# Windows 兼容
if platform.system() == 'Windows':
    Colors.END = ''
    Colors.GREEN = ''
    Colors.BLUE = ''
    Colors.YELLOW = ''
    Colors.RED = ''
    Colors.CYAN = ''
    Colors.BOLD = ''

def get_desktop_path() -> str:
    """获取桌面路径 - 跨平台"""
    home = Path.home()
    if platform.system() == 'Windows':
        return str(home / 'Desktop')
    elif platform.system() == 'Darwin':
        return str(home / 'Desktop')
    else:
        return str(home)

def print_banner():
    """打印欢迎界面"""
    print()
    print(f"{Colors.CYAN}{'='*60}{Colors.END}")
    print(f"{Colors.CYAN}  🛒  京东商品爬虫工具 v1.2{Colors.END}")
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
    """获取店铺URL"""
    print(f"{Colors.BOLD}📝 步骤1: 获取店铺商品列表链接{Colors.END}")
    print()
    print("请按以下步骤操作:")
    print(f"  {Colors.YELLOW}1.{Colors.END} 打开京东店铺首页")
    print(f"       例如: https://mall.jd.com/index-1000462158.html")
    print()
    print(f"  {Colors.YELLOW}2.{Colors.END} 点击店铺页面的 {Colors.CYAN}\"所有商品\"{Colors.END}")
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

def get_max_pages() -> int:
    """获取爬取页数"""
    print(f"{Colors.BOLD}📝 步骤2: 设置爬取页数{Colors.END}")
    print()
    print("注意:")
    print("  - 建议不超过20页（每页约15-20秒）")
    print("  - 页数越多，耗时越长")
    print()
    
    try:
        pages = input(f"{Colors.GREEN}📄 请输入要爬取的页数 (直接回车默认18页):{Colors.END} ").strip()
        if not pages:
            return 18
        return int(pages)
    except:
        return 18

def get_output_path() -> str:
    """获取输出路径"""
    desktop = get_desktop_path()
    print(f"\n{Colors.BOLD}📝 步骤3: 设置输出路径{Colors.END}")
    print()
    print(f"桌面路径: {desktop}")
    print()
    
    path = input(f"{Colors.GREEN}📁 直接回车导出到桌面，或输入自定义路径:{Colors.END} ").strip()
    
    if not path:
        return desktop
    
    return path

# 输出目录（临时存储）
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
DEFAULT_OUTPUT_JSON = OUTPUT_DIR / "jd_products.json"


def click_next_page(page) -> bool:
    """点击下一页按钮"""
    try:
        selectors = [
            'a.next',
            '.next',
            'a[class*="next"]',
            'a:has-text("下一页")',
            '#jTopPage a.next',
        ]
        
        for selector in selectors:
            try:
                next_btn = page.query_selector(selector)
                if next_btn and next_btn.is_visible():
                    cls = next_btn.get_attribute('class') or ''
                    if 'disabled' in cls:
                        return False
                    next_btn.click()
                    return True
            except:
                continue
        
        return False
    except:
        return False


def crawl_shop(url: str, max_pages: int) -> list:
    """爬取店铺商品"""
    print(f"\n{Colors.BOLD}🚀 开始爬取...{Colors.END}")
    print(f"   URL: {url[:50]}...")
    print(f"   页数: {max_pages}页")
    print()
    
    products = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=['--disable-blink-features=AutomationControlled'])
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
        )
        page = context.new_page()
        
        print(f"   🌐 打开第一页...")
        page.goto(url, wait_until="networkidle", timeout=60000)
        
        for page_num in range(1, max_pages + 1):
            print(f"\n{Colors.CYAN}{'='*50}{Colors.END}")
            print(f"📄 正在处理第 {page_num}/{max_pages} 页")
            
            print(f"   ⏳ 等待价格加载 (15秒)...")
            page.wait_for_timeout(15000)
            
            for _ in range(3):
                page.evaluate("window.scrollBy(0, 500)")
                page.wait_for_timeout(2000)
            
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
            
            with open(DEFAULT_OUTPUT_JSON, 'w', encoding='utf-8') as f:
                json.dump(products, f, ensure_ascii=False, indent=2)
            
            if page_num < max_pages:
                print(f"   👉 点击下一页...")
                success = click_next_page(page)
                if not success:
                    print(f"   ⚠️ 已到最后一页")
                    break
                page.wait_for_timeout(3000)
        
        browser.close()
    
    return products


def export_excel(products: list, output_path: str) -> str:
    """导出Excel到指定路径"""
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment
    except ImportError:
        print(f"{Colors.RED}❌ 请先安装 openpyxl: pip install openpyxl{Colors.END}")
        return None
    
    # 创建输出目录
    output_dir = Path(output_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 生成文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"京东商品_{timestamp}.xlsx"
    output_file = output_dir / filename
    
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
    
    wb.save(output_file)
    return str(output_file)


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
    print()
    
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
            
            max_pages = get_max_pages()
            output_path = get_output_path()
            
            products = crawl_shop(url, max_pages)
            
            if products:
                # 导出Excel到指定路径
                excel_file = export_excel(products, output_path)
                
                print()
                print(f"{Colors.GREEN}{'='*60}{Colors.END}")
                print(f"🎉 爬取完成!")
                print(f"{Colors.GREEN}{'='*60}{Colors.END}")
                print(f"   📦 商品总数: {len(products)}")
                has_price = sum(1 for p in products if p.get('price') != '待获取')
                print(f"   💰 有价格: {has_price}")
                if excel_file:
                    print(f"   📁 Excel: {excel_file}")
                print()
        
        elif choice == '2':
            show_existing_data()
        
        elif choice == '3':
            if DEFAULT_OUTPUT_JSON.exists():
                with open(DEFAULT_OUTPUT_JSON, 'r', encoding='utf-8') as f:
                    products = json.load(f)
                
                output_path = get_output_path()
                excel_file = export_excel(products, output_path)
                
                if excel_file:
                    print(f"{Colors.GREEN}✅ Excel导出成功: {excel_file}{Colors.END}")
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

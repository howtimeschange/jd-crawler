#!/usr/bin/env python3
"""
京东商品爬虫工具
功能：爬取京东店铺商品信息，支持自动翻页、价格加载、Excel导出
"""

import sys
import os
import json
import re
import time
from pathlib import Path
from datetime import datetime

# 确保可以导入playwright
try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("❌ 请先安装 playwright: pip install playwright")
    sys.exit(1)

# 输出目录
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

DEFAULT_OUTPUT_JSON = OUTPUT_DIR / "jd_products.json"
DEFAULT_OUTPUT_XLSX = OUTPUT_DIR / "jd_products.xlsx"


def print_progress(page_num: int, total: int, count: int, with_price: int):
    """打印进度"""
    print(f"  📄 第 {page_num}/{total} 页: 获取 {count} 个商品, 有价格: {with_price}")


def crawl_shop(url: str, output_json: Path = DEFAULT_OUTPUT_JSON):
    """爬取店铺商品"""
    print(f"\n🚀 开始爬取店铺商品...")
    print(f"   URL: {url}")
    print()
    
    products = []
    total_pages = 6  # 默认6页
    
    with sync_playwright() as p:
        # 启动浏览器（有头模式）
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
        )
        page = context.new_page()
        
        for page_num in range(1, total_pages + 1):
            print(f"\n{'='*50}")
            print(f"📄 正在处理第 {page_num}/{total_pages} 页")
            print(f"{'='*50}")
            
            # 构建页面URL
            if page_num == 1:
                page_url = url
            else:
                # 替换页码
                page_url = re.sub(r'-\d+-1-60\.html', f'-{page_num}-1-60.html', url)
                if page_url == url:
                    page_url = url.replace('.html', f'-{page_num}.html')
            
            print(f"   打开页面: {page_url[:60]}...")
            
            try:
                page.goto(page_url, wait_until="networkidle", timeout=60000)
            except Exception as e:
                print(f"   ⚠️ 页面加载超时，继续尝试...")
                page.goto(page_url, wait_until="domcontentloaded", timeout=30000)
            
            # 等待价格加载（关键步骤！）
            print("   ⏳ 等待价格加载 (15秒)...")
            page.wait_for_timeout(15000)
            
            # 滚动触发更多价格加载
            for i in range(3):
                page.evaluate("window.scrollBy(0, 500)")
                page.wait_for_timeout(2000)
            
            # 提取商品数据
            print("   🔍 提取商品数据...")
            
            page_products = page.evaluate("""
                () => {
                    const products = [];
                    const items = document.querySelectorAll('.jSubObject');
                    
                    items.forEach(item => {
                        try {
                            // SKU
                            const link = item.querySelector('a[href*="item.jd.com"]');
                            let sku = '';
                            if (link) {
                                const match = link.href.match(/item\\.jd\\.com\\/(\\d+)/);
                                if (match) sku = match[1];
                            }
                            
                            // 名称
                            const nameEl = item.querySelector('.jGoodsInfo a');
                            let name = nameEl ? nameEl.innerText.trim() : '';
                            
                            // 清理名称
                            if (name && name.length > 60) {
                                name = name.substring(0, 60);
                            }
                            
                            // 价格
                            let price = '';
                            const priceEl = item.querySelector('.jdNum');
                            if (priceEl) {
                                price = priceEl.innerText.trim() || '';
                            }
                            
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
            
            # 统计
            has_price = sum(1 for p in page_products if p['price'] != '待获取')
            print_progress(page_num, total_pages, len(page_products), has_price)
            
            # 保存进度
            with open(output_json, 'w', encoding='utf-8') as f:
                json.dump(products, f, ensure_ascii=False, indent=2)
            
            # 翻页间隔
            if page_num < total_pages:
                page.wait_for_timeout(2000)
        
        browser.close()
    
    return products


def export_excel(products: list, output_xlsx: Path = DEFAULT_OUTPUT_XLSX):
    """导出到Excel"""
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment
    except ImportError:
        print("⚠️ openpyxl未安装，跳过Excel导出")
        return None
    
    # 创建工作簿
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
    
    # 列宽
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 55
    
    wb.save(output_xlsx)
    return output_xlsx


def main():
    """主函数"""
    print("=" * 50)
    print("🛒 京东商品爬虫工具 v1.0")
    print("=" * 50)
    
    # 获取URL
    if len(sys.argv) > 1:
        url = sys.argv[1]
    else:
        print("\n📝 请输入店铺商品列表链接:")
        print("   示例: https://mall.jd.com/advance_search-2863474-1000462158-1000462158-0-0-0-1-1-60.html")
        url = input("\n🔗 > ").strip()
    
    if not url:
        print("❌ URL不能为空")
        sys.exit(1)
    
    # 爬取数据
    products = crawl_shop(url, DEFAULT_OUTPUT_JSON)
    
    # 导出Excel
    xlsx_file = export_excel(products)
    
    # 统计
    print(f"\n{'='*50}")
    print(f"🎉 爬取完成！")
    print(f"{'='*50}")
    print(f"   总商品数: {len(products)}")
    has_price = sum(1 for p in products if p['price'] != '待获取')
    print(f"   有价格: {has_price}")
    print(f"   待获取: {len(products) - has_price}")
    print(f"\n📁 输出文件:")
    print(f"   JSON: {DEFAULT_OUTPUT_JSON}")
    if xlsx_file:
        print(f"   Excel: {xlsx_file}")


if __name__ == "__main__":
    main()

# 🛒 京东商品爬虫工具

自动爬取京东店铺商品信息，支持一键安装、自动翻页、Excel导出。

## 功能特点

- ✅ 自动安装依赖（playwright）
- ✅ 自动翻页爬取（支持任意页数）
- ✅ 等待价格加载（解决京东异步加载问题）
- ✅ 支持Excel导出
- ✅ 有头模式（可观察爬取过程）

## 快速开始

### 一键安装运行

```bash
# 克隆项目
git clone https://github.com/xingyicheng/jd-crawler.git
cd jd-crawler

# 一键安装依赖并运行
chmod +x start.sh
./start.sh "https://mall.jd.com/advance_search-2863474-1000462158-1000462158-0-0-0-1-1-60.html"
```

## 使用方法

### 运行爬虫

```bash
# 带参数运行
python3 jd_crawler.py "店铺URL"

# 运行时输入
python3 jd_crawler.py
```

## 获取店铺URL

1. 打开京东店铺页面 (如 https://mall.jd.com/index-1000462158.html)
2. 点击"所有商品"
3. 复制浏览器地址栏URL

## 输出文件

- `jd_products.json` - 商品数据（JSON格式）
- `jd_products.xlsx` - 商品数据（Excel格式）

## 注意事项

1. **需要登录**：运行前请先在Chrome中登录京东
2. **网络代理**：如遇网络问题，请确保代理设置正确
3. **爬取时间**：每页约15-20秒

## 依赖

- Python 3.8+
- playwright
- openpyxl

## License

MIT

# 🛒 京东商品爬虫工具

自动爬取京东店铺商品信息，支持一键安装、友好交互界面、Excel导出、跨平台支持。

## 合作交流与赞助

欢迎联系作者邢易成，交流「京东商品爬虫」的使用与想法、探讨项目合作，或赞助支持项目持续开发。

使用微信扫描下方二维码添加好友，备注「京东商品爬虫 + 合作 / 交流 / 赞助」即可。此二维码与[抓虾官网](https://crawshrimp.com/)使用的联系方式一致。

<a href="https://crawshrimp.com/wechat-qr.jpg">
  <img src="https://crawshrimp.com/wechat-qr.jpg" alt="作者邢易成的微信联系二维码，用于京东商品爬虫合作交流与赞助联系" width="280" />
</a>

## 功能特点

- ✅ **友好交互界面**: 无需记命令，菜单式操作
- ✅ **自动翻页**: 点击下一页按钮，爬取全部页面
- ✅ **可配置页数**: 可自定义爬取页数（默认18页）
- ✅ **自定义导出路径**: 支持导出到桌面或指定目录
- ✅ **跨平台支持**: 支持 macOS、Windows、Linux
- ✅ **价格加载等待**: 解决京东异步加载问题

## 支持系统

- ✅ macOS
- ✅ Windows
- ✅ Linux

## 快速开始

```bash
# 克隆项目
git clone https://github.com/howtimeschange/jd-crawler.git
cd jd-crawler

# 运行
./start.sh
```

## 使用方法

### 1. 运行脚本

```bash
./start.sh
```

### 2. 选择功能

```
🛒 京东商品爬虫工具 v1.2
============================================================

📋 功能菜单:
  1. 爬取京东店铺商品
  2. 查看已有数据
  3. 导出Excel表格
  0. 退出
```

### 3. 爬取流程

1. **输入店铺URL**: 粘贴京东店铺商品列表链接
2. **设置页数**: 输入要爬取的页数（直接回车默认18页）
3. **设置输出路径**: 直接回车导出到桌面，或输入自定义路径

### 获取店铺URL

1. 打开京东店铺首页 (如 https://mall.jd.com/index-1000462158.html)
2. 点击"所有商品"
3. 复制浏览器地址栏URL

## 输出文件

Excel文件会保存到你指定的路径，默认保存到桌面：
- `京东商品_20240316_123456.xlsx`

## 常见问题

### Q: 提示 "externally-managed-environment"？
A: 使用 `--user` 参数：`pip install --user playwright`

### Q: Windows上运行报错？
A: 直接用 `python jd_crawler.py` 或双击运行

### Q: 爬取数量不对？
A: 确保已登录京东，且网络稳定

## 技术栈

- Python 3.8+
- Playwright (浏览器自动化)
- openpyxl (Excel导出)

## License

MIT

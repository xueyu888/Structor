<p align="center">
  <img src="docs/logo.png" width="160" alt="Structor logo"/>
</p>

<h1 align="center">Structor</h1>
<p align="center">
 🐼 A lightweight Excel templating & mapping tool for structured data cleaning and merging.
</p>
<p align="center">
  <em>“让脏乱的表格，变成结构整洁的模版。”</em>
</p>

---

## 🧩 What is Structor?

**Structor** 是一个支持图形界面的开源工具，帮助你将多个结构不一致的 Excel 文件，通过 **表头映射**、**模糊匹配** 和 **模板对齐**，统一整理成干净的数据模板。

非常适合：
- 银行流水、发票、交易记录等来源不一致的数据
- 批量清洗格式不统一的 Excel 报表
- 非技术用户通过 GUI 操作完成结构清洗

---

## ✨ Features

✅ 模板驱动的 Excel 合并  
✅ 支持 `.xlsx`、`.xls`、`.et`（自动 WPS 转换）  
✅ 表头模糊识别 + 多别名映射  
✅ 只保留命中 ≥1 列的有效表格  
✅ GUI 界面，无需编程  
✅ 生成统一结构的 `_filled.xlsx`

---

## 📦 Installation

推荐使用 [Poetry](https://python-poetry.org/) 管理：

```bash
git clone https://github.com/yourname/structor.git
cd structor
poetry install
poetry run python excel_template_merger.py

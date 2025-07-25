<p align="center">
  <img src="docs/logo.png" width="160" alt="Structor logo"/>
</p>

<h1 align="center">Structor</h1>
<p align="center">一个根据模板匹配列名并合并 Excel 的小工具</p>

---

## 简介

**Structor** 是一个带图形界面的工具，用来：

> 把多个结构不一致的 Excel 文件，按照一个模板的列顺序，对齐并合并成一个标准表格。

适用于银行流水、对账单、发票记录等格式不统一的文件批处理。

---

## 功能特点

- 支持 `.xlsx`、`.xls`、`.et` 文件（老 `.et` 会自动用 WPS 转换）
- 支持表头模糊匹配：列名不同也能智能对齐
- 支持表头映射配置：可自定义列名映射关系
- 合并输出结构和模板完全一致，生成 `_filled.xlsx`
- 图形界面，无需写代码

---

## 使用方法

### 第一步：准备 3 个文件类型

1. **模板文件**（必需）  
   定义输出表格的列顺序。

2. **表头映射文件**（推荐）  
   两列结构：模板表头 | 源表头  
   例如：

   | 模板表头 | 源表头         |
   |-----------|----------------|
   | 查询卡号  | 折/卡号/存单号 |
   | 金额      | 发生额         |
   | 日期      | 交易日期       |

3. **源数据文件**（多个）  
   来自不同来源、列名格式不一致的 Excel 文件。

---

### 第二步：打开程序开始处理

```bash
git clone https://github.com/yourname/structor.git
cd structor
poetry install
poetry run python excel_template_merger.py

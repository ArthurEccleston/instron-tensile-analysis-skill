---
name: instron-tensile-analysis
description: This skill should be used when the user asks to process Instron tensile test data, extract mechanical properties from .is_tens_Exports export folders, calculate 抗拉强度 (tensile strength) or 断裂伸长率 (elongation at break), batch process foil tensile test CSV files, or create a 力学性能 summary spreadsheet. Use this whenever the user mentions Instron tensile testing, 拉伸试验, 力学性能, 拉伸应力, 拉伸应变, 箔材拉伸, or wants to process material testing .is_tens_Exports export data.
---

# Instron 箔材拉伸数据分析

处理 Instron 拉伸试验机导出的 `.is_tens_Exports` 文件夹，提取抗拉强度和断裂伸长率，生成带统计汇总的 `力学性能_YYYYMMDD.xlsx`。

## 适用场景

- Instron 万能试验机导出的箔材/薄板拉伸数据
- 文件夹命名格式：`a-b.is_tens_Exports`（a = 样品种类编号，b = 样品编号）
- 文件夹内含 CSV 数据文件（GBK/UTF-8 编码）
- CSV 表头含 `拉伸应力`、`拉伸应变` 等列，第 1 行表头，第 2 行单位，第 3 行起数据

## 工作流程

### Step 1: 检查依赖

```bash
python -c "import openpyxl"
```

若未安装则执行 `pip install openpyxl`。

### Step 2: 运行分析脚本

```bash
python <skill-dir>/scripts/analyze_tensile.py "<数据文件夹路径>"
```

输出文件自动保存到数据文件夹内，文件名带日期后缀：`力学性能_YYYYMMDD.xlsx`，避免重复运行时覆盖旧结果。

### Step 3: 输出说明

输出 Excel 包含两个 Sheet：

**Sheet 1 — 原始数据：**
长格式（tidy data），每行一个样品，列：种类 | 编号 | 抗拉强度(MPa) | 断裂伸长率(%)。可直接用于绘图和透视分析。

**Sheet 2 — 统计汇总：**
每种类型一行，包含：原始样品数 | 有效样品数 | 抗拉强度均值±标准差 | 断裂伸长率均值±标准差 | 剔除异常值明细。使用 1.5×IQR 规则自动标记异常值（样品数 ≥ 4 时启用），异常值不计入统计。

### Step 4: 汇报结果

报告处理样品数、种类数、输出文件位置，以及是否有被剔除的异常值。

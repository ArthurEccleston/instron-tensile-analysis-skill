---
name: instron-tensile-analysis
description: This skill should be used when the user asks to process Instron tensile test data, extract mechanical properties from .is_tens_Exports export folders, calculate 抗拉强度 (tensile strength) or 断裂伸长率 (elongation at break), batch process foil tensile test CSV files, or create a 力学性能 summary spreadsheet. Use this whenever the user mentions Instron tensile testing, 拉伸试验, 力学性能, 拉伸应力, 拉伸应变, 箔材拉伸, or wants to process material testing .is_tens_Exports export data.
---

# Instron 箔材拉伸数据分析

处理 Instron 拉伸试验机导出的 `.is_tens_Exports` 文件夹，提取关键力学性能并汇总到 `力学性能.xlsx` 表格中。

## 适用场景

- Instron 万能试验机导出的箔材/薄板拉伸数据
- 文件夹命名格式：`a-b.is_tens_Exports`（a = 样品种类编号，b = 样品编号）
- 文件夹内含 CSV 数据文件（通常为 GBK 编码）
- CSV 表头含 `拉伸应力`、`拉伸应变` 等列，第 1 行表头，第 2 行单位，第 3 行起数据

## 工作流程

### Step 1: 检查依赖

```bash
python -c "import openpyxl"
```

若未安装则执行 `pip install openpyxl`。

### Step 2: 运行分析脚本

```bash
python <skill-dir>/scripts/analyze_tensile.py "<数据文件夹路径>" "<输出路径>/力学性能.xlsx"
```

脚本自动执行：

1. 扫描指定文件夹下所有 `a-b.is_tens_Exports` 目录
2. 读取每个目录内的 CSV 文件（自动识别 GBK/UTF-8 编码）
3. 从目录名解析样品种类 `a` 和样品编号 `b`
4. 对每个样品执行两项操作：

**操作 X — 抗拉强度：**
- 在第 1 行找到 `拉伸应力` 列
- 从第 3 行起读取该列所有数值
- 取最大值，写入 `力学性能.xlsx` 的 **第 a 列第 b 行**

**操作 V — 断裂伸长率：**
- 在第 1 行找到 `拉伸应变` 列
- 从第 3 行起计算相邻行差值 `strain[i+1] - strain[i]`
- 检测该差值是否为负且远低于前 20 个差值的趋势（4 个标准差以下），判定为断裂点
- 若无异常应变下降，取最后一个应变值作为回退
- 写入 `力学性能.xlsx` 的 **第 a 列第 b+10 行**

### Step 3: 汇报结果

报告处理文件数、输出位置、数据预览。

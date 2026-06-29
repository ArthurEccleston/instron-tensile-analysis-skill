# Instron 箔材拉伸数据分析 Skill

处理 Instron 拉伸试验机导出的 `.is_tens_Exports` 数据，自动提取**抗拉强度**和**断裂伸长率**，生成 `力学性能.xlsx` 汇总表。

## 安装

将 `SKILL.md` 和 `scripts/` 目录放入 `.claude/skills/instron-tensile-analysis/` 下即可。Claude Code 会自动发现该 skill。

## 依赖

- Python 3.x
- openpyxl：`pip install openpyxl`

## 数据要求

- 文件夹命名：`a-b.is_tens_Exports`（如 `1-2.is_tens_Exports`）
- 文件夹内含 CSV 文件，第 1 行表头，第 2 行单位，第 3 行起数据
- CSV 表头需包含 `拉伸应力` 和 `拉伸应变` 列
- 编码支持：UTF-8 / GBK / GB2312 / GB18030

## 使用方式

在 Claude Code 中直接对话触发：

> "帮我分析 C:\测试数据\20260617 文件夹里的拉伸数据"

Skill 自动触发后，会运行分析脚本并生成 `力学性能.xlsx`。

## 输出格式

| 位置 | 内容 |
|------|------|
| 第 a 列第 b 行 | 样品种类 a、编号 b 的抗拉强度 (MPa) |
| 第 a 列第 b+10 行 | 样品种类 a、编号 b 的断裂伸长率 (%) |

## 文件结构

```
instron-tensile-analysis/
├── SKILL.md                    # Skill 定义
├── scripts/
│   └── analyze_tensile.py      # 分析脚本
└── README.md
```

# EIS 批量处理工具：从原始数据到 Excel 汇总

批量读取 CHI760E 电化学工作站导出的 EIS（交流阻抗谱）.txt 文件，筛选相位角为负（电容性）的频段，输出为单个 Excel 工作簿（每样本一个 sheet），自动调整列宽。

## 项目结构

```
├── data_to_excel.py          # 主脚本：批量读取 → 过滤 → 输出 Excel
├── lesson_data_to_excel.py   # 教学拆解：逐行注释 + 知识点 + 练习
├── data/                     # 示例数据：36 个 EIS 阻抗谱原始 .txt 文件
├── output/                   # 运行输出（已在 .gitignore 中排除）
├── .gitignore
├── LICENSE
└── README.md
```

## 快速开始

```bash
pip install pandas openpyxl
python data_to_excel.py
```

输出文件：`output/EIS_Phase_Negative.xlsx`，每个样本一个 sheet。

## 数据说明

`data/` 中的 .txt 文件为 CHI760E 电化学工作站实测 EIS 数据。

**文件格式**：
- 前 19 行：仪器元数据（时间、型号、激励电压、频率范围等）
- 第 20 行起：CSV 表格，5 列

| 列名 | 含义 | 单位 |
|------|------|------|
| Freq/Hz | 激励频率 | Hz |
| Z'/ohm | 阻抗实部 | Ω |
| Z"/ohm | 阻抗虚部 | Ω |
| Z/ohm | 阻抗模 | Ω |
| Phase/deg | 相位角 | ° |

**过滤逻辑**：Phase < 0 表示电容性行为，保留；Phase ≥ 0 表示电感性行为，丢弃。

## 处理流程

```
data/*.txt
    │
    ├── read_csv(skiprows=19, header=None, names=[...])
    ├── dropna(how='all')
    ├── df[df["Phase/deg"] < 0]
    │
    └── pd.ExcelWriter → EIS_Phase_Negative.xlsx
        ├── Sheet: "100mah"    (31 字符内截断)
        ├── Sheet: "150mah"
        ├── ...
        └── 自动列宽调整 (openpyxl)
```

## 知识要点

`lesson_data_to_excel.py` 对每一行代码做了逐行拆解，涵盖：

- `glob.glob()` 批量文件发现
- `pd.read_csv()` 的 `skiprows` / `header=None` / `names` 三参数协作
- 布尔索引 + `.copy()` 避免 SettingWithCopyWarning
- `pd.ExcelWriter` 写入多 sheet
- `openpyxl` 底层 API 做自动列宽
- 异常处理、魔法数字、编码兼容性等工程实践

## License

MIT License — 见 [LICENSE](LICENSE) 文件

"""
=============================================================================
  教案：从 EIS 原始数据到 Excel 汇总 — data_to_excel.py 逐行拆解
  日期：2026-05-21
  📖 对应 PDF 章节：Ch5 (pandas 入门), Ch6 §6.1-6.2 (文件读写/Excel)

  教学目标
  ────────
  学完本课，你应该能：
    1. 用 glob + pd.read_csv 批量读取非标准格式文件
    2. 理解 skiprows / header / names 三个参数在 read_csv 中的协作
    3. 用布尔索引做条件过滤，并解释 .copy() 的必要性
    4. 用 pd.ExcelWriter 将多 DataFrame 写入同一工作簿的不同 sheet
    5. 用 openpyxl 底层 API 自动调整列宽

  先修要求
  ────────
  - 熟悉 DataFrame 基本操作 (loc/iloc, 列选择, dropna)
  - 了解 Python 标准库 os.path / glob
=============================================================================
"""

# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  第一阶段：业务理解 (Business Understanding)                              ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
#
# 背景：电化学工作站 (CHI760E) 导出的 EIS 数据为 .txt 格式，每个文件前 19 行
# 是仪器元数据（时间、型号、激励电压、频率范围等），第 20 行起才是 CSV 表格。
#
# 业务目标：从所有样本中提取"相位角为负"的频段（即电容性区域），汇总到一个
# Excel 工作簿，每个样本一个 sheet，便于后续做 Nyquist/Bode 图批量绘图。
#
# 列含义：
#   Freq/Hz   — 激励频率 (Hz)
#   Z'/ohm    — 阻抗实部 (Ω)
#   Z"/ohm    — 阻抗虚部 (Ω)
#   Z/ohm     — 阻抗模 (Ω)
#   Phase/deg — 相位角 (°)，负值表示电容性行为

import pandas as pd
import os
import glob

# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  第二阶段：数据理解 (Data Understanding)                                  ║
# ╚═══════════════════════════════════════════════════════════════════════════╝

# ── 知识点 1: 路径常量化 ──
# 把路径抽成常量而非散落在代码各处，好处：
#   - 换数据源只改一处
#   - IDE 自动补全 / 跳转
#   - 防止字符串拼写错误在运行时才暴露
ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, 'data')
OUTPUT = os.path.join(ROOT, 'output', 'EIS_Phase_Negative.xlsx')

# ── 知识点 2: 用列表定义列名 ──
# 因为原始文件没有 header 行（第 20 行直接是数据），我们需要手动命名。
# 注意 Z"/ohm 中的反斜杠是转义符 —— "Z\"/ohm" 等价于 Z"/ohm
COLUMNS = ["Freq/Hz", "Z'/ohm", "Z\"/ohm", "Z/ohm", "Phase/deg"]

# ── 知识点 3: glob.glob — 批量文件发现 ──
# glob 不是 regex！glob 语法远简单：
#   *     匹配任意字符（不含路径分隔符）
#   **    递归匹配子目录 (需 recursive=True)
#   ?     匹配单个字符
#   [abc] 匹配字符集中的一个
#
# os.path.join 负责跨平台拼接（Windows 用 \，Linux/Mac 用 /）
all_files = sorted(glob.glob(os.path.join(DATA_DIR, "*.txt")))
print(f"Found {len(all_files)} data files")

# ── 严师提问 1 ──
# sorted() 按字符串字典序排列。如果文件名含数字（如 "100mah"、"50mah"），
# 排序结果会是 "100mah" 在 "50mah" 之前（因为 '1' < '5'）。
# 如果需要按数值排序，应该用什么方法？

# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  第三阶段：数据准备 + 建模 (Data Preparation + Modeling)                  ║
# ╚═══════════════════════════════════════════════════════════════════════════╝

# ── 知识点 4: pd.ExcelWriter 上下文管理器 ──
# 使用 with 语句确保文件在写入完毕后正确关闭。
# engine='openpyxl' 是写 .xlsx 所必需的（xlwt 只能写老 .xls 格式）。
# 📖 PDF Ch6 §6.2: "Using openpyxl to write Excel files"
with pd.ExcelWriter(OUTPUT, engine='openpyxl') as writer:
    sheets_written = 0
    total_rows = 0

    # ── 知识点 5: 遍历 + 异常保护 ──
    # 真实数据里总会有坏文件（编码损坏、格式不一致、被其他程序锁定...）
    # try/except 加 continue 是批量处理的标准模式：跳过坏文件但不中断全局
    for fpath in all_files:
        fname = os.path.basename(fpath)

        # ── 知识点 6: Excel sheet 名限制 ──
        # Excel 规定 sheet 名最长 31 个字符。文件名可能更长，所以截断。
        # 注意：如果两个文件截断后重名，后面的会覆盖前面的 —— 这是潜在 bug
        sheet_name = fname.replace(".txt", "")[:31]

        try:
            # ── 知识点 7: read_csv 的三个关键参数 ──
            # skiprows=19  → 跳过前 19 行元数据，从第 20 行开始读
            # header=None  → 不要把第 20 行当列名，全部当数据
            # names=COLUMNS → 手动指定列名（与 header=None 搭配使用）
            # 📖 PDF Ch6 §6.1: "Reading Text Files in Pieces"
            df = pd.read_csv(
                fpath,
                skiprows=19,
                header=None,
                names=COLUMNS,
                encoding='utf-8',
            )
        except Exception as e:
            print(f"  SKIP {fname}: {e}")
            continue

        print(df.shape)
        print(df.head())
        print(df.dtypes)

        # ── 知识点 8: dropna(how='all') ──
        # 只删除"整行全是 NaN"的行。保留部分缺失的行供后续判断。
        # reset_index(drop=True) 重建连续索引，避免过滤后的索引空洞
        # 📖 PDF Ch7 §7.1: "Handling Missing Data"
        df = df.dropna(how='all').reset_index(drop=True)

        if df.empty:
            print(f"  EMPTY: {fname}")
            continue

        # ── 知识点 9: 布尔索引 (Boolean Indexing) ──
        # df["Phase/deg"] < 0   → 返回 bool Series
        # df[bool_series]       → 保留 True 的行
        # .copy()               → 显式复制，防止后续操作触发
        #                          SettingWithCopyWarning
        # 📖 PDF Ch5 §5.1: "Selecting with Boolean Arrays"
        df_neg = df[df["Phase/deg"] < 0].copy()

        if df_neg.empty:
            print(f"  NO NEGATIVE Phase in: {fname} (all {len(df)} rows)")
            # ── 设计决策：没有负相位时写入全部数据 ──
            df_neg = df

        # 只保留前 5 列（实际上 names 只有 5 列，此处的防御性写法）
        df_out = df_neg[COLUMNS]

        # ── 知识点 10: df.to_excel() ──
        # index=False → 不写入行号（Excel 自带行号，再写一遍是冗余）
        # 📖 PDF Ch6 §6.2
        df_out.to_excel(writer, sheet_name=sheet_name, index=False)

        # ── 知识点 11: openpyxl 底层操作 —— 自动列宽 ──
        # writer.sheets[sheet_name] 返回 openpyxl 的 Worksheet 对象
        ws = writer.sheets[sheet_name]
        for col_idx, col_name in enumerate(COLUMNS, 1):
            # 用 astype(str).str.len().max() 找该列最宽的单元格
            # max(列名宽度, 数据最大宽度) 保证表头不被截断
            max_width = max(
                len(col_name),
                df_out[col_name].astype(str).str.len().max() if not df_out.empty else 0
            )
            # column_letter: 1→A, 2→B, 3→C ...
            # 加 3 字符的 padding，且不超过 25（防止某列过宽）
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_width + 3, 25)

        sheets_written += 1
        total_rows += len(df_out)
        print(f"  {fname}: {len(df)} rows → {len(df_out)} rows (Phase < 0)")

print(f"\nDone: {sheets_written} sheets, {total_rows} total rows → {OUTPUT}")


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  第四阶段：严师代码审查 (Code Review)                                     ║
# ╚═══════════════════════════════════════════════════════════════════════════╝

# ── 问题 1: skiprows=19 是魔法数字 ──
# 为什么是 19？文件格式变了怎么办？CHI760E 的不同软件版本导出格式是否一致？
# 改进方案：
#   1) 先用 open() 读文件头，动态找 "Freq/Hz" 那一行，记录行号
#   2) 或者用 comment 参数，但目前 CSV 没有注释符

# ── 问题 2: sheet name 冲突 ──
# sheet_name = fname.replace(".txt", "")[:31]
# 如果两个文件名前 31 字符相同，后写的会静默覆盖先写的 sheet。
# 改进方案：用计数器加后缀，或检查重名并报 warning

# ── 问题 3: 没有负相位时写入全部数据 ──
# 这种行为隐式改变了输出含义：用户打开 Excel 看到某些 sheet 全正相位，
# 会以为是处理成功。应该在 sheet 名中标注（如 "xxx_NO_NEG"）或添加一个
# 标记列。

# ── 问题 4: 列宽计算过于保守 ──
# max_width + 3 的 padding 合理，但上限 25 个字符可能截断科学计数法的长数字。
# EIS 数据如 "9.6680000000e+3" 是 15 字符，在 25 内没问题，
# 但如果有更长的数据，建议把上限调为 35 或去掉上限。

# ── 问题 5: 缺少数据验证 ──
# 没有 assert df.shape[1] == len(COLUMNS)，如果某文件实际列数不对，
# read_csv 会报错（因为 names 数量不匹配），但报错信息对用户不友好。
# 改进：read_csv 成功后 assert 形状。

# ── 问题 6: encoding 硬编码 ──
# 只尝试 utf-8，如果文件是 gb2312 / latin-1 就直接跳过。
# 改进：用 try utf-8 → try gbk → try latin-1 的级联尝试，
# 或者用 chardet 库自动检测。

# ── 问题 7: 缺少样本量变化的记录 ──
# CRISP-DM 要求每一步记录样本量变化。当前仅 print 到控制台。
# 改进：写入一个 summary sheet，记录每个文件的 原始行数 → 清洗后行数 → 过滤后行数

# ── 问题 8: 向量化已经做到，但 .astype(str).str.len() 在列宽计算时重复 ──
# 对每列做了一次 astype(str).str.len().max()，数据量大时可能慢。
# 但考虑到 EIS 文件通常几百行，这完全不是瓶颈 —— 不必过度优化。


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  第五阶段：变式练习 (Exercises)                                           ║
# ╚═══════════════════════════════════════════════════════════════════════════╝

# ── 练习 1 (★ 基础) ──
# 修改过滤条件为 Phase < -5（即只保留强电容性区域），
# 并让输出文件名也反映这个阈值。

# ── 练习 2 (★★ 进阶) ──
# 将所有文件合并到一个 sheet，增加一列 "Sample" 记录来源文件名，形成
# 可用于 seaborn 分组绘图的长表格式 (tidy data)。

# ── 练习 3 (★★ 进阶) ──
# 将过滤条件从 Phase < 0 改为 Z"/ohm < 0（虚部为负同样是电容性的标志），
# 并对比两种过滤方法得到的结果是否一致。如果不一致，用 scatter plot
# 标出差异点。

# ── 练习 4 (★★★ 挑战) ──
# 将代码重构为函数式风格：
#   def read_eis_file(fpath) -> pd.DataFrame:
#   def filter_capacitive(df, phase_col="Phase/deg") -> pd.DataFrame:
#   def write_excel_report(file_dict, output_path) -> None:
# 要求：每个函数的 docstring 含参数说明、返回值、可能抛出的异常。
# 函数内部不能出现硬编码路径。


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  附录：本次涉及的全部 pandas / Python 知识点索引                          ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
#
#  Python 标准库
#  ├─ os.path.join()          路径拼接（跨平台）
#  ├─ os.path.basename()      从完整路径提取文件名
#  ├─ glob.glob()             通配符批量文件发现
#  └─ sorted()                内置排序
#
#  pandas I/O
#  ├─ pd.read_csv()           📖 PDF Ch6 §6.1
#  │   ├─ skiprows=19         跳过前 N 行
#  │   ├─ header=None         不将首行当列名
#  │   ├─ names=COLUMNS       手动指定列名
#  │   └─ encoding='utf-8'    指定文件编码
#  ├─ pd.ExcelWriter()        📖 PDF Ch6 §6.2
#  └─ df.to_excel()           写入 Excel
#
#  pandas 数据处理
#  ├─ df.dropna(how='all')    📖 PDF Ch7 §7.1
#  ├─ df[bool_series]         布尔索引 📖 PDF Ch5 §5.1
#  ├─ df.copy()               显式复制
#  └─ df.empty / df.dtypes    属性检查
#
#  openpyxl 底层
#  ├─ writer.sheets[name]     获取 Worksheet
#  ├─ ws.column_dimensions    列宽字典
#  └─ ws.cell().column_letter 数字→列字母
"""

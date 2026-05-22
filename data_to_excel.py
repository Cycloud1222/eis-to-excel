"""
读取data文件夹下所有 EIS 格式的文本文件，筛选相位小于 0的数据行，保留前五列数据
将每个文件的数据分别输出至同一个 Excel工作簿的不同工作表中
"""
import pandas as pd
import os
import glob

ROOT = os.path.dirname(os.path.abspath(__file__))#读取该文件目录下的所有文件（绝对路径）
DATA_DIR = os.path.join(ROOT, 'data')#读取文件名为data的数据
OUTPUT   = os.path.join(ROOT, 'output', 'EIS_Phase_Negative.xlsx')#创建输出excel文件

COLUMNS = ["Freq/Hz", "Z'/ohm", "Z\"/ohm", "Z/ohm", "Phase/deg"]#数据结构列名

all_files = sorted(glob.glob(os.path.join(DATA_DIR, "*.txt")))#批量遍历目标路径下的.txt文件
print(f"Found {len(all_files)} data files")#数据文件量

with pd.ExcelWriter(OUTPUT, engine='openpyxl') as writer:
    sheets_written = 0
    total_rows = 0

    for fpath in all_files:
        fname = os.path.basename(fpath)
        # Sheet name max 31 chars in Excel
        sheet_name = fname.replace(".txt", "")[:31]

        try:
            df = pd.read_csv(
                fpath,
                skiprows=19,#跳过设备参数及其他无用数据
                header=None,#请求头
                names=COLUMNS,#列名
                encoding='utf-8',
            )
        except Exception as e:
            print(f"  SKIP {fname}: {e}")
            continue
        #删除空值行how='all'：只有当一行的所有列都是空值（NaN）时，才会删除这一行。
        #对比一下：如果写成 how='any'，只要一行里有任何一个空值，就会把整行删掉
        df = df.dropna(how='all').reset_index(drop=True)

        if df.empty:
            print(f"  EMPTY: {fname}")
            continue

     
        df_neg = df[df["Phase/deg"] < 0].copy()

        if df_neg.empty:
            print(f"  NO NEGATIVE Phase in: {fname} (all {len(df)} rows)")
            # Still write the full data so they can see
            df_neg = df
            #空值检测，防写入excel报错

        df_out = df_neg[COLUMNS]

        df_out.to_excel(writer, sheet_name=sheet_name, index=False)
        #给工作簿命名

        ws = writer.sheets[sheet_name]
        for col_idx, col_name in enumerate(COLUMNS, 1):
            max_width = max(
                len(col_name),
                df_out[col_name].astype(str).str.len().max() if not df_out.empty else 0
            )
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_width + 3, 25)
            #设置列宽（min限制）
        
        
        sheets_written += 1
        total_rows += len(df_out)
        print(f"  {fname}: {len(df)} rows → {len(df_out)} rows (Phase < 0)")

print(f"\nDone: {sheets_written} sheets, {total_rows} total rows → {OUTPUT}")

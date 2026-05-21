"""
Read all EIS .txt files from data/, filter rows where Phase < 0,
keep first 5 columns, output each file to a separate sheet in one Excel workbook.
"""
import pandas as pd
import os
import glob

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, 'data')
OUTPUT   = os.path.join(ROOT, 'output', 'EIS_Phase_Negative.xlsx')

COLUMNS = ["Freq/Hz", "Z'/ohm", "Z\"/ohm", "Z/ohm", "Phase/deg"]

all_files = sorted(glob.glob(os.path.join(DATA_DIR, "*.txt")))
print(f"Found {len(all_files)} data files")

with pd.ExcelWriter(OUTPUT, engine='openpyxl') as writer:
    sheets_written = 0
    total_rows = 0

    for fpath in all_files:
        fname = os.path.basename(fpath)
        # Sheet name max 31 chars in Excel
        sheet_name = fname.replace(".txt", "")[:31]

        # Read CSV lines — skip 19 header lines, no real header row
        try:
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

        # Drop any completely empty rows
        df = df.dropna(how='all').reset_index(drop=True)

        if df.empty:
            print(f"  EMPTY: {fname}")
            continue

        # Filter: keep only Phase < 0
        df_neg = df[df["Phase/deg"] < 0].copy()

        if df_neg.empty:
            print(f"  NO NEGATIVE Phase in: {fname} (all {len(df)} rows)")
            # Still write the full data so they can see
            df_neg = df

        # Keep only first 5 columns
        df_out = df_neg[COLUMNS]

        # Write to sheet
        df_out.to_excel(writer, sheet_name=sheet_name, index=False)

        # Auto-adjust column widths
        ws = writer.sheets[sheet_name]
        for col_idx, col_name in enumerate(COLUMNS, 1):
            max_width = max(
                len(col_name),
                df_out[col_name].astype(str).str.len().max() if not df_out.empty else 0
            )
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_width + 3, 25)

        sheets_written += 1
        total_rows += len(df_out)
        print(f"  {fname}: {len(df)} rows → {len(df_out)} rows (Phase < 0)")

print(f"\nDone: {sheets_written} sheets, {total_rows} total rows → {OUTPUT}")

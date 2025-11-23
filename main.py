import pandas as pd
from tkinter import Tk, filedialog, Button, Entry, Label
from tkinterdnd2 import DND_FILES, TkinterDnD

def process_files(folder, delete_cols, filter_col, filter_val, split_col):
    # フォルダ内の Excel/CSV を読み込む
    import os
    all_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls', '.csv'))]
    df_list = []
    for file in all_files:
        if file.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        df_list.append(df)
    
    combined = pd.concat(df_list, ignore_index=True)
    
    # 列削除
    if delete_cols:
        cols_to_delete = [c.strip() for c in delete_cols.split(',')]
        combined.drop(columns=cols_to_delete, errors='ignore', inplace=True)
    
    # フィルタ
    if filter_col and filter_val:
        combined = combined[combined[filter_col] == filter_val]
    
    # Excel 出力
    output_file = 'formatted_output.xlsx'
    if split_col:
        with pd.ExcelWriter(output_file) as writer:
            for val, group in combined.groupby(split_col):
                group.to_excel(writer, sheet_name=str(val), index=False)
    else:
        combined.to_excel(output_file, index=False)

def run_gui():
    root = TkinterDnD.Tk()
    root.title("Excel Automation GUI Tool")

    Label(root, text="フォルダを選択").grid(row=0, column=0)
    folder_entry = Entry(root, width=50)
    folder_entry.grid(row=0, column=1)
    
    Button(root, text="選択", command=lambda: folder_entry.insert(0, filedialog.askdirectory())).grid(row=0, column=2)
    
    Label(root, text="削除する列 (カンマ区切り)").grid(row=1, column=0)
    delete_entry = Entry(root, width=50)
    delete_entry.grid(row=1, column=1)

    Label(root, text="フィルタ列").grid(row=2, column=0)
    filter_col_entry = Entry(root, width=50)
    filter_col_entry.grid(row=2, column=1)

    Label(root, text="フィルタ値").grid(row=3, column=0)
    filter_val_entry = Entry(root, width=50)
    filter_val_entry.grid(row=3, column=1)

    Label(root, text="シート分割列").grid(row=4, column=0)
    split_entry = Entry(root, width=50)
    split_entry.grid(row=4, column=1)

    Button(root, text="実行", command=lambda: process_files(
        folder_entry.get(),
        delete_entry.get(),
        filter_col_entry.get(),
        filter_val_entry.get(),
        split_entry.get()
    )).grid(row=5, column=1)

    root.mainloop()

if __name__ == "__main__":
    run_gui()

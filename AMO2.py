import pandas as pd
from openpyxl import load_workbook

# 定义文件路径和目标 sheet 名称
summary_filepath = "5. AMO\\Summary_AMO.xlsx"
output_sheetname = "data2"
target_sheetname = "data"

try:
    # 读取名为 "data" 的 sheet 数据
    df_data = pd.read_excel(summary_filepath, sheet_name=target_sheetname, usecols=['A', 'C'], nrows=3)
    
    # 检查 DataFrame 是否为空
    if df_data.empty:
        print(f"No data found in '{target_sheetname}' sheet.")
    else:
        # 写入新的 sheet
        with pd.ExcelWriter(summary_filepath, mode='a', engine='openpyxl') as writer:
            writer.book = load_workbook(summary_filepath)
            df_data.to_excel(writer, sheet_name=output_sheetname, index=False)
        
        print(f"Successfully copied columns A and C from '{target_sheetname}' sheet to '{output_sheetname}'.")
    
except FileNotFoundError:
    print(f"File '{summary_filepath}' not found.")
except KeyError:
    print(f"Sheet '{target_sheetname}' not found in '{summary_filepath}'.")
except Exception as e:
    print(f"An error occurred: {str(e)}")

import pandas as pd

# 定义文件路径和目标 sheet 名称
summary_filepath = "5. AMO\\Summary_AMO.xlsx"

target_sheetname1 = "Sheet1"
target_sheetname2 = "data1"

try:
    # 读取两个 sheet 的数据
    df1 = pd.read_excel(summary_filepath, sheet_name=target_sheetname1)
    df2 = pd.read_excel(summary_filepath, sheet_name=target_sheetname2)

    # 合并两个 DataFrame，假设根据 'SKU' 列进行合并
    merged_data = pd.merge(df1, df2, on='SKU', how='left')

    # 写入新的 sheet
    output_sheetname = "merged_data2"
    with pd.ExcelWriter(summary_filepath, mode='a', engine='openpyxl') as writer:
        merged_data.to_excel(writer, sheet_name=output_sheetname, index=False)

    print(f"Successfully merged data from '{target_sheetname1}' and '{target_sheetname2}' into '{output_sheetname}'.")

except FileNotFoundError:
    print(f"File '{summary_filepath}' not found.")
except Exception as e:
    print(f"An error occurred: {str(e)}")

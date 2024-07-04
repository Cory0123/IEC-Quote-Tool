import os
import pandas as pd

def get_unique_sheetname(writer, base_name):
    """Generate a unique sheet name by appending a number to the base name if it already exists."""
    existing_sheets = writer.book.sheetnames
    if base_name not in existing_sheets:
        return base_name
    i = 1
    while f"{base_name}{i}" in existing_sheets:
        i += 1
    return f"{base_name}{i}"

def process_and_append_data(input_filepath, summary_filepath, output_sheetname):
    # 读取输入 Excel 文件
    xl = pd.ExcelFile(input_filepath)
    
    # 打开 "SKU Summary" 的 sheet
    df = xl.parse("SKU Summary")
    
    # 从第二行开始抓取数据（pandas 默认是从 0 开始索引，所以要从第 1 行开始）
    data = df.iloc[1:, :4]
    
    # 过滤第四列的值为 "Adicora" 或 "HOOK 2.0" 的行
    filtered_data = data[(data.iloc[:, 3] == "Adicora") | (data.iloc[:, 3] == "HOOK 2.0")]
    
    # 检查 Summary_AMO.xlsx 是否存在
    if os.path.exists(summary_filepath):
        # 如果存在，打开已有文件并创建新的 sheet
        with pd.ExcelWriter(summary_filepath, mode='a', engine='openpyxl') as writer:
            unique_sheetname = get_unique_sheetname(writer, output_sheetname)
            filtered_data.to_excel(writer, sheet_name=unique_sheetname, index=False)
    else:
        # 如果不存在，报告错误
        raise FileNotFoundError(f"The file '{summary_filepath}' does not exist. Cannot append data.")
    
    print(f"Filtered data saved to {summary_filepath} in sheet '{output_sheetname}'.")
    

# 定义输入和输出文件路径
input_filepath = "5. AMO\\New folder\\Inventec CQ OPP & cNB & Docking July 2024 Initial Quote - Copy.xlsx"
summary_filepath  = "5. AMO\\Summary_AMO.xlsx"
output_sheetname = "data"

# 调用函数处理 Excel 文件并附加数据
process_and_append_data(input_filepath, summary_filepath, output_sheetname)

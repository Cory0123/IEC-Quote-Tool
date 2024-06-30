import os
import openpyxl

def update_master_data_sheets():
    # 定義要處理的目錄名稱
    directory_name = "Quote"
    current_directory = os.path.join(os.getcwd(), directory_name)
    
    for filename in os.listdir(current_directory):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(current_directory, filename)
            
            try:
                wb = openpyxl.load_workbook(file_path)
                
                # 忽略大小寫比較工作表名稱
                master_data_sheet = None
                for sheet_name in wb.sheetnames:
                    if sheet_name.lower() == 'master data':
                        master_data_sheet = wb[sheet_name]
                        break
                
                if master_data_sheet:
                    master_data_sheet['D1'] = 'Current Month'
                    wb.save(file_path)
                    print(f"Updated {filename}: 'Master Data' sheet D1 cell.")
                else:
                    print(f"Skipped {filename}: No 'Master Data' sheet found.")
            
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")

if __name__ == "__main__":
    update_master_data_sheets()

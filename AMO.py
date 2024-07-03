import os
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook, Workbook
from PyQt5.QtWidgets import QApplication, QFileDialog, QWidget, QMessageBox, QPushButton, QVBoxLayout

from datetime import datetime
from shutil import copyfile
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

class Get_QuoteForm(QWidget):
    def __init__(self, name='Get_QuoteForm'):
        super(Get_QuoteForm, self).__init__()
        self.setWindowTitle(name)
        self.resize(400, 150)   # 設定彈出視窗的大小

        # 按鈕1
        self.btn_chooseFile1 = QPushButton(self)  
        self.btn_chooseFile1.setObjectName("btn_chooseFile")  
        self.btn_chooseFile1.setText("[---Choose Quote file's Path---]")
        
        # 設定視窗的佈局
        layout = QVBoxLayout()
        layout.addWidget(self.btn_chooseFile1)
        self.setLayout(layout)

        # 設定按鈕的信號與槽函數連接
        self.btn_chooseFile1.clicked.connect(self.slot_btn_chooseFile1)
        
        self.file_paths = []

    def slot_btn_chooseFile1(self):
        global QuoteName_choose
        self.file_paths, _ = QFileDialog.getOpenFileNames(self,  
                                    "Choose Quote file's Path",  
                                    "", # start path
                                    "Excel File (*.xlsx *.xls *.xlsb);;All Files (*)")   

        if self.file_paths:
            self.close()


def find_summary_sheet(wb):
    for sheet_name in wb.sheetnames:
        if "summary" in sheet_name.lower():
            return wb[sheet_name]
    raise ValueError("No sheet containing 'summary' found in the workbook")


def process_quote_file(file_path):
    # Load the workbook from the provided file path
    wb = load_workbook(filename=file_path)
    
    # Usage example
    wb = load_workbook(filename=file_path)
    sheet = find_summary_sheet(wb)
    
    # Create a new workbook to paste the data
    new_wb = Workbook()
    new_sheet = new_wb.active
    
    # Iterate through all rows in the summary sheet and copy to new workbook
    for row in sheet.iter_rows(values_only=True):
        new_sheet.append(row)
    
    # Save the new workbook with today's date as the file name
    today_date = datetime.today().strftime('%Y-%m-%d')
    output_filename = f"Summary_{today_date}.xlsx"
    output_path = os.path.join(os.path.dirname(file_path), output_filename)
    new_wb.save(output_path)
    
    # Return the path where the new file is saved
    return output_path

def create_docking_av_sku():
    print("hello world")
    app = QApplication([])  # 初始化 PyQt5 應用程式
    quote_form = Get_QuoteForm()  # 創建取得報價表單的視窗
    quote_form.show()  # 顯示報價表單視窗
    app.exec_()  # 開始 PyQt5 的事件迴圈，等待使用者操作完畢後返回控制

    if quote_form.file_paths:
        for file_path in quote_form.file_paths:
            try:
                output_path = process_quote_file(file_path)
                show_message_box(f"Updated file saved to:\n{output_path}\n\nDone")
            except Exception as e:
                show_message_box(f"Error processing file {file_path}:\n{str(e)}")

def show_message_box(message):
    msg_box = QMessageBox()
    msg_box.setWindowTitle("Operation Complete")
    msg_box.setText(message)
    msg_box.exec_()

if __name__ == "__main__":
    create_docking_av_sku()

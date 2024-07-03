import os
import openpyxl
import pandas as pd
import numpy as np
import pyxlsb 
import os
import decimal
import xlwings as xw
import datetime
import re
import sys
from PyQt5.QtWidgets import *

def update_master_data_sheets():
      ## 2-1.function2-1: Let user chooses the Quote files ---Start-------------------------------------------------------------------------
class Get_QuoteForm(QWidget):
    def __init__(self, name = 'Get_QuoteForm'):
        super(Get_QuoteForm,self).__init__()
        self.setWindowTitle(name)
        self.resize(400,150)   # set the pop up widget's size

        # btn 1
        self.btn_chooseFile1 = QPushButton(self)  
        self.btn_chooseFile1.setObjectName("btn_chooseFile")  
        self.btn_chooseFile1.setText("[---Choose Quote file's Path---]")
        
        # set the widget's layout
        layout = QVBoxLayout()
        layout.addWidget(self.btn_chooseFile1)
        self.setLayout(layout)

        # set the widget's signal
        self.btn_chooseFile1.clicked.connect(self.slot_btn_chooseFile1)
    
    def slot_btn_chooseFile1(self):
        global QuoteName_choose
        QuoteName_choose, filetype = QFileDialog.getOpenFileNames(self,  
                                    "Choose Quote file's Path",  
                                    "", # start path
                                    "Excel File (*.xlsx *.xls *.xlsb);;All Files (*)")   

        if QuoteName_choose != "":
            self.close()

def get_quoteform():
    app = QApplication(sys.argv)
    get_quoteform = Get_QuoteForm('Get_QuoteForm')
    get_quoteform.show()
    try:
        sys.exit(app.exec_())
    except SystemExit:
        print('Closing Window...')
        
## 2-1.function2-1: Let user chooses the Quote files ---End-------------------------------------------------------------------------
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
                    master_data_sheet['G1'] = 'Month total price'
                    wb.save(file_path)
                    print(f"Updated {filename}: 'Master Data' sheet D1 cell.")
                else:
                    print(f"Skipped {filename}: No 'Master Data' sheet found.")
            
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")

if __name__ == "__main__":
    update_master_data_sheets()
    
 
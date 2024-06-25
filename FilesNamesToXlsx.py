import os
from openpyxl import Workbook

def list_files_to_excel(folder_path, excel_file_path):
    file_list = os.listdir(folder_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "ST36 New Format"

    ws.append(["File Name"])

    for file_name in file_list:
        ws.append([file_name])

    wb.save(excel_file_path)
    print(f"File list saved to {excel_file_path}")

folder_path = 'C:/Users/dasa/Documents/36_newFormat'
excel_file_path = 'C:/Users/dasa/Documents/36_newFormat.xlsx'

list_files_to_excel(folder_path, excel_file_path)

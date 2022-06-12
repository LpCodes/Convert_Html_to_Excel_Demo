import csv

import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

file = pd.read_html("./Test Report_2021-08-18_12-45-00.html")
path = "./your_csv_name.csv"
xlpath = 'name.xlsx'


def write_html_csv():
    for index, data in enumerate(file):
        # print(index)
        if index:
            # print("printing values for index " + str(index))
            # print(data)
            # print(type(data))
            data.to_csv("./your_csv_name.csv", mode='a+', header=True)

    wb = Workbook()
    ws = wb.active
    with open(path, 'r') as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save(xlpath)


def modify_excel():
    wb_obj = openpyxl.load_workbook(xlpath)

    sheet_obj = wb_obj.active

    rows = sheet_obj.max_row
    cols = sheet_obj.max_column

    print(rows, cols)

    for i in range(1, rows + 1):
        for j in range(1, cols + 1):
            if ("Test_Cases" in str(sheet_obj.cell(i, j).value)) or ("Status" in str(sheet_obj.cell(i, j).value)):
                x = sheet_obj.cell(i, j).coordinate
                y = sheet_obj.cell(i, j).row
                print(x)
                print(y)

                # sheet_obj.cell(i, j, value="Hello test This is a Replace Text")
                sheet_obj[x].font = Font(bold=True)

    wb_obj.save(xlpath)


print("Starting task one")
write_html_csv()
print("Task one over")
print("Starting task two")
modify_excel()
print("Task  two over")

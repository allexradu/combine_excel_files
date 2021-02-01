from openpyxl import Workbook
import csv
from os import listdir
from os.path import isfile, join
import platform
import os

excel_path = os.path.join(os.getcwd(), 'excel')
csv_path = os.path.join(os.getcwd(), 'csvs')
all_files = [f for f in listdir(csv_path) if isfile(join(csv_path, f))]

print(all_files)
for file in all_files:
    print(file)
    wb = Workbook()
    ws = wb.active
    with open(os.path.join(csv_path, file), 'r') as f:
        for row in csv.reader(f, delimiter = ';'):
            ws.append(row)

    wb.save(os.path.join(excel_path, file.split('.')[0] + '.xlsx'))

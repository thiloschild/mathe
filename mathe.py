import pandas as pd
import numpy as np
import xlrd
import openpyxl

csv = xlrd.open_workbook("exam.xlsx")
table = csv.sheet_by_name("table")

list1 = []

num_col = table.ncols

for x in range(1, num_col):
	list1.append(table.cell(1, x).value)

median = np.median(list1)
mean = np.mean(list1)

print(mean)




print(list1)

#############################################
#write to excel

wb = openpyxl.load_workbook("exam.xlsx")
sh = wb.get_sheet_by_name("sol")

sh['B1'] = 6

wb.save()
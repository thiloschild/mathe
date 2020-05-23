
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
variance = np.var(list1)
std = np.std(list1)
corrcoef = np.corrcoef(list1)

print(median)
print(mean)
print(variance)
print(std)
print(corrcoef)


#############################################
#write to excel

wb = openpyxl.load_workbook("exam.xlsx")
sh = wb["sol"]
dest_filename = "exam.xlsx"

cell = sh.cell(2,2)
cell.value = median
cell = sh.cell(3,2)
cell.value = mean
cell = sh.cell(4,2)
cell.value = variance
cell = sh.cell(5,2)
cell.value = std
cell = sh.cell(9,2)
cell.value = corrcoef


wb.save(filename=dest_filename)



import pandas as pd
import numpy as np
import xlrd

csv = xlrd.open_workbook("exam.xlsx")
table = csv.sheet_by_name("table")

list1 = []

num_col = table.ncols

for x in num_col:
	list1 = table.cell(1, x)


print(list1)

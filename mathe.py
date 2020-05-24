import numpy as np
import xlwings as xw
import pandas as pd



#import from the excel file
def get_data(file):

	wb.xw.Book(file)
	lastcol = countcol(2)
	lastrow = countrow(2)







#count the horizontal columns in the excel sheet
def countcol(row_num):
    lastcol = 0
    while True:
        cell_value = xw.Range((row_num,lastcol+1)).value
        if cell_value is not None:

            lastcol = lastcol + 1
        else:
            break
    return lastcol +1

#count the vertical columns in the excel sheet
def countrow(col_num):
    lastrow = 0
    while True:
        cell_value = xw.Range((lastrow+1,col_num)).value
        if cell_value is not None:

            lastrow = lastrow + 1
        else:
            break
    return lastrow +1


##################################################################################






#xw.Range("B17").value = a
import numpy as np
import xlwings as xw
import pandas as pd






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
#import from the excel file

xw.Book('4felder.xlsm')
lastcol = countcol(1)
lastrow = countrow(1)




#############
PvonA = xw.Range((lastrow,2)).value
PvonB = xw.Range((lastrow,3)).value
if(lastcol > 3):
    PvonC = xw.Range((lastrow,4)).value
#############
PvonX = xw.Range((2,lastcol)).value
PvonY = xw.Range((3,lastcol)).value
if(lastrow > 3):
    PvonZ = xw.Range((4,lastcol)).value
#############
P_AX = xw.Range((2,2)).value
P_AY = xw.Range((3,2)).value
if(lastrow > 3):
    P_AZ = xw.Range((4,2)).value
#############
P_BX = xw.Range((2,3)).value
P_BY = xw.Range((3,3)).value
if(lastrow > 3):
    P_BZ = xw.Range((4,3)).value
#############
P_CX = xw.Range((2,4)).value
P_CY = xw.Range((3,4)).value
if(lastrow > 3):
    P_CZ = xw.Range((4,4)).value
#############


list1 = [PvonA, PvonB, PvonC, PvonX, PvonY, PvonZ, P_AX, P_AY, P_AZ, P_BX, P_BY, P_BZ, P_CX, P_CY, P_CZ]
list_known = []

for x in list1:
    if x != None:
        list_known.append(x)



print(list_known)





#xw.Range("B17").value = a
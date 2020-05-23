import numpy as np
import xlwings as xw

def countcol(row_num):
    lastcol = 0
    while True:
        cell_value = xw.Range((row_num,lastcol+1)).value
        if cell_value is not None:

            lastcol = lastcol + 1
        else:
            break
    return lastcol +1


wb = xw.Book("2variable.xlsm")

num_col = countcol(2)
print(num_col)
list1 = []

for x in range(2, num_col):
    list1.append(xw.Range((2,x)).value)

print(list1)

median = np.median(list1)
mean = np.mean(list1)
variance = np.var(list1)
stda = np.std(list1)
varcoef = stda/np.abs(mean)

xw.Range("B5").value = median
xw.Range("B6").value = mean
xw.Range("B7").value = variance
xw.Range("B8").value = stda
xw.Range("B9").value = varcoef
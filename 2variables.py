import numpy as np
import xlwings as xw

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

def cova(x,y):
    meanxy = 0
    lenx = len(x)
    for i in range(lenx):
        meanxy += x[i] * y[i]
    cov = meanxy / lenx - (meanx * meany)
    return cov

def corrcoef(x,y):
    cov = cova(x,y)
    stdax = np.std(x)
    stday = np.std(y)
    corrcoef = cov / (stdax * stday)
    return corrcoef


wb = xw.Book("2variables.xlsm")

num_col = countcol(2)

x = []
y = []

for i in range(2, num_col):
    x.append(xw.Range((2,i)).value)

for i in range(2, num_col):
    y.append(xw.Range((3,i)).value)

medianx = np.median(x)
mediany = np.median(y)
meanx = np.mean(x)
meany = np.mean(y)
variancex = np.var(x)
variancey = np.var(y)
stdax = np.std(x)
stday = np.std(y)
varcoefx = stdax/np.abs(meanx)
varcoefy = stday/np.abs(meany)
cov = cova(x,y)
corrcoef = corrcoef(x,y)

xw.Range("B6").value = medianx
xw.Range("C6").value = mediany
xw.Range("B7").value = meanx
xw.Range("C7").value = meany
xw.Range("B8").value = variancex
xw.Range("C8").value = variancey
xw.Range("B9").value = stdax
xw.Range("C9").value = stday
xw.Range("B10").value = varcoefx
xw.Range("C10").value = varcoefy
xw.Range("B11").value = cov
xw.Range("B12").value = corrcoef

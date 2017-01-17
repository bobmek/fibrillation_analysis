# -*- coding: utf-8 -*-
"""
Created on Wed Jan 11 09:40:26 2017

@author: bobmek
"""

import openpyxl
import numpy as np
import pylab
from scipy.optimize import curve_fit
import string

def sigmoid(x, x0, k, a, c):
     y =c+(a / (1 + np.exp(-k*(x-x0))))
     return y

wb1 = openpyxl.load_workbook('datasample.xlsx')
s1=wb1.get_sheet_by_name('Sheet1')
s2=wb1.get_sheet_by_name('Sheet2')
writesheet=s2

RawData=np.array([[cell.value for cell in col] for col in s1['AM69':'AU107']])
TRawData=np.transpose(RawData)
time=TRawData[0]
numsamples=len(TRawData)
samples=TRawData[1:numsamples]
popt=np.zeros((numsamples-1, 4))
p0=([0,0,0,0])


for m in range(0, numsamples-1):

    ydata=samples[m]
    pcov=np.zeros((4, 4))


    popt[m], pcov = curve_fit(sigmoid, time, ydata, p0)
    x = np.linspace(1,46000,50)
    y = sigmoid(x, *popt[m])

    pylab.plot(time, ydata, 'o', label='data')
    pylab.plot(x,y, label='fit')
    pylab.ylim(0, 3500)
    pylab.legend(loc='best')
    pylab.show()


tot_tableshape=np.shape(popt)
alph=list(string.ascii_uppercase)

for i in range(tot_tableshape[0]):
    for j in range(tot_tableshape[1]):
        writesheet[alph[i]+str(j+1)] = popt[i, j]

wb1.save('analyzed.xlsx')
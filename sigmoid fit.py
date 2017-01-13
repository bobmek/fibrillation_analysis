# -*- coding: utf-8 -*-
"""
Created on Wed Jan 11 09:40:26 2017

@author: bobmek
"""

import openpyxl
import numpy as np
import pylab
from scipy.optimize import curve_fit

def sigmoid(x, x0, k, a, c):
     y =c+(a / (1 + np.exp(-k*(x-x0))))
     return y

wb1 = openpyxl.load_workbook('datasample.xlsx')
s1=wb1.get_sheet_by_name('Sheet1')
s2=wb1.get_sheet_by_name('Sheet2')

RawData=np.array([[cell.value for cell in col] for col in s1['D3':'X104']])
TRawData=np.transpose(RawData)
time=TRawData[0]
numsamples=len(TRawData)
samples=TRawData[1:numsamples]





xdata=time
ydata=samples[0]

#xdata = np.array([879.00, 1779.00, 2679.00, 3579.00, 4479.00, 5379.00, 6279.00, 7179.00, 8079.00, 8979.00, 9879.00, 10779.00, 11679.00,	12579.00,	13479.00,	14379.00,	15279.00,	16179.00,	17079.00,	17979.00,	18879.00,	19779.00,	20679.00,	21579.00,	22479.00,	23379.00,	24279.00,	25179.00,	26079.00,	26979.00,	27879.00,	28779.00,	29679.00,	30579.00,	31479.00,	32379.00,	33279.00,	34179.00,	35079.00,	35979.00,	36879.00,	37779.00,	38679.00,	39579.00,	40479.00,	41379.00,	42279.00,	43179.00])
#ydata = np.array([16,	15,	22,	8,	8,	12,	1,	53,	593,	1252,	1714,	1937,	1939,	1925,	1913,	1933,1922,	1877,	1897,	1910,	1857,	1839,	1840,	1795,	1774,	1753,	1730,	1768,	1708,	1697,	1677,	1659,	1644,	1612,	1571,	1573,	1562,	1576,	1533,	1517,	1550,	1514,	1539,	1509,	1522,	1536,	1483,	1517])
#ydatapercent = np.array([0.008251676,	0.007735946,	0.011346055,	0.004125838,	0.004125838,	0.006188757,	0.00051573,	0.027333677,	0.305827746,	0.645693657,	0.883960805,	0.99896854,	1,	0.992779783,	0.986591026,	0.996905621,	0.991232594,	0.968024755,	0.97833935,	0.985043837,	0.95771016,	0.948427024,	0.948942754,	0.925734915,	0.91490459,	0.904074265,	0.892212481,	0.911810211,	0.880866426,	0.875193399,	0.864878804,	0.855595668,	0.847859722,	0.831356369,	0.810211449,	0.811242909,	0.805569881,	0.812790098,	0.790613718,	0.782362042,	0.799381124,	0.780814853,	0.793708097,	0.778236204,	0.784940691,	0.792160908,	0.764827231,	0.782362042])
p0=([0,0,0,0])

popt, pcov = curve_fit(sigmoid, xdata[0:22], ydata[0:22], p0)
print "Fit:"
print "x0 =", popt[0]
print "k  =", popt[1]
print "a  =", popt[2]
print "c  =", popt[3]
print "Asymptotes are", popt[3], "and", popt[3] + popt[2]

#print xdata
#print ydata


x = np.linspace(1,46000,50)
y = sigmoid(x, *popt)

pylab.plot(xdata, ydata, 'o', label='data')
pylab.plot(x,y, label='fit')
pylab.ylim(0, 2000)
pylab.legend(loc='best')
pylab.show()
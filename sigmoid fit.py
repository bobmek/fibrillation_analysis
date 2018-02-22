"""
Created on Wed Jan 11 09:40:26 2017

@author: bobmek
"""

import openpyxl
import numpy as np
import pylab
from scipy.optimize import curve_fit
import xlsxwriter
#import panda as pd

#sigmoid funtion 
def sigmoid(x, x0, k, a, c):
     y =c+(a / (1 + np.exp(-k*(x-x0))))
     return y

#commands to open notebook for I/O from Excel
wb1 = openpyxl.load_workbook('rawdata.xlsx')
s1=wb1.get_sheet_by_name('Sheet1')
s2=wb1.get_sheet_by_name('Sheet2')
writesheet=s2

workbook = xlsxwriter.Workbook('Testdata.xlsx')
worksheet = workbook.add_worksheet()


#Actually reading the data in the spreadsheet and transposing it 
RawData=np.array([[cell.value for cell in col] for col in s1['A3':'CS300']])
Analogs=np.array([[cell.value for cell in col] for col in s1['B1':'CS1']])
TRawData=np.transpose(RawData)
time=TRawData[0]
numsamples=len(TRawData)
samples=TRawData[1:numsamples]
popt=np.zeros((numsamples-1, 4))
p0=([0,0,0,0])

def my_range(start, end, step):
    while start <= end:
        yield start
        start += step

#Going throught the "Columns" in the spreadsheet and fitting them each to a sigmoid

for m in my_range(0, numsamples-1,5):

    ydata=samples[m:m+4]
    pcov=np.zeros((4, 4))
    ydatalength=len(ydata)

    #p0[2]=max(ydata[0:100])
    #print(p0)
    try:
        popt[m], pcov = curve_fit(sigmoid, time, ydata, p0)
        x = np.linspace(1,150000,2000)
        y = sigmoid(x, *popt[m])

        pylab.plot(time, ydata, '+', label='Raw Data')
        pylab.plot(x,y, label='Sigmoid Fit')
        pylab.plot(popt[m,0], popt[m,2]/2, 'o', label='EC50')
        #pylab.ylim(0, 25000)
        pylab.legend(loc='best')
        pylab.title(Analogs[0,m])
        count=str(m)
#        savefig = Analogs[0,m] + '.png'
#    
#        pylab.savefig(savefig)
#    
        pylab.show()
    except:
#        for n in range (0,len(time)):
            try:
                ptry=([0,0,2500,0])
                popt[m], pcov = curve_fit(sigmoid, time[0:len(time)/2], ydata, p0)
                #print(n)
            except:
                print("nope")
                pass
#                try:
#                    p1=()
    print(m)
                

FinishedData=np.empty([numsamples-1], dtype=[('analog_id', 'U16'), ('EC50','f4'), ('k-value', 'f4'),('amplitude','f4'),('baseline','f4') ])

FinishedData['EC50']=popt[:,0]
FinishedData['k-value']=popt[:,1]
FinishedData['amplitude']=popt[:,2]
FinishedData['baseline']=popt[:,3]
FinishedData['analog_id']=Analogs
#na_csv_output = np.zeros((len(FinishedData),), dtype=('U16,f4,f4,f4,f4'))
#df=pd.DataFrame(na_csv_output)

np.savetxt('Finished analyzed.txt', FinishedData, delimiter=",", fmt='%r %f %f %f %f')

worksheet.write('A1', 'Analog ID')
worksheet.write('B1', 'EC50')
worksheet.write('C1', 'K-Value')
worksheet.write('D1', 'Amplitude')
worksheet.write('E1', 'Baseline')
row = 1 
col = 0

for m in range(0, numsamples-1):
    worksheet.write (row, col, FinishedData[m][col])
    worksheet.write (row, col + 1 , FinishedData[m][col+1])
    worksheet.write (row, col + 2 , FinishedData[m][col+2])
    worksheet.write (row, col + 3 , FinishedData[m][col+3])
    worksheet.write (row, col + 4 , FinishedData[m][col+4])   
    row +=1


workbook.close()
#Actually write the analyzed data into the analyzed.xls spreadsheet
#for i in range(tot_tableshape[0]):
#    for j in range(tot_tableshape[1]):
#        writesheet[alph[i]+str(j+1)] = popt[i, j]
#
#wb1.save('analyzed.xlsx')
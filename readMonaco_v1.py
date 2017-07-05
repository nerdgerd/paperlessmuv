__author__ = 'Gerd'
import numpy as np
from numpy import genfromtxt
from Tkinter import *
from tkFileDialog import askopenfilename
import ttk
from collections import deque
from itertools import islice
import pandas as pd
import csv
import datetime
import xlwt


# Find data point in DVH closest to DVx
def find_nearest(array, value):
    idx = (np.abs(array-value)).argmin()
    return array[idx]

# Start - General data
Date = datetime.date.today()

prescribed_dose = 54                # get from DICOM in the future????

with open('MonacoTest.csv') as f:
    reader = csv.reader(f)
    DVH = list(reader)


my_data = genfromtxt('MonacoTest.csv', delimiter=',',skip_header=3,skip_footer=2)

# making numpy of doses and volumes
#d = np.asmatrix(my_data)

#dose = d[:,2]
#volume = d[:,1]

# extract names of all occurring OARs
names = [i[0] for i in DVH[3:-3]]
dose = [i[1]for i in DVH[3:-3]]
volume = [i[2] for i in DVH[3:-3]]


DVH_df = pd.DataFrame(
    {'OAR':         names,
     'Dose' :      dose,
     'Volume' :    volume
     })

# List of structures in the DVH
nameslist = list(set(names))
n = len(nameslist)

value1 = nameslist[6]

DVX1 = 50


#print value1


## Find volume that receives x dose
def find_DVH_Vx(DVX,OARX):
    dd = DVH_df.ix[(DVH_df['OAR']==OARX)]

    doseOAR1 = dd['Dose'].tolist()
    doseOAR1_array = np.array(doseOAR1,dtype=float)
    DX1 =find_nearest(doseOAR1_array, DVX)
    index_OAR1 = np.where(doseOAR1_array ==DX1)

    DVX_volume_list = dd['Volume'].tolist()
    DVX_volume_array = np.array(DVX_volume_list,dtype=float)
    DVX_volume = DVX_volume_array[index_OAR1[0]]
    return DVX_volume

# Find dose that is received by certain volume x
def find_DVH_Dx(DVX,OARX):
    dd = DVH_df.ix[(DVH_df['OAR']==OARX)]

    DVX_volume_list = dd['Volume'].tolist()
    DVX_volume_array = np.array(DVX_volume_list,dtype=float)
    DX1 =find_nearest(DVX_volume_array, DVX)
    index_OAR1 = np.where(DVX_volume_array == DX1)

    doseOAR1 = dd['Dose'].tolist()
    doseOAR1_array = np.array(doseOAR1,dtype=float)
    DVX_doseall = doseOAR1_array[index_OAR1[0]]
    DVX_dose = DVX_doseall[-1]
    return DVX_dose

# find max dose
def find_Dmax(OARX):
    dd = DVH_df.ix[(DVH_df['OAR']==OARX)]

    DVX_volume_list = dd['Volume'].tolist()
    DVX_volume_array = np.array(DVX_volume_list,dtype=float)
    DX1 =np.where(DVX_volume_array>0,DVX_volume_array,DVX_volume_array.max()).min()     # find non-zero minimum volume
    index_OAR1 = np.where(DVX_volume_array == DX1)
    doseOAR1 = dd['Dose'].tolist()
    doseOAR1_array = np.array(doseOAR1,dtype=float)
    Dmaxall = doseOAR1_array[index_OAR1[0]]
    Dmax = Dmaxall[-1]  # just return the largest Dmax (in case of multiples)
    return Dmax

# find min dose
def find_Dmin(OARX):
    dd = DVH_df.ix[(DVH_df['OAR']==OARX)]
    doseOAR1 = dd['Dose'].tolist()
    doseOAR1_array = np.array(doseOAR1,dtype=float)
    Dmin = np.where(doseOAR1_array>0,doseOAR1_array,doseOAR1_array.max()).min()
    return Dmin

# find mean dose
def find_Dmean(OARX):
    dd = DVH_df.ix[(DVH_df['OAR']==OARX)]
    doseOAR1 = dd['Dose'].tolist()
    doseOAR1_array = np.array(doseOAR1,dtype=float)
    Dmean = np.mean(doseOAR1_array)
    return Dmean

book = xlwt.Workbook()
sh = book.add_sheet('sheet')
r = 1

for i in range(0,n):
    OAR1 = nameslist[i]

    if 'nerve' and 'Optic' in OAR1:
        optic_nerve_Dmax = find_Dmax(OAR1)
        print OAR1.strip(),': Dmax =',optic_nerve_Dmax
        sh.write(r,0,OAR1)
        sh.write(r,1,'Dmax')
        sh.write(r,2,optic_nerve_Dmax)
        sh.write(r,3,'Gy')
        r += 1

    elif 'PTV-2mm' in OAR1:
        PTV_V95 = find_DVH_Vx(prescribed_dose*.95, OAR1)
        PTV_Dmax =  find_Dmax(OAR1)
        PTV_Dmin =  find_Dmin(OAR1)
        PTV_Dmean =  find_Dmean(OAR1)
        sh.write(r,0,OAR1)
        sh.write(r,1,'V95')
        sh.write(r,2,PTV_V95[0])
        sh.write(r,3,'%')
        r += 1
        sh.write(r,1,'Dmax')
        sh.write(r,2,PTV_Dmax)
        sh.write(r,3,'Gy')
        r += 1
        sh.write(r,1,'Dmin')
        sh.write(r,2,PTV_Dmin)
        sh.write(r,3,'Gy')
        r += 1
        sh.write(r,1,'Dmean')
        sh.write(r,2,PTV_Dmean)
        sh.write(r,3,'Gy')
        r += 1

    elif 'BrainStemPRV' in OAR1:
        brainstem_Dmax = find_Dmax(OAR1)
        brainstem_V40 = find_DVH_Vx(40,OAR1)
        sh.write(r,0,'brainstem')
        sh.write(r,1,'Dmax')
        sh.write(r,2,brainstem_Dmax)
        sh.write(r,3,'Gy')
        r += 1
        sh.write(r,1,'V40')
        sh.write(r,2,brainstem_V40[0])
        sh.write(r,3,'%')
        r += 1

eee = find_DVH_Vx(DVX1,value1)
fff = find_DVH_Dx(DVX1,value1)

#print value1.strip(),': ',eee,'% of the volume get',DVX1,'Gy'
#print value1.strip(),': ',DVX1,'% of the volume get',fff,'Gy'


col1_name = 'Structure'
col2_name = 'Parameter'
col3_name = 'value'

sh.write(0, 0, col1_name)
sh.write(0, 1, col2_name)
sh.write(0, 2, col3_name)



book.save('test1.xls')




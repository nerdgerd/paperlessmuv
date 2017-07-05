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
def find_nearest(array,value):
    idx = (np.abs(array-value)).argmin()
    return array[idx]

# Start - General data
Date = datetime.date.today()

prescribed_dose1 = 50                # get from DICOM in the future????
prescribed_dose2 = 74

filename = askopenfilename()

with open(filename) as f:
    reader = csv.reader(f)
    DVH = list(reader)

my_data = genfromtxt(filename, delimiter=',',skip_header=3,skip_footer=2)

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
    Dmin = np.where(doseOAR1_array>0,doseOAR1_array,doseOAR1_array.max()).min()     # find non-zero minimum volume
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
xlwt.add_palette_colour("custom_colour", 0x21)
book.set_colour_RGB(0x21, 251, 200, 200)
style1 = xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour')
style2 = xlwt.easyxf('font: bold on')

first_col = sh.col(0)
first_col.width = 256 * 20

second_col = sh.col(1)
second_col.width = 256 * 12

third_col = sh.col(2)
third_col.width = 256 * 12

fourth_col = sh.col(3)
fourth_col.width = 256 * 12

fifth_col = sh.col(4)
fifth_col.width = 256 * 12


r = 1

for i in range(0,n):
    OAR1 = nameslist[i]

    if 'Darm' in OAR1:
        darm_Dmax = find_Dmax(OAR1)
        darm_V50 = find_DVH_Vx(50, OAR1)
        darm_V45 = find_DVH_Vx(45, OAR1)
        darm_V40 = find_DVH_Vx(40, OAR1)
        print OAR1.strip(), ': Dmax =',darm_Dmax
        sh.write(r, 0, OAR1)
        sh.write(r, 1, 'Dmax')
        sh.write(r, 2, '56 Gy')
        sh.write(r, 3, darm_Dmax)
        sh.write(r, 4, 'Gy')
        if darm_Dmax > 56:
            sh.write(r, 5, 'exceeding', style1)
        r += 1
        sh.write(r, 1, 'V50')
        sh.write(r, 2, '10 %')
        sh.write(r, 3, darm_V50[0])
        sh.write(r, 4, '%')
        if darm_V50[0] > 10:
            sh.write(r, 5, 'exceeding', style1)
        r += 1
        sh.write(r,1,'V45')
        sh.write(r,2,'15 %')
        sh.write(r,3,darm_V45[0])
        sh.write(r,4,'%')
        if darm_V45[0] > 15:
            sh.write(r,5,'exceeding',style1)
        r += 1
        sh.write(r,1,'V40')
        sh.write(r,2,'20 %')
        sh.write(r,3,darm_V40[0])
        sh.write(r,4,'%')
        if darm_V40[0] > 20:
            sh.write(r,5,'exceeding',style1)
        r += 1

    if 'Blase' in OAR1:
        blase_Dmax = find_Dmax(OAR1)
        blase_V65 = find_DVH_Vx(65,OAR1)
        blase_V55 = find_DVH_Vx(55,OAR1)
        blase_V50 = find_DVH_Vx(50,OAR1)
        blase_V35 = find_DVH_Vx(35,OAR1)
        print OAR1.strip(),': Dmax =',blase_Dmax
        sh.write(r,0,OAR1)
        sh.write(r,1,'Dmax')
        sh.write(r,2,'79 Gy')
        sh.write(r,3,blase_Dmax)
        sh.write(r,4,'Gy')
        if blase_Dmax > 79:
            sh.write(r,5,'exceeding',style1)
        r += 1
        sh.write(r,1,'V65')
        sh.write(r,2,'20 %')
        sh.write(r,3,blase_V65[0])
        sh.write(r,4,'%')
        if blase_V65[0] > 20:
            sh.write(r,5,'exceeding',style1)
        r += 1
        sh.write(r,1,'V55')
        sh.write(r,2,'40 %')
        sh.write(r,3,blase_V55[0])
        sh.write(r,4,'%')
        if blase_V55[0] > 40:
            sh.write(r,5,'exceeding',style1)
        r += 1
        sh.write(r,1,'V50')
        sh.write(r,2,'50 %')
        sh.write(r,3,blase_V50[0])
        sh.write(r,4,'%')
        if blase_V50[0] > 50:
            sh.write(r,5,'exceeding',style1)
        r += 1
        sh.write(r,1,'V35')
        sh.write(r,2,'80 %')
        sh.write(r,3,blase_V35[0])
        sh.write(r,4,'%')
        if blase_V35[0] > 80:
            sh.write(r,5,'exceeding',style1)
        r += 1

    if 'Rektum' in OAR1:
        rektum_Dmax = find_Dmax(OAR1)
        rektum_V65 = find_DVH_Vx(65,OAR1)
        rektum_V60 = find_DVH_Vx(60,OAR1)
        rektum_V55 = find_DVH_Vx(55,OAR1)
        rektum_V50 = find_DVH_Vx(50,OAR1)
        print OAR1.strip(),': Dmax =',rektum_Dmax
        sh.write(r,0,OAR1)
        sh.write(r,1,'Dmax')
        sh.write(r,2,'79')
        sh.write(r,3,rektum_Dmax)
        sh.write(r,4,'Gy')
        if rektum_Dmax > 79:
            sh.write(r,5,'exceeding',style1)
        r += 1
        sh.write(r,1,'V65')
        sh.write(r,2,'20 %')
        sh.write(r,3,rektum_V65[0])
        sh.write(r,4,'%')
        if rektum_V65[0] > 20:
            sh.write(r,5,'exceeding',style1)
        r += 1
        sh.write(r,1,'V60')
        sh.write(r,2,'40 %')
        sh.write(r,3,rektum_V60[0])
        sh.write(r,4,'%')
        if rektum_V65[0] > 40:
            sh.write(r,5,'exceeding',style1)
        r += 1
        sh.write(r,1,'V55')
        sh.write(r,2,'45 %')
        sh.write(r,3,rektum_V55[0])
        sh.write(r,4,'%')
        if rektum_V65[0] > 45:
            sh.write(r,5,'exceeding',style1)
        r += 1
        sh.write(r,1,'V50')
        sh.write(r,2,'50 %')
        sh.write(r,3,rektum_V50[0])
        sh.write(r,4,'%')
        if rektum_V65[0] > 50:
            sh.write(r,5,'exceeding',style1)
        r += 1

    elif 'PTV_Becken' in OAR1:
        PTV_V95 = find_DVH_Vx(prescribed_dose1*.95, OAR1)
        PTV_Dmax =  find_Dmax(OAR1)
        PTV_Dmin =  find_Dmin(OAR1)
        PTV_Dmean =  find_Dmean(OAR1)
        sh.write(r,0,OAR1)
        sh.write(r,1,'V95')
        sh.write(r,3,PTV_V95[0])
        sh.write(r,4,'%')
        r += 1
        sh.write(r,1,'Dmax')
        sh.write(r,3,PTV_Dmax)
        sh.write(r,4,'Gy')
        r += 1
        sh.write(r,1,'Dmin')
        sh.write(r,3,PTV_Dmin)
        sh.write(r,4,'Gy')
        r += 1
        sh.write(r,1,'Dmean')
        sh.write(r,3,PTV_Dmean)
        sh.write(r,4,'Gy')
        r += 1

    elif 'PTV_Prostata' in OAR1:
        PTV_V95 = find_DVH_Vx(prescribed_dose2*.95, OAR1)
        PTV_Dmax =  find_Dmax(OAR1)
        PTV_Dmin =  find_Dmin(OAR1)
        PTV_Dmean =  find_Dmean(OAR1)
        sh.write(r,0,OAR1)
        sh.write(r,1,'V95')
        sh.write(r,3,PTV_V95[0])
        sh.write(r,4,'%')
        r += 1
        sh.write(r,1,'Dmax')
        sh.write(r,3,PTV_Dmax)
        sh.write(r,4,'Gy')
        r += 1
        sh.write(r,1,'Dmin')
        sh.write(r,3,PTV_Dmin)
        sh.write(r,4,'Gy')
        r += 1
        sh.write(r,1,'Dmean')
        sh.write(r,3,PTV_Dmean)
        sh.write(r,4,'Gy')
        r += 1

    if 'Hueftkopf' in OAR1:
        femoral_Dmax = find_Dmax(OAR1)
        femoral_V45 = find_DVH_Vx(45,OAR1)
        print OAR1.strip(),': Dmax =',femoral_Dmax
        sh.write(r,0,OAR1)
        sh.write(r,1,'Dmax')
        sh.write(r,2,'55 Gy')
        sh.write(r,3,femoral_Dmax)
        sh.write(r,4,'Gy')
        r += 1
        sh.write(r,1,'V45')
        sh.write(r,2,'5 %')
        sh.write(r,3,femoral_V45[0])
        sh.write(r,4,'%')
        r += 1

eee = find_DVH_Vx(DVX1,value1)
fff = find_DVH_Dx(DVX1,value1)

#print value1.strip(),': ',eee,'% of the volume get',DVX1,'Gy'
#print value1.strip(),': ',DVX1,'% of the volume get',fff,'Gy'


col1_name = 'Structure'
col2_name = 'Parameter'
col3_name = 'constraint'
col4_name = 'value'

sh.write(0, 0, col1_name, style2)
sh.write(0, 1, col2_name, style2)
sh.write(0, 2, col3_name, style2)
sh.write(0, 3, col4_name, style2)

book.save('Prostata6050_14Gy_patient2.xls')




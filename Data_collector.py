# Modules for the analysis

import pandas as pd
import time
from datetime import datetime
import xlsxwriter
import numpy as np
import math

#######################################################################################################################
# Reading the excel File
start_time = time.clock()
path1 = 'SEA_TAC_0517.xlsx'
xls_file = pd.ExcelFile(path1)

df = xls_file.parse('Sheet1')

num_rows = len(df[ ' Temp']) # Counting num of rows
#######################################################################################################################
# Data Formatting

Date = []
HrMn = []
Temp = []
Date_time= []
for rows in range(0,num_rows):
    dummy1 = str(df[' Date'][rows]) # unformatted date
    dummy2 = str(df[' HrMn'][rows]) # unformatted time
    Date.append(dummy1[4]+dummy1[5]+'/'+dummy1[6]+dummy1[7]+'/'+dummy1[0:4])
    if (len(dummy2) == 3):
        HrMn.append((dummy2[0])+':'+(dummy2[1:3]))
        Date_time.append((dummy1[4] + dummy1[5] + '/' + dummy1[6] + dummy1[7] + '/' + dummy1[0:4] + ' '+'0'+(dummy2[0])+':'+(dummy2[1:3])))
    elif(len(dummy2)== 2):
        HrMn.append('00' + ':' + (dummy2[1:3]))
        Date_time.append((dummy1[4] + dummy1[5] + '/' + dummy1[6] + dummy1[7] + '/' + dummy1[0:4] + ' ' +'00:'+ (dummy2[0:2])))
    elif(len(dummy2)==1):
        HrMn.append('00:00')
        Date_time.append((dummy1[4] + dummy1[5] + '/' + dummy1[6] + dummy1[7] + '/' + dummy1[0:4] + ' '+'00:00'))
    elif(len(dummy2)==4):
        HrMn.append((dummy2[0:2])+':'+(dummy2[1:3]))
        Date_time.append((dummy1[4] + dummy1[5] + '/' + dummy1[6] + dummy1[7] + '/' + dummy1[0:4] + ' '+(dummy2[0:2])+':'+(dummy2[2:4])))
    # Date_time.append(dummy1[4]+dummy1[5]+'/'+dummy1[6]+dummy1[7]+'/'+dummy1[0:4]+''+)
    Temp.append(df[' Temp'][rows]*9/5+32)

Date_time_form= []
for rows in range(0,num_rows):
    Date_time_form.append(datetime.strptime(Date_time[rows], '%m/%d/%Y %H:%M'))

#######################################################################################################################
# Heating Degree Minutes & Cooling Degree Minutes

T_base_H = list(range(45,76)) # Cooling Base Temperatures
T_base_C = list(range(55,76)) # HEating Base Temperatures

HDM = np.zeros((num_rows, len(T_base_H))) #Heating Degree Minutes
CDM = np.zeros((num_rows, len(T_base_C))) #Cooling Degree Minutes

for col in range(0,len(T_base_H)):
    for rows in range(0,len(Temp)):
        if(rows==0):
            HDM[rows][col]=60*max((T_base_H[col]-Temp[rows]),0)
        else:
            D = Date_time_form[rows]-Date_time_form[rows-1]
            HDM[rows][col] = (D.days*24*60+D.seconds/60) * max((T_base_H[col] - Temp[rows]), 0)

for col in range(0,len(T_base_C)):
    for rows in range(0,len(Temp)):
        if(rows==0):
            CDM[rows][col]=60*max((Temp[rows]-T_base_C[col]),0)
        else:
            D = Date_time_form[rows]-Date_time_form[rows-1]
            CDM[rows][col] = (D.days*24*60+D.seconds/60)* max((Temp[rows] - T_base_C[col]), 0)

# print(HDM)
# #
# print(CDM)
# # print('\n')
# print(HDM.shape)
# print(CDM.shape)
######################################################################################################################

# workbook = xlsxwriter.Workbook('date_time_data.xlsx')
# worksheet = workbook.add_worksheet()
# row = 0
#
# for col, data in enumerate(HDM):
#     worksheet.write_column(col, row, data)
#
# workbook.close()
# col = 0
#
# row = 1
# for item in Date_time:
#    worksheet.write(row,col,item)
#    row+=1
#
# workbook.close()
#######################################################################################################################
# Importing the tester file and Getting the HDD days and cooling degree days for the specified periods:

def HDD_calc(SD,ED,HDM):
    for values in range(0,len(Date)):
        if SD==Date[values]:
            start = values
        elif ED ==Date[values]:
            end = values
    list_HDD =np.zeros(len(T_base_H))
    for rows in range(start,end+1):
        list_HDD = (list_HDD+HDM[rows][:])
    return list(list_HDD/(60*24))

def CDD_calc(SD,ED,CDM):
    start = 0
    end = 0
    for values in range(0,len(Date)):
        if SD==Date[values]:
            start = values
        elif ED ==Date[values]:
            end = values
    list_CDD =np.zeros(len(T_base_C))
    for rows in range(start,end+1):
        list_CDD = (list_CDD+CDM[rows][:])
    return list(list_CDD/(60*24))

def MAX_MIN_AVE(SD,ED):
    for values in range(0,len(Date)):
        if SD==Date[values]:
            start = values
        elif ED ==Date[values]:
            end = values
    max_temp = max(Temp[start:end+1])
    min_temp = min(Temp[start:end+1])
    avg_temp = sum(Temp[start:end+1])/len(Temp[start:end+1])
    return [avg_temp, max_temp, min_temp]

path2 = 'Tester_1.xlsx'
xls_file = pd.ExcelFile(path2)
df_2= xls_file.parse('Test')
max_rows = len(df_2[ 'Start']) # Counting num of rows


list_HDD = []
list_CDD = []
list_Min_Max_Ave = []
for rows in range(0,max_rows):
    # print(df_2['Start'][rows],df_2['End'][rows])
    start_date = datetime.strptime(str(df_2['Start'][rows]), '%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y')
    end_date   = datetime.strptime(str(df_2['End'][rows]), '%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y')
    list_HDD.append(HDD_calc(start_date,end_date,HDM))
    list_CDD.append(CDD_calc(start_date,end_date,CDM))
    list_Min_Max_Ave.append(MAX_MIN_AVE(start_date,end_date))

# print(list_HDD)
# print(len(T_base_C))

path3 = 'DataDump.xlsx'
workbook_1 = xlsxwriter.Workbook(path3)

worksheet_1 = workbook_1.add_worksheet('HDD_datadump')
for rows in range(0,max_rows):
    worksheet_1.write_row(rows+1,2,list_HDD[rows])

worksheet_3 = workbook_1.add_worksheet('Average_min_max')
for rows in range(0,max_rows):
    worksheet_3.write_row(rows+1,2,list_Min_Max_Ave[rows])

worksheet_2 = workbook_1.add_worksheet('CDD_datadump')
for rows in range(0,max_rows):
    worksheet_2.write_row(rows+1,2,list_CDD[rows])
workbook_1.close()

#######################################################################################################################

# workbook_1.close()
# path4 = 'CDD.xlsx'
# workbook_2 = xlsxwriter.Workbook(path4)
# worksheet_2 = workbook_2.add_worksheet('CDD_datadump')
#
# for rows in range(0,max_rows):
#     worksheet_2.write_row(rows+1,2,list_CDD[rows])
# workbook_2.close()

###################################################################
# path3 = 'datadump.xlsx'
# workbook_1 = xlsxwriter.workbook(path3)
# worksheet_1 = workbook_1.add_worksheet('HDD_datadump')
# worksheet_2 = workbook_1.add_worksheet('CDD_datadump')
#
# for rows in range(0,max_rows):
#     worksheet_1.write_row(rows+1,2,list_HDD[rows])
#
# for rows in range(0,max_rows):
#     worksheet_2.write_row(rows+1,2,list_CDD[rows])
#
# workbook_1.close()


print('Time Elapsed: %0.2f s '%(time.clock() - start_time))
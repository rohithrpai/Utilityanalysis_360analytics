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

# print('Time Elapsed: %0.2f s '%(time.clock() - start_time))
######################################################################################################################
# Collecting the user specified base temperature for Heating and Cooling
path2 = 'Data_collector.xlsx'
xls_file = pd.ExcelFile(path2)
df_2= xls_file.parse('Data_utility_analysis')
T_base_H_specified = df_2['T_base_H'][0]
T_base_C_specified = df_2['T_base_C'][0]
# print(T_base_H_specified,T_base_C_specified)

#######################################################################################################################
# Heating Degree Minutes & Cooling Degree Minutes
HDM = []
CDM = []
for rows in range(0, len(Temp)):
    if (rows == 0):
        HDM.append(60 * max((T_base_H_specified - Temp[rows]), 0))
    else:
        D = Date_time_form[rows] - Date_time_form[rows - 1]
        HDM.append((D.days * 24 * 60 + D.seconds / 60) * max((T_base_H_specified - Temp[rows]), 0))

for rows in range(0,len(Temp)):
    if(rows==0):
        CDM.append(60*max((Temp[rows]-T_base_C_specified),0))
    else:
        D = Date_time_form[rows]-Date_time_form[rows-1]
        CDM.append((D.days*24*60+D.seconds/60) * max((Temp[rows] - T_base_C_specified), 0))

#######################################################################################################################
# Computing the Heating Degree Days and Cooling Degree Days per day
list_HDD = []
list_CDD = []
list_Min_Max_Ave = []


def HDDperday_calc(SD,ED,HDM,d):
    for values in range(0,len(Date)):
        if SD==Date[values]:
            start = values
        elif ED ==Date[values]:
            end = values
    list_HDD = 0
    for rows in range(start,end+1):
        list_HDD = list_HDD+ HDM[rows]
    return (list_HDD/(60*24*d))

def CDDperday_calc(SD,ED,CDM,d):
    for values in range(0,len(Date)):
        if SD==Date[values]:
            start = values
        elif ED ==Date[values]:
            end = values
    list_CDD = 0
    for rows in range(start,end+1):
        list_CDD= list_CDD+ (CDM[rows])
    return (list_CDD/(60*24*d))

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

max_rows = len(df_2['Start']) # Computes the number of user entries
start_col = []
end_col = []
for rows in range(0,max_rows):
    start_date = datetime.strptime(str(df_2['Start'][rows]), '%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y')
    end_date = datetime.strptime(str(df_2['End'][rows]), '%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y')
    start_col.append(start_date)
    end_col.append(end_date)
    diff = datetime.strptime(end_date, '%m/%d/%Y') - datetime.strptime(start_date, '%m/%d/%Y')
    list_HDD.append(HDDperday_calc(start_date, end_date, HDM, diff.days))
    list_CDD.append(CDDperday_calc(start_date, end_date, CDM, diff.days))
    list_Min_Max_Ave.append(MAX_MIN_AVE(start_date, end_date))

# print(len(list_HDD),len(list_CDD))
path3 = 'Data_dump_new.xlsx'
workbook_1 = xlsxwriter.Workbook(path3)
#
worksheet_1 = workbook_1.add_worksheet('Average_min_max_HDD&CDD')
for rows in range(0,max_rows):
    write = [start_col[rows],end_col[rows],list_Min_Max_Ave[rows][0],list_Min_Max_Ave[rows][1],list_Min_Max_Ave[rows][2],list_HDD[rows],list_CDD[rows]]
    worksheet_1.write_row(rows+1,0,write)

workbook_1.close()
# worksheet_2 = workbook_1.add_worksheet('HDD_T_base_H_specified')
# for rows in range(0,max_rows):
#     worksheet_2.write(rows+1,2,list_HDD[rows])
# #
# worksheet_3 = workbook_1.add_worksheet('HDD_T_base_C_specified')
# for rows in range(0,max_rows):
#     worksheet_3.write(rows+1,1,list_CDD[rows])

#######################################################################################################################
print('Time Elapsed: %0.2f s '%(time.clock() - start_time))

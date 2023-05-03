import pandas as pd
import numpy as np
from pathlib import Path
import openpyxl

# 0 = 'PLANNED WORK OVERVIEW - REPORT'
# 1 = 'PLANNED WORK OVERVIEW'
# 2 = 'BACKLOG'
# 3 = 'CONSOLIDATED HOURS'
# 4 = 'TIME'
# 5 = 'BL_UPDATE'
# 6 = 'ActiveWorkOrders'
try:
    loc = Path(r'C:\Users\104535brbo\Desktop\PLANNED WORK OVERVIEW 4 (003).xlsx')
    pwo_report = pd.read_excel(loc, sheet_name=0)
    planned_wo = pd.read_excel(loc, sheet_name=1)
    backlog = pd.read_excel(loc, sheet_name=2)
    conso_hr = pd.read_excel(loc, sheet_name=3)
    time = pd.read_excel(loc, sheet_name=4)
    BL_UPDATE = pd.read_excel(loc, sheet_name=5)
    active_work_orders = pd.read_excel(loc, sheet_name=6)
    pwo_report = pwo_report[pwo_report['ID'] == 'x']
    pwo_report['WORK ORDER'] = pwo_report['WORK ORDER'].astype('int64')
    pwo_report['WORK TASK SEQUENCE NO'] = pwo_report['WORK TASK SEQUENCE NO'].astype('int64')
    conso_hr['Task No'] = conso_hr['Task No'].astype('int64')
    col = [col for col in pwo_report if col.startswith('Unnamed:')]
    pwo_report.drop(col, axis=1, inplace=True)
    col = [col for col in planned_wo if col.startswith('Unnamed:')]
    planned_wo.drop(col, axis=1, inplace=True)
    col = [col for col in backlog if col.startswith('Unnamed:')]
    backlog.drop(col, axis=1, inplace=True)
    col = [col for col in time if col.startswith('Unnamed:')]
    time.drop(col, axis=1, inplace=True)
    col = [col for col in BL_UPDATE if col.startswith('Unnamed:')]
    BL_UPDATE.drop(col, axis=1, inplace=True)
except:
    pass

for work_order in list(pwo_report['WORK ORDER'].unique()):
    try:
        if work_order in list(planned_wo['WORK ORDER NO']):
            pwo_report.loc[pwo_report['WORK ORDER'] == work_order, 'IN PROGRAM'] = 'TRUE'
            pwo_report.loc[pwo_report['WORK ORDER'] == work_order, 'PROGRAM WO STATUS'] = \
            list(planned_wo.loc[planned_wo['WORK ORDER NO'] == work_order, 'PROGRAM WO STATUS'])[0]
        else:
            pwo_report.loc[pwo_report['WORK ORDER'] == work_order, 'IN PROGRAM'] = 'FALSE'
    except:
        pass

for work_order in list(pwo_report['WORK ORDER'].unique()):
    try:
        if work_order in list(planned_wo['WORK ORDER NO']):
            pwo_report.loc[pwo_report['WORK ORDER'] == work_order, 'RESOURCES PLANNED HOURS'] = \
            list(planned_wo.loc[planned_wo['WORK ORDER NO'] == work_order, 'RESOURCE PLANNED HOURS'])[0]
        else:
            pwo_report.loc[pwo_report['WORK ORDER'] == work_order, 'RESOURCES PLANNED HOURS'] = ' '
    except:
        pass

for work_order in list(pwo_report['WORK ORDER'].unique()):
    try:
        if work_order in list(backlog['WORK ORDER']):
            pwo_report.loc[pwo_report['WORK ORDER'] == work_order, 'PREVIOUSLY REPORTED STATUS'] = 'TRUE'
    except:
        pass

for work_order in list(planned_wo['WORK ORDER NO']):
    try:
        if work_order in list(pwo_report['WORK ORDER']):
            planned_wo.loc[planned_wo['WORK ORDER NO'] == work_order, 'IN PROGRAM'] = 'TRUE'
        else:
            planned_wo.loc[planned_wo['WORK ORDER NO'] == work_order, 'IN PROGRAM'] = 'FALSE'
    except:
        pass

add_false_in_program = planned_wo[planned_wo['IN PROGRAM'] == 'FALSE']
add_false_in_program['ID'] = 'y'
new_column = pwo_report.columns

pwo_report = pd.DataFrame(np.concatenate((pwo_report.values, add_false_in_program.values), axis=0))
pwo_report.columns = new_column

planned_wo['IN PROGRAM'] = 'TRUE'
pwo_report.loc[pwo_report['ID'] == 'y', 'IN PROGRAM'] = 'TRUE'

# #WORKING ON THE TIME AND CONSOLIDATED TIME SHEET

for task_number in list(pwo_report['WORK TASK SEQUENCE NO'].unique()):
    if task_number in list(conso_hr['Task No']):
        pwo_report.loc[((pwo_report['WORK TASK SEQUENCE NO'] == task_number) & (
                    pwo_report['DISTINCT WORK TASK'] == 1)), 'REPORTED RESOURCE MANHOURS'] = \
        list(conso_hr.loc[conso_hr['Task No'] == task_number, 'Hours'])[0]
        pwo_report.loc[((pwo_report['WORK TASK SEQUENCE NO'] == task_number) & (
                    pwo_report['DISTINCT WORK TASK'] == 1)), 'REPORTED  RESOURCE GROUP DESCRIPTION'] = \
        list(conso_hr.loc[conso_hr['Task No'] == task_number, 'Resource Group Description'])[0]
        pwo_report.loc[((pwo_report['WORK TASK SEQUENCE NO'] == task_number) & (
                    pwo_report['DISTINCT WORK TASK'] == 1)), ' REPORTED WORK TASK MAINTENANCE ORG'] = \
        list(conso_hr.loc[conso_hr['Task No'] == task_number, 'Task Maint. Org. Description'])[0]
# add code to identify distinct work task and set to 1s and 0s
# REPLICATING ROWS OF TASK NUMBER WHOSE MANHOUR IS GREATER THAN 75 BY COMPARING IT TO THE ORIGINAL TIME SHEET

great_75 = list(pwo_report[pwo_report['REPORTED RESOURCE MANHOURS'] >= 75]['WORK TASK SEQUENCE NO'])
for task_no in great_75:
    number_dup = time[time['Task No'] == task_no].shape[0] - \
                 pwo_report[pwo_report['WORK TASK SEQUENCE NO'] == task_no].shape[0]
    if number_dup > 0:
        new = pwo_report[pwo_report['WORK TASK SEQUENCE NO'] == task_no]
        pwo_report = pwo_report.append([new] * number_dup, ignore_index=True)
    else:
        pass

for task_no in great_75:
    num_of_rows = time[time['Task No'] == task_no].shape[0] - \
                  pwo_report[pwo_report['WORK TASK SEQUENCE NO'] == task_no].shape[0]
    used_data = list(time.loc[time['Task No'] == task_no, 'Hours'])
    if num_of_rows == 0:
        pwo_report.loc[pwo_report['WORK TASK SEQUENCE NO'] == task_no, 'REPORTED RESOURCE MANHOURS'] = used_data
    elif num_of_rows < 0:
        num_of_rows = pwo_report[pwo_report['WORK TASK SEQUENCE NO'] == task_no].shape[0] - \
                      time[time['Task No'] == task_no].shape[0]
        for i in range(num_of_rows):
            used_data.append(0)
        pwo_report.loc[pwo_report['WORK TASK SEQUENCE NO'] == task_no, 'REPORTED RESOURCE MANHOURS'] = used_data

for work_order in list(pwo_report['WORK ORDER'].unique()):
    if work_order in list(BL_UPDATE['WO No']):
        pwo_report.loc[pwo_report['WORK ORDER'] == work_order, 'PREVIOUSLY REPORTED'] = "TRUE"

    # backlog work on
try:
    for work_order in list(backlog['WORK ORDER'].unique()):
        if work_order in list(BL_UPDATE['WO No']):
            backlog.loc[backlog['WORK ORDER'] == work_order, 'WORK TASK ACTUAL START'] = \
            list(BL_UPDATE.loc[BL_UPDATE['WO No'] == work_order, 'Actual Start'])[0]
            backlog.loc[backlog['WORK ORDER'] == work_order, 'WORK TASK ACTUAL FINISH'] = \
            list(BL_UPDATE.loc[BL_UPDATE['WO No'] == work_order, 'Actual Finish'])[0]
            backlog.loc[backlog['WORK ORDER'] == work_order, 'REPORTED WO STATUS'] = \
            list(BL_UPDATE.loc[BL_UPDATE['WO No'] == work_order, 'Status'])[0]
            backlog.loc[backlog['WORK ORDER'] == work_order, 'BACKLOG  STATUS'] = \
            list(BL_UPDATE.loc[BL_UPDATE['WO No'] == work_order, 'Status'])[0]
            backlog.loc[backlog['WORK ORDER'] == work_order, 'PREVIOUSLY REPORTED'] = "TRUE"
except:
    pass

writer = pd.ExcelWriter('plan overview report.xlsx', engine='xlsxwriter')

pwo_report.to_excel(writer, sheet_name='PLANNED WORK OVERVIEW - REPORT', index=False)
planned_wo.to_excel(writer, sheet_name='PLANNED WORK OVERVIEW', index=False)
backlog.to_excel(writer, sheet_name='BACKLOG', index=False)
conso_hr.to_excel(writer, sheet_name='CONSOLIDATED TIME', index=False)
time.to_excel(writer, sheet_name='TIME', index=False)
BL_UPDATE.to_excel(writer, sheet_name='BL UPDATE', index=False)

writer.save()
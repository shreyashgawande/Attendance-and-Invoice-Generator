import pandas as pd
from openpyxl import *
from openpyxl.styles import *
from datetime import datetime, timedelta
import calendar
# from win32com import client
import os



def generate_attendance(holidays:list, name:str):
    month31 = load_workbook(r"Attendance31.xlsx")
    month30 = load_workbook(r"Attendance30.xlsx")
    
    year = datetime.now().year
    month = datetime.now().month
    month_name = datetime.now().strftime('%b')
    last_month_date = datetime.now() - timedelta(days=datetime.now().day) 
    last_month_name = last_month_date.strftime('%b') 
    last_month = last_month_date.month
    last_month_days = calendar.monthrange(year, last_month)[1]
    ascii_value = 74
    ascii_value1 = 65
    ascii_value_new = None
    presentDays, absentDays = 0, 0
    
    sheet = None
    
    if last_month_days==31:
        sheet = month31.active
    else:
        sheet = month30.active
    
    sheet['E7'] = name
    sheet['F7'] = f'15th-{last_month_name}-{str(year)[2:]}'    
    sheet['G7'] = f'14th-{month_name}-{str(year)[2:]}'
    
    for day in range(15, last_month_days+1):
        sheet[f'{chr(ascii_value)}5'] = f'{day}-{last_month_name}'        
        date_object = datetime(year,last_month,day)
        day_name = date_object.strftime('%A')
        sheet[f'{chr(ascii_value)}6'] = day_name
        
        if day_name in ['Sunday', 'Saturday']:
            sheet[f'{chr(ascii_value)}7'] = 'WO'
            
        elif day in holidays:
            sheet[f'{chr(ascii_value)}7'] = 'A'
            absentDays += 1
        
        else:
            sheet[f'{chr(ascii_value)}7'] = 'P'
            presentDays += 1
        
        ascii_value += 1       
        
    
    for day in range(1, 15):
        if ascii_value == 90:
            ascii_value_new = 'Z'
        if ascii_value > 90:            
            ascii_value_new = f'A{chr(ascii_value1)}'
            ascii_value1 += 1
        
        sheet[f'{ascii_value_new}5'] = f'{day}-{month_name}'        
        date_object = datetime(year,month,day)
        day_name = date_object.strftime('%A')
        sheet[f'{ascii_value_new}6'] = day_name
        
        if day_name in ['Sunday', 'Saturday']:
            sheet[f'{ascii_value_new}7'] = 'WO'
            
        elif day in holidays:
            sheet[f'{ascii_value_new}7'] = 'A'
            absentDays += 1
        
        else:
            sheet[f'{ascii_value_new}7'] = 'P'
            presentDays += 1
        
        ascii_value += 1  
    
    sheet['H7'] = presentDays
    sheet['I7'] = absentDays
    
    if last_month_days==31:
         month31.save(filename=f'Attendance {last_month_name} - {month_name}.xlsx') 
    else:
         month30.save(filename=f'Attendance {last_month_name} - {month_name}.xlsx') 

    return f'Attendance {last_month_name} - {month_name}.xlsx'

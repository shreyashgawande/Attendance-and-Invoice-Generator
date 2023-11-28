import pandas as pd
import openpyxl
from openpyxl import *
from openpyxl.styles import *
from datetime import datetime, timedelta
import calendar
from win32com import client
import os


def generate_invoice_excel(invoice, leavesTaken:int, salary:int):
    year = datetime.now().year
    month_name = datetime.now().strftime('%b')
    month_name1 = datetime.now().strftime('%B')
    month = datetime.now().month
    days = calendar.monthrange(year, month)[1]
    last_month_date = datetime.now() - timedelta(days=datetime.now().day) 
    last_month_name = last_month_date.strftime('%B')
    
    date = f'15th {month_name} {year}'    
    amount_deducted = round(salary/days)*leavesTaken
    
    total_amount = salary - amount_deducted
    
    invoice = openpyxl.load_workbook(invoice)
    sheet = invoice.active
    
    sheet['H4'] = date    
    sheet['H20'] = total_amount
    sheet['G24'] = total_amount
    sheet['B20'] = f" Consultation / Professional Charges for the Period of 1st {month_name} {year} to 30th {month_name} {year}"
    invoice.save(filename=f'Consultant Invoice {last_month_name} - {month_name1}.xlsx')
    
    excel = client.Dispatch("Excel.Application") 
    path = os.path.join(os.getcwd(), f'Consultant Invoice {last_month_name} - {month_name1}.xlsx')
    sheets = excel.Workbooks.Open(path) 
    work_sheets = sheets.Worksheets[0] 
    work_sheets.ExportAsFixedFormat(0, os.path.join(os.getcwd(), f'Consultant Invoice {last_month_name} - {month_name1}.pdf'))
    sheets.Close()
    os.remove(f'Consultant Invoice {last_month_name} - {month_name1}.xlsx')
    return f'Consultant Invoice {last_month_name} - {month_name1}.pdf'
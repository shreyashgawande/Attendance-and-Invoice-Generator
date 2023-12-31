import pandas as pd
import inflect
import openpyxl
from openpyxl import *
from openpyxl.styles import *
from datetime import datetime, timedelta
import calendar
import os
import io
from io import BytesIO


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
    ingine = inflect.engine()
    total_amount_in_words = ingine.number_to_words(total_amount)
    total_amount_in_words="INR: "+total_amount_in_words + " only"

    invoice = openpyxl.load_workbook(invoice)
    sheet = invoice.active
    
    sheet['H4'] = date    
    sheet['H20'] = total_amount
    sheet['G24'] = total_amount
    sheet['C28']= total_amount_in_words
    sheet['B20'] = f" Consultation / Professional Charges for the Period of 1st {month_name} {year} to 30th {month_name} {year}"
    invoice_bytes = io.BytesIO()

    invoice.save(invoice_bytes)
    # invoice.save(filename=f'Consultant Invoice {last_month_name} - {month_name1}.xlsx')
    
    # excel = client.Dispatch("Excel.Application") 
    # path = os.path.join(os.getcwd(), f'Consultant Invoice {last_month_name} - {month_name1}.xlsx')
    # sheets = excel.Workbooks.Open(path) 
    # work_sheets = sheets.Worksheets[0] 
    # work_sheets.ExportAsFixedFormat(0, os.path.join(os.getcwd(), f'Consultant Invoice {last_month_name} - {month_name1}.pdf'))
    # sheets.Close()
    # os.remove(f'Consultant Invoice {last_month_name} - {month_name1}.xlsx')

    # return f'Consultant Invoice {last_month_name} - {month_name1}.pdf'
    return (f'Consultant Invoice {last_month_name} - {month_name1}.xlsx',invoice_bytes)

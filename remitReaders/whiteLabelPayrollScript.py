# -*- coding: utf-8 -*-
"""
Created on Tue Mar  8 09:03:06 2022

@author: jacob.sterling
"""

import pandas as pd
from datetime import date
import datetime
from pathlib import Path
from pdfminer.high_level import extract_text

datetime.datetime.now()
today = date.today()
Year = today.year

yearppp = Year - 3
yearpp = Year - 2
yearp = Year - 1
yearc = Year
yearcc = Year + 1

if today.isocalendar()[1] > 39:
    Yearpp = (f'{yearpp} - {yearp}')
    Yearp = (f'{yearp} - {yearc}')
    Year = (f'{yearc} - {yearcc}')
else:
    Yearpp = (f'{yearppp} - {yearpp}')
    Yearp = (f'{yearpp} - {yearp}')
    Year = (f'{yearp} - {yearc}')
    
Week = input('Enter Week Number: ')

file_path = Path(rf'C:\Users\jacob.sterling\advance.online\J Drive - Operations\Remittances and invoices\White Label\Week {Week}')

result = pd.DataFrame([],columns = ['Name','Date','Hours','Rate','Total','File Name'])

pdf_list = list()

for pdf in file_path.glob('*'):
    if pdf.is_file():
       if pdf.suffix in [".PDF",".pdf"]:
           pdf_list.append(pdf.name)
           print(f'Reading {pdf.name}......')
           df_pdf = pd.DataFrame(extract_text(pdf).split(' '),columns=['values'])
           for i, row in df_pdf.iterrows():
               if 'Total:£' in row['values']:
                   Total = float(row['values'].replace('Total','').replace(':£','').replace(',',''))
               if 'for:' in row['values']:   
                   Name = row['values'].replace('for:','') + ' ' + df_pdf.at[i+1,'values'].replace('Timesheet','')
               if 'Date:' in row['values'] and 'Invoice' in row['values']:
                   Date = row['values'].replace('Date:','').replace('Invoice','')
               if '(Hours):' in row['values']:
                   Hours, Minutes = row['values'].replace('Notes:No','').replace('Notes:','').replace('(Hours):','').split(':')
                   try:
                       Hours = int(Hours) + int(Minutes)/60
                   except ValueError:
                       Minutes = int(Minutes[0:2])
                       Hours = int(Hours) + int(Minutes)/60
           Rate = Total/Hours
           File_Name = pdf.name
           result = pd.concat([result,pd.DataFrame([[Name,Date,Hours,Rate,Total,File_Name]],columns = ['Name','Date','Hours','Rate','Total','File Name'])]).reset_index(drop = True)

pdf_notread = result.loc[result['File Name'].isin(pdf_list) == False,'File Name']
if len(pdf_notread) > 0:
    pdf_notread.to_csv('White Label PDFs NOT Read Week {Week}',index=False)
result.to_csv('White Label Import Week {Week}',index=False)
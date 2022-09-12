# -*- coding: utf-8 -*-
"""
Created on Tue Mar 22 14:00:40 2022

@author: jacob.sterling
"""

import pandas as pd
from datetime import date
import datetime
from pathlib import Path
from tabula import read_pdf

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

file_path = Path(
    rf'C:\Users\jacob.sterling\advance.online\J Drive - Operations\Remittances and invoices\Fircroft\Tax Year {Year}\Week {Week}')

df_result = pd.DataFrame([], columns=['Name', 'Date', 'Description',
                         'Quantity', 'Unit Cost', 'Amount', 'File Name'])

def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)

for pdf in file_path.glob('*'):
    if pdf.is_file():
        if pdf.suffix in [".PDF", ".pdf"]:
            print(f'reading {pdf.name}....')
            df_pdf = read_pdf(pdf, pages='all', guess=False)
            for df_page in df_pdf:
                df_page.columns = list(range(0, len(df_page.columns), 1))
                n = 0
                for i, row in df_page.iterrows():
                    if n > 0:
                        break
                    for j, item in row.iteritems():
                        if str(item).__contains__('Description'):
                            items = list()
                            Description = ''
                            if pdf.name.__contains__('Invoice'):
                                n = 1
                                Name = ''
                            elif pdf.name.__contains__('Remittance'):
                                Name = df_page.at[i-1,0]
                                n = 2
                            for k in range(j, len(df_page.columns)):
                                items.append(str(df_page.at[i+n,k]))

                for item in items:
                    if n <= 2:
                        for x in item.split(' '):
                            k = 0
                            for char in x:
                                if char == '/':
                                    k += 1
                                if k == 2:
                                    Date = x
                                    Description = item.replace(x + ' ', '')
                                    break
                        
                        if Description.__contains__('Mr ') or Name.__contains__('Mr '):
                            prefix = 'Mr '
                        elif Description.__contains__('Mrs ') or Name.__contains__('Mrs '):
                            prefix = 'Mrs '
                        elif Description.__contains__('Miss ') or Name.__contains__('Miss '):
                            prefix = 'Miss '
                        
                        if n == 1:
                            Description, Name = Description.split(prefix)
                            n = 3
                            continue
                        elif n == 2:
                            Name = Name.replace(prefix,'')
                            n = 3
                            continue
                        
                    elif n >= 3:
                        
                        vals = str(item).replace(',','').replace('£','').split(' ')

                        for val in vals:
                            try:
                                value = float(val)
                            except ValueError:
                                continue
                            
                            if n == 3:
                                Quantity = value
                                n += 1
                                continue
                            elif n == 4:
                                Unit_Cost = value
                                n += 1
                                continue
                            elif n == 5:
                                Amount = value
                                n += 1
                                continue
                            elif n >= 6:
                                df_result = pd.concat([df_result, pd.DataFrame([[Name, Date, Description, Quantity, Unit_Cost, Amount, pdf.name]], 
                                                                                  columns=['Name',
                                                                                            'Date', 
                                                                                            'Description',
                                                                                            'Quantity',
                                                                                            'Unit Cost', 
                                                                                            'Amount',
                                                                                            'File Name'])]).reset_index(drop=True)

df_result.to_excel(file_path / f'Fircoft Merit Import Week {Week}.xlsx',index = False)
# -*- coding: utf-8 -*-
"""
Created on Wed Feb 23 16:20:45 2022

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

def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)

file_path = Path(rf'C:\Users\jacob.sterling\advance.online\J Drive - Operations\Remittances and invoices\Lignum Recruitment\Tax Year {Year}\Week {Week}')

for pdf in file_path.glob('*'):
    if pdf.is_file():
        if pdf.suffix in [".PDF",".pdf"]:
            df_pdf = read_pdf(pdf,pages='all',guess=False)
            df = 1
            for df_page in df_pdf:
                if isinstance(df, pd.DataFrame):
                    break
                df_page.columns = list(range(0,len(df_page.columns)))
                for i, row in df_page.iterrows():
                    for j, item in row.iteritems():
                        if item == 'Date':
                            df = df_page[i:]
                            for k in range(0,j+1):
                                df = df.drop(k,axis =1).reset_index(drop=True)
                            df.columns = list(range(0,len(df.columns)))
                            df = df.drop(list(range(4,len(df.columns))),axis =1).reset_index(drop=True)
                            df = df.drop(0,axis =0).reset_index(drop=True)
                            df = df.dropna(subset=[0,len(df.columns)-1],axis=0)
                            break
            
            m = len(df.columns)
            for i, row in df.iterrows():
                for item in row[0].split(' '):
                    if has_numbers(item):
                        try:
                            df.at[i, 'QUANTITY'] = float(item)
                            df.at[i, 'DESCRIPTION'] = row[0].replace(' ' + item,'')
                            n = 1
                        except ValueError:
                            df.at[i, 'DESCRIPTION'] = row[0]
                    else:
                        df.at[i, 'DESCRIPTION'] = row[0]
                        n = 1
                        for j in range(n,m,1):
                            try:
                                df.at[i, 'QUANTITY'] = float(row[j])
                                n = 2
                                break
                            except ValueError:
                                pass
                
                for j in range(n,m,1):    
                    try:
                        items = str(row[j]).replace(' ','').replace(',','').split('£')
                        del items[0]
                        df.at[i, 'UNIT PRICE'] = float(items[0])
                        df.at[i, 'TOTAL'] = float(items[1])
                        break
                    except IndexError:
                        if pd.isnull(row[j]) == False:
                            df.at[i, 'UNIT PRICE'] = float(str(row[j]).replace('£','').replace(' ','').replace(',',''))
                            n = j+1
                            break
                
                for j in range(n,m,1):    
                    try:
                        if pd.isnull(row[j]) == False:
                            df.at[i, 'TOTAL'] = float(str(row[j]).replace('£','').replace(' ','').replace(',',''))
                            break
                    except ValueError:
                        continue
                    
            df = df.drop(list(range(0,m)),axis=1).dropna(axis=0)
            df.to_csv(file_path / f'{pdf.name} Merit Import Week {Week}.csv',index = False)

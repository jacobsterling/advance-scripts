# -*- coding: utf-8 -*-
"""
Created on Tue Mar  8 16:32:38 2022

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

file_path = Path(rf'C:\Users\jacob.sterling\advance.online\J Drive - Operations\Remittances and invoices\Gold Group\Tax Year {Year}\Week {Week}')

result = pd.DataFrame([],columns = ['Name','Date','Hours','Rate','Rate Value','Total','TSheet Total','TS Difference','File Name'])
remittance = pd.DataFrame([],columns = ['Name','Date','TS','Amount','File Name'])

pdf_list = list()

for pdf in file_path.glob('*'):
    if pdf.is_file():
        if (pdf.name).__contains__('Remittance') and pdf.suffix in ['.PDF','.pdf']:
           print(f'Reading {pdf.name}......')
           df_Remittance = read_pdf(pdf,pages = 'all',guess = False)[0].fillna('')
           col = str(df_Remittance.columns[0])
           for i, row in df_Remittance.iterrows():
               if row[col] == 'Date Ref Details' or row[col] =='Date PO. No. Ref. No.':
                   df_Remittance.columns = row
                   df_Remittance = df_Remittance.iloc[i+1:,:].reset_index(drop=True)
                   break
           for i, row in df_Remittance.iterrows():
               try:
                   items = row['Date Ref Details'].split(' ')
               except KeyError:
                   items = row['Date PO. No. Ref. No.'].split(' ') + row['Details'].split(' ')
               Name = 'Totals'
               Date = ''
               TS = ''
               for j in range(0,len(items)):
                   k = 0
                   for l in items[j]:
                       if l == '/':
                           k += 1
                       if k == 2:
                          Date = items[j]
                   if items[j].__contains__('tsId='):
                       TS = items[j].replace('tsId=','')
                   elif items[j].upper().isupper():
                       if items[j-1].upper().__contains__('('):
                           Name = items[j].upper() + ' ' + items[j-2].upper()
                       else:
                           Name = items[j].upper() + ' ' + items[j-1].upper()
               try:
                   try:
                       Amount = float(row['Debit Credit'].replace(',','').replace('£',''))
                   except KeyError:
                       Amount = float(row['Credit'].replace(',','').replace('£',''))
                   remittance = pd.concat([remittance,pd.DataFrame([[Name,Date,TS,Amount,pdf.name]],columns = ['Name','Date','TS','Amount','File Name'])]).reset_index(drop=True)
               except ValueError:
                   pass
        elif pdf.suffix in [".PDF",".pdf"]:
           pdf_list.append(pdf.name)
           print(f'Reading {pdf.name}......')
           df_pdf = read_pdf(pdf,pages = 1,guess=False)[0].fillna('0')
           col = str(df_pdf.columns[0])
           n= 0
           items = list()
           Sheet_Total = 0
           Hours = 0
           Name_list = list()
           for i, row in df_pdf.iterrows():
               if 'Sheet: TS_' in row[col]:
                   n = 1
                   TS = row[col].replace('Sheet: TS_','')
               if 'Total For Sheet TS_' in row[col]:
                   n = 0
               if row[col] == 'Net':
                   Sheet_Total = float(row['Unnamed: 0'].replace(',',''))
               if n == 1 and 'Sheet: TS_' not in row[col]:
                   item = row[col].split(' ')
                   for j in item:
                       if j != TS:
                           items.append(j.upper())
                           if j.upper() not in ['UNITS','DAILY','DAY','HRS','HOURLY','BASIC']:
                               Name_list.append(j.upper())
                   
           for i in range(0,len(items)):
               k = 0
               for j in items[i]:
                   if j == '/':
                       k += 1
                   if k == 2:
                      Date = items[i]
               if items[i].isupper() == False and k == 0:
                   try:
                       value = float(items[i])
                       k = 1
                   except ValueError:
                       try:
                           value = float(items[i].replace(':','.'))
                           k = 1
                       except ValueError:
                           pass
                   if Hours == 0 and  i + 2 < len(items) and k == 1:
                       if items[i+1] == 'DAILY' or items[i+1] == 'UNITS' or items[i+1] == 'HRS':
                           Hours = value
                           Rate = items[i+1:]
                           Rate_List = list()
                           for n in Rate:
                               try:
                                   if float(n) > 1:
                                       Rate_Value = float(n)
                               except ValueError:
                                   Rate_List.append(n)
                                   Rate_List.append(' ')
                           Rate = ''.join(Rate_List)
                           Rate = Rate[:-1]
           Name = list()
           for characters in Name_list:
               for char in characters:
                   if char.isupper():
                       Name.append(char)
                       b = 0
                   else:
                       b = 1
               if b == 0:
                   Name.append(' ')
           Name = "".join(Name)
           if Name != '':
               if Name[-1] == ' ':
                   Name = Name[:-1]
           for j in Name.split(' '):
               if j not in ['UNITS','DAILY','DAY','HRS','HOURLY','BASIC']:
                   Rate = Rate.replace(j + ' ','')
           Rate = Rate.replace(Date + ' ','')
           Total = Rate_Value*Hours
           result = pd.concat([result,pd.DataFrame([[Name,Date,Hours,Rate,Rate_Value,Total,Sheet_Total,Total-Sheet_Total,pdf.name]],columns = ['Name','Date','Hours','Rate','Rate Value','Total','TSheet Total','TS Difference','File Name'])]).reset_index(drop=True)

remittance['Amount'] = remittance['Amount']/1.2

for i, row in remittance.iterrows():
    if row['Name'] == 'Totals':
        if remittance.loc[remittance['File Name'] == row['File Name'],'Amount'].sum() - row['Amount'] != row['Amount']:
            remittance.to_csv(f'Gold Group Remittance Amount Error Detected Week {Week} {Year}.csv',index=False)
    else:
        Name = row['Name'].split(' ')
        First_Name = Name[0]
        Last_Name = Name[-1]
        result.loc[result['Name'].str.contains(Last_Name),'Name'] = row['Name']
        result.loc[result['Name'].str.contains(Last_Name),'Remittance'] = row['Amount']
        result.loc[result['Name'].str.contains(Last_Name),'Remittance File Name'] = row['File Name']
        
        result.loc[result['Name'].str.contains(First_Name),'Name'] = row['Name']
        result.loc[result['Name'].str.contains(First_Name),'Remittance'] = row['Amount']
        result.loc[result['Name'].str.contains(First_Name),'Remittance File Name'] = row['File Name']
        
result = result.fillna(0)
result['Remittance Difference'] = result['Total'] - result['Remittance']
            
result.to_csv(f'Gold Group Import Week {Week} {Year}.csv',index=False)
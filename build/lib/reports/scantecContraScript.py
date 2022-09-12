# -*- coding: utf-8 -*-
"""
Created on Wed Mar 30 14:22:54 2022

@author: jacob.sterling
"""
from pathlib import Path
from scantecPayrollScript import scantecPdfReader
import pandas as pd
from Formats import taxYear
from Formats import day
from Functions import tax_calcs
import numpy as np

Year = taxYear().Year_format1('/')
Year_format = taxYear().Year(' - ')
Week = input('Enter Week Number: ')

df = pd.concat([scantecPdfReader(Week,'TU'),scantecPdfReader(Week,'F')]).drop(['Description','Hours','Rate'],axis = 1).rename(columns={'Worker Name':'Description'})

df.loc[df['UF1'] == 'NA' ,'UF1'] = 0

Period = "0" + str(tax_calcs().tax_month_calc()) if tax_calcs().tax_month_calc() < 10 else str(tax_calcs().tax_month_calc())

df['UF1'] = df['UF1'].astype(str) + Period

df['UF1'] = df['UF1'].astype(int)

df = df.groupby(['Description']).agg({'UF1':'max','Net':np.sum,'File Name':'first'}).reset_index()

df.loc[df['UF1'] == int("0" +  Period), 'UF1'] = 'NA'

df = pd.concat([pd.DataFrame([['Scantec Contra Entry',0000,-1*df['Net'].sum(), 'Total']], columns = ['Description','UF1','Net','File Name']),df])

seq = list(['JRNL' for i in range(0,len(df))])
year = list([Year for i in range(0,len(df))])

df = pd.concat([pd.Series(year, index=df.index, name='Year'), df], axis=1)
df = pd.concat([pd.Series(seq, index=df.index, name='Document Type'), df], axis=1)

df['Nominal'] = 5310
df['Period'] = Period
df['Date'] = day().dayToday()
df['Account'] = 'SCAN01'
#df['Value'] = '£ ' + (df['Net']*-1.2).round(2).astype(str)
df['Value'] = (df['Net']*-1.2).round(2)
df = df.drop('Net',axis = 1)

df = df.reindex(columns=['Document Type', 'Year','Period', 'Date', 'Nominal', 'Account', 'Value', 'Description','UF1'])

df_path = Path.home() / rf"advance.online/J Drive - Operations/Remittances and invoices/Scantec/Tax Year {Year_format}"

print('=======================================================================')
print(f'Week {Week}: ', df['Value'].sum())

df.to_excel(df_path / rf'Scantec Week {Week} Py Contra Import.xlsx',sheet_name = 'Contra Import', index = False)
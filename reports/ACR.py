# -*- coding: utf-8 -*-
"""
Created on Thu Oct 14 14:40:15 2021

@author: jacob.sterling
"""

import pandas as pd
import numpy as np
from datetime import date
from utils.functions import PAYNO_Check, age

# refer to MCR for annotations

PAY_DESC = dict([('Company Income',1),
                 ('Basic',1),
                 ('Basic Pay', 1),
                 ('Day rate EDU - TEN',6),
                 ('Day Rate EDU - TEN',6),
                 ('Daily Rate',7.5),
                 ('Overtime',1),
                 ('Overtime 1',1),
                 ('Overtime 2',1),
                 ('Day rate EDU',6.5),
                 ('Day Rate EDU',6.5),
                 ('Day Rate EDU - GSL',7),
                 ('Day Rate EDU - Coba',6),
                 ('RSS DAY',10),
                 ('Standard',1),
                 ('Standard Rate',1),
                 ('Standard rate',1),
                 ('Standard hourly rate',1),
                 ('Holiday Pay from Archive',0),
                 ('Income Other',0),
                 ('Expense.',0),
                 ('Bonus',0),
                 ('Rate Adj',0),
                 ('Accomodatin non VAT',0),
                 ('Expense No VAT',0),
                 ('Holiday Pay',0),
                 ('SSP',0),
                 ('SNP',0),
                 ('SMP',0),
                 (np.nan,0),
                 ('RSS day',0),
                 ('Margin Refund',0)])

PAY_DESC_DAY_RATE = ['Day rate EDU - TEN',
                     'Day Rate EDU - TEN',
                     'Daily Rate',
                     'Day rate EDU',
                     'Day Rate EDU',
                     'Day Rate EDU - GSL',
                     'Day Rate EDU - Coba',
                     'RSS DAY']

df_MCR = pd.read_csv("Timesheets Hist.csv",
                     encoding = 'latin')


df_Joiners_Error_Report = pd.read_csv("Joiners Error Report.csv",
                                      encoding = 'latin',
                                      usecols = ['Pay No','Sdc Option', 'Type', 'Date of Birth','DOJ','FREQ'])
        
df_Joiners_Error_Report = df_Joiners_Error_Report[df_Joiners_Error_Report['Pay No'].apply(lambda x: PAYNO_Check(x))]


df_Salary_Sacrifice = pd.read_csv('Salary Sacrifice.csv',
                                  encoding = 'latin')

df_MCR['PAYNO'] = df_MCR['PAYNO'].apply(pd.to_numeric, errors='ignore', downcast='float')
df_Joiners_Error_Report['Pay No'] = df_Joiners_Error_Report['Pay No'].apply(pd.to_numeric, errors='ignore', downcast='float')
df_Salary_Sacrifice ['PAYNO'] = df_Salary_Sacrifice ['PAYNO'].apply(pd.to_numeric, errors='ignore', downcast='float')

df_Joiners_Error_Report['DOJ'] = pd.to_datetime(df_Joiners_Error_Report['DOJ'], format='%d/%m/%Y')

df_Salary_Sacrifice = pd.DataFrame(df_Salary_Sacrifice, columns= ['PAYNO', 'DED_ONGOING'])
df_Salary_Sacrifice = df_Salary_Sacrifice.dropna()
df_Joiners_Error_Report['DOJ'] = pd.to_datetime(df_Joiners_Error_Report['DOJ'], format='%d/%m/%Y')
df_Salary_Sacrifice['DED_ONGOING']=df_Salary_Sacrifice['DED_ONGOING'].str.replace(',', '').astype(int)/100
df_Salary_Sacrifice = df_Salary_Sacrifice.apply(pd.to_numeric, errors='ignore', downcast='float')

df_MCR = pd.merge(df_MCR, df_Salary_Sacrifice, left_on = 'PAYNO', right_on = 'PAYNO', how = 'left')

df_MCR['DED_ONGOING'] = df_MCR['DED_ONGOING'].fillna(0)

df = pd.DataFrame([] ,columns = ['PAYNO', 'PROCESSED DATE', 'T/S NUMBER', 'TEMPNAME', 'COMPNAME', 
                                                    'TOTAL HOURS', 'TOTAL PAY','CONTRACTING RATE', 
                                                    'COMPANY INCOME TOTAL','DAY RATE TOTAL','DAY RATE TYPE',
                                                    'SALARY SACRIFICE'])

n = -1

for i, item in df_MCR.iterrows():
    if item['PAY_DESC'] in PAY_DESC:
        TOTAL_HOURS = item['HOURS']*PAY_DESC[item['PAY_DESC']]
        if TOTAL_HOURS > 0:
            TOTAL_PAY = item['HOURS']*item['PAY_RATE']
            CONTRACTING_RATE = (TOTAL_PAY - item['DED_ONGOING'])/TOTAL_HOURS
        else:
            TOTAL_PAY = 0
            CONTRACTING_RATE = 0
        if pd.isnull(item['PAYNO']) == False:
                payno = item['PAYNO']
                if item['PAY_DESC'] == 'Company Income':
                    COMPANY_INCOME_TOTAL = TOTAL_PAY
                else:
                    COMPANY_INCOME_TOTAL = 0
                if item['PAY_DESC'] in PAY_DESC_DAY_RATE:
                    DAY_RATE_TOTAL = item['PAY_RATE']
                    DAY_RATE_TYPE = PAY_DESC[item['PAY_DESC']]
                else:
                    DAY_RATE_TOTAL = 0
                    DAY_RATE_TYPE = 0
                ls = [payno, item['PROCESSED DATE'], item['T/S Number'], item['TEMPNAME'], 
                      item['COMPNAME'], TOTAL_HOURS, TOTAL_PAY, CONTRACTING_RATE, 
                      COMPANY_INCOME_TOTAL, DAY_RATE_TOTAL, DAY_RATE_TYPE, item['DED_ONGOING']]
                row = pd.Series(ls, index=df.columns)
                df = df.append(row, ignore_index=True)
                n = n + 1
        else:
            TOTAL_HOURS = df.at[n, 'TOTAL HOURS'] + TOTAL_HOURS
            if TOTAL_HOURS > 0:
                TOTAL_PAY = df.at[n, 'TOTAL PAY'] + TOTAL_PAY
                df.at[n, 'CONTRACTING RATE'] = (TOTAL_PAY - df.at[n, 'SALARY SACRIFICE'])/TOTAL_HOURS
            else:
                TOTAL_PAY = 0
                df.at[n, 'CONTRACTING RATE'] = 0
                
            df.at[n, 'TOTAL HOURS'] = TOTAL_HOURS
            df.at[n, 'TOTAL PAY'] = TOTAL_PAY
            
            if item['PAY_DESC'] == 'Company Income':
                df.at[n, 'COMPANY INCOME TOTAL'] = df.at[n, 'COMPANY INCOME TOTAL'] + TOTAL_PAY 

            if item['PAY_DESC'] in PAY_DESC_DAY_RATE:
                if item['PAY_RATE'] < df.at[n, 'DAY RATE TOTAL'] or df.at[n, 'DAY RATE TOTAL'] == 0:
                    df.at[n, 'DAY RATE TOTAL'] = item['PAY_RATE']
                    df.at[n, 'DAY RATE TYPE'] = PAY_DESC[item['PAY_DESC']]
    else:
        print('Error - Undefined Pay Description : ',item['PAY_DESC'], ' for PAYNO: ',item['PAYNO'])

df['COMPANY INCOME TOTAL'] = df['COMPANY INCOME TOTAL']/df['TOTAL HOURS']

df = df.drop(0 ,axis = 0)
df.reset_index(drop=True)

# filter the data here

df_negative = pd.concat([df[df['TOTAL HOURS'] <= 0], df[df['TOTAL PAY'] <= 0]])
df = df[(df['TOTAL HOURS'] > 0) & (df['TOTAL PAY'] > 0)]

df['TOTAL HOURS'] = df['TOTAL HOURS'].round(decimals=1)
df[['TOTAL PAY','CONTRACTING RATE','SALARY SACRIFICE']] = df[['TOTAL PAY','CONTRACTING RATE','SALARY SACRIFICE']].round(decimals=2)

df = pd.merge(df, df_Joiners_Error_Report, left_on = 'PAYNO', right_on = 'Pay No', how = 'left').drop(['Pay No'], axis = 1)
df['PAYNO'] = df['PAYNO'].astype(int).round()

df = df.sort_values(by=['COMPNAME','PAYNO']).reset_index(drop=True)
df_MEAN = df.groupby(df['PAYNO']).mean()
df_MEAN = df_MEAN['CONTRACTING RATE'].reset_index().rename(columns={'CONTRACTING RATE':'Average CR','index':'PAYNO'})
df_MEAN['PAYNO'] = df_MEAN['PAYNO'].astype(int).round()

over_91 = df[df['TOTAL HOURS'] >= 91]

with pd.ExcelWriter('ACR Report.xlsx') as writer:
    
    #add any fiters to the outpu here
    df_MEAN.to_excel(writer, sheet_name = "Summary", index = False)
    
    over_91.to_excel(writer, sheet_name = "Over 91", index = False)
    
    df.to_excel(writer, sheet_name = "Data", index = False)

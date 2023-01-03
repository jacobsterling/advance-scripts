# -*- coding: utf-8 -*-
"""
Created on Thu Jan 20 16:57:19 2022

@author: jacob.sterling
"""

import pandas as pd
import numpy as np
from utils.functions import age
from utils.functions import PAYNO_Check

PAY_DESC = dict([('Company Income',1),
                 ('Day rate EDU - TEN',6),
                 ('Day Rate EDU - TEN',6),
                 ('Daily Rate',7.5),
                 ('Overtime',1),
                 ('Basic', 1),
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
                 ('Income Other',0),
                 ('Expense.',0),
                 ('Bonus',0),
                 ('Rate Adj',0),
                 ('Accomodatin non VAT',0),
                 ('Expense No VAT',0),
                 ('Holiday Pay',0),
                 ('Holiday Pay from Archive',0),
                 ('SSP',0),
                 ('SNP',0),
                 ('SMP',0),
                 ('RSS day',0),
                 ('Margin Refund',0),
                 (np.nan,0)])

PAY_DESC_DAY_RATE = ['Day rate EDU - TEN',
                     'Day Rate EDU - TEN',
                     'Daily Rate',
                     'Day rate EDU',
                     'Day Rate EDU',
                     'Day Rate EDU - GSL',
                     'Day Rate EDU - Coba',
                     'RSS DAY']

#reads csv file
df_MCR = pd.read_csv("Timesheets Hist.csv",
                     encoding = 'latin')
df_MCR['PAYNO'] = df_MCR['PAYNO'].astype(float)

df_Joiners_Error_Report = pd.read_csv("Joiners Error Report io.csv",
                                      encoding = 'latin',
                                      low_memory=False,
                                      usecols = ['Pay No',"Forenames", "Surname","STATUS",'Sdc Option', 'Type', 'Date of Birth','DOJ','FEE_TYPE','REWARDS'])

df_Joiners_Error_Report = df_Joiners_Error_Report[df_Joiners_Error_Report["Pay No"].apply(lambda x: PAYNO_Check(x))]
df_Joiners_Error_Report['Pay No'] = df_Joiners_Error_Report['Pay No'].astype(float)

df_Salary_Sacrifice = pd.read_csv('Salary Sacrifice.csv',
                                  encoding = 'latin', usecols=['PAYNO', 'DED_ONGOING']).dropna()
df_Salary_Sacrifice = df_Salary_Sacrifice[df_Salary_Sacrifice["PAYNO"].apply(lambda x: PAYNO_Check(x))]
df_Salary_Sacrifice ['PAYNO'] = df_Salary_Sacrifice['PAYNO'].astype(float)

df_Salary_Sacrifice['DED_ONGOING']=df_Salary_Sacrifice['DED_ONGOING'].str.replace(',', '').astype(int)/100

df_MCR = pd.merge(df_MCR, df_Salary_Sacrifice, how = 'left')

df_MCR['DED_ONGOING'] = df_MCR['DED_ONGOING'].replace(np.nan, 0)

df = pd.DataFrame([[0,0,0,0,0,0,0,0,0,0,0,0]] ,columns = ['PAYNO', 'T/S NUMBER', 'TEMPNAME', 'COMPNAME', 
                                                    'TOTAL HOURS', 'TOTAL PAY','CONTRACTING RATE', 
                                                    'COMPANY INCOME TOTAL','DAY RATE TOTAL','DAY RATE TYPE',
                                                    'SALARY SACRIFICE','Week'])
n = 0

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
                    
                row = pd.DataFrame([[
                    item['PAYNO'], item['T/S Number'], item['TEMPNAME'], 
                    item['COMPNAME'], TOTAL_HOURS, TOTAL_PAY, CONTRACTING_RATE, 
                    COMPANY_INCOME_TOTAL, DAY_RATE_TOTAL, DAY_RATE_TYPE, 
                    item['DED_ONGOING'], item['Week']]], columns = df.columns)
                
                df = pd.concat([df, row], ignore_index=True)
                
                n += 1
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
        print('Error - Undefined Pay Description : ',item['PAY_DESC'])
df = df.drop(0 ,axis = 0)
df.reset_index(drop=True)

df['COMPANY INCOME TOTAL'] = df['COMPANY INCOME TOTAL']/df['TOTAL HOURS']
df = df[df['TOTAL HOURS'] > 0]
df = df[df['TOTAL PAY'] > 0]

df['TOTAL HOURS'] = df['TOTAL HOURS'].round(decimals=1)
df['TOTAL PAY'] = df['TOTAL PAY'].round(decimals=2)
df['CONTRACTING RATE'] = df['CONTRACTING RATE'].round(decimals=2)
df['SALARY SACRIFICE'] = df['SALARY SACRIFICE'].round(decimals=2)

df = pd.merge(df, df_Joiners_Error_Report, left_on = 'PAYNO', right_on = 'Pay No', how = 'left').drop(['Pay No'], axis = 1)

df["Full Name"] = df["Forenames"] + " " + df["Surname"]

df_pivot = pd.pivot_table(df, values=['CONTRACTING RATE'], index=["Full Name", 'PAYNO' , "STATUS"], aggfunc={'CONTRACTING RATE': np.mean}, fill_value=0, margins=True)

df_pivot.to_csv("ACR Week.csv")


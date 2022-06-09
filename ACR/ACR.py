# -*- coding: utf-8 -*-
"""
Created on Thu Jan 20 16:57:19 2022

@author: jacob.sterling
"""

import pandas as pd
import numpy as np
from datetime import date
import win32com.client as client

def age(birthdate):
    DOB = str(birthdate).split('/')
    today = date.today()
    age = today.year - int(DOB[2]) - ((today.month, today.day) < (int(DOB[1]), int(DOB[0])))
    return age

PAY_DESC = dict([('Company Income',1),
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


df_Joiners_Error_Report = pd.read_csv("Joiners Error Report.csv",
                                      encoding = 'latin',
                                      usecols = ['Pay No','Sdc Option', 'Type', 'Date of Birth','DOJ','FEE_TYPE','REWARDS'])



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



df_MCR['DED_ONGOING'] = df_MCR['DED_ONGOING'].replace(np.nan, 0)

df = pd.DataFrame([(0,0,0,0,0,0,0,0,0,0,0,0)] ,columns = ['PAYNO', 'T/S NUMBER', 'TEMPNAME', 'COMPNAME', 
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
                ls = [item['PAYNO'], item['T/S Number'], item['TEMPNAME'], 
                      item['COMPNAME'], TOTAL_HOURS, TOTAL_PAY, CONTRACTING_RATE, 
                      COMPANY_INCOME_TOTAL, DAY_RATE_TOTAL, DAY_RATE_TYPE, item['DED_ONGOING'],item['Week']]
                row = pd.Series(ls, index=df.columns)
                df = df.append(row, ignore_index=True)
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

df = pd.merge(df, df_Joiners_Error_Report, left_on = 'PAYNO', right_on = 'Pay No', how = 'left')
df = df.drop(['Pay No'], axis = 1)
df['PAYNO'] = df['PAYNO'].astype(int).round()

df_path: str = r"C:\Users\jacob.sterling\OneDrive - advance.online\Documents\ACR\fees retained ts.csv"
df_margins = pd.read_csv(df_path)[['TSNO','Management Fee']]
df_margins[['OFFNO','TSNO']] = df_margins['TSNO'].str.split('*',expand = True)
df_margins = df_margins.drop('OFFNO',axis=1)
df_margins = df_margins.drop_duplicates(subset='TSNO', keep='first')
df = pd.merge(df, df_margins, left_on='T/S NUMBER', right_on ='TSNO', how = 'left')

# df_CRM = pd.read_csv("Accounts Fee Code.csv", encoding = 'latin')
# df_CRM.columns = df_CRM.iloc[0, :]
# df_CRM = df_CRM.drop(0, axis = 0)
# df_CRM = df_CRM[:-3]
# df_CRM = df_CRM.drop(['ACCOUNTID'], axis = 1)
# for i, row in df_CRM.iterrows():
#     df.loc[(df['COMPNAME'] == row['Account Name']) & (df['Type'] == 'CIS'),'Fee Code'] = row['Retained Margin CIS']
#     df.loc[(df['COMPNAME'] == row['Account Name']) & (df['Type'] == 'SE'),'Fee Code'] = row['Retained Margin Non CIS (SE)']
#     df.loc[(df['COMPNAME'] == row['Account Name']) & (df['Type'] == 'PAYE'),'Fee Code'] = row['Retained Margin Umbrella no Expenses']
#     df.loc[(df['COMPNAME'] == row['Account Name']) & (df['Sdc Option'] == 'Fixed Expenses'),'Fee Code'] = row['Retained Margin Umbrella with Expenses']
#     df.loc[(df['COMPNAME'] == row['Account Name']) & (df['Sdc Option'] == 'Fixed Expenses Under'),'Fee Code'] = row['Retained Margin Umbrella with Expenses']
#     df.loc[(df['COMPNAME'] == row['Account Name']) & (df['Sdc Option'] == 'Mileage Only'),'Fee Code'] = row['Retained Margin Umbrella with Expenses']

df = df.sort_values(by=['COMPNAME','PAYNO','Week']).reset_index(drop=True)
df.to_excel('ACR df.xlsx',index=False)

# df_MEAN = df.groupby(df['PAYNO']).mean()
# df_MEAN = df_MEAN['CONTRACTING RATE'].reset_index().rename(columns={'CONTRACTING RATE':'Average CR','index':'PAYNO'})
# df_MEAN['PAYNO'] = df_MEAN['PAYNO'].astype(int).round()

# df1 = pd.merge(df,df_MEAN)
# df1 = df1[df1['Type'] == 'PAYE'].groupby(['PAYNO'], as_index=False).agg({'COMPNAME':'first','Average CR':np.mean})

# df1['< £14.50'] = 0
# df1['>= £14.50'] = 0

# df1.loc[df1['Average CR'] >= 14.5, '>= £14.50'] = 1
# df1.loc[df1['Average CR'] < 14.5, '< £14.50'] = 1

# df_pivot = pd.pivot_table(df1, values=['Average CR','>= £14.50','< £14.50'], index=['COMPNAME'],aggfunc={'Average CR': np.mean, '>= £14.50': np.sum, '< £14.50': np.sum}, fill_value=0, margins=True)

# df_pivot['Total'] = df_pivot['< £14.50'] + df_pivot['>= £14.50']
# df_pivot['% below £14.50'] = (df_pivot['< £14.50'] / df_pivot['Total'])*100


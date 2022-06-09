# -*- coding: utf-8 -*-
"""
Created on Thu Oct 14 14:40:15 2021

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

outlook = client.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.To = 'payroll@advance.online'
email.CC = 'jacob.sterling@advance.online ; joshua.richards@advance.online'
#email.To = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = ('MCR Report')

# email.SentOnBehalfOfName = 'email address here'

PAY_DESC = dict([('Company Income',1),
                 ('Basic',1),
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

#reads csv file
df_MCR = pd.read_csv("MCR.csv",
                     encoding = 'latin')


df_Joiners_Error_Report = pd.read_csv("Joiners Error Report.csv",
                                      encoding = 'latin',
                                      usecols = ['Pay No','Sdc Option', 'Type', 'Date of Birth','DOJ'])


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

df = pd.DataFrame([(0,0,0,0,0,0,0,0,0,0,0)] ,columns = ['PAYNO', 'T/S NUMBER', 'TEMPNAME', 'COMPNAME', 
                                                    'TOTAL HOURS', 'TOTAL PAY','CONTRACTING RATE', 
                                                    'COMPANY INCOME TOTAL','DAY RATE TOTAL','DAY RATE TYPE',
                                                    'SALARY SACRIFICE'])

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
                ls = [payno, item['T/S Number'], item['TEMPNAME'], 
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
        print('Error - Undefined Pay Description : ',item['PAY_DESC'], ' for PAYNO: ',payno)

df['COMPANY INCOME TOTAL'] = df['COMPANY INCOME TOTAL']/df['TOTAL HOURS']

df = df.drop(0 ,axis = 0)
df.reset_index(drop=True)
        
df_negative = pd.concat([df[df['TOTAL HOURS'] <= 0], df[df['TOTAL PAY'] <= 0]])
df = df[(df['TOTAL HOURS'] > 0) & (df['TOTAL PAY'] > 0)]

df['TOTAL HOURS'] = df['TOTAL HOURS'].round(decimals=1)
df[['TOTAL PAY','CONTRACTING RATE','SALARY SACRIFICE']] = df[['TOTAL PAY','CONTRACTING RATE','SALARY SACRIFICE']].round(decimals=2)

df = pd.merge(df, df_Joiners_Error_Report, left_on = 'PAYNO', right_on = 'Pay No', how = 'left').drop(['Pay No'], axis = 1)
df['PAYNO'] = df['PAYNO'].astype(int).round()

df_company_income_above_75_low_hours = df[(df['COMPANY INCOME TOTAL'] >= 75)]
df_company_income_above_75_low_hours = df_company_income_above_75_low_hours[(df_company_income_above_75_low_hours['TOTAL HOURS'] <= 10)]

df_day_rate_2_low = pd.concat([df[(df['DAY RATE TOTAL'] < 86.25) & (df['DAY RATE TYPE'] == 7.50)], 
                               df[(df['DAY RATE TOTAL'] < 69.00) & (df['DAY RATE TYPE'] == 6.00)], 
                               df[(df['DAY RATE TOTAL'] < 74.75) & (df['DAY RATE TYPE'] == 6.50)], 
                               df[(df['DAY RATE TOTAL'] < 115.00) & (df['DAY RATE TYPE'] == 10.00)]])

df_day_rate_over_7d = df[(df['DAY RATE TYPE'] > 0)]
df_day_rate_over_7d = df_day_rate_over_7d[(df_day_rate_over_7d['TOTAL HOURS']/df_day_rate_over_7d['DAY RATE TYPE'] > 7)]

df = df.drop(['DAY RATE TOTAL','DAY RATE TYPE','COMPANY INCOME TOTAL'], axis = 1)

df_missing_DOB = df.loc[pd.isnull(df['Date of Birth'])]
df = df.loc[~pd.isnull(df['Date of Birth'])]
df_under18 = df.loc[df['Date of Birth'].apply(age) < 18]

JOIN_DATE = pd.to_datetime('01/10/2021', format='%d/%m/%Y')

df_fixed_expenses_u14 = df[(df['Sdc Option'] == 'Fixed Expenses') & (df['CONTRACTING RATE'] < 14.00)]

def ADJ_SS(dataframe,cr):
    for i, item in dataframe.iterrows():
        if item['SALARY SACRIFICE'] > 0:
            dataframe.at[i, 'ADJ SS'] = (item['TOTAL PAY'] - cr*item['TOTAL HOURS'])
        else:
            dataframe.at[i, 'ADJ SS'] = 0
    return dataframe

df_fixed_expenses_u14 = ADJ_SS(df_fixed_expenses_u14,14)

df_CIS_u13 = df[(df['Type'] == 'CIS') & (df['CONTRACTING RATE'] < 13)]

# for i, item in df_CIS_u13.iterrows(): ######################removing makwana arron
#     if item['PAYNO'] == 97465 and item['DOJ'] < JOIN_DATE:
#         df_CIS_u13 = df_CIS_u13.drop(i)
    
df_CIS_u13 = ADJ_SS(df_CIS_u13,13)
#df_CIS_u12 = ADJ_SS(df_CIS_u12,12)
        
df_uSDC_o23 = df[(df['Sdc Option'] == 'Under SDC') & (df['CONTRACTING RATE'] < 12.50) & (df['Date of Birth'].apply(age) >= 23)]
exceptions = ['PROMAN RECRUITMENT LTD','DANIEL OWEN LTD','JAMES GRAY TRADES LTD', 'JAMES GRAY RECRUITMENT LTD', 'SEARCH CONSULTANCY LIVERPOOL', 'SEARCH CONSULTANCY DUNDEE', 'SEARCH CONSULTANCY MANCHESTER','SEARCH CONSULTANCY LEEDS']
df_uSDC_o23_exceptions = df_uSDC_o23[(df['CONTRACTING RATE'] < 12.19) & (df_uSDC_o23.COMPNAME.isin(exceptions))]
df_uSDC_o23_exceptions = ADJ_SS(df_uSDC_o23_exceptions,12.50)

df_uSDC_o23 = df_uSDC_o23[~df_uSDC_o23.COMPNAME.isin(exceptions)]
df_uSDC_o23 = ADJ_SS(df_uSDC_o23,12.50)
        
df_o21_u22 = df[(df['Date of Birth'].apply(age) >= 21) & (df['Date of Birth'].apply(age) < 23) & (df['CONTRACTING RATE'] < 12.06)]
df_o21_u22 = ADJ_SS(df_o21_u22,12.06)

df_o18_u21 = df[(df['Date of Birth'].apply(age) >= 18) & (df['Date of Birth'].apply(age) < 21) & (df['CONTRACTING RATE'] < 8.92)]
df_o18_u21 = ADJ_SS(df_o18_u21,8.92)
        
df_u18 = df[(df['Date of Birth'].apply(age) <= 18) & (df['CONTRACTING RATE'] < 6.22)]
df_u18 = ADJ_SS(df_u18,5.71)
        
df_o75_hours = df[df['TOTAL HOURS'] > 75]

reports = [
df_uSDC_o23,
df_uSDC_o23_exceptions,
df_o21_u22,
df_o18_u21,
df_u18,
df_CIS_u13,
df_fixed_expenses_u14,
df_o75_hours,
df_company_income_above_75_low_hours,
df_day_rate_2_low,
df_day_rate_over_7d]


def get_df_name(df):
    name = [x for x in globals() if globals()[x] is df][0].replace('df_','')
    return name[:31] if len(name) >= 31 else name

with pd.ExcelWriter('MCR Report.xlsx') as writer:
    for report in reports:
        report.to_excel(writer, sheet_name = get_df_name(report), index = False)
        
        
html = """
    <div> 
    </div><br>
        See the below workers which have been highlighted on the MCR;<br><br>
    </div>
    </div>
        <b> Over 23 + Under SDC  w/ Rate Under £12.50 <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
    </div>
    </div>
        <b> Over 23 + Under SDC  w/ Rate Under £12.19 & agency in {exceptions} <b><br><br>
    </div>
    <div>
        {table2}<br><br><br>
    </div>
    </div>
    <div>
        <b> Over 21 + Under 22 w/ Rate Under £12.06 <b><br><br>
    </div>
    <div>
        {table3}<br><br><br>
    </div>
    </div>
    <div>
        <b> Over 18 + Under 21 w/ Rate Under £8.92 <b><br><br>
    </div>
    <div>
        {table4}<br><br><br>
    </div>
    </div>
    <div>
        <b> Under 18 w/ Rate Under £6.22 <b><br><br>
    </div>
    <div>
        {table5}<br><br><br>
    </div>
    </div>
    <div>
        <b> CIS w/ Under Minimum Rate of £13 <b><br><br>
    </div>
    <div>
        {table6}<br><br><br>
    </div>
    </div>
    <div>
        <b> Fixed Expenses w/ Under Minimum Rate of £14 <b><br><br>
    </div>
    <div>
        {table7}<br><br><br>
    </div>
    </div>
    <div>
        <b> Over 75 Hours <b><br><br>
    </div>
    <div>
        {table8}<br><br><br>
    </div>
        </div>
    </div>
    <div>
        <b> High Company Income w/ Low Hours <b><br><br>
    </div>
    <div>
        {table9}<br><br><br>
    </div>
    </div>
    <div>
        <b> Day Rate Too Low <b><br><br>
    </div>
    <div>
        {table10}<br><br><br>
    </div>
    </div>
    <div>
        <b> Day Rate w/ Over 7 Days Worked <b><br><br>
    </div>
    <div>
        {table11}<br><br><br>
    </div>
"""

email.HTMLBody = html.format(table1 = df_uSDC_o23.to_html(index=False),
                             table2 = df_uSDC_o23_exceptions.to_html(index=False),
                             table3 = df_o21_u22.to_html(index=False),
                             table4 = df_o18_u21.to_html(index=False),
                             table5 = df_u18.to_html(index=False),
                             table6 = df_CIS_u13.to_html(index=False),
                             table7 = df_fixed_expenses_u14.to_html(index=False),
                             table8 = df_o75_hours.to_html(index=False),
                             table9 = df_company_income_above_75_low_hours.to_html(index=False),
                             table10 = df_day_rate_2_low.to_html(index=False),
                             table11 = df_day_rate_over_7d.to_html(index=False),
                             exceptions = exceptions)

email.Send()

email = outlook.CreateItem(0)
email.To = 'enquiries@advance.online; hannah.jarvis@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'
#email.To = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = ('MCR Report - Under 18')

html = """
    </div>
    <div>
        <b> Under 18 <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
"""

email.HTMLBody = html.format(table1 = df_under18.to_html(index=False))
email.Send()

email = outlook.CreateItem(0)
email.To = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = ('MCR Report - Missing DOB')

html = """
    </div>
    <div>
        <b> Missing DOB or Not In Joiners Error <b><br><br>
    </div>
    <div>
        {table2}<br><br><br>
    </div>
    
"""

email.HTMLBody = html.format(table2 = df_missing_DOB.to_html(index=False))
email.Send()
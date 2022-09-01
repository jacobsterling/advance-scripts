# -*- coding: utf-8 -*-
"""
Created on Wed Feb 16 14:14:00 2022

@author: jacob.sterling
"""

import pandas as pd
import numpy as np
from datetime import date
import xlsxwriter

def PAYNO_Check(payno):
    try:
        int(payno)
        return True
    except ValueError:
        return False

df1_path: str = r"C:\Users\jacob.sterling\OneDrive - advance.online\Exec Reports\Margins Reports\Margins 2020-2021\Margins Report 2020-2021.xlsx"
df_path: str = r"C:\Users\jacob.sterling\OneDrive - advance.online\Exec Reports\Margins Reports\Margins 2021-2022\Margins Report 2021-2022.xlsx"
df_path_new: str = rf"C:\Users\jacob.sterling\OneDrive - advance.online\Documents\Data\repot.xlsx"

df_joiners_io_path: str = r"C:\Users\jacob.sterling\OneDrive - advance.online\Exec Reports\Margins Reports\Margins 2021-2022\Data\Week 45\Joiners Error Report io"
df_joiners_axm_path: str = r"C:\Users\jacob.sterling\OneDrive - advance.online\Exec Reports\Margins Reports\Margins 2021-2022\Data\Week 45\Joiners Error Report axm"

df1 = pd.read_excel(df1_path, sheet_name= 'Core Data',usecols = ['PAYNO','Client Name','Surname','Forename','Margins','CHQDATE']).dropna(subset=['PAYNO'])
df1_payno_error = df1[df1['PAYNO'].apply(lambda x: PAYNO_Check(x)) == False]
df1 = df1[df1['PAYNO'].apply(lambda x: PAYNO_Check(x)) == True]
df1['PAYNO'] = df1['PAYNO'].astype(int)
df1['Client Name'] = df1['Client Name'].str.upper()

df1_axm = df1[df1['Client Name'] == 'ALEXANDER MANN']
df1_io = df1[df1['Client Name'] != 'ALEXANDER MANN']

df = pd.read_excel(df_path, sheet_name= 'Core Data',usecols = ['PAYNO','Client Name','Surname','Forename','Margins','CHQDATE']).dropna(subset=['PAYNO'])
df_payno_error = df[df['PAYNO'].apply(lambda x: PAYNO_Check(x)) == False]
df = df[df['PAYNO'].apply(lambda x: PAYNO_Check(x)) == True]
df['PAYNO'] = df['PAYNO'].astype(int)
df['Client Name'] = df['Client Name'].str.upper()

df_axm = df[df['Client Name'] == 'ALEXANDER MANN']
df_io = df[df['Client Name'] != 'ALEXANDER MANN']


df1_joiners_io = pd.read_excel('Joiners 2020-2021 io.xlsx',usecols = ['Pay No','Sdc Option', 'Type','POST_CODE']).dropna(subset=['Pay No'])
df1_joiners_io = df1_joiners_io[df1_joiners_io['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
df1_joiners_io['Pay No'] = df1_joiners_io['Pay No'].astype(int)

df1_joiners_axm = pd.read_excel('Joiners 2020-2021 axm.xlsx',usecols = ['Pay No','Sdc Option', 'Type','POST_CODE']).dropna(subset=['Pay No'])
df1_joiners_axm = df1_joiners_axm[df1_joiners_axm['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
df1_joiners_axm['Pay No'] = df1_joiners_axm['Pay No'].astype(int)

df_joiners_io = pd.read_csv('Joiners Error Report io.csv',usecols = ['Pay No','Sdc Option', 'Type','POST_CODE'],encoding = 'latin').dropna(subset=['Pay No'])
df_joiners_io = df_joiners_io[df_joiners_io['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
df_joiners_io['Pay No'] = df_joiners_io['Pay No'].astype(int)

df_joiners_axm = pd.read_csv('Joiners Error Report axm.csv',usecols = ['Pay No','Sdc Option', 'Type','POST_CODE'],encoding = 'latin').dropna(subset=['Pay No'])
df_joiners_axm = df_joiners_axm[df_joiners_axm['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
df_joiners_axm['Pay No'] = df_joiners_axm['Pay No'].astype(int)

df1_io = pd.merge(df1_io.drop_duplicates(subset=['PAYNO']),df1_joiners_io.drop_duplicates(subset=['Pay No']),left_on = 'PAYNO',right_on='Pay No',how='left').drop('Pay No',axis=1)
df1_axm = pd.merge(df1_axm.drop_duplicates(subset=['PAYNO']),df1_joiners_axm.drop_duplicates(subset=['Pay No']),left_on = 'PAYNO',right_on='Pay No',how='left').drop('Pay No',axis=1)
df_io = pd.merge(df_io.drop_duplicates(subset=['PAYNO']),df_joiners_io.drop_duplicates(subset=['Pay No']),left_on = 'PAYNO',right_on='Pay No',how='left').drop('Pay No',axis=1)
df_axm = pd.merge(df_axm.drop_duplicates(subset=['PAYNO']),df_joiners_axm.drop_duplicates(subset=['Pay No']),left_on = 'PAYNO',right_on='Pay No',how='left').drop('Pay No',axis=1)

df = pd.concat([pd.concat([df_io,df1_io]),pd.concat([df_axm,df1_axm])])

df = df[df['Margins'] > 0]
df['Margins'] = 1
df['Area Code'] = df['POST_CODE'].str[:3]

df_CIS = df[df['Type'] == 'CIS']
df = df[df['Type'] != 'CIS']

table = pd.pivot_table(df, values='Margins', index=['Area Code'], columns=['Sdc Option'],aggfunc=np.sum,fill_value=0,margins = True)
table_CIS = pd.pivot_table(df_CIS, values='Margins', index=['Area Code'], columns=['Type'],aggfunc=np.sum,fill_value=0)

with pd.ExcelWriter('FCSA Data.xlsx') as writer:
    table.to_excel(writer,sheet_name='Umbrella Summary')
    df.to_excel(writer,sheet_name='Umbrella Data',index=False)
    table_CIS.to_excel(writer,sheet_name='CIS Summary')
    df_CIS.to_excel(writer,sheet_name='CIS Data',index=False)

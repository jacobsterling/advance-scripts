# -*- coding: utf-8 -*-
"""
Created on Mon Feb 21 12:11:18 2022

@author: jacob.sterling
"""

import pandas as pd
from Formats import taxYear
from Functions import PAYNO_Check

Year = taxYear().Year('-')
    
Week = input('Enter Week Number: ')
    
fees_io = pd.read_csv(rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}\fees retained io.csv",usecols=['PAYNO','Surname','Forename','CHQDATE'], encoding = 'latin')
fees_io['PAYNO'] = fees_io['PAYNO'].astype(int)
fees_io = fees_io.drop_duplicates(subset='PAYNO')
fees_axm = pd.read_csv(rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}\fees retained axm.csv",usecols=['PAYNO','Surname','Forename','CHQDATE'], encoding = 'latin')
fees_axm['PAYNO'] = fees_axm['PAYNO'].astype(int)
fees_axm = fees_axm.drop_duplicates(subset='PAYNO')

joiners_io  = pd.read_csv(rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}\Joiners Error Report io.csv",usecols=['Pay No','Email Address','WEEKS_PAID'], encoding = 'latin')
joiners_io = joiners_io[joiners_io['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
joiners_io['Pay No'] = joiners_io['Pay No'].astype(int)

joiners_axm  = pd.read_csv(rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}\Joiners Error Report axm.csv",usecols=['Pay No','Email Address','WEEKS_PAID'], encoding = 'latin')
joiners_axm = joiners_axm[joiners_axm['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
joiners_axm['Pay No'] = joiners_axm['Pay No'].astype(int)

df_LAST_PAID_IO = pd.merge(fees_io, joiners_io, left_on = 'PAYNO',right_on = 'Pay No', how='left').rename(columns={'CHQDATE':'Last Paid'}).drop('Pay No',axis=1).drop_duplicates(subset=['Email Address'])
df_LAST_PAID_AXM = pd.merge(fees_axm, joiners_axm, left_on = 'PAYNO',right_on = 'Pay No', how='left').rename(columns={'CHQDATE':'Last Paid'}).drop('Pay No',axis=1).drop_duplicates(subset=['Email Address'])

for i, item in df_LAST_PAID_IO.iterrows():
    count = str(item['WEEKS_PAID']).count(',') + 1
    df_LAST_PAID_IO.at[i,'WEEKS_PAID'] = count

for i, item in df_LAST_PAID_AXM.iterrows():
    count = str(item['WEEKS_PAID']).count(',') + 1
    df_LAST_PAID_AXM.at[i,'WEEKS_PAID'] = count
    
df_LAST_PAID_IO.to_csv('last paid io.csv', encoding='utf-8', index=False)
df_LAST_PAID_AXM.to_csv('last paid axm.csv', encoding='utf-8',index=False)

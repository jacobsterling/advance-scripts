# -*- coding: utf-8 -*-
"""
Created on Mon Feb 21 12:11:18 2022

@author: jacob.sterling
"""

import pandas as pd

from utils import formats
from utils import functions

Year = formats.taxYear().Year('-')
    
Week = input('Enter Week Number: ')
    
fees_io = pd.read_csv(rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}\fees retained io.csv",usecols=['Client Name', 'PAYNO','Surname','Forename','CHQDATE','Solution'], encoding = 'latin')
fees_io['PAYNO'] = fees_io['PAYNO'].astype(int)
fees_io = fees_io.drop_duplicates(subset='PAYNO')
fees_axm = pd.read_csv(rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}\fees retained axm.csv",usecols=['Client Name','PAYNO','Surname','Forename','CHQDATE','Solution'], encoding = 'latin')
fees_axm['PAYNO'] = fees_axm['PAYNO'].astype(int)
fees_axm = fees_axm.drop_duplicates(subset='PAYNO')

joiners_io  = pd.read_csv(rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}\Joiners Error Report io.csv",usecols=['OFFNO','Pay No','Email Address','WEEKS_PAID','NI_NO'], encoding = 'latin')
joiners_io = joiners_io[joiners_io['Pay No'].apply(lambda x: functions.PAYNO_Check(x)) == True]
joiners_io['Pay No'] = joiners_io['Pay No'].astype(int)
joiners_io['OFFNO'] = joiners_io['OFFNO'].astype(str)


joiners_axm  = pd.read_csv(rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}\Joiners Error Report axm.csv",usecols=['OFFNO','Pay No','Email Address','WEEKS_PAID','NI_NO'], encoding = 'latin')
joiners_axm = joiners_axm[joiners_axm['Pay No'].apply(lambda x: functions.PAYNO_Check(x)) == True]
joiners_axm['Pay No'] = joiners_axm['Pay No'].astype(int)
joiners_axm['OFFNO'] = joiners_axm['OFFNO'].astype(str)


df_LAST_PAID_IO = pd.merge(fees_io, joiners_io, left_on = 'PAYNO',right_on = 'Pay No', how='left').rename(columns={'CHQDATE':'Last Paid'}).drop('Pay No',axis=1).drop_duplicates(subset=['Email Address'])
df_LAST_PAID_AXM = pd.merge(fees_axm, joiners_axm, left_on = 'PAYNO',right_on = 'Pay No', how='left').rename(columns={'CHQDATE':'Last Paid'}).drop('Pay No',axis=1).drop_duplicates(subset=['Email Address'])

for i, row in df_LAST_PAID_IO.iterrows():
    payroll = str(row['PAYNO']) if pd.isnull(row['PAYNO']) else str(row['OFFNO']) + "*" + str(row['PAYNO'])
    if row["Solution"] == "PAYE":
        df_LAST_PAID_IO.loc[i, "Merit Payroll Number (Umbrella IO_Database)"] = payroll
    elif row["Solution"] in ["CIS","SE"]:
       df_LAST_PAID_IO.loc[i, "Merit Payroll Number (Self-Employed and CIS)"] = payroll
    else:
        raise Exception("No Solution for payno: {row['PAYNO']}")

df_LAST_PAID_AXM["Merit Payroll Number - Alexander Mann"] = df_LAST_PAID_AXM['OFFNO'] + "*" + df_LAST_PAID_AXM['PAYNO'].astype(str)


df_unmerged = pd.concat([df_LAST_PAID_IO, df_LAST_PAID_AXM])

df_unmerged = df_unmerged[df_unmerged['Email Address'].isna()].reset_index(drop=True)

df_missingNI = pd.concat([df_LAST_PAID_IO, df_LAST_PAID_AXM])

df_missingNI = df_missingNI[df_missingNI['NI_NO'].isna()].reset_index(drop=True)


df_LAST_PAID_IO = df_LAST_PAID_IO.drop("NI_NO", axis= 1)
df_LAST_PAID_AXM = df_LAST_PAID_AXM.drop("NI_NO", axis= 1)

for i, item in df_LAST_PAID_IO.iterrows():
    count = str(item['WEEKS_PAID']).count(',') + 1
    df_LAST_PAID_IO.at[i,'WEEKS_PAID'] = count

for i, item in df_LAST_PAID_AXM.iterrows():
    count = str(item['WEEKS_PAID']).count(',') + 1
    df_LAST_PAID_AXM.at[i,'WEEKS_PAID'] = count
    
df_LAST_PAID_IO.to_csv('last paid io.csv', encoding='utf-8', index=False)
df_LAST_PAID_AXM.to_csv('last paid axm.csv', encoding='utf-8',index=False)

import win32com.client as client
email = client.Dispatch('Outlook.Application').CreateItem(0)
email.To = 'enquiries@advance.online; hannah.jarvis@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'
#email.To = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = ('Paid Last Week w/ Missing NI Numbers')

html = """
    </div>
    <div>
        <b> Missing NI Numbers <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
"""

email.HTMLBody = html.format(table1 = df_missingNI.to_html(index=False))
email.Send()



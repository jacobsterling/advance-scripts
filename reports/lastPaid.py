# -*- coding: utf-8 -*-
"""
Created on Mon Feb 21 12:11:18 2022

@author: jacob.sterling
"""

import pandas as pd

from utils.formats import taxYear
from utils.functions import age, PAYNO_Check
from pathlib import Path

Year = taxYear().Year('-')
    
Week = input('Enter Week Number: ')

homePath = Path.home() / "advance.online"

dataPath = homePath / rf"J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}"

fees_io = pd.read_csv(dataPath / "fees retained io.csv",usecols=['Client Name', 'PAYNO','Surname','Forename','CHQDATE','Solution'], encoding = 'latin')
fees_io['PAYNO'] = fees_io['PAYNO'].astype(int)
fees_io = fees_io.drop_duplicates(subset='PAYNO')
fees_axm = pd.read_csv(dataPath / "fees retained axm.csv",usecols=['Client Name','PAYNO','Surname','Forename','CHQDATE','Solution'], encoding = 'latin')
fees_axm['PAYNO'] = fees_axm['PAYNO'].astype(int)
fees_axm = fees_axm.drop_duplicates(subset='PAYNO')

joiners_io  = pd.read_csv(dataPath / "Joiners Error Report io.csv",usecols=['OFFNO','Pay No','Email Address','WEEKS_PAID','NI_NO','Date of Birth',"Sdc Option"], encoding = 'latin')
joiners_io = joiners_io[joiners_io['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
joiners_io['Pay No'] = joiners_io['Pay No'].astype(int)
joiners_io['OFFNO'] = joiners_io['OFFNO'].astype(str)


joiners_axm  = pd.read_csv(dataPath / "Joiners Error Report axm.csv",usecols=['OFFNO','Pay No','Email Address','WEEKS_PAID','NI_NO','Date of Birth'], encoding = 'latin')
joiners_axm = joiners_axm[joiners_axm['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
joiners_axm['Pay No'] = joiners_axm['Pay No'].astype(int)
joiners_axm['OFFNO'] = joiners_axm['OFFNO'].astype(str)


last_paid_io = pd.merge(fees_io, joiners_io, left_on = 'PAYNO',right_on = 'Pay No', how='left').rename(columns={'CHQDATE':'Last Paid'}).drop('Pay No',axis=1).drop_duplicates(subset=['Email Address'])
last_paid_axm = pd.merge(fees_axm, joiners_axm, left_on = 'PAYNO',right_on = 'Pay No', how='left').rename(columns={'CHQDATE':'Last Paid'}).drop('Pay No',axis=1).drop_duplicates(subset=['Email Address'])

for i, row in last_paid_io.iterrows():
    if row["Solution"] == "PAYE":
        last_paid_io.loc[i, "Merit Payroll Number (Umbrella IO_Database)"] = str(row['OFFNO']) + "*" + str(row['PAYNO'])
    elif row["Solution"] in ["CIS","SE"]:
       last_paid_io.loc[i, "Merit Payroll Number (Self-Employed and CIS)"] = str(row['OFFNO']) + "*" + str(row['PAYNO'])
    else:
        raise Exception("No Solution for payno: {row['PAYNO']}")

last_paid_axm["Merit Payroll Number - Alexander Mann"] = last_paid_axm['OFFNO'] + "*" + last_paid_axm['PAYNO'].astype(str)

last_paid = pd.concat([last_paid_io, last_paid_axm])

unmerged = last_paid[last_paid['Email Address'].isna()].reset_index(drop=True)

missingNI = last_paid[last_paid['NI_NO'].isna()].reset_index(drop=True)

under18 = last_paid[last_paid['Date of Birth'].apply(age) < 18]

last_paid_io = last_paid_io.drop(columns = ["NI_NO", "Date of Birth", "Sdc Option"])
last_paid_axm = last_paid_axm.drop(columns = ["NI_NO", "Date of Birth"])

for i, item in last_paid_io.iterrows():
    count = str(item['WEEKS_PAID']).count(',') + 1
    last_paid_io.at[i,'WEEKS_PAID'] = count

for i, item in last_paid_axm.iterrows():
    count = str(item['WEEKS_PAID']).count(',') + 1
    last_paid_axm.at[i,'WEEKS_PAID'] = count
    
last_paid_io.to_csv('last paid io.csv', encoding='utf-8', index=False)
last_paid_axm.to_csv('last paid axm.csv', encoding='utf-8',index=False)

import win32com.client as client
email = client.Dispatch('Outlook.Application').CreateItem(0)
email.To = 'enquiries@advance.online; hannah.jarvis@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = ('Paid Last Week - Enquiries Checks')

html = """
    </div>
    <div>
        <b> Missing NI Numbers <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
    <div>
        <b> Under 18 <b><br><br>
    </div>
    <div>
        {table2}<br><br><br>
    </div>
"""

email.HTMLBody = html.format(table1 = missingNI.to_html(index=False), table2 = under18.to_html(index=False))
email.Send()



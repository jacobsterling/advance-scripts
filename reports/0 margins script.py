# -*- coding: utf-8 -*-
"""
Created on Mon Mar 21 12:23:37 2022

@author: jacob.sterling
"""

import pandas as pd
from datetime import date
import datetime
datetime.datetime.now()

from Formats import taxYear

Year = taxYear().Year('-')
    
Week = input('Enter Week Number: ')

def PAYNO_Check(payno):
    try:
        int(payno)
        return True
    except ValueError:
        return False
    
file_path: str = rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}"
result_path: str = rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\0 Margins"

df_0_MARGINS_IO: str = f"{file_path}\\£0 Margins io.csv"
df_0_MARGINS_AXM: str = f"{file_path}\\£0 Margins axm.csv"
df_FEES_RETAINED_IO: str = f"{file_path}\\fees retained io.csv"
df_FEES_RETAINED_AXM: str = f"{file_path}\\fees retained axm.csv"

df_FEES_RETAINED_IO = pd.read_csv(df_FEES_RETAINED_IO,encoding = 'latin')

df_FEES_RETAINED_IO = df_FEES_RETAINED_IO.rename(columns={"Management Fee":"Margins"})
df_FEES_RETAINED_IO = df_FEES_RETAINED_IO.rename(columns={"OFF_PAYNO":"PAYNO"})
df_FEES_RETAINED_IO['PAYNO'] = df_FEES_RETAINED_IO['PAYNO'].astype(int)
df_FEES_RETAINED_IO['Client Name'] = df_FEES_RETAINED_IO['Client Name'].str.upper()
df_FEES_RETAINED_IO = df_FEES_RETAINED_IO[df_FEES_RETAINED_IO['Margins'] > 0]
df_FEES_RETAINED_IO['CHQDATE'] = pd.to_datetime(df_FEES_RETAINED_IO['CHQDATE'], format='%d/%m/%Y')

try:
    df_FEES_RETAINED_AXM = pd.read_csv(df_FEES_RETAINED_AXM,encoding = 'latin')

    df_FEES_RETAINED_AXM = df_FEES_RETAINED_AXM.rename(columns={"Management Fee":"Margins"})
    df_FEES_RETAINED_AXM = df_FEES_RETAINED_AXM.rename(columns={"OFF_PAYNO":"PAYNO"})
    df_FEES_RETAINED_AXM['PAYNO'] = df_FEES_RETAINED_AXM['PAYNO'].astype(int)
    df_FEES_RETAINED_AXM['Client Name'] = df_FEES_RETAINED_AXM['Client Name'].str.upper()
    df_FEES_RETAINED_AXM = df_FEES_RETAINED_AXM[df_FEES_RETAINED_AXM['Margins'] > 0]
    df_FEES_RETAINED_AXM['CHQDATE'] = pd.to_datetime(df_FEES_RETAINED_AXM['CHQDATE'], format='%d/%m/%Y')

    df_FEES_RETAINED = pd.concat([df_FEES_RETAINED_IO,                           
                          df_FEES_RETAINED_AXM])
    
    df_0_MARGINS = pd.concat([pd.read_csv(df_0_MARGINS_IO,encoding = 'latin'),                           
                              pd.read_csv(df_0_MARGINS_AXM,encoding = 'latin')])
    
except FileNotFoundError:
    df_FEES_RETAINED = df_FEES_RETAINED_IO
    df_0_MARGINS = pd.read_csv(df_0_MARGINS_IO,encoding = 'latin')

df_0_MARGINS = df_0_MARGINS.rename(columns={"OFF_PAYNO":"PAYNO"})

df_0_MARGINS.drop_duplicates(subset ="PAYNO",
                     keep = "first", inplace = True)

df_0_MARGINS = df_0_MARGINS[df_0_MARGINS['PAYNO'].notna()]
df_0_MARGINS['CHQDATE'] = pd.to_datetime(df_0_MARGINS['CHQDATE'], format='%d/%m/%y')

df_0_MARGINS['PAYNO'] = df_0_MARGINS['PAYNO'].astype(int)
df_0_MARGINS["Reason"] = 0
df_0_MARGINS.loc[df_0_MARGINS["SMP"] > 0, "Reason"] = 'SMP'
df_0_MARGINS.loc[df_0_MARGINS["Spp"] > 0, "Reason"] = 'Spp'
df_0_MARGINS.loc[df_0_MARGINS["SSP"] > 0, "Reason"] = 'SSP'
df_0_MARGINS.loc[df_0_MARGINS["Paydesc"] == "Expenses", "Reason"] = "Expenses"
df_0_MARGINS.loc[df_0_MARGINS["Fee Code"] == 0, "Reason"] = "£0 Margin Agreed"
df_0_MARGINS.loc[df_0_MARGINS["Paydesc"] == "Additional Pay", "Reason"] = "Additional Pay"
df_0_MARGINS.loc[df_0_MARGINS["Fee Code"] == "GSL", "Reason"] = 'GSL'
df_0_MARGINS.loc[df_0_MARGINS["GROSS"] < 100, "Reason"] = "Low Pay"

df_0_MARGINS = df_0_MARGINS[df_0_MARGINS['Reason'].isin([0,"£0 Margin Agreed"])]
df_0_MARGINS = df_0_MARGINS[df_0_MARGINS['Paydesc'] == 'Basic Pay']
df_0_MARGINS = df_0_MARGINS[df_0_MARGINS['Pay Rates'] >= 0]
df_0_MARGINS = df_0_MARGINS.rename(columns={'COMPNAME':'Client Name'})
df_0_MARGINS = df_0_MARGINS.rename(columns={'SURNAME':'Surname'})
df_0_MARGINS = df_0_MARGINS.rename(columns={'FIRSTNAME':'Forename'})
df_0_MARGINS = df_0_MARGINS.rename(columns={'MANAGEMENT FEE':'Margins'})
df_0_MARGINS = df_0_MARGINS.rename(columns={'TYPE':'Type'})
df_0_MARGINS = df_0_MARGINS.rename(columns={'Sdc Option':'Solution'})

df_PREV_WEEK = df_FEES_RETAINED[df_FEES_RETAINED['Margins'] > 0]
df_0_MARGINS = df_0_MARGINS[~df_0_MARGINS["PAYNO"].isin(df_PREV_WEEK['PAYNO'])]
df_0_MARGINS['REWARDS'] = 'No'
df_0_MARGINS_ADDED = pd.DataFrame([],columns=df_0_MARGINS.columns)
    
for i, item in df_0_MARGINS.iterrows():
    print('___________________________________________________________________')
    print('')
    print(f'Listing {len(df_0_MARGINS)} £0 Margins...')
    print('___________________________________________________________________')
    print('')
    print(item)
    print('___________________________________________________________________')
    option1 =input('Type "y" if is a Agreed 0 margin, Type "add" to add as a margin manually (for invoiced margins) and type "n" to remove from £0 margins reports (Other reasons)("y" / "n" / "add"): ')
    if option1  == 'n':
        df_0_MARGINS = df_0_MARGINS.drop(i ,axis = 0)
    elif option1  == 'add':
        print('')
        Quantity = input('Enter number of margins this entry is equal too: ')
        item['Margins'] = 1
        for j in range(0, int(Quantity), 1):
            df_FEES_RETAINED = df_FEES_RETAINED.append(item[['Client Name','PAYNO','Surname','Forename','Margins','CHQDATE','Type','REWARDS']])
            df_0_MARGINS_ADDED = df_0_MARGINS_ADDED.append(item[['Client Name','PAYNO','Surname','Forename','Margins','CHQDATE','Type','REWARDS']])
        df_0_MARGINS = df_0_MARGINS.drop(i ,axis = 0)
    else:
        continue

df_0_MARGINS.to_excel(f'{result_path}\£0 Margins - Week {Week}.xlsx',index=False)
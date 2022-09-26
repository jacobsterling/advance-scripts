# -*- coding: utf-8 -*-
"""
Created on Tue Feb  1 16:20:16 2022

@author: jacob.sterling
"""

import pandas as pd
import datetime
import numpy as np
from functions import tax_calcs
from formats import taxYear
from pathlib import Path

Year = taxYear().Year('-')

#change month num
month_num = '07'

Week = tax_calcs().tax_week_calc() - 1

Year_format1 = taxYear().Year_format1('-')
Year_format2 = taxYear().Year_format2()
datetime_object = datetime.datetime.strptime(month_num, "%m")
month_name = datetime_object.strftime("%b")
full_month_name = datetime_object.strftime("%B")

file_path: str = rf"C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}"
margins_path: str = rf"{file_path}\Margins Report {Year}.xlsx"
groups_path: str = r"data\Groups.xlsx"
df_path: str = rf"data\Net Margin {month_name} {Year}.xlsx"
df_rebate_path: str = rf"C:\Users\jacob.sterling\advance.online\J Drive - Finance\Rebates Reports\Rebates {Year_format1}\April Rebates Import.xlsx"
 #{month_name} {Year_format2}

df_CLIENTS_IO: str = rf"{file_path}\Data\Week {Week}\clients io.csv"
df_CLIENTS_AXM: str = rf"{file_path}\Data\Week {Week}\clients axm.csv"

df_CLIENTS = pd.concat([pd.read_csv(df_CLIENTS_IO,encoding = 'latin',
                                    usecols = ['Company Name                   ','OFFNO','Account','Stat A/C']),
                        pd.read_csv(df_CLIENTS_AXM,encoding = 'latin',
                                    usecols = ['Company Name                   ','OFFNO','Account','Stat A/C'])])
    
df_CLIENTS.columns = ['Client Name','OFFNO','Account','Account Code']
df_CLIENTS.loc[df_CLIENTS['Account Code'].isna(), 'Account Code'] = df_CLIENTS.loc[df_CLIENTS['Account Code'].isna(),'Account']
df_CLIENTS['Client Name'] = df_CLIENTS['Client Name'].str.upper()
df_CLIENTS.sort_values(["OFFNO"], inplace = True)
df_CLIENTS = df_CLIENTS.reset_index(drop = True)
df_CLIENTS.drop_duplicates(subset ="OFFNO",keep = "last", inplace = True)
df_CLIENTS.drop_duplicates(subset ="Client Name",keep = "last", inplace = True)

df_groups = pd.read_excel(groups_path)
df_groups['Client Name']= df_groups['Client Name'].str.upper()
df_groups['Name Change'] = df_groups['Name Change'].str.upper()

CORE_ACCOUNTS = ['Adam Shaw','Dave Levenston','Gerry Hunnisett','Sam Amos']

for i, row in df_groups.iterrows():
    df_CLIENTS.loc[df_CLIENTS['Client Name'] == row['Client Name'], 'Group Name'] = row['Name Change']

df_COREDATA = pd.read_excel(margins_path, sheet_name= 'Core Data',usecols = ['Client Name','PAYNO','Margins','CHQDATE','Solution','Count of', "CRM"]).rename(columns={"CRM":"Account Owner"})

df_COREDATA = df_COREDATA[df_COREDATA['CHQDATE'].dt.month == int(month_num)]
df_COREDATA = df_COREDATA[df_COREDATA['CHQDATE'].dt.year == taxYear().yearc]
df_COREDATA['CHQDATE'] = df_COREDATA['CHQDATE'].dt.strftime('%d/%m/%Y')
df_COREDATA['Client Name'] = df_COREDATA['Client Name'].str.upper()

df_margins = df_COREDATA[['Client Name','Margins','CHQDATE','Count of',"Account Owner"]]
df_margins = pd.merge(df_margins, df_CLIENTS, how = 'left',validate= 'many_to_one')

for i, row in df_margins[~df_margins["Group Name"].isna()].iterrows():
    if row["Account Owner"] in CORE_ACCOUNTS:
        df_margins.loc[df_margins["Group Name"] == row["Group Name"], "Account Owner"] = row["Account Owner"]
        
df_margins = df_margins[df_margins['Margins'] != 0]
df_margins['Average Margin'] = df_margins['Margins']

df_margins.loc[~df_margins['Group Name'].isna(), 'Client Name'] = df_margins.loc[~df_margins['Group Name'].isna(), 'Group Name']
df_margins.loc[~df_margins['Account Owner'].isin(CORE_ACCOUNTS), 'Account Owner'] = 'Other Account'

df_rebate = pd.read_excel(df_rebate_path ,usecols=['Merit Name','Account Code','Group Sum']).rename(columns={'Group Sum':'Rebate'})

df_rebate['Account Code'] = df_rebate['Account Code'].astype(str)

df_pivot = pd.pivot_table(df_margins, values=['Count of','Margins','Average Margin'],
                          index=['Client Name','Account Code','OFFNO','Account Owner'], 
                          columns=['CHQDATE'],
                          aggfunc={'Count of': np.sum, 'Margins': np.sum,'Average Margin': np.mean}, 
                          fill_value=0, 
                          margins=True)

df_revenue = df_pivot.loc[:,('Margins','All')].reset_index().droplevel(1,axis=1).rename(columns={'Margins':'Revenue'})
df_average_margin = df_pivot.loc[:,('Average Margin','All')].reset_index().droplevel(1,axis=1)
df_pivot = df_pivot.drop(('Margins'),axis=1).drop(('Average Margin'),axis=1).droplevel(0,axis=1).reset_index().rename(columns={'index':'Client Name'})
df_pivot = df_pivot[df_pivot['Client Name'] != 'All']
df_pivot = pd.merge(df_pivot, df_average_margin, validate= 'one_to_one')
df_pivot = pd.merge(df_pivot, df_revenue, validate= 'one_to_one')
df_pivot['Average Margin'] = df_pivot['Average Margin'].round(2)

d = dict({'Client Name':'first','OFFNO':'first','Account Owner':'first'})
for i in list(df_margins['CHQDATE'].unique()):
    d.update(dict({i:sum}))
d.update(dict({'All':sum,'Average Margin':np.mean,'Revenue':sum}))
df_pivot = df_pivot.groupby(['Account Code']).agg(d).reset_index()

df_pivot = pd.merge(df_pivot, df_rebate, how='outer', validate= 'one_to_one')
df_unmerged_rebates = df_pivot[df_pivot['Client Name'].isnull() == True][['Merit Name','Rebate','Account Code']]
df_pivot = df_pivot[df_pivot['Client Name'].isnull() == False].drop(['Merit Name'],axis = 1).fillna(0)
    

d = dict({'OFFNO':'first','Account Code':'first','Account Owner': "first"})
for i in list(df_margins['CHQDATE'].unique()):
    d.update(dict({i:sum}))
d.update(dict({'All':sum,'Average Margin':np.mean,'Revenue':sum,'Rebate':sum}))
df_pivot = df_pivot.groupby(['Client Name']).agg(d).reset_index()

df_pivot['Profit'] = df_pivot['Revenue'] - df_pivot['Rebate']

df_pivot_1 = df_pivot[df_pivot['Account Owner'] == 'Other Account']
df_pivot = df_pivot[df_pivot['Account Owner'] != 'Other Account']

df_pivot = df_pivot.append(df_pivot.sum(numeric_only=True), ignore_index=True)
df_pivot_1 = df_pivot_1.append(df_pivot_1.sum(numeric_only=True), ignore_index=True)

df_pivot.loc[df_pivot['Client Name'].isnull() == True, 'Client Name'] = 'Totals'
df_pivot.loc[df_pivot['Client Name'] == 'Totals', 'OFFNO'] = ''
df_pivot.loc[df_pivot['Client Name'] == 'Totals', 'Average Margin'] = df_pivot.loc[df_pivot['Client Name'] != 'Totals', 'Average Margin'].mean()

df_pivot_1.loc[df_pivot_1['Client Name'].isnull() == True, 'Client Name'] = 'Totals'
df_pivot_1.loc[df_pivot_1['Client Name'] == 'Totals', 'OFFNO'] = ''
df_pivot_1.loc[df_pivot_1['Client Name'] == 'Totals', 'Average Margin'] = df_pivot_1.loc[df_pivot_1['Client Name'] != 'Totals', 'Average Margin'].mean()

writer = pd.ExcelWriter(df_path, engine='xlsxwriter')

df_pivot.to_excel(writer, index=False, sheet_name='Net Margins Core Accounts')
df_pivot_1.to_excel(writer, index=False, sheet_name='Net Margins Other Accounts')
df_COREDATA.to_excel(writer, index=False, sheet_name='Margin Data')
df_unmerged_rebates.to_excel(writer, index=False, sheet_name='Unmerged Rebates')

wb = writer.book
money_fmt = wb.add_format({'num_format': '£#,##0'})

ws = writer.sheets['Net Margins Core Accounts']
ws.set_column('J:M', 12, money_fmt)
ws = writer.sheets['Net Margins Other Accounts']
ws.set_column('J:M', 12, money_fmt)

writer.save()
writer.close()
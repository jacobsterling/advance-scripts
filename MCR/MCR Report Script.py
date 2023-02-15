# -*- coding: utf-8 -*-
"""
Created on Thu Oct 14 14:40:15 2021

@author: jacob.sterling
"""

from pathlib import Path

import numpy as np
import pandas as pd
import win32com.client as client
from utils.functions import PAYNO_Check, age

# converts the name of a variable into a string, used to name the sheets on the MCR file
def get_var_name(res):
    name = [x for x in globals() if globals()[x] is res][0]
    return name[:31] if len(name) >= 31 else name

# Maps a merit pay discription to an hours multiplier
payDesc = dict([('Company Income',1),
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
                 ('Standard week',1),
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
                 ('Margin Refund',0),
                 ('SPP',0)
                 ])

# identifies the day rates in the above map
payDescDayRate = ['Day rate EDU - TEN',
                'Day Rate EDU - TEN',
                'Daily Rate',
                'Day rate EDU',
                'Day Rate EDU',
                'Day Rate EDU - GSL',
                'Day Rate EDU - Coba',
                'RSS DAY']

# Reads the MCR file from the current working directory (cwd), parses the start date and end date column into a datetime object
mcr = pd.read_csv("MCR.csv", encoding = 'latin', parse_dates=["START.DATE","END.DATE"], date_parser=lambda x: pd.to_datetime(x, format="%d.%m.%Y"))

# Assures the payno column is of type float
mcr['PAYNO'] = mcr['PAYNO'].astype(float)

# Reads joiners error from cwd
joiners = pd.read_csv("Joiners Error Report.csv",# name of joiners file
                    encoding = 'latin',
                    usecols = ['Pay No','Sdc Option', 'Type', 'Date of Birth','DOJ', "Job Title","Ws Id Proof"],# can specify excel column range e.g. "A:B"
                    low_memory=False)

# Removes all rows from joiners with invalid payroll numbers
joiners = joiners[joiners["Pay No"].apply(lambda x: PAYNO_Check(x))]
joiners['Pay No'] = joiners['Pay No'].astype(float)

#reads salary sacrifice (ss) and drops all rows with no ss
ss = pd.read_csv('Salary Sacrifice.csv', encoding = 'latin', usecols=['PAYNO', 'DED_ONGOING']).dropna()

ss = ss[ss["PAYNO"].apply(lambda x: PAYNO_Check(x))]
ss ['PAYNO'] = ss['PAYNO'].astype(float)

#converts into pounds
ss['DED_ONGOING']=ss['DED_ONGOING'].str.replace(',', '').astype(float)/100

# for backdated hours check
UmbflipsExpenseCheck = pd.read_excel("Umbrella flips expenses.xlsx", sheet_name="Sheet1")

# merges mcr to the salary sacrifice report
mcr = mcr.merge(ss, how = "left")

# any row without a salary sacrifice will be taken as 0
mcr['DED_ONGOING'] = mcr['DED_ONGOING'].fillna(0)

# each row of the MCR will be compressed into single rows of this dataframe
res = pd.DataFrame([], columns = [
    'PAYNO', 
    'T/S NUMBER',
    'TEMPNAME',
    'COMPNAME',
    'TOTAL HOURS',
    'TOTAL PAY',
    'CONTRACTING RATE',
    'COMPANY INCOME TOTAL',
    'DAY RATE TOTAL',
    'DAY RATE TYPE',
    'SALARY SACRIFICE',
    "START DATE","END DATE"
])

# index of operating row (mcr row where a payno is not null)
n = -1

# iterates through each row of the MCR
for i, item in mcr.iterrows():
    
    # checks for unmapped pay descriptions
    if item['PAY_DESC'] in payDesc:
        
        #calculates hours on row
        hours = item['HOURS']*payDesc[item['PAY_DESC']]
        
        #calculates pay on row and assumes no pay if hours < 0
        pay = item['HOURS']*item['PAY_RATE'] if hours > 0 else 0
        
        # contracting rate calculation
        cr = (pay - item['DED_ONGOING'])/hours if hours > 0 else 0

        # identifies if the current row is a operating row or not
        if not pd.isnull(item['PAYNO']):
            
            # creates a new row for the res dataframe
            row = pd.DataFrame([[
                item['PAYNO'], item['T/S Number'], item['TEMPNAME'], item['COMPNAME'], hours, pay, cr, 
                pay if item['PAY_DESC'] == 'Company Income' else 0,
                item['PAY_RATE'] if item['PAY_DESC'] in payDescDayRate else 0,
                payDesc[item['PAY_DESC']] if item['PAY_DESC'] in payDescDayRate else 0,
                item['DED_ONGOING'] , item["START.DATE"], item["END.DATE"]
                ]], columns = res.columns)
            
            # adds it to the res dataframe
            res = pd.concat([res, row]).reset_index(drop = True)
            
            # since the current row was identified as a operating row, its index is assigned to n
            n += 1
        else:
            # since non operating row, hours are added to the operating row
            hours += res.at[n, 'TOTAL HOURS']
            
            # if hours are positive, pay is added to total
            if hours > 0:
                pay += res.at[n, 'TOTAL PAY']
                
                # contracting rate recalculated with new total pay
                res.at[n, 'CONTRACTING RATE'] = (pay - res.at[n, 'SALARY SACRIFICE'])/hours
            else:
                pay = 0
                # not sure this shoud be here
                # res.at[n, 'CONTRACTING RATE'] = 0
            
            # assign summed hours and pay to the result
            res.at[n, 'TOTAL HOURS'] = hours
            res.at[n, 'TOTAL PAY'] = pay
            
            # if company income, add to company income total
            if item['PAY_DESC'] == 'Company Income':
                res.at[n, 'COMPANY INCOME TOTAL'] = res.at[n, 'COMPANY INCOME TOTAL'] + pay 

            # if day rate, pay rate < day rate total or day rate total is = 0, set new pay rate
            if item['PAY_DESC'] in payDescDayRate and (item['PAY_RATE'] < res.at[n, 'DAY RATE TOTAL'] or res.at[n, 'DAY RATE TOTAL'] == 0):
                res.at[n, 'DAY RATE TOTAL'] = item['PAY_RATE']
                res.at[n, 'DAY RATE TYPE'] = payDesc[item['PAY_DESC']]
    else:
        # flags a unmapped pay description, if you get this error, just add it to the payDesc var
        print('Error - Undefined Pay Description : ',item['PAY_DESC'], ' for PAYNO: ', item['PAYNO'])

# copy to result for later reference
result = res

#remove negative pay workers from result
negative = res[res['TOTAL PAY'] <= 0]
res = res[res['TOTAL PAY'] > 0]

# final calculation on company income
res['COMPANY INCOME TOTAL'] = res['COMPANY INCOME TOTAL']/res['TOTAL HOURS']

# round numbers
res['TOTAL HOURS'] = res['TOTAL HOURS'].round(decimals=1)
res[['TOTAL PAY','CONTRACTING RATE','SALARY SACRIFICE']] = res[['TOTAL PAY','CONTRACTING RATE','SALARY SACRIFICE']].round(decimals=2)

# merge to joiners
res = pd.merge(res, joiners, left_on = 'PAYNO', right_on = 'Pay No', how = 'left').drop(['Pay No'], axis = 1)
res['PAYNO'] = res['PAYNO'].astype(int).round()

# fill all empty values with a empty string
res["Ws Id Proof"] = res["Ws Id Proof"].fillna("")

# gets all restricted ids
retricted_ids =   res[res["Ws Id Proof"].str.contains("estrict")]

# gets company income over 75/hour and hours < 10
company_income_above_75_low_hours = res[(res['COMPANY INCOME TOTAL'] >= 75) & (res['TOTAL HOURS'] <= 10)]

# flags day rates that are too low for multiple conditions:
day_rate_2_low = pd.concat([res[(res['DAY RATE TOTAL'] < 86.25) & (res['DAY RATE TYPE'] == 7.50)], 
                               res[(res['DAY RATE TOTAL'] < 69.00) & (res['DAY RATE TYPE'] == 6.00)], 
                               res[(res['DAY RATE TOTAL'] < 74.75) & (res['DAY RATE TYPE'] == 6.50)], 
                               res[(res['DAY RATE TOTAL'] < 115.00) & (res['DAY RATE TYPE'] == 10.00)]])

# flags day rates over 7 days
day_rate_over_7d = res[(res['DAY RATE TYPE'] > 0)]
day_rate_over_7d = day_rate_over_7d[(day_rate_over_7d['TOTAL HOURS']/day_rate_over_7d['DAY RATE TYPE'] > 7)]

# all day rate checks completed, so columns are dropped
res = res.drop(['DAY RATE TOTAL','DAY RATE TYPE','COMPANY INCOME TOTAL'], axis = 1)

# creates missing DOB email
missing_DOB = res.loc[pd.isnull(res['Date of Birth'])]

# removes above workers from result
res = res.loc[~pd.isnull(res['Date of Birth'])]

# workers under 18 flagged for complience
under18 = res.loc[res['Date of Birth'].apply(age) < 18]

# join date defined for later use
joinDate = pd.to_datetime('01/10/2021', format='%d/%m/%Y')

#flags fixed expenses under givenn contracting rate
fixed_expenses_u14 = res[(res['Sdc Option'] == 'Fixed Expenses') & (res['CONTRACTING RATE'] < 14.00)]

#calculation for adjusting the salary sacrifice amount to suit a given minimum contracting rate (cr)
def ADJ_SS(dataframe, cr):
    for i, item in dataframe.iterrows():
        if item['SALARY SACRIFICE'] > 0:
            dataframe.at[i, 'ADJ SS'] = (item['TOTAL PAY'] - cr*item['TOTAL HOURS'])
        else:
            dataframe.at[i, 'ADJ SS'] = 0
    return dataframe

# adjusts ss for following workers to acheive contracting rate of £14
fixed_expenses_u14 = ADJ_SS(fixed_expenses_u14, 14)

# flags CIS u13
CIS_u13 = res[(res['Type'] == 'CIS') & (res['CONTRACTING RATE'] < 13)]
    
CIS_u13 = ADJ_SS(CIS_u13,13)

# flags workers under sdc with a contracting rate over 12.50 and age over 23
uSDC_o23 = res[(res['Sdc Option'] == 'Under SDC') & (res['CONTRACTING RATE'] < 12.50) & (res['Date of Birth'].apply(age) >= 23)]

# exceptions to the above rule are taken out of the data and applied there own minimum contracting rate
exceptions = ['PROMAN RECRUITMENT LTD','DANIEL OWEN LTD','JAMES GRAY TRADES LTD', 'JAMES GRAY RECRUITMENT LTD', 'SEARCH CONSULTANCY LIVERPOOL', 'SEARCH CONSULTANCY DUNDEE', 'SEARCH CONSULTANCY MANCHESTER','SEARCH CONSULTANCY LEEDS']
uSDC_o23_exceptions = uSDC_o23[(uSDC_o23['CONTRACTING RATE'] < 12.19) & (uSDC_o23.COMPNAME.isin(exceptions))]
uSDC_o23_exceptions = ADJ_SS(uSDC_o23_exceptions,12.50)

# removing workers in exeptions from the following data
uSDC_o23 = uSDC_o23[~uSDC_o23.COMPNAME.isin(exceptions)]
uSDC_o23 = ADJ_SS(uSDC_o23,12.50)

# workers over 21 and under 23 with a contracting rate of 12.06
o21_u22 = res[(res['Date of Birth'].apply(age) >= 21) & (res['Date of Birth'].apply(age) < 23) & (res['CONTRACTING RATE'] < 12.06)]
o21_u22 = ADJ_SS(o21_u22,12.06)

# workers over 18 and under 21 with a contracting rate of 8.92
o18_u21 = res[(res['Date of Birth'].apply(age) >= 18) & (res['Date of Birth'].apply(age) < 21) & (res['CONTRACTING RATE'] < 8.92)]
o18_u21 = ADJ_SS(o18_u21,8.92)

# workers under 18 and cr under 6.22
u18 = res[(res['Date of Birth'].apply(age) <= 18) & (res['CONTRACTING RATE'] < 6.22)]
u18 = ADJ_SS(u18,5.71)

# workers over 90 hours
o90_hours = res[res['TOTAL HOURS'] > 90]

#calculating backdated hours check
backdated = result[(result["END DATE"] <= pd.to_datetime("02/10/2022", format="%d/%m/%Y")) & (result["PAYNO"].isin(UmbflipsExpenseCheck["Payno"].values))]
fordated = result[(result["END DATE"] > pd.to_datetime("02/10/2022", format="%d/%m/%Y")) & (result["PAYNO"].isin(UmbflipsExpenseCheck["Payno"].values))]
both = result[(result["PAYNO"].isin(backdated["PAYNO"].values)) & (result["PAYNO"].isin(fordated["PAYNO"].values))].sort_values(by= "PAYNO")

backdated = backdated[~backdated["PAYNO"].isin(both["PAYNO"].values)].sort_values(by= "PAYNO")
fordated = fordated[~fordated["PAYNO"].isin(both["PAYNO"].values)].sort_values(by= "PAYNO")
              
              
with pd.ExcelWriter('backdated hours.xlsx') as writer:
    for report in [backdated, both]:
        report.to_excel(writer, sheet_name = get_var_name(report), index = False)

# any new reports you want in the excel attachment of the MCR need to go in here
reports = [
uSDC_o23,
uSDC_o23_exceptions,
o21_u22,
o18_u21,
u18,
CIS_u13,
fixed_expenses_u14,
o90_hours,
company_income_above_75_low_hours,
day_rate_2_low,
day_rate_over_7d,
missing_DOB,
res
]

# creating MCR excel file
with pd.ExcelWriter('MCR Report.xlsx') as writer:
    for report in reports:
        report.to_excel(writer, sheet_name = get_var_name(report), index = False)

# connects to outlook, outlook needs to be runnning
outlook = client.Dispatch('Outlook.Application')

#creating email templates
email = outlook.CreateItem(0)
email.To = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = 'MCR Report - Missing DOB'

html = """
    </div>
    <div>
        <b> Missing DOB or Not In Joiners Error <b><br><br>
    </div>
    <div>
        {table2}<br><br><br>
    </div>
    
"""

# places the table in the email body
email.HTMLBody = html.format(table2 = missing_DOB.to_html(index=False))

#displays the email on your screen, can use email.Send() to automatically send
email.Display()
        
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
        <b> Over 90 Hours <b><br><br>
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


email = outlook.CreateItem(0)
email.To = 'payroll@advance.online; grace.webber@advance.online'
email.CC = 'jacob.sterling@advance.online ; joshua.richards@advance.online'
email.Subject = 'MCR Report'

email.HTMLBody = html.format(table1 = uSDC_o23.to_html(index=False),
                             table2 = uSDC_o23_exceptions.to_html(index=False),
                             table3 = o21_u22.to_html(index=False),
                             table4 = o18_u21.to_html(index=False),
                             table5 = u18.to_html(index=False),
                             table6 = CIS_u13.to_html(index=False),
                             table7 = fixed_expenses_u14.to_html(index=False),
                             table8 = o90_hours.to_html(index=False),
                             table9 = company_income_above_75_low_hours.to_html(index=False),
                             table10 = day_rate_2_low.to_html(index=False),
                             table11 = day_rate_over_7d.to_html(index=False),
                             exceptions = exceptions)

email.Attachments.Add(Source=str(Path().absolute() / "MCR Report.xlsx"))
email.Display()

email = outlook.CreateItem(0)
email.To = 'enquiries@advance.online; hannah.jarvis@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = 'MCR Report - Under 18'

html = """
    </div>
    <div>
        <b> Under 18 <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
"""

email.HTMLBody = html.format(table1 = under18.to_html(index=False))
email.Display()

email = outlook.CreateItem(0)
email.To = 'hannah.jarvis@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = 'MCR Report - Restricted Ids'

html = """
    </div>
    <div>
        <b> Restricted Ids <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
"""

email.HTMLBody = html.format(table1 = retricted_ids.to_html(index=False))
email.Display()


email = outlook.CreateItem(0)
email.To = 'payroll@advance.online; hannah.jarvis@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = 'MCR Report - Backdated Hours'

html = """
    </div>
    <div>
        <b> Backdated <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
    <div>
        <b> Both <b><br><br>
    </div>
    <div>
        {table2}<br><br><br>
    </div>
"""

email.Attachments.Add(Source=str(Path().absolute() / "backdated hours.xlsx"))
email.HTMLBody = html.format(table1 = backdated.to_html(index=False), table2 = both.to_html(index=False))
email.Display()


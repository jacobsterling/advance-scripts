
from pathlib import Path

import pandas as pd
import win32com.client as client
import xlsxwriter
from openpyxl.utils import get_column_letter
from utils.formats import taxYear
from utils.functions import PAYNO_Check, tax_calcs

year = taxYear().Year("-")
week = tax_calcs().tax_week()

homePath = Path.home() / "advance.online"


marginsPath = homePath / rf"J Drive - Exec Reports/Margins Reports/Margins {year}"

dataPath = marginsPath / rf"Data/Week {week - 1}"

marginsReport = pd.read_excel(marginsPath / rf"Margins Report 2022-2023.xlsx", sheet_name=['Core Data', "Accounts 2", "Clients"])

core_accounts = ["Martin Brown", "Gerry Hunnisett", "Adam Shaw", "Sam Amos", "Dave Levenston"]

accounts = marginsReport["Accounts 2"][["Office Number", "Account Owner"]]
accounts = accounts.dropna(subset=['Office Number'])
accounts['Office Number'] = accounts['Office Number'].astype(int)
accounts = accounts.drop_duplicates(subset=['Office Number'], keep='first')
accounts.loc[~accounts["Account Owner"].isin(core_accounts), "Account Owner"] = "Unmanned Account"

core = marginsReport["Core Data"]

clients = marginsReport["Clients"][["Company Name                   ","OFFNO"]].rename(columns = {"Company Name                   ":"Client Name"})
clients.sort_values("Client Name", inplace = True)
clients["Client Name"] = clients["Client Name"].str.upper()
clients.drop_duplicates(subset ="Client Name",
                        keep = "last", inplace = True)

prevWeek = core[(core["Week Number"] == week - 1) & (core["Margins"] > 0)]

prevPivot = pd.pivot_table(data= prevWeek, values="Margins", index=["CRM", "Client Name"], aggfunc={"Margins": len}).rename(columns={"Margins": "Prev Week"}).reset_index()

prevWeek = prevWeek[prevWeek['PAYNO'].apply(lambda x: PAYNO_Check(x))]
prevWeek['PAYNO'] = prevWeek['PAYNO'].astype(int)

date = tax_calcs().chqdate(week)

io = pd.read_csv("margins.csv")
axm = pd.read_csv("margins axm.csv")
paye = pd.read_csv("margins paye.csv")

paye.loc[paye["FREQ"] == "W", "MANAGEMENT FEE"] = 1

margins = pd.concat([io, axm, paye])
margins = margins[(margins['PAYNO'].apply(lambda x: PAYNO_Check(x))) & (margins["MANAGEMENT FEE"] > 0)]
margins['PAYNO'] = margins['PAYNO'].astype(int)

missingWorkers = prevWeek[~prevWeek["PAYNO"].isin(margins["PAYNO"].unique())][["Client Name", "PAYNO","Surname","Forename","Solution.1","CRM"]].drop_duplicates(subset="PAYNO").reset_index(drop=True).fillna('')
missingAgencies = prevPivot[~prevPivot["Client Name"].isin(margins["COMPNAME"].unique())][["Client Name", "CRM"]].drop_duplicates(subset="Client Name").reset_index(drop=True).fillna('')

pivot = pd.pivot_table(data= margins, values="MANAGEMENT FEE", index="COMPNAME", aggfunc={"MANAGEMENT FEE": len}).reset_index().rename(columns={"MANAGEMENT FEE": "Total", "COMPNAME": "Client Name"})

pivot = pivot.merge(prevPivot, how="left").drop(columns = ["CRM"])

pivot["Difference"] = pivot["Total"] - pivot["Prev Week"]

pivot = pivot.merge(clients, how="left")

pivot = pivot.merge(accounts, left_on="OFFNO", right_on="Office Number", how="left").drop(columns=["OFFNO", "Office Number"])

pivot.loc[len(pivot)] = pivot.sum(numeric_only=True)

pivot.at[len(pivot) - 1, "Client Name"] = "Total"

marginsTotal = round(pivot.at[len(pivot) - 1, "Total"])

pivot = pivot.fillna('')

wb = xlsxwriter.Workbook("margins.xlsx")

format1 = wb.add_format({'bg_color': '#FFC7CE',
                        'font_color': '#9C0006'})
        
format2 = wb.add_format({'bg_color': '#C6EFCE',
                        'font_color': '#006100'})
        
cell_format_column = wb.add_format({'font_size' : 16,
                                    'align': 'center',
                                    'bg_color': '#FFFF00',
                                    'border':1})

ws = wb.add_worksheet('Summary')

for j, column in enumerate(pivot.columns.values):
    col = get_column_letter(j + 1)
    rowend = len(pivot)+1
    REF = f'{col}{1}'
    ws.write(REF,column,cell_format_column)
    ws.set_column(f'{col}:{col}', 15)
    if column in ["Difference", "Variance"]:
        ws.conditional_format(f'{col}{2}:{col}{rowend}', {'type': 'cell',
                                        'criteria': '<',
                                        'value': 0,
                                        'format': format1})
        
        ws.conditional_format(f'{col}{2}:{col}{rowend}', {'type': 'cell',
                                        'criteria': '>',
                                        'value': 0,
                                        'format': format2})
        
for i, row in pivot.iterrows():
    j = 0
    for item in row:
        REF = rf'{get_column_letter(j + 1)}{i + 2}'
        ws.write(REF,item)
        j += 1

ws = wb.add_worksheet('Missing Workers')

for j, column in enumerate(missingWorkers.columns.values):
    rowend = len(missingWorkers) + 1
    REF = f'{get_column_letter(j + 1)}{1}'
    ws.write(REF,column,cell_format_column)
    ws.set_column(f'{get_column_letter(j + 1)}:{get_column_letter(j + 1)}', 15)
        
for i, row in missingWorkers.iterrows():
    j = 0
    for item in row:
        REF = f'{get_column_letter(j + 1)}{i + 2}'
        ws.write(REF,item)
        j += 1
        
ws = wb.add_worksheet('Missing Agencies')

for j, column in enumerate(missingAgencies.columns.values):
    rowend = len(missingWorkers)+1
    REF = f'{get_column_letter(j + 1)}{1}'
    ws.write(REF,column,cell_format_column)
    ws.set_column(f'{get_column_letter(j + 1)}:{get_column_letter(j + 1)}', 15)
        
for i, row in missingAgencies.iterrows():
    j = 0
    for item in row:
        REF = f'{get_column_letter(j + 1)}{i + 2}'
        ws.write(REF,item)
        j += 1

wb.close()

outlook = client.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.To = 'Joshua.Richards@advance.online; Adam.Shaw@advance.online; Gerry.Hunnisett@advance.online; Sam.Amos@advance.online; Dave.Levenston@advance.online; Martin.Brown@advance.online; Jodie.Beeston@advance.online; Alwyn.Barrow@advance.online; harriet.murray@advance.online; jake.price@advance.online; anna.sills@advance.online;'
email.CC = 'jacob.sterling@advance.online'
email.Subject = ('Margins per agency')

email.HTMLBody = rf"""
    <div>
        Hi All, 
        <br>
        <br>
        See the updated margins report
        <br>
        <br>
        Margins: {marginsTotal}
        <br>
        <br>
        <b>Kind regards</b>,
        <br>
        <br>
        <font color='#3F43AD'>
        <b>Jacob Sterling</b>
        <br>
        Reporting and IT Trainee 
        <br>
        <br>
        Office: 01244 564 564
        <br>
        Email: jacob.sterling@advance.online
        <br>
        Visit: <a href='https://www.advance.online/'>www.advance.online</a>
        <br>
        <br>
        <em>Service is important to us, and we value your feedback. Please tell us how we did today by clicking <a href='https://www.google.com/search?rlz=1C1GCEA_enGB894GB894&ei=jRPJX56yDr2p1fAP0piSqAw&q=advance+contracting&gs_ssp=eJzj4tFP1zfMSDYsNsyxLDdgtFI1qDCxME9MNjY1MDG3SEw1tDC3MqhINjSyTE00SE1MszAzSExN9RJOTClLzEtOVUjOzyspSkwuycxLBwARHBbA&oq=advance+contract&gs_lcp=CgZwc3ktYWIQAxgAMgsILhDHARCvARCTAjIICC4QxwEQrwEyCAguEMcBEK8BMgIIADIICC4QxwEQrwEyBQgAEMkDMggILhDHARCvATICCAAyAggAMgIIADoKCAAQsQMQgwEQQzoOCC4QsQMQgwEQxwEQowI6CAgAELEDEIMBOgQIABBDOgUIABCxAzoLCC4QsQMQxwEQowI6CAguEMcBEKMCOgoILhDHARCjAhBDOg0ILhDHARCjAhBDEJMCOgsILhDHARCvARCRAjoQCC4QsQMQxwEQowIQQxCTAjoFCAAQkQI6BwgAELEDEEM6CwguELEDEMcBEK8BOg0ILhCxAxDHARCjAhBDOgoILhDHARCvARAKOgQIABAKOgIILlCFFFjCOGDAQGgCcAF4AIABxQGIAZIYkgEEMS4xN5gBAKABAaoBB2d3cy13aXrAAQE&sclient=psy-ab\#lrd=0x487ac350478ae187:0xc129ea0eaf860aee,3,,,https://www.google.com/search?rlz=1C1GCEA_enGB894GB894&ei=jRPJX56yDr2p1fAP0piSqAw&q=advance+contracting&gs_ssp=eJzj4tFP1zfMSDYsNsyxLDdgtFI1qDCxME9MNjY1MDG3SEw1tDC3MqhINjSyTE00SE1MszAzSExN9RJOTClLzEtOVUjOzyspSkwuycxLBwARHBbA&oq=advance+contract&gs_lcp=CgZwc3ktYWIQAxgAMgsILhDHARCvARCTAjIICC4QxwEQrwEyCAguEMcBEK8BMgIIADIICC4QxwEQrwEyBQgAEMkDMggILhDHARCvATICCAAyAggAMgIIADoKCAAQsQMQgwEQQzoOCC4QsQMQgwEQxwEQowI6CAgAELEDEIMBOgQIABBDOgUIABCxAzoLCC4QsQMQxwEQowI6CAguEMcBEKMCOgoILhDHARCjAhBDOg0ILhDHARCjAhBDEJMCOgsILhDHARCvARCRAjoQCC4QsQMQxwEQowIQQxCTAjoFCAAQkQI6BwgAELEDEEM6CwguELEDEMcBEK8BOg0ILhCxAxDHARCjAhBDOgoILhDHARCvARAKOgQIABAKOgIILlCFFFjCOGDAQGgCcAF4AIABxQGIAZIYkgEEMS4xN5gBAKABAaoBB2d3cy13aXrAAQE&sclient=psy-ab'>here</a> to leave us a review on Google</em>
        </font>
        <br>
        <br>
        <img src="{str(Path.home() / 'OneDrive - advance.online/Documents/signature.png')}">
    <\div>
"""

email.Attachments.Add(Source=str(Path().absolute() / "margins.xlsx"))

email.Display()
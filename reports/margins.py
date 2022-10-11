from turtle import home
import pandas as pd
from pathlib import Path
from utils.formats import taxYear
from utils.functions import tax_calcs
from utils.functions import PAYNO_Check
from openpyxl.utils import get_column_letter
import xlsxwriter
        
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

# joiners_io = pd.read_csv(homePath / "J Drive - Operations/Reports/MCR/Joiners Error Report.csv", encoding="latin")
# joiners_io = joiners_io[joiners_io['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
# joiners_io['Pay No'] = joiners_io['Pay No'].astype(int)

# joiners_axm = pd.read_csv(dataPath / "Joiners Error Report axm.csv", encoding="latin")
# joiners_axm = joiners_axm[joiners_axm['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
# joiners_axm['Pay No'] = joiners_axm['Pay No'].astype(int)

# joiners_paye = pd.read_csv(dataPath / "Joiners Error Report paye.csv", encoding="latin")
# joiners_paye = joiners_paye[joiners_paye['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
# joiners_paye['Pay No'] = joiners_paye['Pay No'].astype(int)

io = pd.read_csv("margins.csv")#.merge(joiners_io, left_on = 'PAYNO',right_on = 'Pay No', how='left')
axm = pd.read_csv("margins axm.csv")#.merge(joiners_axm, left_on = 'PAYNO',right_on = 'Pay No', how='left')
paye = pd.read_csv("margins paye.csv").drop_duplicates(subset=["PAYNO"]).reset_index(drop=True)#.merge(joiners_paye, left_on = 'PAYNO',right_on = 'Pay No', how='left')

margins = pd.concat([io, axm, paye])
margins = margins[(margins['PAYNO'].apply(lambda x: PAYNO_Check(x))) & (margins["MANAGEMENT FEE"] > 0)]
margins['PAYNO'] = margins['PAYNO'].astype(int)

# margins["Solution"] = margins["TYPE"]
# margins.loc[margins["Type"] == "Under SDC", "Solution"] = "Umbrella"
# margins.loc[(margins["Type"] == "Not Under SDC") & (margins["TYPE"] == "PAYE"), "Solution"] = "Umbrella"

missingWorkers = prevWeek[~prevWeek["PAYNO"].isin(margins["PAYNO"].unique())][["Client Name", "PAYNO","Surname","Forename","Solution.1","CRM"]].drop_duplicates(subset="PAYNO").reset_index(drop=True).fillna('')
missingAgencies = prevPivot[~prevPivot["Client Name"].isin(margins["COMPNAME"].unique())][["Client Name", "CRM"]].drop_duplicates(subset="Client Name").reset_index(drop=True).fillna('')

pivot = pd.pivot_table(data= margins, values="MANAGEMENT FEE", index="COMPNAME", aggfunc={"MANAGEMENT FEE": len}).reset_index().rename(columns={"MANAGEMENT FEE": "Total", "COMPNAME": "Client Name"})

pivot = pivot.merge(prevPivot, how="left").drop(columns = ["CRM"])

pivot["Difference"] = pivot["Total"] - pivot["Prev Week"]

pivot = pivot.merge(clients, how="left")

pivot = pivot.merge(accounts, left_on="OFFNO", right_on="Office Number", how="left").drop(columns=["OFFNO", "Office Number"])

pivot.loc[len(pivot)] = pivot.sum(numeric_only=True)

pivot.loc[len(pivot) - 1, "Client Name"] = "Total"

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
    rowend = len(missingWorkers)+1
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


from pathlib import Path

import pandas as pd
from utils.functions import PAYNO_Check

homePath = Path.home() / "advance.online"

mcrPath = homePath / rf"J Drive - Operations\Reports\MCR"

def get_var_name(res):
    name = [x for x in globals() if globals()[x] is res][0]
    return name[:31] if len(name) >= 31 else name

joiners  = pd.read_csv(mcrPath / "Joiners Error Report.csv",usecols=['Pay No',"Sdc Option", "Fixed Expense Value", "Type"], encoding = 'latin', low_memory=False)
joiners = joiners[joiners['Pay No'].apply(lambda x: PAYNO_Check(x))].drop_duplicates(subset = 'Pay No')
joiners['Pay No'] = joiners['Pay No'].astype(int)
core_data = pd.read_csv(mcrPath / "Expenses Created.csv", encoding = 'latin', low_memory=False).merge(joiners, left_on = 'PAYNO',right_on = 'Pay No', how='left').drop(columns = ['Pay No'])

core_data.loc[core_data["Receipt Required"].isna(), "Receipt Required"] = "No"

unmerged = core_data[core_data["Sdc Option"].isna()]
expenses = core_data[~core_data["Sdc Option"].isna()]

fixed_expenses = expenses[expenses["Sdc Option"] == "Fixed Expenses"]
undersdc = expenses[expenses["Sdc Option"] == "Under SDC"]
mileage = expenses[expenses["Sdc Option"] == "Mileage"]
cis = expenses[expenses["Type"] == "CIS"]
other = expenses[(~expenses["Sdc Option"].isin(["Fixed Expenses", "Under SDC", "Mileage"])) | (expenses["Type"] != "CIS")]

undersdc_w_null_rr = undersdc[undersdc["Receipt Required"] == "No"]

fixed_expenses_w_value_0 = expenses[(expenses["Receipt Required"] == "No") & (fixed_expenses["EXPENSE_DESC"] != "Mileage") & (expenses["Fixed Expense Value"] == 0)]

fixed_expenses_w_value_25 = fixed_expenses[(fixed_expenses["Fixed Expense Value"] == 25) & (fixed_expenses["EXPENSE_DESC"] != "Subsistence £25") & (fixed_expenses["Receipt Required"] == "No")]

fixed_expenses_w_value_10 = fixed_expenses[(fixed_expenses["Fixed Expense Value"] == 10) & (fixed_expenses["EXPENSE_DESC"] != "Subsistence £10") & (fixed_expenses["Receipt Required"] == "No")]

fixed_expenses_w_value_5 = fixed_expenses[(fixed_expenses["Fixed Expense Value"] == 5) & (fixed_expenses["EXPENSE_DESC"] != "Subsistence £5") & (fixed_expenses["Receipt Required"] == "No")]

fixed_expenses_w_value_other = fixed_expenses[~fixed_expenses["Fixed Expense Value"].isin([25, 10, 5, 0])]

mileage_claiming_other = mileage[(mileage["Receipt Required"] == "No") & (mileage["EXPENSE_DESC"] != "Mileage")]

fixed_expenses_claiming_accomodation = fixed_expenses[(fixed_expenses["Receipt Required"] == "No") & (expenses["EXPENSE_DESC"] == "Accommodation/Rent")]

with pd.ExcelWriter('Expenses Created Report.xlsx') as writer:
    for report in [undersdc_w_null_rr, fixed_expenses_w_value_25, fixed_expenses_w_value_10, fixed_expenses_w_value_5, fixed_expenses_w_value_0, mileage_claiming_other, fixed_expenses_w_value_other, fixed_expenses_claiming_accomodation, cis, other]:
        sheet_name = get_var_name(report)
        report.to_excel(writer, sheet_name, index = False)
    unmerged.to_excel(writer, sheet_name = "unmerged", index = False)
    expenses.to_excel(writer, sheet_name = "merged", index = False)

import win32com.client as client

email = client.Dispatch('Outlook.Application').CreateItem(0)
email.To = 'enquiries@advance.online; hannah.jarvis@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = ('Enquiries Checks - Expenses Created')
email.Attachments.Add(Source=str(Path().absolute() / "Expenses Created Report.xlsx"))
email.Display()
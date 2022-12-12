import pandas as pd
from pathlib import Path
from datetime import datetime
from utils.taxYear import taxYear

dataPath = Path().home() / "advance.online"

year = datetime.today().year

monthNum = int(input(rf"Enter Month Number ({year}): "))

if monthNum < 1:
    year -= 1
    monthNum += 12
    
month = datetime.strptime(monthNum, "%m")

accountCodes = pd.read_csv("Accounts.csv")

expenseNominals = pd.read_csv("Nominals for Expenses.csv")

report = pd.read_excel("Expense_Report {month} {year}.xlsx").merge(expenseNominals, how="left")

missingNominals = report.loc[report["Expense Nominal"].isna()]["Expense Catagory"].unique()

for catagory in missingNominals.values:
    nominal = int(input("Enter nominal code for missing expense catagory {row['Expense Catagory']}: "))
    while len(report.loc[report["Expense Nominal"] == nominal]) > 0:
        nominal = int(input("Nominal {nominal} is already enter a another code for missing expense catagory {row['Expense Catagory']}: "))
    report.loc[report['Expense Catagory'] == catagory] = nominal
    expenseNominals = pd.concat([expenseNominals, [nominal, catagory]])
    
total = sum(report[report["Expense.CF.Cash or Credit Card?"] == "Credit Card"]["Expense Amount"])

print("Total Sum: {total}" )


# accountCodes.to_csv("Accounts.csv", index = False)
# expenseNominals.to_csv("Nominals for Expenses.csv", index = False)
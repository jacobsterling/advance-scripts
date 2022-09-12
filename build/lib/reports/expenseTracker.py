# -*- coding: utf-8 -*-
"""
Created on Tue Jul 12 09:33:18 2022

@author: jacob.sterling
"""

import pandas as pd 
from pathlib import Path
from functions import tax_calcs
from functions import PAYNO_Check
from formats import taxYear
from datetime import datetime
#latestOpp["End Date"] = latestOpp["End Date"].apply(lambda x: pd.to_datetime(x)) # format="%d %b, %Y")

homePath = Path().home() / "advance.online"

dataPath = homePath / rf"J Drive - Exec Reports\Margins Reports\Margins {taxYear().Year('-')}"

Week = input("Enter week number: ")

latestOpp = pd.read_csv(dataPath / rf"Data/Week {Week}/Expense+Tracker.csv", parse_dates=["Full Name", "Created Time", "Start Date on Site", "End Date", "Latest Start Date on Site", "Date Last Paid"]).sort_values("Created Time", ascending=False).drop_duplicates(subset="Email (Contact Name)").reset_index(drop=True)

latestOpp.loc[latestOpp["Start Date on Site"].isna(), "Start Date on Site"] = latestOpp.loc[latestOpp["Start Date on Site"].isna(), "Latest Start Date on Site"]

data = pd.read_excel(dataPath / rf"Margins Report 22-23.xlsx", sheet_name=["Core Data", "Joiners Compliance"])
#change to 6
#{taxYear().Year('-')}

df = data["Core Data"][data["Core Data"]["Week Number"].astype(int) >= tax_calcs().tax_week_calc() - 7][["Client Name","PAYNO","CHQDATE", "Type","Email", "Solution", "Solution.1"]]


df = df[(df["Type"] == "Fixed Expenses") | (df["Type"] == "Mileage Only")].sort_values("CHQDATE", ascending=False).drop_duplicates(subset=["PAYNO"]).reset_index(drop=True)

joiners = data["Joiners Compliance"][["Pay No", "WEEKS_PAID", "FIXED_EXPENSE_FREQ", "FIXED_EXPENSE_VALUE"]]
joiners = joiners[joiners["Pay No"].apply(lambda x: PAYNO_Check(x))]

df = df.merge(joiners, left_on="PAYNO", right_on="Pay No", how="left").drop("Pay No", axis = 1)
df = df.merge(latestOpp, left_on="Email", right_on="Email (Contact Name)", how="left").drop("Email (Contact Name)", axis = 1)

df_1 = df[df["Start Date on Site"] + pd.to_timedelta(2*365, unit='d') < datetime.now()]

old_df = pd.read_excel(Path().absolute() / "data/Expense Tracker Data.xlsx", sheet_name="Sheet1")

new = df_1[~df_1["PAYNO"].isin(old_df["PAYNO"])]

with pd.ExcelWriter("data/Expense Tracker Data.xlsx") as writer:
    df_1.to_excel(writer, index= False)
    df.to_excel(writer, sheet_name="Data", index= False)

writer.save()

import win32com.client as client
outlook = client.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.To = "joshua.richards@advance.online"
email.Subject = ('Expense Tracker')
email.HTMLBody = rf"{new.to_html(index=False)}"
email.Attachments.Add(Source=str(Path().absolute() / "data/Expense Tracker Data.xlsx"))
email.Display(True)
email.Send()
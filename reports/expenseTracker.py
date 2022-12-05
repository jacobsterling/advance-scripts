# -*- coding: utf-8 -*-
"""
Created on Tue Jul 12 09:33:18 2022

@author: jacob.sterling
"""

from datetime import datetime
from pathlib import Path

import pandas as pd
from utils.formats import taxYear
from utils.functions import PAYNO_Check, tax_calcs

#latestOpp["End Date"] = latestOpp["End Date"].apply(lambda x: pd.to_datetime(x)) # format="%d %b, %Y")

homePath = Path().home() / "advance.online"

old_df_path = Path().home() / "OneDrive - advance.online/Documents/data/Expense Tracker Data.xlsx"

dataPath = homePath / rf"J Drive - Exec Reports\Margins Reports\Margins {taxYear().Year('-')}"

Week = int(input("Enter week number: "))

prevOpp = pd.read_csv(dataPath / rf"Data/Week {Week - 1}/Expense+Tracker.csv", parse_dates=["Created Time", "Start Date on Site", "End Date", "Latest Start Date on Site", "Date Last Paid"], na_values="-").sort_values("Created Time", ascending=False).drop_duplicates(subset="Email (Contact Name)").reset_index(drop=True)
latestOpp = pd.read_csv(dataPath / rf"Data/Week {Week}/Expense+Tracker.csv", parse_dates=["Created Time", "Start Date on Site", "End Date", "Latest Start Date on Site", "Date Last Paid"], na_values="-", skiprows=6).sort_values("Created Time", ascending=False).drop_duplicates(subset="Email (Contact Name)").reset_index(drop=True)

tracker = pd.concat([latestOpp, prevOpp]).drop_duplicates(subset="Record Id")

tracker.loc[tracker["Start Date on Site"].isna(), "Start Date on Site"] = tracker.loc[tracker["Start Date on Site"].isna(), "Latest Start Date on Site"]

data = pd.read_excel(dataPath / rf"Margins Report 2022-2023.xlsx", sheet_name=["Core Data", "Joiners Compliance"])

df = data["Core Data"][data["Core Data"]["Week Number"].astype(int) >= int(Week) - 7][["Client Name","PAYNO","CHQDATE", "Type","Email", "Solution", "Solution.1"]]


df = df[(df["Type"] == "Fixed Expenses") | (df["Type"] == "Mileage Only")].sort_values("CHQDATE", ascending=False).drop_duplicates(subset=["PAYNO"], keep="last").reset_index(drop=True)

joiners = data["Joiners Compliance"][["Pay No", "WEEKS_PAID", "FIXED_EXPENSE_FREQ", "FIXED_EXPENSE_VALUE"]]
joiners = joiners[joiners["Pay No"].apply(lambda x: PAYNO_Check(x))]

df = df.merge(joiners, left_on="PAYNO", right_on="Pay No", how="left").drop("Pay No", axis = 1)
df = df.merge(tracker, left_on="Email", right_on="Email (Contact Name)", how="left").drop("Email (Contact Name)", axis = 1)

df_1 = df[df["Start Date on Site"] + pd.to_timedelta(2*365, unit='d') < datetime.now()]

old_df = pd.read_excel(old_df_path, sheet_name="24 Months")

new = df_1[~df_1["PAYNO"].isin(old_df["PAYNO"])]

with pd.ExcelWriter(old_df_path) as writer:
    df_1.to_excel(writer, index= False, sheet_name="24 Months")
    df.to_excel(writer, sheet_name="Data", index= False)

writer.save()

import win32com.client as client

outlook = client.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.To = "joshua.richards@advance.online"
email.Subject = ('Expense Tracker')
email.HTMLBody = rf"""
New Workers To The Report
<br>
{
    new.to_html(index=False)
}"""
email.Attachments.Add(str(old_df_path))
email.Display(True)
#email.Send()
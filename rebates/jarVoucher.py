# -*- coding: utf-8 -*-
"""
Created on Wed May 18 10:40:54 2022

@author: jacob.sterling
"""

#run Jar Opportunities - Incentives

import pandas as pd
from pathlib import Path
from datetime import datetime
from utils import formats

#Advance2018!

Year = formats.taxYear().Year("-")

Week = input("Enter Week to update with: ")

rootPath = Path.home() / rf"advance.online/J Drive - Exec Reports/Margins Reports/Margins {Year}"

dataPath = rootPath / rf"Data/Week {Week}"

reportPath = rootPath / rf"Margins Report {formats.taxYear().Year('-')}.xlsx"

previouslyPaid = pd.read_csv("Jar Voucher Paid.csv")

jarOpportunites = pd.read_csv(dataPath / "Jar+Opportunities+-+Incentives.csv", na_values="-") #, skiprows=6
jarOpportunites["Margin Accrual"] = jarOpportunites["Margin Accrual"].fillna("0")

jarOpportunites["Margin Accrual"] = jarOpportunites["Margin Accrual"].str.replace("£ ","").astype(float)

report = pd.read_excel(reportPath, sheet_name="Core Data", parse_dates=["CHQDATE"])
report = report[report["Client Name"] == "JAR SOLUTIONS"]

latestOpp = pd.read_csv(dataPath / "Expense+Tracker.csv", parse_dates=["Created Time", "Start Date on Site", "End Date", "Latest Start Date on Site", "Date Last Paid"]).sort_values("Created Time", ascending=False).drop_duplicates(subset="Email (Contact Name)").reset_index(drop=True)

df = report.merge(latestOpp, left_on="Email", right_on="Email (Contact Name)", how="left").drop("Email (Contact Name)", axis = 1)

df = df[(df["CHQDATE"] >= df["Created Time"]) & (~df["PAYNO"].isin(previouslyPaid["PAYNO"]))]

df = df.groupby(['PAYNO']).agg({"Margins":sum, "Email":"first", "Record Id":"first"}).reset_index(drop=True)

df.to_csv("Backup.csv", index=False)

jarFeesRetained = None
for file in dataPath.glob("*"):
    if file.name.__contains__("io"):
        if file.name.__contains__("ees retained"):
            feesRetained = pd.read_csv(file)
        if file.name.__contains__("oiners"):
            joiners = pd.read_csv(file, usecols=["Pay No", "Email Address"], encoding = "latin").rename(columns={"Pay No": "PAYNO", "Email Address":"Email (Contact Name)"})

feesRetained["PAYNO"] = feesRetained["PAYNO"].astype(int)

def isInt(x):
    try:
        int(x)
        return True
    except ValueError:
        return False

joiners = joiners[joiners["PAYNO"].apply(lambda x: isInt(x))]

joiners["PAYNO"] = joiners["PAYNO"].astype(int)

feesRetained = feesRetained[feesRetained["Client Name"] == "JAR SOLUTIONS"]

feesRetained = feesRetained.groupby("PAYNO").agg({"Management Fee":sum, "Solution":"first"}).reset_index()

feesRetained = feesRetained.merge(joiners, how = "left", validate="one_to_many")

crmimport = feesRetained.merge(jarOpportunites, how = "left")

crmimport["Margin Accrual"] = crmimport["Margin Accrual"] + crmimport["Management Fee"]

crmimport.loc[(crmimport["Margin Accrual"] >= 102) & ((crmimport["Solutions"] == "Umbrella") | (crmimport["Solutions"] == "Umbrella no Expenses")) ,"Consultant Voucher Received"] = "£50 Voucher"
crmimport.loc[(crmimport["Margin Accrual"] >= 102) & ((crmimport["Solutions"] == "Umbrella with Mileage") | (crmimport["Solutions"] == "Umbrella with Expenses")) ,"Consultant Voucher Received"] = "£75 Voucher"

crmimport.loc[ (crmimport["PAYNO"].isin(previouslyPaid["PAYNO"])) & (crmimport["Consultant Voucher Received On"].isna()),"Consultant Voucher Received On"] = datetime.now().strftime("%d/%m/%Y")
crmimport.loc[ crmimport["PAYNO"].isin(previouslyPaid["PAYNO"]) & ((crmimport["Solutions"] == "Umbrella") | (crmimport["Solutions"] == "Umbrella no Expenses")) ,"Consultant Voucher Received"] = "£50 Voucher"
crmimport.loc[ crmimport["PAYNO"].isin(previouslyPaid["PAYNO"]) & ((crmimport["Solutions"] == "Umbrella with Mileage") | (crmimport["Solutions"] == "Umbrella with Expenses")) ,"Consultant Voucher Received"] = "£75 Voucher"

crmimport = crmimport.dropna(subset=["Record Id"])

crmimport.to_csv("Voucher Opportunity Import.csv", index=False)
# -*- coding: utf-8 -*-
"""
Created on Wed May 18 10:40:54 2022

@author: jacob.sterling
"""

#run Jar Opportunities - Incentives

import pandas as pd
from pathlib import Path
from datetime import datetime
from Formats import taxYear

Year = taxYear().Year("-")

Week = input("Enter Week to update with: ")

dataPath = Path.home() / rf"advance.online/J Drive - Exec Reports/Margins Reports/Margins {Year}/Data/Week {Week}"

previouslyPaid = pd.read_csv("Previously Paid.csv")

jarOpportunites = pd.read_csv("Jar+Opportunities+-+Incentives.csv", skiprows=6, na_values="-")
jarOpportunites["Margin Accrual"] = jarOpportunites["Margin Accrual"].fillna("0")
jarOpportunites["Margin Accrual"] = jarOpportunites["Margin Accrual"].str.replace("£ ","").astype(float)


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
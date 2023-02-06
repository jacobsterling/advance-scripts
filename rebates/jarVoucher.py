# -*- coding: utf-8 -*-
"""
Created on Wed May 18 10:40:54 2022

@author: jacob.sterling
"""

#run Jar Opportunities - Incentives

from datetime import datetime
from pathlib import Path

import pandas as pd
from utils import formats

#1120

Year = formats.taxYear().Year("-")

Week: int = input("Enter Week to update with: ")

rootPath = Path.home() / rf"advance.online/J Drive - Exec Reports/Margins Reports/Margins {Year}"

dataPath = rootPath / rf"Data/Week {Week}"

reportPath = rootPath / rf"Margins Report {formats.taxYear().Year('-')}.xlsx"

previouslyPaid = pd.read_csv("jarPaid.csv")

jarOpportunites = pd.read_csv(dataPath / "Jar+Opportunities+-+Incentives.csv", na_values="-", skiprows=6)

report = pd.read_excel(reportPath, sheet_name="Core Data", parse_dates=["CHQDATE"])
report = report[report["Client Name"] == "JAR SOLUTIONS"]

df = report.merge(jarOpportunites, left_on="Email", right_on="Email (Contact Name)", how="left").drop("Email (Contact Name)", axis = 1)

df = df[(df["CHQDATE"] >= df["Created Time"]) & (~df["PAYNO"].isin(previouslyPaid["PAYNO"]))]

df = df.groupby(['PAYNO']).agg({"Margins":sum, "Email":"first", "Record Id":"first", "Solutions":"first", "PAYNO": "first", "Consultant Voucher Received On": "first" }).reset_index(drop=True)

df.loc[(df["Margins"] >= 102) & ((df["Solutions"] == "Umbrella") | (df["Solutions"] == "Umbrella no Expenses")) ,"Consultant Voucher Received"] = "£50 Voucher"

df.loc[(df["Margins"] >= 102) & ((df["Solutions"] == "Umbrella with Mileage") | (df["Solutions"] == "Umbrella with Expenses")) ,"Consultant Voucher Received"] = "£75 Voucher"

df.loc[ (df["PAYNO"].isin(previouslyPaid["PAYNO"])) & (df["Consultant Voucher Received On"].isna()),"Consultant Voucher Received On"] = datetime.now().strftime("%d/%m/%Y")

df.to_csv("jarImport.csv", index=False)

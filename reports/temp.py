# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from pathlib import Path
import pandas as pd
from Formats import taxYear

Week = input("Enter Week Number: ")
Year = taxYear().Year('-')

homePath = Path.home() / rf"advance.online/J Drive - Exec Reports/Margins Reports/Margins {Year}/Data/Week {Week}"

joinersAXM = None
joinersColumns = None

for file in homePath.glob('*'):
    if file.isfile() and file.suffix == ".xlsx":
        if file.name.__contains__("fees retained"):
            if file.name.__contains__("axm"):
                feesRetainedAXM = pd.read_excel(file)
            else:
                feesRetainedIO = pd.read_excel(file)
        if file.name.__contains__('oiners'):
            if file.name.__contains__("axm"):
                joinersAXM = pd.read_excel(file)
            else:
                joinersIO = pd.read_excel(file)

feesRetained = pd.concat([feesRetainedIO,feesRetainedAXM])

if joinersAXM:
    joiners = pd.concat([joinersIO, joinersAXM])
else:
    joiners = joinersIO

df = pd.merge(feesRetained, joiners, how="left", left_on="", right_on="", validate='one-one')
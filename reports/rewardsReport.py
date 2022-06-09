# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from pathlib import Path
import pandas as pd
from Formats import taxYear
import numpy as np
from Functions import PAYNO_Check

Week = input("Enter Week Number: ")
Year = taxYear().Year('-')

homePath = Path.home() / rf"advance.online/J Drive - Exec Reports/Margins Reports/Margins {Year}/Data/Week {Week}"

joinersColumns = ['Pay No','MOBILE', 'Email Address','FEE_TYPE', 'REWARDS']

for file in homePath.glob('*'):
    if file.suffix == ".csv":
        if file.name.__contains__("fees retained"):
            if file.name.__contains__("axm"):
                feesRetainedAXM = pd.read_csv(file, encoding="latin")
            else:
                feesRetainedIO = pd.read_csv(file, encoding="latin")
        if file.name.__contains__('oiners'):
            if file.name.__contains__("axm"):
                joinersAXM = pd.read_csv(file, encoding="latin", usecols=joinersColumns)
            else:
                joinersIO = pd.read_csv(file, encoding="latin", usecols=joinersColumns)
        if file.name.__contains__('Paid+wast+week+w_+rewards'):
            rewards = pd.read_csv(file, skiprows=(6),usecols=['Email','ADVANCE Rewards'])

joinersIO = joinersIO[joinersIO['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
joinersIO['Pay No'] = joinersIO['Pay No'].astype(int)

joinersAXM = joinersAXM[joinersAXM['Pay No'].apply(lambda x: PAYNO_Check(x)) == True]
joinersAXM['Pay No'] = joinersAXM['Pay No'].astype(int)

df_io = pd.merge(feesRetainedIO, joinersIO, how="left", left_on="PAYNO", right_on="Pay No")#, validate='many_to_one'
df_axm = pd.merge(feesRetainedAXM, joinersAXM, how="left", left_on="PAYNO", right_on="Pay No")#, validate='many_to_one'

df = pd.concat([df_io,df_axm]).drop('Pay No', axis=1)

df = df.merge(rewards, how="left", left_on="Email Address", right_on="Email", validate='many_to_one').drop('Email', axis=1)

df.loc[df['ADVANCE Rewards'] == np.nan , 'ADVANCE Rewards'] = df.loc[df['ADVANCE Rewards']== np.nan , 'REWARDS']

df = df[df['ADVANCE Rewards'] == 'Yes'].drop('REWARDS', axis = 1)

df = df[df['Management Fee'] > 0].reset_index(drop=True)


invalidChars = ['£', ' ', 'PT', 'AM', 'P',]
invalidFee = []

for i, row in df.iterrows():
    feeType = row['FEE_TYPE']
    if pd.isnull(feeType):
        df.at[i, 'Fee Type'] = row['Management Fee'] - 1.99
    else:
        for char in invalidChars:
            feeType = feeType.replace(char,'')
        try:
            feeType = float(feeType)
            df.at[i, 'Fee Type'] = feeType
        except ValueError:
            invalidFee.append(feeType)
            df.at[i, 'Fee Type'] = row['Management Fee'] - 1.99
    if df.at[i, 'Fee Type'] > 120:
        feeType = str(df.at[i, 'Fee Type'])
        try:
            pre0 = int(feeType[0])
            pre1 = int(feeType[1])
            post0 = int(feeType[2])
            post1 = int(feeType[3])
            df.at[i, 'Fee Type'] = float(f"{pre0}{pre1}.{post0}{post1}")
        except ValueError:
            invalidFee.append(feeType)
            df.at[i, 'Fee Type'] = row['Management Fee'] - 1.99
    df.at[i, 'Margin w/out Rewards'] = row['Management Fee'] - 1.99
    df.at[i, 'Number Of Margins'] = df.at[i, 'Margin w/out Rewards']/df.at[i, 'Fee Type']

df.to_excel(f'Rewards Report Week {Week}.xlsx',sheet_name='Rewards',index=False)

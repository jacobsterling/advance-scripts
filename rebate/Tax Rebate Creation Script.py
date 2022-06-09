# -*- coding: utf-8 -*-
"""
Created on Fri Feb 11 12:23:26 2022

@author: jacob.sterling
"""

import pandas as pd
import numpy as np

df = pd.read_csv("Agent Authorities.csv", encoding = 'latin')
df_CRM = pd.read_csv("CRM Report.csv", encoding = 'latin')

df_CRM.columns = df_CRM.iloc[0, :]
df_CRM = df_CRM.drop(0, axis = 0)
df_CRM = df_CRM[:-3]
df_CRM = df_CRM.drop(['CONTACTID'], axis = 1)
df_CRM['UTR'] = df_CRM['UTR'].astype('int64')

df = pd.merge(df,df_CRM,left_on='Tax Reference number', right_on='UTR',how='left').fillna(0).drop('UTR',axis=1)

df_unadded = df[df['Email'] == 0]
df = df[df['Email'] != 0]

df['Tax Year'] = '20/21'
df['Tax Rebate Name'] = df['Full Name'] + ' ' + df['Tax Year']
df['Bank Details on Booklet'] = 'Not Checked'
df['Tax Rebate Service Status'] = 'Awaiting Agent Authority'

df.to_csv('CRM Import.csv',index=False)
df_unadded.to_csv('Tax Rebates Not Created.csv',index=False)
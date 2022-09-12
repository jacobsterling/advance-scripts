# -*- coding: utf-8 -*-
"""
Created on Tue May 10 14:07:18 2022

@author: jacob.sterling
"""

from experisPayrollScript import experisPDFReader
import pandas as pd

Week = input("Enter Week Number: ")

df = pd.DataFrame([],columns = ['Worker Name', 'Date', 'Description', 'Hours', 'Rate', 'Amount', 'File Name'])
totals = pd.DataFrame([],columns = ['Amount', 'Gross', 'PDF Total', 'Difference'])

for i in range(1, 4):
    batch, total = experisPDFReader(Week, str(i))
    totals = pd.concat([totals, total])
    df = pd.concat([df, batch])


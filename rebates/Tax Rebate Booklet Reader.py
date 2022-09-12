# -*- coding: utf-8 -*-
"""
Created on Fri Feb 18 16:37:55 2022

@author: jacob.sterling
"""

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
import pandas as pd
from pathlib import Path
import numpy as np

df = pd.DataFrame([],columns=['Fornames', 'Surname', 'Date of Birth',
       'National Insurance Number', 'Unique Tax Reference', 'Nationality',
       'Postcode', 'Mobile Number', 'Email Address', 'Address',
       'Variable Textbox 1', 'Variable Textbox 2',
       'Income - Employer ID 1', 'Income - Employer PAYE 1',
       'Income - Total Gross 1', 'Income - Total Tax 1',
       'Income - Employer ID 2', 'Income - Employer PAYE 2',
       'Income - Total Gross 2', 'Income - Total Tax 2',
       'Income - Employer ID 3', 'Income - Employer PAYE 3',
       'Income - Total Gross 3', 'Income - Total Tax 3',
       'Income - Employer ID 4', 'Income - Employer PAYE 4',
       'Income - Total Gross 4', 'Income - Total Tax 4',
       'Income - Employer ID 5', 'Income - Employer PAYE 5',
       'Income - Total Gross 5', 'Income - Total Tax 5',
       'Income - Employer ID 6', 'Income - Employer PAYE 6',
       'Income - Total Gross 6', 'Income - Total Tax 6',
       'Income - Employer ID 7', 'Income - Employer PAYE 7',
       'Income - Total Gross 7', 'Income - Total Tax 7',
       'Income - Employer ID 8', 'Income - Employer PAYE 8',
       'Income - Total Gross 8', 'Income - Total Tax 8',
       'Income - Employer PAYE 9', 'Income - Employer ID 9',
       'Income - Employer ID 10', 'Income - Employer PAYE 10',
       'Income - Total Gross 9', 'Income - Total Gross 10',
       'Income - Total Tax 9', 'Variable Textbox 3',
       'SE Income - Payer 1', 'SE Income - Gross 1',
       'SE Income - CIS Tax 1', 'SE Income - Payer 2',
       'SE Income - Gross 2', 'SE Income - CIS Tax 2',
       'SE Income - Payer 3', 'SE Income - Gross 3',
       'SE Income - CIS Tax 3', 'SE Income - Payer 4',
       'SE Income - Gross 4', 'SE Income - CIS Tax 4',
       'SE Income - Payer 5', 'SE Income - Gross 5',
       'SE Income - CIS Tax 5', 'SE Income - Total Gross',
       'SE Income - Total CIS Tax', 'SE Income - Payer 6',
       'SE Income - Payer 7', 'SE Income - Gross 6',
       'SE Income - Gross 7', 'SE Income - Gross 8',
       'SE Income - CIS Tax 6', 'SE Income - CIS Tax 7',
       'SE Income - CIS Tax 8', 'Variable Textbox 4', 'SEISS Amount',
       'Other Taxable Benefits - Type 1',
       'Other Taxable Benefits - Type 2',
       'Other Taxable Benefits - Type 3', 'Paid by ADVANCE',
       'Oher Taxable Benefits - Gross 1',
       'Oher Taxable Benefits - Gross 2',
       'Oher Taxable Benefits - Gross 3', 'Oher Taxable Benefits - Tax 1',
       'Oher Taxable Benefits - Tax 2', 'Oher Taxable Benefits - Tax 3',
       'Other Income - Type 1', 'Other Income - Type 2',
       'Other Income - Type 3', 'Other Income - Gross 1',
       'Other Income - Gross 2', 'Other Income - Gross 3',
       'Other Income - Tax 1', 'Other Income - Tax 2',
       'Other Income - Tax 3', 'Variable Textbox 7', 'Variable Textbox 8',
       'Variable Textbox - Small Date', 'Expenses - Car Miles',
       'Expenses - Motorcycle Miles', 'Expenses - Vehicle Registration',
       'Expenses - Bicycle Miles', 'Expenses - Road Tax',
       'Expenses - Vehicle Insurance', 'Expenses - Fuel',
       'Expenses - Repairs', 'Expenses - Servicing', 'Expenses - Bus',
       'Expenses - Taxi', 'Expenses - Flights', 'Expenses - Train',
       'Expenses - Parking & Tolls', 'Expenses - Meals',
       'Expenses - Accommodation', 'Expenses - Clothing',
       'Expenses - Uniform Cleaning', 'Expenses - Materials',
       'Expenses - Tools', 'Expenses - Work Phone',
       'Expenses - Stationary', 'Expenses - Computer', 'Expenses - Bank',
       'Expenses - Rentals', 'Expenses - Fees',
       'Expenses - Subscriptions', 'Expenses - Training',
       'Expenses - Licenses', 'Expenses - Bank Interest',
       'Expenses - Insurances', 'Student Loan Deductions by Employers',
       'Marriage Allowance - Received', 'Marriage Allowance - Given',
       'Expenses - Other', 'Bank Details - Bank Name',
       'Bank Details - Account Name', 'Bank Details - Sort Code',
       'Bank Details - Account Number', 'Claimed SEISS Radio',
       'Outstanding Student Loan', 'Student Loan Paid Off 2Y',
       'Variable Textbox 5', 'Copyright', 'Title', 'Marital Status',
       'Variable Textbox 6', 'Tax Rebate Year', 'Student Loan Plan',
       'Variable Textbox 9', 'Mileage 1', 'Mileage 2', 'Mileage 3'])

pdf_list = list()

#file_path = Path(r'C:\Users\jacob.sterling\advance.online\J Drive - Advance Accounting Solutions\Tax Rebate Service Current\4) 20-21 Booklets\Ready to Submit\1) Finalised Booklets\2. Outstanding')
#file_path = Path(r'C:\Users\jacob.sterling\OneDrive - advance.online\Documents\Read Booklet')
file_path = Path(r'C:\Users\jacob.sterling\OneDrive - advance.online\Documents')


for pdf in file_path.glob('*'):
    if pdf.is_file():
        if pdf.suffix in [".PDF",".pdf"] and pdf.name.__contains__('Rebate Booklet'):
            pdf_list.append(pdf.name)
            print(f'Reading {pdf.name}......')
            fp = open(pdf, 'rb')
            parser = PDFParser(fp)
            doc = PDFDocument(parser)
            fields = resolve1(doc.catalog['AcroForm'])['Fields']
            df_temp = pd.DataFrame([],columns=['Name','Value'])
            for i in fields:
                field = resolve1(i)
                name, value = field.get('T'), field.get('V')
                try:
                   value = value.name
                except AttributeError:
                   if value is not None:
                       try:
                           value = value.decode('utf-8')
                       except UnicodeDecodeError:
                           value = value.decode('iso-8859-1')
                   else:
                       value = 'None'
                value = value.replace('None','').replace('"',"").replace("'","").replace(r"\r",", ")
                
                name = name.decode('utf-8').replace('None','').replace('"',"").replace("'","").replace(r"\r",", ")
                
                df_temp = pd.concat([df_temp,pd.DataFrame([[name,value]],columns=['Name','Value'])]).reset_index(drop=True)
            df_temp = df_temp.T
            df_temp.columns = df_temp.loc['Name',:]
            dfdfdf = df_temp.columns
            df = pd.concat([df, df_temp.drop(['Name'],axis = 0)]).fillna('').reset_index(drop=True)

df.to_excel('output.xlsx',index = False)

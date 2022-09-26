# -*- coding: utf-8 -*-
"""
Created on Tue Apr 26 09:30:36 2022

@author: jacob.sterling
"""

class armPayroll:
    def __init__(self, Week = None):
        import pandas as pd
        from pathlib import Path
        from utils.formats import taxYear
        self.pd = pd
        
        if not Week:
            Week = input('Enter Week Number: ')

        Year = taxYear().Year(' - ')
        
        self.homePath = Path.home() / rf"advance.online/J Drive - Operations/Remittances and invoices/ARM - Advanced Resource Managers"
        self.filePath = self.homePath / rf"Tax Year {Year}/Week {Week}"
        
        pdf_columns = ['Candidate', 'TS Ref', 'Period End', 'Description', 'Net', 'VAT', 'VAT %', 'Total', 'Invoice Number', 'File Name']
        self.df_result = pd.DataFrame([], columns=["Remittance Ref","Invoice Number"])
    
    def readPDF(self):
        import pdfplumber
        import re
        invoicePattern = re.compile(r"^(.*) Invoice No: ([0-9]{6})$")
        remittancePattern = re.compile(r"^(.*) Remittance Ref: ([0-9]{6})$")
        for file in self.filePath.glob('*'):
            if file.is_file() and file.suffix in [".PDF", ".pdf"]:
                print(f'reading {file.name}....')
                pdf = pdfplumber.open(file)
                for page in pdf.pages:
                    text = page.extract_text()
                    for line in text.split("\n"):
                        if remittancePattern.match(line):
                            remittanceRef = re.search(remittancePattern,line).groups()[-1]
                        elif invoicePattern.match(line):
                            invoiceNo = re.search(invoicePattern,line).groups()[-1]
                            self.df_result = self.pd.concat([self.df_result,self.pd.DataFrame([[remittanceRef, invoiceNo]],columns = self.df_result.columns)]).reset_index(drop=True)
                            remittanceRef, invoiceNo = None, None
        return self.df_result.reset_index(drop=True)
                # df_list.append(df_text[j:i])


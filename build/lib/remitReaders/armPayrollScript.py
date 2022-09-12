# -*- coding: utf-8 -*-
"""
Created on Tue Apr 26 09:30:36 2022

@author: jacob.sterling
"""

class armPayroll:
    def __init__(self, Week = None):
        import pandas as pd
        from pathlib import Path
        from Formats import taxYear
        self.pd = pd
        
        if not Week:
            Week = input('Enter Week Number: ')

        Year = taxYear().Year(' - ')
        
        self.homePath = Path.home() / rf"advance.online/J Drive - Operations/Remittances and invoices/ARM - Advanced Resource Managers"
        self.filePath = self.homePath / rf"Tax Year {Year}/Week {Week}"
        
        pdf_columns = ['Candidate', 'TS Ref', 'Period End', 'Description', 'Net', 'VAT', 'VAT %', 'Total', 'Invoice Number', 'File Name']
        self.df_result = pd.DataFrame([], columns=["Remittance Ref","Invoice Number"])
    
    @staticmethod
    def has_numbers(inputString):
        return any(char.isdigit() for char in inputString)
    
    def readPDF(self):
        #from tabula import read_pdf
        from pdfminer.high_level import extract_text

        for pdf in self.filePath.glob('*'):
            if pdf.is_file() and pdf.suffix in [".PDF", ".pdf"]:
                print(f'reading {pdf.name}....')
                df_text, df_list, i, j, n = extract_text(pdf).split('\n'), [], 0, 0, 0
                for row in df_text:
                    if row.__contains__('REMITTANCE') and i != 0:
                        df_list.append(df_text[j:i])
                        j = i
                    i += 1
                
                j = 0
                for row in df_list[j]:
                    if n == 0:
                        n = 1 if row.__contains__("Remittance Ref") else n
                    elif n == 1:
                        try:
                            remittanceRef = int(row)
                            n = 2 
                        except ValueError:
                            pass
                    elif n == 2:
                        try:
                            invoiceNo = int(row)
                            break
                        except ValueError:
                            pass
                self.df_result = self.pd.concat([self.df_result, self.pd.DataFrame([[remittanceRef,invoiceNo]], columns = self.df_result.columns)])
                j += 1
        return self.df_result
        
                # df_list.append(df_text[j:i])
                # df_pdf, j = read_pdf(pdf,pages="all"), 0
                # for df in df_pdf:
                #     n = 0
                    
                #     #paste text here
                            
                #     for i, row in df.iterrows():
                #         for column in df.columns:
                #             if self.pd.isnull(row[column]):
                #                 try:
                #                     df.at[i, column] = df.at[i + 1, column]
                #                 except KeyError:
                #                     pass
                #             try:
                #                 df.at[i, column] = float(df.at[i, column])
                #             except ValueError:
                #                 pass
                    
                    # df['Remittance Ref'] = remittanceRef
                    # df['Invoice Number'] = invoiceNo
                    # df['File Name'] = pdf.name
                    # df.dropna(axis = 0)
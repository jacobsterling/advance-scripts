# -*- coding: utf-8 -*-
"""
Created on Thu Feb 24 09:44:11 2022

@author: jacob.sterling
"""


def scantecPdfReader(Week,Day):
    import pandas as pd
    from pathlib import Path
    from tabula import read_pdf
    from Formats import taxYear
    from Formats import day
    from Functions import has_numbers
    
    Year = taxYear().Year(' - ')
    Day = day().dayFormat(Day)
    
    file_path = Path.home() / rf"advance.online\J Drive - Operations\Remittances and invoices\Scantec\Tax Year {Year}\Week {Week}\{Day}"
    
    df_result = pd.DataFrame([],columns=['Worker Name','UF1','Description','Hours','Rate','Net','File Name'])
    
    df_pdf_collection = list()
    
    df_temp_collection = pd.DataFrame([],columns=[0])
    
    for file in file_path.glob('*'):
        if file.is_file() and file.suffix == ".PDF":
            print(f'Reading {file.name}....')
            df_pdf = read_pdf(file,pages='all',guess=False)
            for df in df_pdf:
                df_pdf_collection.append(df)
                df.columns, n = list(range(0,len(df.columns))), 0
                for i, row in df.iterrows():
                    if n > 0:
                        break
                    for j in df.columns:
                        if str(row[j]).__contains__('WORKER'): 
                            df_temp, n = df.iloc[i+1:,j:], 1
                            break
                df_temp, n = pd.DataFrame(df_temp).reset_index(drop=True), 0
                for i, row in df_temp.iterrows():
                    if n > 0:
                        break
                    for j in df_temp.columns:
                        if str(row[j]).__contains__('Remittance') or str(row[j]).__contains__('CONTINUED'):
                            n = i
                            break
                        elif j > 0:
                            df_temp.at[i, 0] = df_temp.at[i, 0] + ' ' + str(row[j])
                            
                df_temp = pd.DataFrame(df_temp.iloc[:n,0]).reset_index(drop=True)
                
                df_temp_collection = pd.concat([df_temp_collection,df_temp])
                #chqDate = '**/**/**'
                for i, row in df_temp.iterrows():
                    items = row[0].split(' ')
                    del items[0]
                    n = 0
                    name = list()
                    for item in items:
                        if n == 0:
                            UF1 = 'NA'
                            for char in item:
                                if char.isupper():
                                    name.append(char)
                                else:
                                    n = 1
                                if n == 1:
                                    break
                            name.append(' ')
                        if n == 1:
                            if (item.count('/') > 0 or item.count('[') > 0 or item.count(']') > 0) and has_numbers(item):
                                if item.count('[') == 1:
                                    UF1 = list()
                                    m = 0
                                    for char in item:
                                        if char.isnumeric():
                                            if item.count('/') > 0:
                                                UF1 = 'NA'
                                                break
                                            UF1.append(char)
                                        elif char == ']':
                                            m = 1
                                        if m == 1:
                                            break
                                elif item.count('/') == 2 and item.isupper() == False and len(item) == 8:
                                    #chqDate = item
                                    pass
                                continue
                            elif item.isupper() and has_numbers(item):
                                continue
                            else:
                                n = 2
                                
                        if n in [2,3]:
                            if has_numbers(item) == False or n == 3:
                                if n == 2:
                                    desc = item
                                    n = 3
                                elif item.__contains__('.') and has_numbers(item):
                                    try:
                                        hours = float(str(item).replace(',','').replace('£',''))
                                        n = 4
                                    except ValueError:
                                        desc = desc + ' ' + item
                                else:
                                    desc = desc + ' ' + item
                        elif n == 4:
                            try:
                                rate = float(str(item).replace(',','').replace('£',''))
                                n = 5
                            except ValueError:
                                desc = desc + ' ' + str(hours) + ' ' + item
                                n = 3
                                continue
                        elif n == 5 and item.__contains__('.') and has_numbers(item):
                            name = "".join(name)
                            if name[-1] == ' ':
                                name = name[:-1]
                            try:
                                UF1 = "".join(UF1)
                            except TypeError:
                                pass
                            #if len(UF1) > 4:
                            #    UF1 = UF1[:4]
                            try:
                                UF1 = int(UF1)
                            except ValueError:
                                UF1 = 'NA'
                            try:
                                net = float(str(item).replace('L1','').replace('T1','').replace(',','').replace('£',''))
                                df_result = pd.concat([df_result,pd.DataFrame([[name,UF1,desc,hours,rate,net,file.name]],columns=['Worker Name','UF1','Description','Hours','Rate','Net','File Name'])]).reset_index(drop = True)
                            except ValueError:
                                continue
    print('-------------------------------------------------------------------')
    print(f'{Day} Gross: ',df_result['Net'].sum()*1.2)
    print('-------------------------------------------------------------------')
    #df_result.to_excel(file_path / f'Scantec Py Import Week {Week} {Day}.xlsx', index = False)
    df_temp_collection.to_excel(f'temp collection {Week} {Day}.xlsx', index = False)
    return df_result
# -*- coding: utf-8 -*-
"""
Created on Fri Jan 28 14:21:33 2022

@author: jacob.sterling
"""

class intermediaryReportScript():
    def __init__(self):
        import numpy as np
        import pandas as pd
        from pathlib import Path
        from Functions import tax_calcs
        from Formats import taxYear
        self.pd = pd
        self.np = np
        self.QuarterRange, self.Quarter, self.Quarterp = tax_calcs().Quarter_Calc()
        
        #Weekc = tax_calcs().tax_week_calc()
        Weekc = 1
        #Week = self.QuarterRange[-1]
        YearcFormat = taxYear().Year('-') #change to Year()
        self.YearFormat = taxYear().Yearp('-')
        
        YearcFormat1 = taxYear().Yearp_format1('-')#Change to YearcFormat1()
        
        self.report_prefix = 'IntermediaryReport_083_FA45839_'
        trackerName = f'{self.YearFormat}-{self.Quarter} Inter Report Tracker Python New.xlsx'
        
        homePath = Path.home() / "advance.online/J Drive - Operations/Reports/Intermediary Reports"
        dataPath = Path.home() / rf"advance.online/J Drive - Exec Reports/Margins Reports/Margins {YearcFormat}/Data/Week {Weekc}"

        filePathP = homePath / rf"{YearcFormat1}/{self.Quarterp}"
        self.filePath = homePath / rf"{YearcFormat1}/{self.Quarter}"
        self.reportsPath = self.filePath / "Outstanding Reports"
        self.trackerPath = self.filePath / trackerName
        
        
        for file in filePathP.glob('*.xlsx'):
            if file.name.__contains__('racker'):
                self.trackerP = pd.read_excel(file, usecols=['Office Number','Email to be sent to '])
                break#change column names to OFFNO and Merit Email
        
        self.c = 1 if input('Type "y" make Changes to files ?: ') == "y" else 0
        
        df_CLIENTS_IO = dataPath / "clients io.csv"
        df_CLIENTS_AXM = dataPath / "clients axm.csv"
        df_ACCOUNTS = dataPath / "accounts office.csv"
        groups_path = Path.home() / "OneDrive - advance.online/Documents/Data/Groups.xlsx"
        
        df_CLIENTS = pd.concat([pd.read_csv(df_CLIENTS_IO,encoding = 'latin',
                                            usecols = ['Company Name                   ','OFFNO','EMAIL_DETAILS_INTER']),
                                pd.read_csv(df_CLIENTS_AXM,encoding = 'latin',
                                            usecols = ['Company Name                   ','OFFNO','EMAIL_DETAILS_INTER'])])
        
        df_CLIENTS.columns = ['Client Name','OFFNO','Email']
        df_CLIENTS['Client Name'] = df_CLIENTS['Client Name'].str.upper()
        df_CLIENTS.sort_values("Client Name", inplace = True)
        df_CLIENTS.drop_duplicates(subset ="Client Name",
                             keep = "last", inplace = True)
        self.df_CLIENTS = df_CLIENTS
        
        df_ACCOUNTS = pd.read_csv(df_ACCOUNTS, na_values = '-')
        df_ACCOUNTS = df_ACCOUNTS[['Office Number','Account Owner','Send Intermediary Report to','Account Type']]
        df_ACCOUNTS = df_ACCOUNTS.dropna(subset=['Office Number'])
        df_ACCOUNTS['Office Number'] = df_ACCOUNTS['Office Number'].astype(int)
        self.df_ACCOUNTS = df_ACCOUNTS.drop_duplicates(subset=['Office Number'], keep='first')
        
        df_groups = pd.read_excel(groups_path)
        df_groups['Client Name'] = df_groups['Client Name'].str.upper()
        df_groups['Name Change'] = df_groups['Name Change'].str.upper()
        self.df_groups = df_groups
        
        self.missingData = pd.DataFrame([],columns = ['Name','Missing Info'])
        
    def nullQuery(self, query, row, k, df_report, notif = None):
        if self.pd.isnull(row[query]) or row[query] == 'N/A':
            self.missingInfo += 1
            self.details += f'missing {query}, ' if not notif else notif
            if self.c == 1:
                try:
                    user = input(f"Enter missing {query} for {row['Worker forename']} {row['Worker surname']}: ")
                    if user == '':
                        self.missingData.append([[f"{row['Worker forename']} {row['Worker surname']}", query]])
                    else:
                        df_report.at[k, query] = user
                        self.changes += 1
                except KeyError:
                    user = input(f"Enter {notif} for {df_report.columns[1]}: ")
                    if user != '':
                        df_report.at[k, query] = user
                        self.changes += 1
                    else:
                        self.missingInfo += 1
                    self.missingInfo -= 1
        return df_report
                
    def rangeQuery(self, startCol, endCol, df_report_row, k, values, df_report, notif):
        if (df_report.loc[k,startCol:endCol].values != values).all():
            self.missingInfo += 1
            self.details += f'missing {notif}, '
            if self.c == 1:
                df_report.loc[k,startCol:endCol] = values
                self.changes += 1
                print('')
                print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
                print(f"Changed {notif} for {df_report_row['Worker forename']} {df_report_row['Worker surname']}.")
                print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
        return df_report
    
    def matchingQuery(self, condition1, condition2, notif = None):
        if condition1 != condition2:
            self.details += f'{condition1} does not match {condition2}, '
            return False
        return True
    
    def readReport(self,file):
        try:
            df_info = self.pd.read_csv(file,nrows=5, on_bad_lines='skip').iloc[:,:2]
            df_info[df_info.columns[1]] = df_info[df_info.columns[1]].astype(str)
            df_info.loc[df_info[df_info.columns[1]] == 'nan',df_info.columns[1]] = self.np.nan
            return df_info, self.pd.read_csv(file,skiprows=7, on_bad_lines='skip')
        except UnicodeDecodeError:
            df_info = self.pd.read_csv(file,nrows=5,encoding='latin', on_bad_lines='skip').iloc[:,:2]
            df_info[df_info.columns[1]] = df_info[df_info.columns[1]].astype(str)
            df_info.loc[df_info[df_info.columns[1]] == 'nan',df_info.columns[1]] = self.np.nan
            return df_info, self.pd.read_csv(file,skiprows=7,encoding='latin', on_bad_lines='skip')

    def matchReport(self, offno):
        try:
            clientName = self.df_CLIENTS.loc[self.df_CLIENTS['OFFNO']==offno,'Client Name'].values[0]
            clientEmail = self.df_CLIENTS.loc[self.df_CLIENTS['OFFNO']==offno,'Email'].values[0]
        except IndexError:
            clientName = self.np.nan
            clientEmail = self.np.nan
            self.missingInfo += 2
            print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
            print(f'{offno} is missing from clients.csv')
            self.details += f'{offno} is missing from clients.csv, '
            print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
            
        try:
            accountOwner = self.df_ACCOUNTS.loc[self.df_ACCOUNTS['Office Number']==offno,'Account Owner'].values[0]
            accountType = self.df_ACCOUNTS.loc[self.df_ACCOUNTS['Office Number']==offno,'Account Type'].values[0]
            crmEmail = self.df_ACCOUNTS.loc[self.df_ACCOUNTS['Office Number']==offno,'Send Intermediary Report to'].values[0]
        except IndexError:
            accountOwner = self.np.nan
            accountType = self.np.nan
            crmEmail = self.np.nan
    
        self.nullVar('account owner',accountOwner, 1)
        self.nullVar('account type',accountType, 1)
        self.nullVar('crm email',crmEmail)
        if self.nullVar('client email',clientEmail, 1):
            try:
                clientEmail = self.trackerP.loc[self.trackerP['Office Number']==offno,'Email to be sent to '].values[0]
                if clientEmail == 0:
                    clientEmail = self.np.nan 
                    self.details += f'Cannot find email in previous {self.Quarterp} tracker, '
                else:
                    self.details += f'Used email from previous {self.Quarterp} tracker, '
            except IndexError:
                self.missingInfo += 1
                self.details += f'Cannot find email in previous {self.Quarterp} tracker, '
            
        elif clientEmail != crmEmail and not self.pd.isnull(crmEmail):
            self.missingInfo += 1
            self.details += 'Merit Email does not match CRM, '
        
        for i, row in self.df_groups.iterrows():
            if row['Client Name'] == clientName or clientName.__contains__(row['Name Change']):
                clientName = row['Name Change']
                accountType = 'Group'
                self.missingInfo += 1
                self.details += 'Needs to be Grouped, '
                break
    
        return clientName, accountOwner, accountType, clientEmail, crmEmail

    def queryReport(self, df_info, df_report):
        df_iter_info = df_info
        for j, info_row in df_iter_info.iterrows():
            if j == 0:
                df_info = self.nullQuery(df_info.columns[1], info_row, j, df_info,'Missing Address, ')
        df_info = self.nullQuery(df_info.columns[1], info_row, j, df_info,'Missing Post Code, ')

        df_iter_report = df_report
        
        for k, df_report_row in df_iter_report.iterrows():
            
            df_report = self.nullQuery('Worker date of birth', df_report_row, k, df_report)
            
            df_report = self.nullQuery('Worker gender', df_report_row, k, df_report)
            
            df_report = self.nullQuery('Worker National Insurance number', df_report_row, k, df_report)

            df_report = self.nullQuery('Start date of engagement',df_report_row, k, df_report)

            if df_report_row["Worker engagement details where intermediary didn't operate PAYE"] == 'D' or df_report_row["Worker engagement details where intermediary didn't operate PAYE"] == 'B':
                
                df_report = self.rangeQuery("Name of party paid by intermediary for worker's services",
                                "Postcode of party paid by intermediary for worker's services",
                                df_report_row, k,
                                ["Advance Contracting Solutions Ltd","First Floor VISTA","St David's Park","Ewloe","Chester","CH5 3DT"],
                                df_report,
                                'Intermediary Details, ')

                df_report = self.nullQuery("Amount paid for the worker's services" , df_report_row, k, df_report)

                df_report = self.nullQuery("Companies House registration number of party paid by intermediary for worker's services", df_report_row, k, df_report)
                    
            if df_report_row["Worker engagement details where intermediary didn't operate PAYE"] == 'A':
                
                df_report = self.rangeQuery("Name of party paid by intermediary for worker's services",
                                "Postcode of party paid by intermediary for worker's services",
                                df_report_row, k, 
                                ["Advance Contracting Solutions Ltd","First Floor VISTA","St David's Park","Ewloe","Chester","CH5 3DT"],
                                df_report,
                                'Intermediary Details, ')
                        
                df_report = self.nullQuery("Amount paid for the worker's services", df_report_row, k, df_report)

                df_report = self.nullQuery("Worker unique taxpayer reference (UTR)", df_report_row, k, df_report)
                    #add groups together
                    #email potentially ?
                    
        if len(df_report[df_report["Amount paid for the worker's services"].astype(float) == 0]) > 0:
            self.missingInfo += 1
            self.details += "Amount paid for the worker's services <= 0"
            if self.c == 1:
                if input('Amount paid for the workers services <= 0, Remove row ? ("y"): ') == 'y':
                    df_report = df_report[df_report["Amount paid for the worker's services"].astype(float) != 0]
                    self.changes += 1
                    
        return df_info, df_report
    
    def createTracker(self):
        self.df = self.pd.DataFrame([],columns=['File','OFFNO','Client','Account','Account Type','Merit Email','CRM Email','Changes Made','Missing Info','A','F','B','D', 'Details'])
        for file in self.reportsPath.glob('*'):
            if file.is_file():
                if file.suffix == ".csv" and self.report_prefix in file.name: 
                    offno = int(file.name.split('_')[3])
                    
                    print('___________________________________________________________________')
                    print('')
                    print(f'Reading {file.name}')
                    
                    
                    df_info, df_report = self.readReport(file)
                    
                    self.changes, self.missingInfo, self.details  = 0, 0, ''
                    
                    clientName, accountOwner, accountType, clientEmail, crmEmail = self.matchReport(offno)

                    if self.pd.isnull(clientName):
                        clientName = df_info.columns[1].upper()
                        self.details += 'Client name used from file, '
                        
                    if not self.matchingQuery(df_info.columns[1].upper(), clientName):
                        if self.c == 1 and input(f"Miss Matching Client Name {df_info.columns[1].upper()}, type 'y' to change to {clientName}: ") == 'y':
                            self.changes += 1
                            df_info.columns = ['Employment intermediary name',clientName.title()]
                    
                    print('')
                    print(f'Showing Client - {df_info.columns[1].upper()}')
                    print('')
                    
                    df_info, df_report = self.queryReport(df_info, df_report)
                                
                    self.saveChanges(file, df_info, df_report)

                    if accountType != 'Agency':
                        self.missingInfo += 1
                        self.details += 'Not agency, '

                    df_temp = self.pd.DataFrame([[file.name,
                                             offno,
                                             clientName,
                                             accountOwner,
                                             accountType,
                                             clientEmail,
                                             crmEmail,
                                             self.changes,
                                             self.missingInfo,
                                             len(df_report[df_report["Worker engagement details where intermediary didn't operate PAYE"] == 'A']),
                                             len(df_report[df_report["Worker engagement details where intermediary didn't operate PAYE"] == 'F']),
                                             len(df_report[df_report["Worker engagement details where intermediary didn't operate PAYE"] == 'B']),
                                             len(df_report[df_report["Worker engagement details where intermediary didn't operate PAYE"] == 'D']),
                                             self.details]],
                                             columns=['File','OFFNO','Client','Account','Account Type','Merit Email','CRM Email','Changes Made','Missing Info','A','F','B','D','Details'])
                    
                    self.df = self.df.append(df_temp)
                    
        self.exportTracker()
        if input('Type "y" to edit the tracker ?: ') == 'y':
            self.editTracker()
        else:
            return self.df, self.missingData
    
    def nullVar(self, name, var, i = 0):
        if self.pd.isnull(var):
            self.missingInfo += i
            self.details += f'Missing {name}, '
            return True
        else:
            return False
    
    def editTracker(self):
        try:
            self.df = self.pd.read_excel(self.trackerPath)
            for i, row in self.df.iterrows():
                self.changes = row['Changes Made']
                self.details = '' if self.pd.isnull(row['Details']) else row['Details']
                file = self.reportsPath / f"{row['File']}"
                
                print('___________________________________________________________________')
                print('')
                print(f'Reading {file.name}')
                
                clientName = self.df.at[i, 'Client']
                accountOwner = self.df.at[i,'Account']
                accountType = self.df.at[i,'Account Type']
                clientEmail = self.df.at[i,'Merit Email']
                crmEmail = self.df.at[i,'CRM Email']
                self.missingInfo = self.df.at[i,'Missing Info']
                
                if self.changes not in ['Sent']: #,'Send'
                    
                    df_info, df_report = self.readReport(file)
                    self.missingInfo = 0
                    
                    self.changes = 0 if self.changes == 'Send' else self.changes
                    
                    if accountOwner == 'OFFNO not found in accounts office':
                        self.missingInfo += 1
                        self.details += 'Missing Account Owner, '
                    
                    self.nullVar('account type', accountType, 1)
                        
                    if self.nullVar('client email',clientEmail, 1) and not self.nullVar('crm email',crmEmail) and input(f'Type "y" to update {clientName} merit email {clientEmail} with {crmEmail}: ') == 'y':
                        clientEmail = crmEmail
                        self.missingInfo -= 1
                        self.details = self.details.replace('Missing Merit Email, ','')
                        
                    if accountType == 'Group' and self.c == 1:
                        group = self.df.loc[self.df['Client'] == clientName]
                        if len(group) > 1:
                            if input(f"Type 'y' to group {clientName} with groupee's: ") == 'y':
                                changes = self.changes
                                for j, groupee in group.iterrows():
                                    gPath = self.filePath / f"Outstanding Reports\\{groupee['File']}"
                                    if gPath != file:
                                        g_info, g_report = self.readReport(gPath)
                                        self.changes = groupee['Changes Made']
                                        self.changes = 0 if self.changes == 'Send' else self.changes
                                        
                                        print('')
                                        print(f'Showing Groupee - {g_info.columns[1].upper()}')
                                        print('')
                                        
                                        g_info, g_report = self.queryReport(g_info, g_report)
                                        
                                        self.df.at[j, 'Account Type'] = 'Groupee'
                                        
                                        self.saveChanges(file, g_info, g_report)
    
                                        df_report = self.pd.concat([df_report,g_report])
                                        
                                self.changes = changes + 1
                                accountType = 'Grouped'
                            elif input("Type 'y' if already grouped: ") == 'y':
                                accountType = 'Grouped'
                                self.changes += 1
                                
                    if not self.matchingQuery(df_info.columns[1].upper(), clientName):
                        if self.c == 1 and input(f"Miss Matching Client Name {df_info.columns[1].upper()}, type 'y' to change to {clientName}: ") == 'y':
                            self.changes += 1
                            df_info.columns = ['Employment intermediary name',clientName.title()]
                    
                    print('')
                    print(f'Showing {accountType} - {df_info.columns[1].upper()}')
                    print('')
                    
                    df_info, df_report = self.queryReport(df_info, df_report)
                    
                    if accountType != 'Agency':
                        if accountType != 'Grouped':
                            self.missingInfo += 1
                            self.details += 'Not agency, '
                        else:
                            self.details += 'Grouped, '
                            
                    self.saveChanges(file, df_info, df_report)
                            
                    self.df.at[i, ['Client','Account','Account Type','Merit Email','CRM Email','Changes Made','Missing Info','A','F','B','D','Details']] = [clientName,
                    accountOwner,
                    accountType,
                    clientEmail,
                    crmEmail,
                    self.changes,
                    self.missingInfo,
                    len(df_report[df_report["Worker engagement details where intermediary didn't operate PAYE"] == 'A']),
                    len(df_report[df_report["Worker engagement details where intermediary didn't operate PAYE"] == 'F']),
                    len(df_report[df_report["Worker engagement details where intermediary didn't operate PAYE"] == 'B']),
                    len(df_report[df_report["Worker engagement details where intermediary didn't operate PAYE"] == 'D']),
                    self.details]
    
                if self.changes == 'Send' and self.c == 1:
                    if self.missingInfo > 0:
                        s = 1 if input('Type "y" to if you want to compose email despite missing information: ') == 'y' else 0
                    elif input(f'Type "y" to compose email to {clientName}: ') == 'y':
                        s = 1
                    else:
                        s = 0
                    if s == 1:
                        print('')
                        print(f"A: {self.df.at[i,'A']}, F: {self.df.at[i,'F']}, B: {self.df.at[i,'B']}, D: {self.df.at[i,'D']}")
                        print('')
                        self.emailReport(clientEmail,file)
                        self.df.at[i,'Changes Made'] = 'Sent'
                elif self.changes == 'Sent':
                    print('')
                    print('Sent')
        except KeyboardInterrupt:       
            pass
        self.exportTracker()
        return self.df, self.missingData
    
    def saveChanges(self, file, info, report):
        if self.c == 1 and self.changes > 0:
            if input('Type "y" to remove duplicates on NI number: ') == 'y':
                report.drop_duplicates(subset=['Worker National Insurance number',"Worker engagement details where intermediary didn't operate PAYE"],keep = "first", inplace = True)
                
            if input(f'Type "y" to save {self.changes} changes to {info.columns[1].upper()} - {file.name} with {self.missingInfo} missing info: ') == 'y':
                
                reportCol = self.pd.DataFrame([report.columns])
                
                report.columns = range(0, len(report.columns))
                
                emptyRow = self.pd.DataFrame([[self.np.nan for i in report.columns]])
                
                infoCol = self.pd.DataFrame([info.columns])
                
                for i in range(len(info.columns), len(report.columns)):
                    info[i] = self.np.nan
                
                info.columns = report.columns
                
                df = self.pd.concat([infoCol, info, emptyRow, reportCol, report]).reset_index(drop=True)
                
                df.to_csv(file, index=False, header=False)
                
                info = info.iloc[:,:2]
                
                info.columns = infoCol.iloc[0]
                
                report.columns = reportCol.iloc[0]
                
                self.changes = 'Send' if input(f'Type "y" to allow file with {self.changes} changes to be sent to agency (does not send): ') == 'y' else self.changes
                
        else:
            self.changes = 0
            if self.missingInfo == 0:
                self.changes = 'Send'
            else:
                self.changes = 'Send' if input(f'Type "y" to allow file with {self.missingInfo} missing info to be sent to agency (does not send): ') == 'y' else self.changes
            
    def emailReport(self, clientEmail, report):
        import win32com.client as client
        outlook = client.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)
        email.To = clientEmail
        email.Subject = (f'Intermediate Report - {self.Quarter}')
        #email._oleobj_.Invoke(*(64209, 0, 8, 0, info))
        email.SentOnBehalfOfName = 'info@advance.online'
        #info
        html = rf"""
            Hi,
            <br>
            <br>
            Please find attached your intermediary report for {self.Quarter}.
            <br>
            <br>
            Please remember to report on any limited companies you may have contracted with in {self.Quarter}. 
            <br>
            <br>
            I have also attached a PDF document to that you may find useful to review on intermediary reporting.
            <br>
            <br>
            <u><i>Assumptions</i></u>
            <br>
            <br>
            Start date is the date we set the contractor up
            <br>
            <br>
            End date is the day the contractor informed us they are no longer working with ADVANCE
            <br>
            <br>
            If you have any questions please let me know. 
            <br>
            <br>
            Kind regards,
            <br>
            <br>
            ADVANCE
        """
        email.HTMLBody = html
        email.Attachments.Add(Source=str(report))
        email.Attachments.Add(Source=str(self.filePath / "Onshore employment intermediaries - FAQs.pdf"))
        email.Display(True)
        
    def exportTracker(self):
        from openpyxl.utils import get_column_letter
        import xlsxwriter
        
        df = self.df.fillna('').reset_index(drop=True)
        
        wb = xlsxwriter.Workbook(self.trackerPath)
        
        format1 = wb.add_format({'bg_color': '#FFC7CE',
                                       'font_color': '#9C0006'})
        
        format2 = wb.add_format({'bg_color': '#C6EFCE',
                                       'font_color': '#006100'})
        
        cell_format_column = wb.add_format({'font_size' : 16,
                                            'align': 'center',
                                            'bg_color': '#FFFF00',
                                            'border':1})
        
        ws = wb.add_worksheet('Tracker')
        
        for j, column in enumerate(df.columns.values):
            col = get_column_letter(j + 1)
            row = 1
            rowend = len(df)+1
            ws.write(f'{col}{row}',column,cell_format_column)
            ws.set_column(f'{col}:{col}', 15)
            if column == 'Missing Info':
                ws.conditional_format(f'{col}{row + 1}:{col}{rowend}', {'type': 'cell',
                                              'criteria': '>',
                                              'value': 0,
                                              'format': format1})
                
        for i, row in df.iterrows():
            j = 0
            for item in row:
                REF_1 = ('{col}{row}').format(col = get_column_letter(j + 1), row = i + 2)
                if row.index[j] == 'File':
                    #ws.write_url(REF_1, str(self.filePath / "Outstanding Reports") + '\\' + item,string=item)
                    ws.write_url(REF_1, "Outstanding Reports/" + item, string=item)
                elif row.index[j] == 'Changes Made':
                    if item == 'Sent':
                        ws.write(REF_1,item,format2)
                    else:
                        ws.write(REF_1,item)
                        
                elif row.index[j] == 'Account' or row.index[j] == 'Account Type':
                    if item == '':
                        ws.write(REF_1,item,format1)
                    else:
                        ws.write(REF_1,item)
                else:
                    ws.write(REF_1,item)
                j += 1
        wb.close()
        
#df, missingData = intermediaryReportScript().createTracker()
df, missingData = intermediaryReportScript().editTracker()

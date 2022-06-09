# -*- coding: utf-8 -*-
"""
Created on Mon Jan 31 16:59:28 2022

@author: jacob.sterling
"""

import pandas as pd
from pathlib import Path
import win32com.client as client
from Formats import taxYear

Year = taxYear().Year('-')
    
Week = input('Enter Week Number: ')


#run 0joiners_doj

df_path = Path(rf'C:\Users\jacob.sterling\advance.online\J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}\Last Week Setup.csv')
df_path_sample = Path(r'C:\Users\jacob.sterling\OneDrive - advance.online\Documents\Data\Last Week Setup Sample.xlsx')
df = pd.read_csv(df_path,encoding='latin')

df_sample = df[df['WS_ID_RECEIVED'] == 'Yes'].sample(n = round(len(df)*0.1))
df_sample.to_excel(df_path_sample,index=False)

outlook = client.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.Display()
email.To = 'hannah.jarvis@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = ('Random Sample of Last Weeks Setups for Audit')

html = """
    </div>
    <div>
        <b> 10% of Last Weeks Setups with RTW checked <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
"""

email.Attachments.Add(Source=str(df_path_sample))

email.HTMLBody = html.format(table1 = df_sample.to_html(index=False))

email.Send()


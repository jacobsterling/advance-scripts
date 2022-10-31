
import pandas as pd

from utils.formats import taxYear
from pathlib import Path

Year = taxYear().Year('-')

Week = input('Enter Week Number: ')

homePath = Path.home() / "advance.online"

dataPath = homePath / rf"J Drive - Exec Reports\Margins Reports\Margins {Year}\Data\Week {Week}"

joiners  = pd.read_csv(dataPath / "Joiners Error Report io.csv",usecols=['Pay No',"Sdc Option"], encoding = 'latin')
expenses = pd.read_csv("Expenses Created.csv", encoding = 'latin').merge(joiners, left_on = 'PAYNO',right_on = 'Pay No', how='left').drop(columns = ['Pay No'])

expenses.to_csv('Expenses Created.csv', encoding='utf-8',index=False)

len(expenses["PAYNO"].unique())

import win32com.client as client
email = client.Dispatch('Outlook.Application').CreateItem(0)
email.To = 'enquiries@advance.online; hannah.jarvis@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'
email.Subject = ('Enquiries Checks - Expenses Created')

html = """
    </div>
    <div>
        <b> Missing NI Numbers <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
    <div>
        <b> Under 18 <b><br><br>
    </div>
    <div>
        {table2}<br><br><br>
    </div>
"""

email.Attachments.Add(Source=str(Path().absolute() / "Expenses Created.csv"))
email.Display()
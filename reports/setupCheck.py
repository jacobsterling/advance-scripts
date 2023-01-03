from pathlib import Path
import pandas as pd

df = pd.read_csv("Last Week Setup.csv", encoding='latin')

sample_path_1 = "Setup Sample 1.xlsx"
sample_path_2 = "Setup Sample 2.xlsx"

sample = df.sample(n = round(len(df)*0.2))

sample1 = sample.sample(n = round(len(sample)*0.5))
sample2 = sample[~sample["Pay No"].isin(sample1["Pay No"])]

sample1.to_excel(sample_path_1, index=False)
sample2.to_excel(sample_path_2, index=False)

import win32com.client as client

outlook = client.Dispatch('Outlook.Application')

email = outlook.CreateItem(0)

email.Subject = 'Random Sample of Yesterdays Setups for Audit'

html = """
    </div>
    <div>
        <b> 10% of Yesterdays Setups with RTW checked <b><br><br>
    </div>
    <div>
        {table1}<br><br><br>
    </div>
"""

email.To = 'daniel.higginson@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'

email.Attachments.Add(Source=str(Path().absolute() / sample_path_1))

email.HTMLBody = html.format(table1 = sample1.to_html(index=False))

email.Display()

email = outlook.CreateItem(0)

email.To = 'cameron.hill@advance.online'
email.CC = 'jacob.sterling@advance.online; joshua.richards@advance.online'

email.Subject = 'Random Sample of Yesterdays Setups for Audit'

email.Attachments.Add(Source=str(Path().absolute() / sample_path_2))

email.HTMLBody = html.format(table1 = sample2.to_html(index=False))

email.Display()

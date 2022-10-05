# -*- coding: utf-8 -*-
"""
Created on Fri Jul 29 09:05:40 2022

@author: jacob.sterling
"""

import win32com.client as client
import pandas as pd

account = "hello@advance.online"
iterFolder = "Sent Items"

outlook = client.Dispatch("Outlook.Application")

received = pd.DataFrame([], columns=["Sent to", "Received Time"])

namespace = outlook.GetNamespace("MAPI")
folder = namespace.Folders[account].Folders[iterFolder]

for i in range(folder.Items.Count, 0, -1):
    email = folder.Items[i]
    
    if email.Subject.__contains__("Your Feedback is Invaluable!"):
        pd.concat(received, [email.Receiver , email.ReceivedTime])

received.to_csv("received.csv")
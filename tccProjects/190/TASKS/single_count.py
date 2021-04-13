## PYTHON IMPORTS ##

import datetime
import os
import os.path
from os import listdir
import random
import sys
import time
import webbrowser
import win32com.client

## CLASSES ##

## OPEN INBOX
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

for r in outlook.Folders:
    if r.name == "info@animalregister.co.nz":

        inbox = r.Folders.Item("**REGISTRATIONS").Folders.Item("SINGLE REGISTRATIONS")

## 
count = 0
dates = ["2021-03-10", "2021-03-11", "2021-03-12", "2021-03-13", "2021-03-14", "2021-03-15", "2021-03-16"]
for email in inbox.Items:
    try:
        if str(email.ReceivedTime).split()[0] in dates:
            count += 1
            print(count)
        if str(email.ReceivedTime).split()[0] == "2021-03-09":
            break
    except:
        continue

print(count)


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

## FUNCTIONS

## FILEPATH##
dirpath = "G:\\Customer Reporting\\190 - NZCAR\\Automation\\PDF\\"


## OPEN INBOX
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = None
complete = None
human = None

for r in outlook.Folders:
    if r.name == "info@animalregister.co.nz":

        inbox = r.Folders.Item("**REGISTRATIONS").Folders.Item("SINGLE REGISTRATIONS")
        complete = r.Folders.Item("**REGISTRATIONS").Folders.Item("COMPLETED")
        human = r.Folders.Item("**HUMAN CHECK")


if inbox == None:
    sys.exit(0)

## 
snip = "NO"
email_address = "NONE"
email_body = "NONE"
print(inbox.Items.GetLast().Body)


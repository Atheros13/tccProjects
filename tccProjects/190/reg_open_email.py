"""
Works onclick in 190 Registrations Form for "OPEN EMAIL" button. 

1) Opens outlook, gets variable for PDF Split and PDF Complete folders
2) Deletes all 1 hour long saved PDF files and creates a random filename
3) Opens 1 email in PDF Split, saves the pdf attachment as PDF(RandomNumber).pdf inside 
    the g://...//190 Canz//Automation//PDF folder
4) Moves the email into the PDF Complete Folder
5) Opens the PDF to view
6) Returns the path to the PDF so that Remedy can save it to the form
"""

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

for r in outlook.Folders:
    if r.name == "info@animalregister.co.nz":

        inbox = r.Folders.Item("**REGISTRATIONS").Folders.Item("SINGLE REGISTRATIONS")
        complete = r.Folders.Item("**REGISTRATIONS").Folders.Item("FILED SINGLE")
        human = r.Folders.Item("**HUMAN CHECK")


if inbox == None:
    sys.exit(0)

## 
while True:
    snip = "NO"

    e_first = inbox.Items.GetFirst()
    e_last = inbox.Items.GetLast()

    try:
        if e_first.ReceivedTime < e_last.ReceivedTime:
            email = e_first
        else:
            email = e_last
    except:
        try:
            e_first.ReceivedTime
            e_last.Move(human)
            continue
        except:
            e_first.Move(human)
            continue

    email_body = email.Body
    if 'ORIGINAL ADDRESS: ' in email.Body:
        email_address = email.Body.split('ORIGINAL ADDRESS: ')[1].split('\n')[0].rstrip()
    else:
        if email.SenderEmailType=='EX':
            email_address = email.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            email_address = email.SenderEmailAddress

    correct_file_type = False
    for att in email.Attachments:

        ext = ".pdf"
        if '.pdf' in att.FileName:
            pass
        elif '.jpg' in att.FileName:
            ext = ".jpg"
        elif '.jpeg' in att.FileName:
            ext = ".jpeg"
        elif '.png' in att.FileName:
            ext = ".png"
        elif '.gif' in att.FileName:
            ext = ".gif"
        elif '.tif' in att.FileName:
            ext = ".tif"
        else:
            continue

        if "SNIP" in email.Body.upper():
            snip = "SNIP"
        elif "SPCA SNC" in email.Body.upper():
            snip = "SNIP"
        else:
            snip = "PDF"
        
        unique = False
        filename = "%s-%s-%s%s" % (snip, email_address, random.randrange(1000), ext)
        while not unique:
            if filename not in listdir(dirpath):
                unique = True
            else:
                filename = "%s-%s-%s%s" % (snip, email_address, random.randrange(1000), ext)

        # Save PDF
        att.SaveAsFile("G:\\Customer Reporting\\190 - NZCAR\\Automation\\PDF\\%s" % filename)
        correct_file_type = True
        break

    if correct_file_type:
        email.Move(complete)

        os.utime("G:\\Customer Reporting\\190 - NZCAR\\Automation\\PDF\\%s" % filename)
        webbrowser.open("G:\\Customer Reporting\\190 - NZCAR\\Automation\\PDF\\%s" % filename)
        break
    else:
        email.UnRead = True
        email.Move(human)
        continue

sys.stdout.write("G:\\Customer Reporting\\190 - NZCAR\\Automation\\PDF\\%s#&1%s#&2%s" % (filename, email_address, email_body))
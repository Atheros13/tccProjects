from PyPDF2 import PdfFileWriter, PdfFileReader

## PYTHON IMPORTS ##
import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select

from selenium.webdriver.common.keys import Keys

from os import listdir
from os.path import isfile, join

import csv
import datetime
import img2pdf
import os
import sys
import shutil
import string
import random

import win32com.client

import webbrowser

## CLASSES

class SendEmail():

    #some constants (from http://msdn.microsoft.com/en-us/library/office/aa219371%28v=office.11%29.aspx)
    olFormatHTML = 2
    olFormatPlain = 1
    olFormatRichText = 3
    olFormatUnspecified = 0
    olMailItem = 0x0

    def __init__(self, attachment, message, *args, **kwargs):

        self.attachment = attachment
        self.message = message
        self.outlook = win32com.client.Dispatch("Outlook.Application")

        for account in self.outlook.Session.Accounts:
            if account.DisplayName == "info@animalregister.co.nz":
                self.newMail = self.outlook.CreateItem(0)
                self.newMail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                self.newMail.Subject = "AUTOMATION Registration"

        self.send_email()

    def send_email(self):

        self.newMail.BodyFormat = self.olFormatRichText
        self.newMail.Body = self.message
        self.newMail.To = "info@animalregister.co.nz"

        # carbon copies and attachments (optional)

        self.newMail.Attachments.Add(self.attachment)

        # open up in a new window and allow review before send
        #self.newMail.display()

        # or just use this instead of .display() if you want to send immediately
        self.newMail.Send()


## FUNCTIONS

def is_integer(n):
    try:
        float(n)
    except ValueError:
        return False
    else:
        return float(n).is_integer()

## OPEN INBOX
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = None
complete = None
human = None
for r in outlook.Folders:
    if r.name == "info@animalregister.co.nz":
        for box in r.Folders:
            if box.name == "**REGISTRATIONS":
                inbox = box
            elif box.name == "**HUMAN CHECK":
                human = box

        complete = r.Folders.Item("**REGISTRATIONS").Folders.Item("ORIGINAL")

if inbox == None:
    sys.exit(0)

email_count = 1
count = 1
while True:

    email = inbox.Items.GetLast()

    try:
        if email.SenderEmailType == "EX":
            address = email.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            address = email.SenderEmailAddress
    except:
        print("Email: %s Moved to Human Check" % email_count, "- Email Error")
        email.UnRead = True
        email.Move(human)
        email_count += 1
        continue

    human_check = False
    for att in email.Attachments:

        if '.tif' in att.FileName:

            att.SaveAsFile('G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\original.tif')
            tif = open('G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\original.tif', "rb")
            pdf = open('G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\original.pdf', "wb")

            try:
                pdf.write(img2pdf.convert(tif))
            except Exception as e:
                print("Email: %s Moved to Human Check" % email_count, e, email.Body)
                human_check = True
                break

            try:
                inputpdf = PdfFileReader(open('G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\original.pdf', "rb"))
            except Exception as e:
                print("Email: %s Moved to Human Check" % email_count, e, email.Body)
                human_check = True
                break

            try:
                for i in range(inputpdf.numPages):
                    output = PdfFileWriter()
                    output.addPage(inputpdf.getPage(i))
                    with open("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.pdf", "wb") as outputStream:
                        output.write(outputStream)
                    
                    att_path = "G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.pdf"

                    message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % (email.ReceivedTime, email.Subject, address, email.Body)

                    print("Email: %s TIF: %s" % (email_count, count))
                    count += 1
                    SendEmail(att_path, message)

            except Exception as e:
                print("Email: %s Moved to Human Check" % email_count, e, email.Body)
                human_check = True
                break

        elif '.pdf' in att.FileName:

            att.SaveAsFile('G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\original.pdf')

            try:
                inputpdf = PdfFileReader(open('G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\original.pdf', "rb"))
            except Exception as e:
                print("Email: %s Moved to Human Check" % email_count, e, email.Body)
                human_check = True
                break

            try:
                for i in range(inputpdf.numPages):
                    output = PdfFileWriter()
                    output.addPage(inputpdf.getPage(i))
                    with open("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.pdf", "wb") as outputStream:
                        output.write(outputStream)
                    
                    att_path = "G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.pdf"

                    message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % (email.ReceivedTime, email.Subject, address, email.Body)

                    print("Email: %s PDF: %s" % (email_count, count))
                    count += 1
                    SendEmail(att_path, message)

            except Exception as e:
                print("Email: %s Moved to Human Check" % email_count, e, email.Body)
                human_check = True
                break

        elif '.jpg' in att.Filename:

            att.SaveAsFile("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.jpg")
            att_path = "G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.jpg"
            try:
                message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % (email.ReceivedTime, email.Subject, address, email.Body)
            except:
                message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % ("Unknown", email.Subject, address, email.Body)

            print("Email: %s JPG: %s" % (email_count, count))
            count += 1
            SendEmail(att_path, message)

        elif '.jpeg' in att.Filename:

            att.SaveAsFile("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.jpeg")
            att_path = "G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.jpeg"
            try:
                message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % (email.ReceivedTime, email.Subject, address, email.Body)
            except:
                message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % ("Unknown", email.Subject, address, email.Body)

            print("Email: %s JPEG: %s" % (email_count, count))
            count += 1
            SendEmail(att_path, message)

        elif '.png' in att.Filename:

            att.SaveAsFile("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.png")
            att_path = "G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.png"
            try:
                message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % (email.ReceivedTime, email.Subject, address, email.Body)
            except:
                message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % ("Unknown", email.Subject, address, email.Body)

            print("Email: %s PNG: %s" % (email_count, count))
            count += 1
            SendEmail(att_path, message)

        elif '.gif' in att.Filename:

            att.SaveAsFile("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.gif")
            att_path = "G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\temp.gif"
            try:
                message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % (email.ReceivedTime, email.Subject, address, email.Body)
            except:
                message = "ORIGINAL DATETIME: %s\nORIGINAL SUBJECT: %s\nORIGINAL ADDRESS: %s\n\nORIGINAL EMAIL: %s" % ("Unknown", email.Subject, address, email.Body)

            print("Email: %s GIF: %s" % (email_count, count))
            count += 1
            SendEmail(att_path, message)
        else:

            print("Email: %s Moved to Human Check" % email_count, att.FileName)
            human_check = True
            continue


    if human_check:
        email.UnRead = True
        email.Move(human)
    else:
        email.UnRead = False
        email.Move(complete)
    
    if email_count >= 30:
        print("ENDING")
        time.sleep(10)
        sys.exit(0)
    else:
        email_count += 1

print("ENDING")
time.sleep(10)
sys.exit(0)

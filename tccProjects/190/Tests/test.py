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

    def __init__(self, *args, **kwargs):

        self.attachment = "G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\att.pdf"
        self.message = "Test"
        self.outlook = win32com.client.Dispatch("Outlook.Application")

        for account in self.outlook.Session.Accounts:
            if account.DisplayName == "info@animalregister.co.nz":
                self.newMail = self.outlook.CreateItem(0)
                self.newMail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                self.newMail.Subject = "AUTOMATION Registration"
                self.newMail.BodyFormat = self.olFormatRichText
                self.newMail.Body = self.message
                self.newMail.Attachments.Add(self.attachment)

                self.newMail.SaveAs(Path="G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\test.msg")

class OpenEmail():

    def __init__(self, *args, **kwargs):

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        msg = outlook.OpenSharedItem("G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\test.msg")
        print(msg.Body)


## ENGINE ##
OpenEmail()
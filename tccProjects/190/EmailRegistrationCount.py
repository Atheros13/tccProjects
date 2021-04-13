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
class EmailRegistrationCount():

    def __init__(self, *args, date_from=None, date_to=None, date_stop=None, **kwargs):

        #
        self.date_from = date_from
        self.date_to = date_to
        self.date_stop = date_stop

        # OUTLOOK
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        for r in self.outlook.Folders:
            if r.name == "info@animalregister.co.nz":

                self.original = r.Folders.Item("**REGISTRATIONS").Folders.Item("ORIGINAL")
                self.single = r.Folders.Item("**REGISTRATIONS").Folders.Item("SINGLE REGISTRATIONS")
                self.complete = r.Folders.Item("**REGISTRATIONS").Folders.Item("FILED SINGLE")

        # ENGINE
        self.process_email_count()

    def process_email_count(self):

        emails = []
        dict = {}

        messages = self.original.Items
        messages.Sort("[ReceivedTime]", True)
        for email in messages:

            try:
                if email.SenderEmailType == "EX":
                    address = email.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    address = email.SenderEmailAddress
                if address not in emails:
                    emails.append(address)
            except:
                address = None

            date = str(email.ReceivedTime).split()[0]
            if date not in dict:
                dict[date] = {'original':0, 'single': 0}
            if date >= self.date_from and date <= self.date_to:
                    dict[date]['original'] += 1
                    #print(date, dict[date]['original'])
            elif date == self.date_stop:
                break
            else:
                break

        messages = self.single.Items
        messages.Sort("[ReceivedTime]", True)
        for email in messages:
            date = str(email.ReceivedTime).split()[0]
            if date not in dict:
                dict[date] = {'original':0, 'single': 0}
            if date >= self.date_from and date >= self.date_to:
                dict[date]['single'] += 1
            elif date == self.date_stop:
                break

        messages = self.complete.Items
        messages.Sort("[ReceivedTime]", True)
        for email in messages:
            date = str(email.ReceivedTime).split()[0]
            if date not in dict:
                dict[date] = {'original':0, 'single': 0}
            if date >= self.date_from and date >= self.date_to:
                dict[date]['single'] += 1
                print(dict[date]['single'])
            elif date == self.date_stop:
                break

        print(dict)
        print("; ".join(emails))

## ENGINE ##
EmailRegistrationCount(date_from='2021-04-07', date_to='2021-04-13', date_stop='2021-04-06')
EmailRegistrationCount(date_from='2021-03-31', date_to='2021-04-06', date_stop='2021-03-30')
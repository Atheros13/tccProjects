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

    def __init__(self, *args, date_from=None, date_to=None, no_spca=False, **kwargs):

        #
        self.date_from = date_from
        self.date_to = date_to
        self.no_spca = no_spca

        # OUTLOOK
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        for r in self.outlook.Folders:
            if r.name == "info@animalregister.co.nz":

                self.original = r.Folders.Item("**REGISTRATIONS").Folders.Item("ORIGINAL") # the original email
                self.single = r.Folders.Item("**REGISTRATIONS").Folders.Item("SINGLE REGISTRATIONS") # the unprocessed single rego
                self.filed = r.Folders.Item("**REGISTRATIONS").Folders.Item("FILED SINGLE") # the processed single regos

        # ENGINE
        self.process_email_count()

    ## PROCESS ##
    def process_email_count(self):

        emails = []
        dict = {}

        # original emails counted by date, and the dict set up with date keys
        original_emails = self.original.Items
        original_emails.Sort("[ReceivedTime]", True)
        for email in original_emails:

            try:
                date = str(email.ReceivedTime).split()[0]
            except Exception:
                print(Exception)
                continue 

            if date >= self.date_from and date <= self.date_to:

                # if the date is in the correct range, make a key in the dict
                # otherwise add 1 to the count of the original emails
                if date not in dict:
                    dict[date] = {'original':1, 'single': 0}
                else:
                    dict[date]['original'] += 1

                # get the address the original email was sent from,
                # if it is not in the emails list, add it. 
                if email.SenderEmailType == "EX":
                    address = email.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    address = email.SenderEmailAddress
                if address not in emails:
                    emails.append(address)

            elif date < self.date_from:
                break

        # single registrations
        single_emails = self.single.Items
        single_emails.Sort("[ReceivedTime]", True)
        for email in single_emails:

            try:
                date = str(email.ReceivedTime).split()[0]
            except Exception:
                print(Exception)
                continue 

            if date >= self.date_from and date <= self.date_to:
                if date not in dict:
                    dict[date] = {'original':0, 'single': 1}
                else:
                    dict[date]['single'] += 1

            elif date < self.date_from:
                print(False)
                break
  
        # filed single registrations
        filed_emails = self.filed.Items
        filed_emails.Sort("[ReceivedTime]", True)
        for email in filed_emails:

            try:
                date = str(email.ReceivedTime).split()[0]
            except Exception:
                print(Exception)
                continue 

            if date >= self.date_from and date <= self.date_to:
                if date not in dict:
                    dict[date] = {'original':0, 'single': 1}
                else:
                    dict[date]['single'] += 1

            elif date < self.date_from:
                break            

        # show results
        self.process_notepad(emails, dict)

    def process_notepad(self, email_list, dict):

        if self.no_spca == True:
            emails = []
            for email in email_list:
                if "SPCA" not in email.upper():
                    emails.append(email)
        else:
            emails = email_list


        notepad = "REGISTRATION COUNT FROM %s TO %s\n\n" % (self.date_from, self.date_to)

        # count
        for date in sorted(dict):
            notepad += "%s\nOriginal Emails: %s\nSingle Regos: %s\n\n" % (date, dict[date]['original'], dict[date]['single'])
        
        # list of emails
        notepad += "EMAIL LIST: these are the emails we have received registrations from during this time period.\n\n"
        notepad += "; \n".join(emails)

        file = open("G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\EmailRegistrationCount.txt", 'w')
        file.write(notepad)
        file.close()

        webbrowser.open("G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\EmailRegistrationCount.txt")

class EmailGeneralCount():

    def __init__(self, *args, date_from=None, date_to=None, **kwargs):

        #
        self.date_from = date_from
        self.date_to = date_to

        # OUTLOOK
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        for r in self.outlook.Folders:
            if r.name == "info@animalregister.co.nz":

                self.inbox = r.Folders.Item("Inbox")
                self.complete = r.Folders.Item("**REGISTRATIONS").Folders.Item("COMPLETED")
                self.completed = r.Folders.Item("**COMPLETED")

        # ENGINE
        self.process_email_count()

    ## PROCESS ##
    def process_email_count(self):

        dict = {}

        # original emails counted by date, and the dict set up with date keys
        inbox_emails = self.inbox.Items
        inbox_emails.Sort("[ReceivedTime]", True)
        for email in inbox_emails:

            try:
                date = str(email.ReceivedTime).split()[0]
            except Exception:
                continue 

            if date >= self.date_from and date <= self.date_to:

                # if the date is in the correct range, make a key in the dict
                # otherwise add 1 to the count of the original emails
                if date not in dict:
                    dict[date] = 1
                else:
                    dict[date] += 1

            elif date < self.date_from:
                break

        complete_emails = self.complete.Items
        complete_emails.Sort("[ReceivedTime]", True)
        for email in complete_emails:

            try:
                date = str(email.ReceivedTime).split()[0]
            except Exception:
                continue 

            if date >= self.date_from and date <= self.date_to:

                # if the date is in the correct range, make a key in the dict
                # otherwise add 1 to the count of the original emails
                if date not in dict:
                    dict[date] = 1
                else:
                    dict[date] += 1

            elif date < self.date_from:
                break

        completed_emails = self.completed.Items
        completed_emails.Sort("[ReceivedTime]", True)
        for email in completed_emails:

            try:
                date = str(email.ReceivedTime).split()[0]
            except Exception:
                continue 

            if date >= self.date_from and date <= self.date_to:

                # if the date is in the correct range, make a key in the dict
                # otherwise add 1 to the count of the original emails
                if date not in dict:
                    dict[date] = 1
                else:
                    dict[date] += 1

            elif date < self.date_from:
                break

        self.process_notepad(dict)

    def process_notepad(self, dict):

        notepad = "GENERAL DAILY COUNT FROM %s TO %s\n\n" % (self.date_from, self.date_to)

        # count
        for date in sorted(dict):
            notepad += "%s\nDaily Count: %s\n\n" % (date, dict[date])

        file = open("G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\EmailGeneralCount.txt", 'w')
        file.write(notepad)
        file.close()

        webbrowser.open("G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\EmailGeneralCount.txt")

## ENGINE ##

EmailRegistrationCount(date_from='2021-04-28', date_to='2021-05-12', no_spca=True)
#EmailGeneralCount(date_from='2021-03-30', date_to='2021-04-27')

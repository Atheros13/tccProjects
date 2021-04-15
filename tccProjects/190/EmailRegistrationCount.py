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

    def __init__(self, *args, date_from=None, date_to=None, **kwargs):

        #
        self.date_from = date_from
        self.date_to = date_to

        # OUTLOOK
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        for r in self.outlook.Folders:
            if r.name == "info@animalregister.co.nz":

                self.original = r.Folders.Item("**REGISTRATIONS").Folders.Item("ORIGINAL") # the original email
                self.single = r.Folders.Item("**REGISTRATIONS").Folders.Item("SINGLE REGISTRATIONS") # the unprocessed single rego
                self.filed = r.Folders.Item("**REGISTRATIONS").Folders.Item("FILED SINGLE") # the processed single regos

        # ENGINE
        self.process_email_count()
        #self.test()

    ## TEST ##
    def test(self):

        pass

    ## PROCESS ##
    def process_email_count(self):

        emails = []
        dict = {}

        # original emails counted by date, and the dict set up with date keys
        original_emails = self.original.Items
        original_emails.Sort("[ReceivedTime]", True)
        for email in original_emails:

            date = str(email.ReceivedTime).split()[0]

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

            date = str(email.ReceivedTime).split()[0]
            if date >= self.date_from and date <= self.date_to:
                dict[date]['single'] += 1

            elif date < self.date_from:
                print(False)
                break
  
        # filed single registrations
        filed_emails = self.filed.Items
        filed_emails.Sort("[ReceivedTime]", True)
        for email in filed_emails:

            date = str(email.ReceivedTime).split()[0]

            if date >= self.date_from and date <= self.date_to:
                dict[date]['single'] += 1

            elif date < self.date_from:
                break            

        # show results
        self.process_notepad(emails, dict)


    def process_notepad(self, emails, dict):

        notepad = "REGISTRATION COUNT FROM %s TO %s\n\n" % (self.date_from, self.date_to)

        # count
        for date in sorted(dict):
            notepad += "%s\nOriginal Emails: %s\nSingle Regos: %s\n\n" % (date, dict[date]['original'], dict[date]['single'])
        
        # list of emails
        notepad += "EMAIL LIST: these are the emails we have received registrations from during this time period.\n\n"
        notepad += "; ".join(emails)

        file = open("G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\EmailRegistrationCount.txt", 'w')
        file.write(notepad)
        file.close()

        webbrowser.open("G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\EmailRegistrationCount.txt")

## ENGINE ##

#EmailRegistrationCount(date_from='2021-04-07', date_to='2021-04-13')

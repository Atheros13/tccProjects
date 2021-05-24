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
import importlib
import os
import sys
import shutil
import string
import random

import win32com.client

import webbrowser

## CLASSES ##

class SendEmail():

    #some constants (from http://msdn.microsoft.com/en-us/library/office/aa219371%28v=office.11%29.aspx)
    olFormatHTML = 2
    olFormatPlain = 1
    olFormatRichText = 3
    olFormatUnspecified = 0
    olMailItem = 0x0

    def __init__(self, type, subject, chip, pet_name, human_name, to_email, pet=None, *args, **kwargs):

        self.subject = subject
        self.outlook = win32com.client.Dispatch("Outlook.Application")

        for account in self.outlook.Session.Accounts:
            if account.DisplayName == "info@animalregister.co.nz":
                self.newMail = self.outlook.CreateItem(0)
                self.newMail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                self.newMail.Subject = self.subject

        if pet != None:
            self.pet = pet
        self.chip = chip
        self.pet_name = pet_name
        self.human_name = human_name
        self.to_email = to_email

        if type == "Rehome":
            self.send_rehome()
        elif type == "Missing":
            self.send_missing()

    def send_rehome(self):

        message = "Dear %s,\n\n" % self.human_name
        message += "Animal Name: %s\nMicrochip Number: #%s\n\n" % (self.pet_name, self.chip)
        message += "We have received a request to change the primary contact for the above animal, "

        if self.subject == "UPDATED PRIMARY CONTACT - NEW PRIMARY CONTACT":

            message += "and it is now registered to you on the New Zealand Companion Animal Register. "
            message += "You can also update these on the website https://www.animalregister.co.nz/userupdate.aspx or "
            message += "phone us on 0508 LOSTPET (567873) or email info@animalregister.co.nz within the next 48 hours "
            message += "to confirm the new details for our database. "
            message += "Please quote the microchip number and your animal’s name in any correspondence.\n\n"

        elif self.subject == "UPDATED PRIMARY CONTACT - PERMISSION NOT GRANTED":

            message += "which is currently registered to you on the New Zealand Companion Animal Register. "
            message += "An update was sent, but Permission to update was checked as 'No'. "
            message += "Please phone us on 0508 LOSTPET (567873) or email info@animalregister.co.nz within the next 48 hours "
            message += "to confirm the new details for our database. "
            message += "Please quote the microchip number and your animal’s name in any correspondence.\n\n"

        elif self.subject == "UPDATED PRIMARY CONTACT - DETAILS NOT PROVIDED":

            message += "which is currently registered to you on the New Zealand Companion Animal Register. "
            message += "We received no email or phone contact details for the new primary contact. "
            message += "Please phone us on 0508 LOSTPET (567873) or email info@animalregister.co.nz within the next 48 hours "
            message += "to confirm the new details for our database. "
            message += "Please quote the microchip number and your animal’s name in any correspondence.\n\n"
        
        message += "Keeping the contact details up to date is "
        message += "important as in the event of an animal going missing, we want to be able to get that animal home as soon as possible.\n\n"

        message += "Thank you,\n\nNew Zealand Companion Animal Register"

        self.newMail.BodyFormat = self.olFormatRichText
        self.newMail.Body = message
        self.newMail.To = self.to_email

        # carbon copies and attachments (optional)

        #newMail.CC = "moreaddresses here"
        #newMail.BCC = "address"
        #attachment1 = "Path to attachment no. 1"
        #attachment2 = "Path to attachment no. 2"
        #newMail.Attachments.Add(attachment1)
        #newMail.Attachments.Add(attachment2)

        # open up in a new window and allow review before send
        #self.newMail.display()

        # or just use this instead of .display() if you want to send immediately
        self.newMail.Send()

    def send_missing(self):

        message = "Dear %s,\n\n" % self.human_name
        message += "Animal Name: %s\nMicrochip Number: #%s\n\n" % (self.pet_name, self.chip)

        message += "Here at the New Zealand Companion Animal Register (NZCAR), we are passionate about getting missing pets back to their worried families. "
        message += "We understand that when a family member is missing it is a difficult and stressful time, so we have created these tips highlighting how "
        message += "to maximise the chances of getting your missing pet home.\n\nPlease find attached a copy of this document.\n\n"
        message += "Kindest regards,\n\nThe New Zealand Companion Animal Register Team\n"
        message += "0508 LOSTPET (0508 567873)\nwww.animalregister.co.nz\nFind us on facebook"

        self.newMail.BodyFormat = self.olFormatRichText
        self.newMail.Body = message
        self.newMail.To = self.to_email

        #attachments
        if self.pet == "Dog":
            self.newMail.Attachments.Add("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\DogAttach.docx")
        else:
            self.newMail.Attachments.Add("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\CatAttach.docx")

        # open up in a new window and allow review before send
        self.newMail.display()

        # or just use this instead of .display() if you want to send immediately
        #self.newMail.Send()

## FUNCTIONS

def is_integer(n):
    try:
        float(n)
    except ValueError:
        return False
    else:
        return float(n).is_integer()

def is_phone(number):

    count = 0
    for n in number:

        if count > 8:
            return True
        if n == "+":
            continue
        if is_integer(n):
            count += 1
    return False

def open_webpage(address):

    ## OPEN DRIVER

    ## LOGIN TO DVS AND GO TO LOAD PROSPECT
    driver.get(address)  
    try:
        driver.find_element_by_name("ctl00$MainContent$LoginUser$UserName").send_keys(username)
        driver.find_element_by_name("ctl00$MainContent$LoginUser$Password").send_keys(password)
        time.sleep(0.9)
        driver.find_element_by_name('ctl00$MainContent$LoginUser$LoginButton').click()
    except:
        pass

    time.sleep(0.9)    

def process_email(email, complete, unread=False):

    email.UnRead = unread
    email.Move(complete)

def split_email_phone(contact):

    pass

## GLOBAL ##
username = "tccnzcar"
password = "EoJ@V@xb"
driver = webdriver.Chrome(ChromeDriverManager().install())
        
## OPEN INBOX
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = None
other_completed = None
day2_check = None
human_check = None
for r in outlook.Folders:

    if r.name == "info@animalregister.co.nz":
        for box in r.Folders:

            if box.name == "**ANNUAL CANCELLATIONS":
                inbox = box
            elif box.name == "**COMPLETED":
                other_completed = box
            elif box.name == "**HUMAN CHECK":
                human_check = box
            elif box.name == "**FOLLOW UP":
                day2_check = box

if inbox == None:
    sys.exit(0)

## CSV IMPORT
csv_import = open('G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\csv_import.csv', 'a', newline='\n')

## 
for email in list(inbox.Items): # note this works with reverse index, first is the oldest email

    if "Annual Notification Cancellation Request" in email.subject:

        chip = email.subject.split(" : ")[1].split(" (")[0].rstrip()

        if "Pet Is Deceased." in email.body:

            for line in email.body.split('\n'):

                if chip in line[:15]:

                    web_address = (line.replace("%s <" % chip, "")).replace(">", "")

                    open_webpage(web_address)

                    dead_check = driver.find_element_by_id("MainContent_cbDeceased")
                    if not dead_check.is_selected():
                        dead_check.click()

                    driver.find_element_by_id("MainContent_ConfirmSave").click()
                    driver.find_element_by_name("ctl00$MainContent$SaveButton").click()
                    
                    try:
                        csv_import.write('%s,PET HAS DIED ,Michael Atheros,EMAIL - DECEASED PET,%s,%s,%s,%s,%s,%s\n' % (int(chip), "No", "AUTOMATION", "", "", "", "")) 
                    except:
                        csv_import.write('%s,PET HAS DIED ,Michael Atheros,EMAIL - DECEASED PET,%s,%s,%s,%s,%s,%s\n' % (chip, "No", "AUTOMATION", "", "", "", ""))    

            process_email(email, other_completed)

        elif "Pet Has Been Rehomed." in email.body:

            old_email = None
            old_owner = ""

            new_owner = None
            new_owner_full = "NAME NOT PROVIDED"
            new_email = None
            new_phone = None
            pet_name = "NAME NOT PROVIDED"
            permission = None
            
            follow_up = ""
            email_and_phone = False

            for line in email.body.split('\n'):

                if "New Owner Name: " in line:

                    new_owner = line.split("New Owner Name: ")[1].rstrip()
                    new_owner_full = str(new_owner)

                elif "New Owner Phone/Email: " in line:

                    contact = line.split("New Owner Phone/Email: ")[1].strip()

                    if '@' in contact:
                        if is_phone(contact):
                            email_and_phone = True
                        else:
                            new_email = contact
                    elif is_phone(contact):
                        new_phone = contact

                elif "Name :" == line[:6]:
                    pet_name = line[6:].strip()

                elif "Permission Granted: " in line:
                    permission = line.split("Permission Granted: ")[1].strip()

            ## OPEN UPDATE, SEARCH FOR CHIP
            open_webpage("https://www.animalregister.co.nz/Implanters/UpdateMicrochip.aspx") 
            driver.find_element_by_name("ctl00$MainContent$MicrochipNumberToSearchFor").send_keys(chip)
            driver.find_element_by_name("ctl00$MainContent$btnSearch").click()
            time.sleep(0.9)

            ## GET OLD 
            fields = """MainContent_PCTitle
            MainContent_PCFirstName
            MainContent_PCLastName
            MainContent_RoleCheckBoxList_0
            MainContent_RoleCheckBoxList_1
            MainContent_RoleCheckBoxList_2
            MainContent_PCResidentialAddress
            MainContent_PCResidentialAddressCity
            MainContent_PCResidentialAddressPostCode
            MainContent_PCPostalAddress
            MainContent_PCPostalAddressCity
            MainContent_PCPostalAddressPostCode
            MainContent_PCEmailAddress
            MainContent_PCHomePhone
            MainContent_PCWorkPhone
            MainContent_PCMobilePhone
            MainContent_ECTitle
            MainContent_ECFirstName
            MainContent_ECLastName
            MainContent_ECResidentialAddress
            MainContent_ECResidentialAddressCity
            MainContent_ECResidentialAddressPostCode
            MainContent_ECEmailAddress
            MainContent_ECHomePhone
            MainContent_ECWorkPhone
            MainContent_ECMobilePhone
            MainContent_ECFax"""

            values = []
            field_fields = []
            fields = fields.split('\n')
            for i in range(len(fields)):
                
                ## GET FIELD
                field = driver.find_element_by_id(fields[i].strip())

                ## SPECIAL FIELDS
                if i == 1:
                    old_owner = field.get_attribute('value')
                elif i == 2:
                    old_owner += " %s" % field.get_attribute('value')
                if i == 12:
                    old_email = field.get_attribute('value')

                if i in range(3,6):
                    if not field.is_selected:
                        values.append("")
                values.append(field.get_attribute('value'))

            ## WRITE NOTE
            note = "OLD PRIMARY AND ALTERNATE CONTACT DETAILS\n\n"
            note += "PRIMARY: %s %s\n" % (values[1], values[2])
            note += "PHONE: %s - %s - %s\nEMAIL: %s\n" % (values[13], values[14], values[15], old_email)
            note += "RESIDENTIAL: %s,%s,%s\n" % (values[6], values[7], values[8])
            note += "POSTAL: %s, %s, %s\n" % (values[9], values[10], values[11])
            note += "ALTERNATE: %s %s\n" % (values[17], values[18])
            note += "PHONE: %s - %s - %s\nEMAIL: %s\n" % (values[23], values[24], values[25], values[22])
            note += "RESIDENTIAL: %s,%s,%s" % (values[19], values[20], values[21])

            ## SAVE NOTE           
            open_webpage(driver.find_element_by_id("MainContent_AddNewNote").get_attribute("href"))
            driver.find_element_by_id("MainContent_Notes").send_keys(note)
            driver.find_element_by_id("MainContent_Finished").click()
            time.sleep(0.3)

            ## CLEAR ALL 
            for i in range(len(fields)):
                
                ## GET FIELD
                field = driver.find_element_by_id(fields[i].strip())
            
                ## ASSIGN VALUES AS REQUIRED
                
                if i not in [0,3,4,5,16]:
                    try:
                        field.clear()
                        field.send_keys("")
                    except:
                        print(i)

                if i in [0, 16]:
                    try:
                        Select(field).select_by_index(0)
                    except:
                        print(field, i, False)

                # Name
                if i == 1:
                    
                    if new_owner not in [None, ""]:
                        try:
                            print(new_owner.split()[0], True)
                            field.send_keys(new_owner.split()[0])
                            new_owner = (' ').join(new_owner.split()[1:]).rstrip()
                        except:
                            field.send_keys(new_owner)
                            new_owner = None
                    else:
                        field.send_keys("REHOMED: NAME UNKNOWN")
                
                elif i == 2:
                   
                    if new_owner not in [None, ""]:
                        field.send_keys(new_owner)
                    else:
                        field.send_keys("REHOMED: NAME UNKNOWN")

                # Phone
                elif i == 13:
                    
                    if new_phone not in [None, ""]:
                        field.send_keys(new_phone)
                    else:
                        field.send_keys("REHOMED: NUMBER UNKNOWN")

                # Email
                elif i == 12:
                
                    if new_email not in [None, ""]:
                        field.send_keys(new_email)
                    else:
                        field.send_keys("")

                # Address
                elif i == 6 or i == 7:
                    field.send_keys("REHOMED: ADDRESS UNKNOWN")

                ## SPECIAL FIELDS
                if i in range(3,6):
                    if field.is_selected():
                        try:
                            field.click()
                        except:
                            print("click here")

            
            ## UPDATE
            change = driver.find_element_by_name("ctl00$MainContent$ConfirmChanges")
            driver.execute_script("arguments[0].click();", change)

            ## SEND EMAIL
            send_rehome_email = True
            subject = None
            to_email = old_email
            current_pc = old_owner

            if email_and_phone:
                send_rehome_email = False
                follow_up = "Yes"
            elif permission == "No":
                subject = "UPDATED PRIMARY CONTACT - PERMISSION NOT GRANTED"
                follow_up = "Yes"
            elif new_email == None and new_phone == None:
                subject = "UPDATED PRIMARY CONTACT - DETAILS NOT PROVIDED"
                follow_up = "Yes"
            elif new_email != None:
                subject = "UPDATED PRIMARY CONTACT - NEW PRIMARY CONTACT"
                to_email = new_email
                current_pc = new_owner_full
                follow_up = "Yes"
            elif new_phone != None:
                send_rehome_email = False
                follow_up = "Yes"
                
            if send_rehome_email:
                SendEmail("Rehome", subject, chip, pet_name, current_pc, to_email)
                time.sleep(1)
                call_log = "%s Emailed %s %s - Automation Email" % (datetime.datetime.now().strftime("%H:%M"), current_pc, to_email)                 
                process_email(email, day2_check, unread=True)
            else:
                to_email = None
                call_log = ""
                process_email(email, human_check, unread=True)

            try:
                csv_import.write('%s,TRANSFERRING PRIMARY CONTACT ,Michael Atheros,EMAIL,%s,%s - AUTOMATION,%s,%s,"%s",%s\n' % (int(chip), follow_up, current_pc, 
                                                                                                                                new_phone, to_email, subject, call_log))  
            except:
                csv_import.write('%s,TRANSFERRING PRIMARY CONTACT ,Michael Atheros,EMAIL,%s,%s - AUTOMATION,%s,%s,"%s",%s\n' % (chip, follow_up, current_pc, 
                                                                                                                            new_phone, to_email, subject, call_log))  
            update = driver.find_element_by_name("ctl00$MainContent$SaveButton")
            driver.execute_script("arguments[0].click();", update)
            time.sleep(1)

        elif "Other." in email.Body and "Reason:" in email.Body:

            ## FILE EMAIL
            process_email(email, other_completed)

            ## CSV
            try:
                csv_import.write('%s,UPDATING DETAILS OF PET OR CONTACT PERSONS,Michael Atheros,EMAIL,%s,%s,%s,%s,%s,%s\n' % (int(chip), "No", "AUTOMATION - OTHER", "", "", "OTHER", "")) 
            except:
                csv_import.write('%s,UPDATING DETAILS OF PET OR CONTACT PERSONS,Michael Atheros,EMAIL,%s,%s,%s,%s,%s,%s\n' % (chip, "No", "AUTOMATION - OTHER", "", "", "OTHER", ""))    

        elif "Pet Is Missing." in email.body:

            ## OPEN UPDATE, SEARCH FOR CHIP
            open_webpage("https://www.animalregister.co.nz/Implanters/UpdateMicrochip.aspx") 
            driver.find_element_by_name("ctl00$MainContent$MicrochipNumberToSearchFor").send_keys(chip)
            driver.find_element_by_name("ctl00$MainContent$btnSearch").click()
            time.sleep(0.9)

            ## CHECK AS MISSING
            missing = driver.find_element_by_id("MainContent_ReportedMissing")
            if not missing.is_selected():
                driver.execute_script("arguments[0].click();", missing)
            time.sleep(0.2)
            try:
                driver.execute_script("arguments[0].click();", driver.find_element_by_id("MainContent_ReportedMissingConfirm"))
                driver.execute_script("arguments[0].click();", driver.find_element_by_id("MainContent_NotForLPNZ"))
            except:
                pass

            ## OWNER/PET DETAILS
            pet = Select(driver.find_element_by_id("MainContent_SpeciesID")).first_selected_option.text
            pet_name = driver.find_element_by_id("MainContent_AnimalName").get_attribute('value')
            owner = "%s %s" % (driver.find_element_by_id("MainContent_PCFirstName").get_attribute('value'),
                               driver.find_element_by_id("MainContent_PCLastName").get_attribute('value'))
            to_email = driver.find_element_by_id("MainContent_PCEmailAddress").get_attribute('value')

            ## SAVE AS MISSING 
            update = driver.find_element_by_name("ctl00$MainContent$SaveButton")
            driver.execute_script("arguments[0].click();", update)
            time.sleep(1)

            ## EMAIL 
            SendEmail("Missing", "MISSING PET", chip, pet_name, owner, to_email, pet=pet)

            ## FILE EMAIL
            process_email(email, other_completed)

            ## CSV
            try:
                csv_import.write('%s,PET IS LOST AND REGISTERED ,Michael Atheros,EMAIL,%s,%s,%s,%s,%s,%s\n' % (int(chip), "No", "AUTOMATION", "", "", "MISSING", "")) 
            except:
                csv_import.write('%s,PET IS LOST AND REGISTERED ,Michael Atheros,EMAIL,%s,%s,%s,%s,%s,%s\n' % (chip, "No", "AUTOMATION", "", "", "MISSING", ""))    

time.sleep(5)
sys.exit(0)
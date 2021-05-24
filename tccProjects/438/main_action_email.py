"""
This is the primary project program. It is triggered by a button in the 438 Remedy Form
(which passes arguments to the program). The program checks the inbox for a DVS email, 
extracts data, opens and logs into Eaccounts, then processes data, and waits for operator 
actions. Once done, it records a report in a .csv for TCC. 
"""
## PYTHON IMPORTS
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

import datetime
import sys
import time
import webbrowser

## APP IMPORTS 
from app.CSV.csv import CSV
from app.Outlook.outlook438 import Outlook438WebsiteLeads as Outlook438
from app.Selenium.chrome import ChromeDriver

from app.functions import is_integer

## MAIN CLASS ##
class Project438ActionEmail():

    def __init__(self, *args, **kwargs):

        ## ARGUMENTS
        # process arguments from Remedy into useable data
        # [loaded_by, username, password]
        self.remedy_data = self.build_remedy_data()
        self.button_action = self.remedy_data[3]

        ## OUTLOOK
        # access Outlook and reference the required folders
        self.outlook = Outlook438(self.button_action)
        # access an email to work with
        self.email = self.outlook.get_email()
        if self.email == None:
            sys.exit(0)
        # extract data from email in array, example below:
        # [subject, message, notepad, contact, add1, add2, add3, add4, zip, phone, email, wanted1, wanted2, wanted3, best_time]
        self.email_data = self.build_email_data()

        ## NOTEPAD
        self.open_notepad()

        ## DRIVER
        # opens a Chrome driver, processes all website data and extracts operator actions
        # this mainly results in self.customer_code and self.email_type values
        self.email_type = "Call Centre: Sales Lead"
        self.driver = ChromeDriver().driver
        self.run_driver()

        ## PROCESS
        self.process_data()
        time.sleep(300)

    ## BUILD METHODS
    def build_remedy_data(self):
        """Gets data from arguments and returns as an array."""

        if len(sys.argv) > 1:

            argv = sys.argv

            loaded_by = "%s %s" % (argv[1], argv[2])
            username = argv[3]
            password = argv[4]
            button_action = argv[5] # Inbox or Follow or CC

        else:
    
            loaded_by = "Michael Atheros"
            username = "TCC16"
            password = "dvs09"
            button_action = "Inbox"

        return [loaded_by, username, password, button_action]

    def build_email_data(self):
        """Processes an email and returns an array of the data."""
 
        dvs_subjects = ['New submission from Free Consultation - Homepage', 'New submission from DVS Contact Us',
                        'New submission from Book A Consultation', 'New submission from Book Service']        

        msg = self.email
        subject = msg.subject
        if self.button_action == "CC":
            try:
                subject = email.Subject.split(" - ")[1]
            except:
                pass

        message = ""
        notepad = msg.body

        contact = ""
        add1 = ""
        add2 = ""
        add3 = ""
        add4 = ""
        zip = ""

        phone = ""
        email = ""

        wanted1 = "DVS Consultation"
        wanted2 = ""
        wanted3 = ""

        best_time = ""

        lines = msg.body.split('\n')
        for i in range(len(lines)):

            line = lines[i]
    
            if "Name" == line[:4] or "Name:" == line[:5]:
                contact = lines[i+1].strip()
            elif "Email" == line[:5] or "Email:" == line[:6]:
                email = lines[i+1].strip()
                try:
                    email = email.split('<')[0]
                except:
                    pass
            elif "Phone Number" == line[:12] or "Phone Number:" == line[:13] or "Phone" == line[:5] or "Phone:" == line[:6]:
                phone = lines[i+1].strip()
            elif "Address" == line[:7] or "Address:" == line[:7]:

                if subject in dvs_subjects[1:]:
                    for a in range(1,5):
                        txt = lines[i+a].strip()
                        if "Map It" == txt[:6]:
                            break

                        if a > 1:
                            for words in txt.split():
                                try:
                                    if words.is_integer():
                                        zip = words
                                        txt = txt.split(zip)[0]
                                        break
                                except:
                                    pass
                        if a == 1:
                            add1 = txt.strip()
                        elif a == 2:
                            add2 = txt.strip()
                        elif a == 3:
                            add3 = txt.strip()
                        else:
                            add4 = txt.strip()

                elif subject == 'New submission from Free Consultation - Homepage':
                    address = lines[i+1].strip().split(',')
                    for a in range(len(address)):
                        txt = address[a]
                        if a > 1:
                            for words in txt.split():
                                if is_integer(words):
                                    zip = words
                                    txt = txt.split(zip)[0]
                                    break
                        if a == 0:
                            add1 = txt.strip()
                        elif a == 1:
                            add2 = txt.strip()
                        elif a == 2:
                            add3 = txt.strip()
                        else:
                            add4 = txt.strip()

            elif "Your Message (Optional)" in line and subject in ['New submission from Book A Consultation', 'New submission from Book Service']:

                message = lines[i+1].strip()
                if message != "":
                    if len(message) > 70:
                        wanted1 = message[:70]
                        wanted2 = message[70:]
                        if len(message) > 140:
                            wanted2 = message[70:140]
                            wanted3 = message[140:]
                    else:
                        wanted1 = message

            elif 'I have a question about...' in line and subject == 'New submission from DVS Contact Us':

                message = lines[i+1].strip()
                if message != "":
                    wanted1 = message
                
            elif 'Preferred Time Of Day' in line:

                best_time = lines[i+1].strip()

            elif 'Your Message' in line and subject in ['New submission from DVS Contact Us', 'New submission from DVS Upgrade Enquiry']:

                message = lines[i+1].strip()
                if message != "":
                    if wanted1 == "":
                        if len(message) > 70:
                            wanted1 = message[:70]
                            wanted2 = message[70:]
                            if len(message) > 140:
                                wanted2 = message[70:140]
                                wanted3 = message[140:]
                        else:
                            wanted1 = message
                    else:
                        if len(message) > 70:
                            wanted2 = message[:70]
                            wanted3 = message[70:]
                        else:
                            wanted2 = message

        return [subject, message, notepad, contact, 
                add1, add2, add3, add4, zip, 
                phone, email, 
                wanted1, wanted2, wanted3, best_time]

    ## NOTEPAD METHODS
    def open_notepad(self):
        """Writes a .txt file of the email data and opens the file."""

        notepad = self.email_data[2]
        loaded_by = self.remedy_data[0]

        file = open('G:\\Customer Reporting\\438-DVS\\Automation\\Emails\\%s.txt' % loaded_by, 'w')
        try:
            file.write(notepad)
        except:
            file.write("Check email for details - there is a non-english character that won't show up on Notepad\nThe Automation should still work though.")
        file.close()

        webbrowser.open('G:\\Customer Reporting\\438-DVS\\Automation\\Emails\\%s.txt' % loaded_by)

    ## DRIVER METHODS
    def run_driver(self):

        # login to eaccounts, enter data in prospect form
        self.driver_loginToProspect()
        self.driver_prospectInput()

        # Runs a number of checks for the presence of certain elements.
        # Based on the values, the self.email_type may change, and once the code can recognise an 
        # action, it returns a string value
        action = self.driver_actionCheck()

        # Based on the action value, the following functions determine self.email_type and 
        # the self.customer_code value
        if action == "New Account":
            self.driver_NewAccount()
        elif action == "CRM":
            self.driver_CRM_Note()
        elif action == "Follow Up":
            self.driver_FollowUp()

    ## PROSPECT LOGIN AND DATA INPUT
    def driver_loginToProspect(self):
        """Opens Eaccounts, logs in, and navigates to Prospect page."""

        username = self.remedy_data[1]
        password = self.remedy_data[2]

        self.driver.get("https://www7.eaccounts.co.nz/eLogin_Main.asp")  
        self.driver.find_element_by_name("User__Name").send_keys(username)
        self.driver.find_element_by_name("User__Pass").send_keys(password)
        time.sleep(0.5)
        self.driver.find_element_by_name('Login').click()
        time.sleep(0.5)
        self.driver.find_element_by_class_name("MENU-BUTTON").click()

        while True:
            try:
                self.driver.find_element_by_name("Load_Prospect").click()
                break
            except:
                pass

    def driver_prospectInput(self):
        """Enters email_data into Prospect and does a search for the address."""

        loaded_by = self.remedy_data[0]
        contact = self.email_data[3]
        add1 = self.email_data[4]
        add2 = self.email_data[5]
        add3 = self.email_data[6]
        add4 = self.email_data[7]
        zip = self.email_data[8]
        phone = self.email_data[9]
        email = self.email_data[10]
        wanted1 = self.email_data[11]
        wanted2 = self.email_data[12]
        wanted3 = self.email_data[13]
        best_time = self.email_data[14]

        self.driver.find_element_by_name("Prospect_Loaded_By").send_keys(loaded_by)
        self.driver.find_element_by_name("Prospect_Contact").send_keys(contact)
        self.driver.find_element_by_name("Prospect_Name").send_keys(contact)
        self.driver.find_element_by_name("Prospect_Del_Add_1").send_keys(add1)
        self.driver.find_element_by_name("Prospect_Del_Add_2").send_keys(add2)
        self.driver.find_element_by_name("Prospect_Del_Add_3").send_keys(add3)
        self.driver.find_element_by_name("Prospect_Del_Add_4").send_keys(add4)
        self.driver.find_element_by_name("Prospect_Del_Zip").send_keys(zip)
        self.driver.find_element_by_name("Prospect_Ph").send_keys(phone)
        self.driver.find_element_by_name("Prospect_CellPh").send_keys("")
        self.driver.find_element_by_name("Prospect_Email").send_keys(email)
        self.driver.find_element_by_name("Prospect_Source").send_keys("[3] Website Email Lead")
        self.driver.find_element_by_name("Prospect_Note_Type").send_keys("Call Centre: Sales Lead")
        self.driver.find_element_by_name("Prospect_Wanted").send_keys(wanted1)
        self.driver.find_element_by_name("Prospect_Wanted2").send_keys(wanted2)
        self.driver.find_element_by_name("Prospect_Wanted3").send_keys(wanted3)
        self.driver.find_element_by_name("Prospect_Relevant").send_keys("")
        self.driver.find_element_by_name("Prospect_Best_Time").send_keys(best_time)

        ## SEARCH CHECK
        self.driver.find_element_by_name("dCust__Code").send_keys(add1)
        self.driver.find_element_by_name("dCust__Code").send_keys(Keys.ENTER)

    ## GET ACTION
    def driver_actionCheck(self):

        """Runs a constant check on the page to see if certain web elements
        are present, if they are it indicates what the operator has done,
        and what type of email this was
        
        This will continue to run untill it can return a valid action value."""

        while True:

            # Prospect Page - Check if Follow Up is written in the Loaded by field
            # this indicates that the email cannot be processed right now i.e. missing
            # information and needs to be followed up
            try:
                check = self.driver.find_element_by_name("Prospect_Loaded_By").get_attribute("value")
                if "FOLLOW UP" in check.upper():
                    return "Follow Up"
            except:
                pass

            # Prospect Page - If the address is not in Eaccounts, but this is not a sales lead i.e. a Filter Service
            # for a property with a DVS but not in Eaccounts. This process will create a new account.
            # This does not return, it only changes the self.email_type value.
            try:
                note = Select(self.driver.find_element_by_name("Prospect_Note_Type")).first_selected_option.text
                if note != "Call Centre: Sales Lead":
                    self.email_type = note
            except:
                pass

            # Prospect Page - Once a prospect has been created, returns the  action value "New Account"
            try:
                if "New Prospect Saved OK - Code = " in self.driver.find_elements_by_tag_name('h3')[0].text:
                    return "New Account"
            except:
                pass

            # CRM Note - If the operator does not make a new account but is going to have to action the email for a 
            # existing account, they always have to write a CRM note. Once they have moved to this section,
            # this returns the action value "CRM"
            try:
                self.driver.find_element_by_name("Save_CRM") # on CRM note page
                return "CRM"
            except:
                pass

            # After running a check, it waits a third of a second and checks again.
            time.sleep(0.3)

    ## GET DATA BASED ON ACTION ##
    def driver_NewAccount(self):
        """Assigns Customer Code of the new prospect/account."""

        self.customer_code = self.driver.find_elements_by_tag_name('h3')[0].text.split("New Prospect Saved OK - Code = ")[1]

    def driver_CRM_Note(self):
        """Awaits action, and assigns customer code and email type"""

        # wait for action choice
        while True:

            summary = Select(self.driver.find_element_by_id("CRM_Note_Type")).first_selected_option.text
            if "Call Centre:" in summary:
                
                # assign self.email_type
                self.email_type = summary

                # assign self.customer_code
                form_table = self.driver.find_element_by_class_name("form_table")
                self.customer_code = form_table.find_elements_by_tag_name("b")[0].text.split(" ]")[0].replace("[ ", "")

                return

            # If, once the operator has opened the existing account, they find they cannot progress i.e. not enough information,
            # they can type Follow Up in the Summary field and the email will be filed in * WEBSITE - FOLLOW UP
            short_note = self.driver.find_element_by_name("CRM_Short_Note").get_attribute("value").upper()
            if short_note == "FOLLOW UP":
                self.customer_code = "FOLLOW UP"
                self.email_type = "Website: TCC to Follow Up"

                return

            time.sleep(0.3)

    def driver_FollowUp(self):
        """Assigns Follow Up values"""

        self.customer_code = "FOLLOW UP"
        self.email_type = "Website: TCC to Follow Up"

    ## PROCESS METHODS
    def process_data(self):
        """Processes data into a .csv for use in Remedy and Reporting, 
        and files the email. """

        self.process_csv()
        self.process_email()

    def process_csv(self):
        """Append to a .csv recording all DVS email actions."""
        
        subject = self.email_data[0]
        name = self.email_data[3]
        email = self.email_data[10]
        phone = self.email_data[9]
        message = self.email_data[1]
        address = self.email_data[4]
        loaded_by = self.remedy_data[0]
        date = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        csv = CSV("G:\\Customer Reporting\\438-DVS\\Automation\\Emails\\Reports\\", "email_reporting.csv", "a",
                  ["Subject", "Customer Code", "Action", "Name", "Email", "Phone", "Address", "Message", "Date", "TCC Staff"])

        csv.writerow([subject, self.customer_code, self.email_type, name, email, phone, address, message, date, loaded_by])

    def process_email(self):
        """Files the email to the correct folder based on self.email_type."""

        # if this is not a Follow Up, mark as complete and change the Subject to include the CC
        if self.email_type != "Website: TCC to Follow Up" :
            self.email.FlagRequest = "Mark Complete"
            if self.button_action != "CC" or self.customer_code in self.email.Subject:
                self.email.Subject = "%s - %s" % (self.customer_code, self.email.Subject)
            elif self.button_action == "CC":
                self.email.Subject = "%s %s" % (self.customer_code, self.email.Subject)
        # else, just move the email to the Follow Up folder
        else:
            self.email.Move(self.outlook.follow_up)
            self.close()

        # Files the email according to self.email_type
        if "Sales Lead" in self.email_type:
            self.email.Move(self.outlook.sales)
        elif "Technical" in self.email_type:
            self.email.Move(self.outlook.technical)
        elif "Filter" in self.email_type:
            self.email.Move(self.outlook.filters)
        elif "Spare" in self.email_type:
            self.email.Move(self.outlook.spare_parts)
        else:
            self.email.Move(self.outlook.general)

    ## CLOSE METHOD
    def close(self):
        """Closes the browser, notepad, and ends the script - Only used with Follow Up"""

        self.driver.close()
        self.driver.quit()
        sys.exit(0)


## ENGINE ##

Project438ActionEmail()

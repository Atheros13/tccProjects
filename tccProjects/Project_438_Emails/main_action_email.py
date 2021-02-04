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
import webbrowser

## APP IMPORTS 
from app.CSV.csv import CSV
from app.Outlook.outlook438 import Outlook438
from app.Selenium.chrome import ChromeDriver

## MAIN CLASS ##
class Project438ActionEmail():

    def __init__(self, *args, **kwargs):

        ## ARGUMENTS
        # process arguments from Remedy into useable data
        # [loaded_by, username, password]
        self.remedy_data = self.build_remedy_data()

        ## OUTLOOK
        # access Outlook and reference the required folders
        self.outlook = Outlook438()
        # access an email to work with
        self.email = self.outlook.get_email()
        # extract data from email in array, example below:
        # [subject, message, notepad, contact, add1, add2, add3, add4, zip, phone, email, wanted1, wanted2, wanted3, best_time]
        self.email_data = self.build_email_data()

        ## NOTEPAD
        # if the message in the email is too large to include in the prospect
        # page, open a notepad with the full email data
        if len(self.email_data[1]) > 210:
            self.open_notepad()

        ## DRIVER
        # opens a Chrome driver, processes all website data and extracts operator actions
        # this mainly results in self.customer_code and self.email_type values
        self.driver = ChromeDriver().driver
        self.run_driver()

        ## PROCESS
        # if the operator has actioned in 10 minutes, saves a record as a .txt,
        # saves a complete record in a .csv, and files the email
        if self.customer_code == None:
            self.close()
        else:
            self.process_data()
            self.close()

    ## BUILD METHODS
    def build_remedy_data(self):
        """Gets data from arguments and returns as an array."""

        if len(sys.argv) > 1:

            argv = sys.argv

            loaded_by = "%s %s" % (argv[1], argv[2])
            username = argv[3]
            password = argv[4]

        else:
    
            loaded_by = "Michael Atheros"
            username = "TCC16"
            password = "dvs08"

        return [loaded_by, username, password]

    def build_email_data(self):
        """Processes an email and returns an array of the data."""
        
        msg = self.email
        subject = msg.subject
        message = ""
        notepad = ""

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
                notepad += "Name: %s\n" % contact
            elif "Email" == line[:5] or "Email:" == line[:6]:
                email = lines[i+1].strip()
                try:
                    email = email.split('<')[0]
                except:
                    pass
                notepad += "Email: %s\n" % email
            elif "Phone Number" == line[:12] or "Phone Number:" == line[:13] or "Phone" == line[:5] or "Phone:" == line[:6]:
                phone = lines[i+1].strip()
                notepad += "Phone: %s" % phone
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

                notepad += "Address: %s %s %s %s %s\n\n" % (add1, add2, add3, add4, zip) 

            elif "Your Message (Optional)" in line and subject in ['New submission from Book A Consultation', 'New submission from Book Service']:

                message = lines[i+1].strip()
                notepad += "Your Message (Optional)\n\n%s\n\n" % message
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
                notepad += "I have a question about...\n\n%s\n\n" % message
                if message != "":
                    wanted1 = message
                
            elif 'Preferred Time Of Day' in line:

                best_time = lines[i+1].strip()

            elif 'Your Message' in line and subject == 'New submission from DVS Contact Us':

                message = lines[i+1].strip()
                notepad += "Your Message:\n\n%s\n\n" % message
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
        file.write(notepad)
        file.close()

        webbrowser.open('G:\\Customer Reporting\\438-DVS\\Automation\\Emails\\%s.txt' % loaded_by)

    ## DRIVER METHODS
    def run_driver(self):

        # login to eaccounts, enter data in prospect form
        self.driver_loginToProspect()
        self.driver_prospectInput()

        # check for type of email - assuming a website sales lead
        self.email_type = self.driver_prospectSuburbCheck()

        #
        count, action = self.driver_actionCheck()
        self.customer_code = None
        if action == "New Account":
            self.driver_NewAccount()
        elif action == "CRM":
            self.driver_CRM_Note(count)

    def driver_loginToProspect(self):
        """Opens Eaccounts, logs in, and navigates to Prospect page."""

        username = self.remedy_data[1]
        password = self.remedy_data[2]

        self.driver.get("https://www7.eaccounts.co.nz/eLogin_Main.asp")  
        selfdriver.find_element_by_name("User__Name").send_keys(username)
        self.driver.find_element_by_name("User__Pass").send_keys(password)
        self.time.sleep(0.9)
        self.driver.find_element_by_name('Login').click()
        time.sleep(0.9)
        self.driver.find_element_by_class_name("MENU-BUTTON").click()
        self.driver.find_element_by_name("Load_Prospect").click()

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
        driver.find_element_by_name("dCust__Code").send_keys(add1)
        driver.find_element_by_name("dCust__Code").send_keys(Keys.ENTER)

    def driver_prospectSuburbCheck(self):
        """Runs constant check of the "Prospect Suburb" field, if it gets assigned to a 
        non-suburb, non-chinese, non-newbuild, non-commercial value, then it means the 
        address is not in eaccounts and may be a technical issue for John, filter request, 
        or something else. """

        website_lead = 'Website: Sales Lead'
        while True:
            try:
                suburb = Select(self.driver.find_element_by_name("Prospect_Suburb")).first_selected_option.text
                if suburb[0] == ".":
                    if "CHINESE LEADS" in suburb:
                        break
                    elif ".BUILD" in suburb:
                        break
                    elif "COMMERCIAL" in suburb:
                        break
                    elif "JOHN CHING" in suburb:
                        website_lead = "Website: Technical"
                    else:
                        website_lead = "Website: Not in Eaccounts"
            except:
                break

        return website_lead

    def driver_actionCheck(self):
        """Runs a constant check on the page to see if certain web elements
        are present, if they are it indicates what the operator has done,
        and what type of email this was"""

        count = 600
        action = None
        while count > 0:

            try:
                if "New Prospect Saved OK - Code = " in self.driver.find_elements_by_tag_name('h3')[0].text:
                    return count, "New Account"
            except:
                pass

            try:
                self.driver.find_element_by_name("Save_CRM") # on CRM note page
                return count, "CRM"
            except:
                pass

            time.sleep(0.3)
            count -= 0.3

        return count, None

    def driver_NewAccount(self):
        """Assigns Customer Code of the new prospect/account."""

        self.customer_code = self.driver.find_elements_by_tag_name('h3')[0].text.split("New Prospect Saved OK - Code = ")[1]

    def driver_CRM_Note(self, count):
        """Enters text data in CRM Note, awaits action, and assigns customer code and email type"""

        # enter Email Text in Note Field
        email_text = "EMAIL RECEIVED: %s\nMESSAGE: %s" % (self.email_data[0], self.email.Body) # subject & email.Body
        text_box = self.driver.find_element_by_class_name("cke_contents_ltr")
        text_box.click()
        text_box.send_keys(email_text)

        # wait for action choice
        while count > 0:
            summary = Select(self.driver.find_element_by_id("CRM_Note_Type")).first_selected_option.text
            if "Please Specify" not in summary:
                # assign self.email_type
                self.email_type = summary

                # assign self.customer_code
                form_table = self.driver.find_element_by_class_name("form_table")
                self.customer_code = form_table.find_elements_by_tag_name("b")[0].text.split(" ]")[0].replace("[ ", "")

                break

            count -= 0.3

    ## PROCESS METHODS
    def process_data(self):
        """Processes data into a .txt and .csv for use in Remedy and Reporting, 
        and files the email. """

        #self.process_txt() # CURRENTLY ONLY IN IDEA STAGE
        self.process_csv()
        self.process_email()

    def process_txt(self):
        """Writes a .txt file of data, which is used by Remedy to automate the Form."""

        file = open('G:\\Customer Reporting\\438-DVS\\Automation\\Emails\\RemedyResults.txt', 'w')
        file.write("%s\n%s\n%s" % (self.customer_code, self.email_type, self.email.Body))
        file.close()

    def process_csv(self):
        """Append to a .csv recording all DVS email actions."""
        
        subject = self.email_data[0]
        name = self.email_data[3]
        email = self.email_data[10]
        phone = self.email_data[9]
        message = self.email_data[1]
        address = self.email_data[4]

        csv = CSV("G:\\Customer Reporting\\438-DVS\\Automation\\Emails\\Reports\\", "email_reporting.csv", "a",
                  ["Subject", "Customer Code", "Action", "Name", "Email", "Phone", "Address", "Message", "Date"])

        csv.writerow([subject, self.customer_code, self.email_type, name, email, phone, address, message, datetime.datetime.now()])

    def process_email(self):
        """Files the email to the correct folder based on self.email_type."""

        if self.email_type in ['Website: Sales Lead', "Call Centre: Sales Lead"]:
            self.email.Move(self.outlook.sales)
        elif self.email_type in ['Website: Technical', "Call Centre: Technical issue"]:
            self.email.Move(self.outlook.technical)
        elif self.email_type in ["Website: Not in Eaccounts"]:
            self.email.Move(self.outlook.general)
        elif self.email_type in ["Call Centre: Filter Send/Change Request"]:
            self.email.Move(self.outlook.filters)
        else:
            self.email.Move(self.outlook.general)

    ## CLOSE METHOD
    def close(self):
        """Closes the browser and ends the script"""

        self.driver.close()
        self.driver.quit()
        sys.exit(0)

## ENGINE ##
Project438ActionEmail()

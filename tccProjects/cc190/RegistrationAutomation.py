## PYTHON IMPORTS ##
from selenium.webdriver.support.ui import Select

import datetime
import os
import sys
import time
import webbrowser

## APP IMPORTS ##
from app.Selenium.chrome import ChromeDriver

## CLASS ##
class RegistrationAutomation():

    def __init__(self, *args, **kwargs):

        ## ENGINE ##
        self.process_remedy()
        self.website_login()
        self.website_register()

    ## FUNCTIONS
    def is_integer(self, n):
        """Returns True if a value is an integer."""

        try:
            float(n)
        except ValueError:
            return False
        else:
            return float(n).is_integer()

    def is_phone(self, number):
        """Returns True if a number satisfies Phone validation."""

        count = 0
        for n in number:

            if count > 8:
                return True
            if n == "+":
                continue
            if is_integer(n):
                count += 1
        return False

    ## PROCESS ##
    def process_remedy(self):
        """Converts sys.argv values into string values for inputting into the website."""

        if len(sys.argv) > 1:
            
            argv = sys.argv

            raw_text = argv[1].replace("^^^", " ")
            d = raw_text.split("##%%")

            self.username = "info@thecallcentre.co.nz"
            self.password = "TCC123456!"

            self.chip = d[0]
            self.name = d[1]
            self.species = d[2]
            self.gender = d[3]
            self.desex = d[4]

            self.purebred = True
            if d[5] != "":
                self.purebred = False
            self.birth_month = d[6]
            self.birth_year = d[7]
            self.breed_1 = d[8]
            self.breed_2 = d[9]
            self.colour1 = d[10]
            self.colour2 = d[11]
            self.colour = d[12]
            self.animal_notes = d[13]

            self.clinic = d[14]
            self.implanter = d[15]
            self.implantation_date = d[16]

            print("IMPLANTER DATE WILL NEED TO BE MANUALLY ENTERED IN CALENDAR:\n\n%s" % self.implantation_date)

            self.email = d[17]
            self.no_email_reason = d[18]
            self.g_firstname = d[19]
            self.g_lastname = d[20]
            self.g_phone1 = d[21]
            self.g_phone2 = d[22]
            self.g_streetnumber = d[23]
            self.g_streetaddress = d[24]
            self.g_suburb = d[25]
            self.g_city = d[26]
            self.g_region = d[27]
            self.g_postcode = d[28]
            self.s_firstname = d[29]
            self.s_lastname = d[30]
            self.s_phone1 = d[31]
            self.s_phone2 = d[32]

        if False:

            self.username = "info@thecallcentre.co.nz"
            self.password = "TCC123456!"

            self.chip = "666666666666666"
            self.name = "Timmy"
            self.species = "Dog"
            self.gender = "Male"
            self.desex = False

            self.purebred = True
            self.birth_month = "January"
            self.birth_year = "2020"
            self.breed_1 = "German Shepherd"
            self.breed_2 = ""
            self.colour1 = "Black"
            self.colour2 = "Tan"
            self.colour = ""
            self.animal_notes = ""

            self.clinic = "Abyrose"
            self.implanter = "Mike Wazowski"
            self.implantation_date = ""

            self.email = ""
            self.no_email_reason = "Because I said so"
            self.g_firstname = "Michael"
            self.g_lastname = "Atheros"
            self.g_phone1 = "0226472984"
            self.g_phone2 = ""
            self.g_streetnumber = "13"
            self.g_streetaddress = "Queen Street"
            self.g_suburb = "Petone"
            self.g_city = "Lower Hutt"
            self.g_region = "Wellington"
            self.g_postcode = "5012"
            self.s_firstname = "Jaimee"
            self.s_lastname = "Chapman"
            self.s_phone1 = ""
            self.s_phone2 = ""

    def process_address(self, *args):

        pass

    ## WEBSITE ##
    def website_login(self):
        """Open driver, opens registration part of the CANZ website and signs in. """


        # Opens Driver and opens website
        self.driver = ChromeDriver().driver
        self.driver.get("https://www.animalregister.co.nz/implanter/dashboard/register")

        # Waits for demo to be manually entered (while this is still working),
        # and signs in once this is done
        demo = True
        while demo:

            try:
                self.driver.find_element_by_id("MemberLoginForm_LoginForm_Email").send_keys(self.username)
                self.driver.find_element_by_id("MemberLoginForm_LoginForm_Password").send_keys(self.password)
                time.sleep(0.9)
                self.driver.find_element_by_name('action_doLogin').click()
                demo = False
            except:
                pass

            time.sleep(1)    

    def website_register(self):

        self.register_section_basic()
        self.driver.find_elements_by_class_name("c-button")[0].click()

        self.register_section_animal()
        self.driver.find_elements_by_class_name("c-button")[2].click()

        time.sleep(1)
        self.register_section_clinic()
        #self.driver.find_elements_by_class_name("c-button")[4].click()

        time.sleep(1)
        self.register_section_guardian()
        
        #
        time.sleep(100)

    def register_section_basic(self):

        # chip and name
        self.driver.find_element_by_id("microchipNumber").send_keys(self.chip)
        self.driver.find_element_by_id("microchipConfirm").send_keys(self.chip)
        self.driver.find_element_by_id("companionName").send_keys(self.name)

        # species
        species_field = self.driver.find_element_by_id("speciesId")
        Select(species_field).select_by_visible_text(self.species)
        
        # gender
        if self.gender == "Male":
            self.driver.find_element_by_id("gender-option-Male").click()
        elif self.gender == "Female":
            self.driver.find_element_by_id("gender-option-Female").click()
        else:
            self.driver.find_element_by_id("gender-option-Unknown").click()
        
        # desex
        if self.desex == "YES":
            self.driver.find_element_by_id("desexed-option-Yes").click()
        elif self.desex == "NO":
            self.driver.find_element_by_id("desexed-option-No").click()
        else:
            self.driver.find_element_by_id("desexed-option-Unknown").click()

    def register_section_animal(self):

        # breed
        animal = False
        while not animal:
            try:
                if self.purebred:
                    self.driver.find_element_by_id("option-purebred").click()
                else:
                    self.driver.find_element_by_id("option-crossbreed").click()
                animal = True
            except:
                pass        

        breed = self.driver.find_element_by_id("primaryBreedId")
        Select(breed).select_by_visible_text(self.breed_1)
        if self.breed_2 != "":
            breed = self.driver.find_element_by_id("secondaryBreedId")
            Select(breed).select_by_visible_text(self.breed_2)

        # birth
        month = self.driver.find_element_by_xpath('//select[@aria-label="Month of birth"]')
        Select(month).select_by_visible_text(self.birth_month)
        year = self.driver.find_element_by_xpath('//select[@aria-label="Year of birth"]')
        Select(year).select_by_visible_text(self.birth_year)

        # colour & notes
        if self.colour1 != "":
            colour = self.driver.find_element_by_id("primaryColourId")
            Select(colour).select_by_visible_text(self.colour1)
        if self.colour2 != "":
            colour = self.driver.find_element_by_id("secondaryColourId")
            Select(colour).select_by_visible_text(self.colour2)
        if self.colour != "":
            self.driver.find_element_by_id("colourDescription").send_keys(self.colour)
        if self.animal_notes != "":
            self.driver.find_element_by_id("notesGeneral").send_keys(self.animal_notes)

    def register_section_clinic(self):

        #clinic
        clinic = self.driver.find_elements_by_tag_name("select")[8]
        Select(clinic).select_by_visible_text(self.clinic)

        # implanter input13
        self.driver.find_element_by_xpath("""//*[@id="v-Implantation-inputs"]/div/div/div[2]/div[1]/div/div[2]/input""").send_keys(self.implanter)
        
        # implant date: at the moment needs to be manually entered and waited till continue pushed
        check = False
        while not check:
            try:
                self.driver.find_element_by_id("guardianFirstName").send_keys("")
                check = True
            except:
                pass

    def register_section_guardian(self):

        check_account = False

        # email
        if self.no_email_reason == "":
            # enters email
            self.driver.find_element_by_id("guardianEmail").send_keys(self.email)
            # moves to new field - gives time to check if email exists
            self.driver.find_element_by_id("guardianFirstName").send_keys("")
            time.sleep(1)
            # tests to see if email exists, if it does not, the try will succeed
            try:
                self.driver.find_element_by_id("guardianSurname").send_keys("")
            # if the try fails, then the email exists and there is no need to try entering the extra information
            except:
                check_account = True
        else:
            self.driver.find_element_by_id("noEmailAvailable").click()
            self.driver.find_element_by_id("noEmailReason").send_keys(self.no_email_reason)

        # runs if check_account != True
        if not check_account:
            # guardian name and number
            self.driver.find_element_by_id("guardianFirstName").send_keys(self.g_firstname)
            self.driver.find_element_by_id("guardianSurname").send_keys(self.g_lastname)
            self.driver.find_element_by_id("guardianMobileNumber").send_keys(self.g_phone1)
            self.driver.find_element_by_id("guardianAlternativeNumber").send_keys(self.g_phone2)

            # guardian address - enter manually
            self.driver.find_element_by_class_name("address-manual").click()
            self.driver.find_element_by_id("StreetNumber").send_keys(self.g_streetnumber)
            self.driver.find_element_by_id("Street").send_keys(self.g_streetaddress)
            self.driver.find_element_by_id("Suburb").send_keys(self.g_suburb)
            self.driver.find_element_by_id("City").send_keys(self.g_city)
            self.driver.find_element_by_id("Region").send_keys(self.g_region)
            self.driver.find_element_by_id("PostalCode").send_keys(self.g_postcode)

            # alternate details
            self.driver.find_element_by_id("secondaryContactFirstName").send_keys(self.s_firstname)
            self.driver.find_element_by_id("secondaryContactSurname").send_keys(self.s_lastname)
            self.driver.find_element_by_id("secondaryContactMobileNumber").send_keys(self.s_phone1)
            self.driver.find_element_by_id("secondaryContactAlternativeNumber").send_keys(self.s_phone2)

## ENGINE ##
RegistrationAutomation()
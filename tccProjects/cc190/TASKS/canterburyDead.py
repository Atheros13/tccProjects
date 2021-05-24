## IMPORTS ##
from app.Selenium.chrome import ChromeDriver

import csv
import time

## CLASS ##
class canterburyDead():

    def __init__(self, *args, **kwargs):

        self.count = 0
        self.driver = ChromeDriver().driver
        self.open_website()
        self.csv_import = open('G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\csv_canterbury.csv', 'a', newline='\n')
        self.extract_data()

    def open_website(self):

        self.driver.get("https://www.animalregister.co.nz/Account/Login.aspx?ReturnUrl=%2fAdmin%2fDefault.aspx")  
        self.driver.find_element_by_name("ctl00$MainContent$LoginUser$UserName").send_keys("tccnzcar")
        self.driver.find_element_by_name("ctl00$MainContent$LoginUser$Password").send_keys("EoJ@V@xb")
        time.sleep(0.5)
        self.driver.find_element_by_name('ctl00$MainContent$LoginUser$LoginButton').click()
        time.sleep(0.3)    

    def extract_data(self):
        
        with open('G:\\Customer Reporting\\190 - NZCAR\\Automation\\TASKS\\canterbury_chips.csv', "r") as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row[1] == "Deceased - mark in database":
                    chip = row[0]
                    self.mark_deceased(chip)

    def mark_deceased(self, chip):

        self.driver.get("https://www.animalregister.co.nz/Admin/ExtraDetails.aspx")

        self.driver.find_element_by_id("MainContent_MicrochipNumberToSearchFor").send_keys(chip)
        self.driver.find_element_by_id("MainContent_btnSearch").click()
        time.sleep(0.3)
        dead_check = self.driver.find_element_by_id("MainContent_cbDeceased")
        if not dead_check.is_selected():
            dead_check.click()
        self.driver.find_element_by_id("MainContent_ConfirmSave").click()
        self.driver.find_element_by_name("ctl00$MainContent$SaveButton").click()

        self.record_for_remedy(chip)

    def record_for_remedy(self, chip):

        self.csv_import.write('%s,PET HAS DIED ,Michael Atheros,EMAIL - DECEASED PET,%s,%s,%s,%s,%s,%s\n' % (int(chip), "No", "AUTOMATION", "", "", "CANTERBURY SPCA - Marking as deceased", ""))
        self.count += 1
        print(self.count)


## ENGINE ##
canterburyDead()
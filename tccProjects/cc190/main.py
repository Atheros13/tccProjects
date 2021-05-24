import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


driver = webdriver.Chrome(ChromeDriverManager().install())

driver.get("https://www.animalregister.co.nz/Admin/AddMicrochipAsAdmin.aspx")  
time.sleep(0.9)
driver.find_element(By.NAME, "ctl00$MainContent$LoginUser$UserName").send_keys("tccnzcar")
driver.find_element_by_id("MainContent_LoginUser_Password").send_keys("EoJ@V@xb")
time.sleep(0.9)
driver.find_element(By.NAME, "ctl00$MainContent$LoginUser$LoginButton").click()

time.sleep(0.9)

dropdown = driver.find_element(By.NAME, "ctl00$MainContent$ImplanterOrganizationID")

selector = Select(dropdown)

# Waiting for the values to load
element = WebDriverWait(driver, 10).until(EC.element_to_be_selected(selector.options[0]))

options = selector.options

import xlwt
from datetime import datetime

wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')



for index in range(1, len(options)-1):

    break 
    txt = options[index].text

    ws.write(index, 0, txt) 




for species in ['Dog']:
    driver.find_element_by_id("MainContent_SpeciesID").send_keys(species)
    time.sleep(0.9)
    dropdown = driver.find_element(By.ID, "MainContent_PrimaryBreedID")

    selector = Select(dropdown)

    # Waiting for the values to load
    element = WebDriverWait(driver, 
    10).until(EC.element_to_be_selected(selector.options[0]))

    options = selector.options
    for index in range(1, len(options)-1):
        txt = options[index].text
        ws.write(index, 1, txt) 


wb.save('example.xls')
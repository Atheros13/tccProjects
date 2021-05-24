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
import os
import sys
import shutil
import string
import random

import win32com.client

import webbrowser

## SYS ARGUMENTS ##
argv = sys.argv
f =  ' '.join(argv[1:]).split("##%%")\

## TESTING
if False:
    f = []
    for i in range(32):
        f.append("")
    f[1] = "934000090227650"
    f[31] = "800115151"

submitter = f[0]
chip = f[1]
date = f[2]
species = f[3]
breed = f[4]
if breed in ["", None]:
    breed = f[5]
pure = f[6]
colour = f[7]
clinic = f[8]
pet_name = f[9]
imp_firstname = f[10]
imp_lastname = f[11]
gender = f[12]
desexed = f[13]
day = f[14]
if day != "":
    if day[0] == "0":
        day = day[1]
month = f[15]
if month != "":
    if month[0] == "0":
        month = month[1]
year = f[16]
o_title = f[17]
o_firstname = f[18]
o_lastname = f[19]
o_phone1 = f[20]
o_phone2 = f[21]
o_address = f[22]
o_city = f[23]
o_zip = f[24]
o_email = f[25]
a_title = f[26]
a_firstname = f[27]
a_lastname = f[28]
a_phone1 = f[29]
a_phone2 = f[30]
prepaid = f[31]

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

def process_email(email, complete):

    email.UnRead = False
    email.Move(complete)

## GLOBAL ##
username = "tccnzcar"
password = "EoJ@V@xb"
driver = webdriver.Chrome(ChromeDriverManager().install())

## OPEN UPDATE, 
open_webpage("https://www.animalregister.co.nz/Admin/AddMicrochipAsAdmin.aspx") 

# microchip check
driver.find_element_by_id("MainContent_MicrochipNumber").send_keys(chip)
driver.find_element_by_id("MainContent_ChipCrossCheck1").send_keys(chip[:3])
driver.find_element_by_id("MainContent_ChipCrossCheck2").send_keys(chip[3:])

check_chip = driver.find_element_by_id("MainContent_ChipExistsStatus").get_attribute("src")
if "exists" in check_chip:
    driver.execute_script("window.alert('The Microchip already exists, please close this page, make a note in the form and submit the form');")
    time.sleep(600)
    sys.exit(0)

# organisation
if prepaid != "":
    driver.find_element_by_id("MainContent_PrePaidFormNumber").send_keys(prepaid)
    driver.find_element_by_id("MainContent_CheckPrePaidLinkButton").click()
    driver.find_element_by_id("MainContent_MicrochipNumber").send_keys(chip)
    driver.find_element_by_id("MainContent_ChipCrossCheck1").send_keys(chip[:3])
    driver.find_element_by_id("MainContent_ChipCrossCheck2").send_keys(chip[3:])
else:
    clinic_field = driver.find_element_by_id("MainContent_ImplanterOrganizationID")
    Select(clinic_field).select_by_visible_text(clinic)

# clear date first
date_field = driver.find_element_by_id("MainContent_DateChipped")
date_field.clear()
date_field.send_keys("")
date_field.send_keys(date)

driver.find_element_by_id("MainContent_ImplanterName").send_keys(imp_firstname + " " + imp_lastname)
driver.find_element_by_id("MainContent_AnimalName").send_keys(pet_name)

# gender is Male by default
if gender == "FEMALE":
    driver.execute_script("arguments[0].click();", driver.find_element_by_id("MainContent_Gender_1"))

# species
species_field = driver.find_element_by_id("MainContent_SpeciesID")
Select(species_field).select_by_visible_text(species)
time.sleep(0.3)
breed_field = driver.find_element_by_id("MainContent_PrimaryBreedID")
Select(breed_field).select_by_visible_text(breed)

# pet details
if pure != "PUREBRED":
    driver.execute_script("arguments[0].click();", driver.find_element_by_id("MainContent_Purebred_1"))
if desexed != "YES":
    if desexed == "NO":
        driver.execute_script("arguments[0].click();", driver.find_element_by_id("MainContent_Desexed_1"))
    else:
        driver.execute_script("arguments[0].click();", driver.find_element_by_id("MainContent_Desexed_2"))
driver.find_element_by_id("MainContent_PrimaryColour").send_keys(colour)

# born
if day != "":
    Select(driver.find_element_by_id("MainContent_DayBorn")).select_by_value(day)
if month != "":
    Select(driver.find_element_by_id("MainContent_MonthBorn")).select_by_value(month)
if year != "":
    Select(driver.find_element_by_id("MainContent_YearBorn")).select_by_value(year)

# primary contact
driver.execute_script("arguments[0].click();", driver.find_element_by_id("MainContent_RoleCheckBoxList_0"))

# owner
Select(driver.find_element_by_id("MainContent_PCTitle")).select_by_value(o_title)
driver.find_element_by_id("MainContent_PCFirstName").send_keys(o_firstname)
driver.find_element_by_id("MainContent_PCLastName").send_keys(o_lastname)
driver.find_element_by_id("MainContent_PCResidentialAddress").send_keys(o_address)
driver.find_element_by_id("MainContent_PCResidentialAddressCity").send_keys(o_city)
driver.find_element_by_id("MainContent_PCResidentialAddressPostCode").send_keys(o_zip)
driver.find_element_by_id("MainContent_PCHomePhone").send_keys(o_phone1)
driver.find_element_by_id("MainContent_PCWorkPhone").send_keys(o_phone2)

# owner email
driver.find_element_by_id("MainContent_PCEmailAddress").send_keys(o_email)

# alternative
Select(driver.find_element_by_id("MainContent_ECTitle")).select_by_value(a_title)
driver.find_element_by_id("MainContent_ECFirstName").send_keys(a_firstname)
driver.find_element_by_id("MainContent_ECLastName").send_keys(a_lastname)
driver.find_element_by_id("MainContent_ECHomePhone").send_keys(a_phone1)
driver.find_element_by_id("MainContent_ECWorkPhone").send_keys(a_phone2)

# scroll
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

## TESTING
time.sleep(100)

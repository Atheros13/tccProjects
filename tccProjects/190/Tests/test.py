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
    username = "info@thecallcentre.co.nz"
    password = "TCC123456!"

    ## LOGIN TO DVS AND GO TO LOAD PROSPECT
    driver.get(address)  
    demo = True
    while demo:

        try:
            driver.find_element_by_id("MemberLoginForm_LoginForm_Email").send_keys(username)
            driver.find_element_by_id("MemberLoginForm_LoginForm_Password").send_keys(password)
            time.sleep(0.9)
            driver.find_element_by_name('action_doLogin').click()
            demo = False
        except:
            pass

        time.sleep(3)    

def process_email(email, complete):

    email.UnRead = False
    email.Move(complete)

## GLOBAL ##
driver = webdriver.Chrome(ChromeDriverManager().install())

## OPEN UPDATE, 
open_webpage("https://www-uat-animalregister.msapp.co.nz/implanter/dashboard/register") 

# microchip check
driver.find_element_by_id("microchipNumber").send_keys("123451234512345")
#driver.find_element_by_id("MainContent_ChipCrossCheck1").send_keys(chip[:3])

## TESTING
time.sleep(100)

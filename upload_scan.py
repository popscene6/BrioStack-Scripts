from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.touch_actions import TouchActions
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from datetime import *
import unittest
import autoit   
import csv
import os
import sys
import xlrd
from sys import exit
import pandas as pd
from pandas import *
from openpyxl import Workbook
import time 
import glob
import shutil
import pathlib
import subprocess
import os


#import account numbers
    
files = [name for name in os.listdir("Z:") if name.endswith(".pdf")]
acct = []
for name in files:
    number = [name.split("-")[0]]
    acct.append(number)

options = Options()
# 0 - default, 1 - Allow, 2 - Block
#set chrome options
prefs = {"profile.default_content_setting_values.automatic_downloads": 1,}
options.add_experimental_option("prefs", prefs)
options.add_argument("--disable-notifications")
options.add_argument("--start-maximized")
#set driver
driver = webdriver.Chrome(options=options)
#open browser and pull up site
driver.get("https://hmpestcontrol.briostack.com/")
#login 
id_box = driver.find_element_by_name('username')
id_box.send_keys()
pass_box = driver.find_element_by_name('password')
pass_box.send_keys()
driver.find_element_by_id('button-wrapper').click()
#Click into Brio Office
element = WebDriverWait(driver, 30).until(        
    EC.element_to_be_clickable((By.ID, "officeLink")))
element.click()
#If a popup
element = driver.find_elements_by_css_selector('.cz-popover-close .cz-close-popover-trigger')
if len(element) > 0 and element[0].is_enabled():
    element[0].click()
else: 
#Click to Customers
    element = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, ".icon-customers")))
    element.click()
#switch to iframe
element = WebDriverWait(driver, 60).until(
    EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/div[1]/iframe")))
time.sleep(2)
#Click Documents
element = WebDriverWait(driver, 60).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "#BiToolBarRadioButton-144"))) 
element.click()
time.sleep(5)
#Acct/File iteration
    #Enter Acct Number
for a, name in zip(acct, files):
    element = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//html/body/div[2]/div/div[3]/div/div[1]/input[1]")))
    time.sleep(2)
    text_input = driver.find_element_by_xpath("//html/body/div[2]/div/div[3]/div/div[1]/input[1]")
    text_input.clear()
    text_input.send_keys(a)
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    time.sleep(2)
    #customer listing    
    element = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, "//html/body/div[2]/div/div[3]/div/div[3]/div[1]/div[2]/table/tbody/tr[1]")))        
    element.click()
    #document screen
    #Click Add Document
    time.sleep(2)
    element = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//html/body/div[2]/div/div[3]/div/div[3]/div[2]/div[2]/div/div/div[5]/div[2]")))
    element.click()
    #Upload File
    time.sleep(2)
    doc = "Z://" + name
    upload = driver.find_element_by_css_selector("#fileToUpload")
    upload.send_keys(doc)
    #Click OK
    time.sleep(2)       
    element = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#body > div.modal.fade.ng-isolate-scope.in > div > div > div > div.modal-footer > button.btn.btn-primary.btn-small")))
    element.click()
    #Move File
    time.sleep(2)
    os.chdir(r'Z:\\')
    compFolder = 'Upload to Brio Completed'
    oldname = name
    shutil.move(oldname, compFolder)
    time.sleep(2)
browserExe = "chrome.exe"
os.system("taskkill /f /im "+browserExe)
sys.exit()



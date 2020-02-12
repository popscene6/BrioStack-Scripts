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

options = Options()
# 0 - default, 1 - Allow, 2 - Block
#set chrome options
prefs = {"profile.default_content_setting_values.automatic_downloads": 1,
        "download.default_directory" : r"Y:\To be printed"
}
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

#click tickets and appointment
element = WebDriverWait(driver, 60).until(
        EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/div[1]/iframe")))
time.sleep(2)
element = WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located((By.ID,  "BiLabel-143"))) 
element.click()
#click ticket listing
time.sleep(2)
element = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//html/body/div[2]/div/div[3]/div/div[3]/div[2]/div[2]/div/div/div[3]")))
time.sleep(2)
#resize ticket list
ActionChains(driver).click_and_hold(element).perform()
ActionChains(driver).move_by_offset(0, 250).release().perform()
time.sleep(2)

#Import Excel file

with open(r'Y:\BrioStack\Script Uploads\Retreat Letters.csv') as samplefile:
    reader = csv.reader(samplefile)
    next(reader)
    columns = zip(*reader)
    col1, col2, col3 = columns    
    date = []
    for item in col1:
        start = item[4:15]
        checkWords = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        repWords = ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
        for check, rep in zip(checkWords, repWords):
            start = start.replace(check, rep)
        date.append(start.replace( ' ', '/'))

#Iteration for each file

for l, d, a in zip(col2, date, col3):
        time.sleep(5)

        element = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//html/body/div[2]/div/div[3]/div/div[1]/input[1]")))
        time.sleep(2)
        text_input = driver.find_element_by_xpath("//html/body/div[2]/div/div[3]/div/div[1]/input[1]")
        text_input.clear()
        text_input.send_keys(l)
        ActionChains(driver).send_keys(Keys.ENTER).perform()
        #customer listing
        time.sleep(2)
        element = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, "//html/body/div[2]/div/div[3]/div/div[3]/div[1]/div[2]/table/tbody/tr[1]")))
        element.click()
        #select ticket
        element = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div[3]/div/div[3]/div[2]/div[2]/div/div/div[1]/div[2]/table/tbody/tr/td[text()='Job']")))
        element.click()
        #right click in list to generate document
        ActionChains(driver).move_to_element(element).context_click(element).perform()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
        ActionChains(driver).send_keys(Keys.ENTER).perform()
        #generate document screen
        driver.switch_to.default_content()
        #Template Select
        element = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-customers-page/div/brio-generate-document/form/div/div/section[1]/p-card/div/div[2]/div/div[1]/div/div/p-dropdown/div")))
        element.click() 
        ActionChains(driver).move_to_element(element).perform()
        element = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-customers-page/div/brio-generate-document/form/div/div/section[1]/p-card/div/div[2]/div/div[1]/div/div/p-dropdown/div/div[4]/div/ul/li/span[text()='Overdue Re-Treat']")))
        element.click()
        element = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-customers-page/div/brio-generate-document/form/div/div/section[1]/p-card/div/div[2]/div/div[3]/div/div/brio-calendar/span/input")))
        element.click()
        ActionChains(driver).send_keys(d).perform()
        ActionChains(driver).send_keys(Keys.TAB).perform() 
        #element = WebDriverWait(driver, 30).until(
        #        EC.element_to_be_clickable((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-customers-page/div/brio-generate-document/form/div/div/section[1]/p-card/div/div[2]/div/div[3]/div/div/brio-calendar/span/input")))
        #element.click()
        
        element = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-customers-page/div/brio-generate-document/form/div/div/section[1]/p-card/div/div[2]/div/div[3]/div[2]/div/div/input")))
        ActionChains(driver).send_keys(a).perform()
        ActionChains(driver).send_keys(Keys.TAB).perform()
        #element = WebDriverWait(driver, 15).until(
        #       EC.element_to_be_clickable((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-customers-page/div/brio-generate-document/form/div/div/section[1]/p-card/div/div[2]/div/div[3]/div/div/brio-calendar/span/button/span[1]")))
        #element.click()
        element = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-customers-page/div/brio-generate-document/form/div/div/section[1]/p-card/div/div[2]/div/div[5]/div[1]/div/div/div[1]/p-checkbox")))
        element.click()
        element = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-customers-page/div/brio-generate-document/form/div/div/section[2]/div/div/div/div[2]/p-button")))
        element.click()
        element = WebDriverWait(driver, 30).until(
               EC.element_to_be_clickable((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-preview-document/section/div/div/div/div[2]/p-button"))) #- finalize
        element.click()
        time.sleep(2)
        element = WebDriverWait(driver, 60).until(
                EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/div[1]/iframe")))
        
              #cancel out for testing - up to print PDFs
              #EC.element_to_be_clickable((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[13]/brio-customers-page/div/brio-generate-document/form/div/div/section[2]/div/div/div/div[1]/p-button")))
        #element.click()
        #element = WebDriverWait(driver, 30).until(
        #        EC.element_to_be_clickable((By.XPATH, "//html/body/p-confirmdialog[2]/div/div[3]/button[1]")))
        #element.click()
        #element = WebDriverWait(driver, 30).until(
        #        EC.element_to_be_clickable((By.XPATH, "//html/body/brio-office-root/brio-office-navigation/div/div[10]/a[1]/span")))
        #element.click()
                #driver.implicitly_wait(30)

#time.sleep(5)
     


#sys.exit()










                





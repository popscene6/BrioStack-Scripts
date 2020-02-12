import unittest
import csv
import os
import sys
import xlrd
from sys import exit
import pandas as pd
from pandas import *
from openpyxl import Workbook
import glob
from datetime import *
import time 
import shutil
import pathlib
import subprocess

x = 0 
for file in glob.glob(r'Y:\To be printed\*'):
        #print(file)
        if (file != ''):
                os.startfile(file, 'print')
                #print('Printing file >> ' + str(file))
                x = x + 1
                #Sleep for 5 seconds for printing 
        time.sleep(5)
        os.remove(file)
        time.sleep(5)

renewal = pathlib.Path(r'Y:\BrioStack\Script Uploads\Renewal Letters.csv')
retreat = pathlib.Path(r'Y:\BrioStack\Script Uploads\Retreat Letters.csv')
final = pathlib.Path(r'Y:\BrioStack\Script Uploads\Final Letters.csv')

if renewal.exists():
        os.chdir(r'Y:\BrioStack\Script Uploads')
        compFolder = 'Complete'
        oldname = 'Renewal Letters.csv'
        newname = 'Renewal Letters_' +str(date.today())+'.csv'
        shutil.move(oldname, compFolder+ '/' +newname)
        time.sleep(2)

if retreat.exists():
        os.chdir(r'Y:\BrioStack\Script Uploads')
        compFolder = 'Complete'
        oldname = 'Retreat Letters.csv'
        newname = 'Retreat Letters_' +str(date.today())+'.csv'
        shutil.move(oldname, compFolder+ '/' +newname)
        time.sleep(2)     

if final.exists():
        os.chdir(r'Y:\BrioStack\Script Uploads')
        compFolder = 'Complete'
        oldname = 'Final Letters.csv'
        newname = 'Final Letters_' +str(date.today())+'.csv'
        shutil.move(oldname, compFolder+ '/' +newname)
        time.sleep(2)

else: 
    print ("None!")
   
browserExe = "chrome.exe"
os.system("taskkill /f /im "+browserExe)
adobe = "Acrobat.exe"
os.system("taskkill /f /im "+adobe)

sys.exit()


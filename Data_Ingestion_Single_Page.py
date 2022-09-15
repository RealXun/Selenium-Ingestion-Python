#!/usr/bin/env python
# coding: utf-8
# Import Module
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
#Time is important in this code due to sometimes Google Form cannot handle the speed and we need to give it time to process the information.
import time
#We need it to read the excel file
import pandas as pd
# we need it in order to automatically recognize the user of the computer
import os
#we get a warning from the excel file. It is due openpyxl cant handle files with list boxes, it does not interfere the proper runnin of the code. 
import warnings 
#we are gonna create a log file with it. The log file will be re-written everytime we use the code.
import sys
#############################   Opening/creation of the log file   ###############################################
sys.stdout=open("LogFile.txt","w")
#############################   Excel File Path and Read Excel   ###############################################
## Somtimes Openpyxl will give a waning using Selenium. That warning does not affect the work, so we just ifnore it.
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
##To get the username of the system
username = os.getlogin()
##In this way we do not assign a windows user, it will look for it.# Creating a Path variable and Use username var in f string
path= os.path.expanduser(F'C:/Users/{username}/Downloads/Calls.xlsx')
#The tab where the incidents/processes are must be called "Results"
df=pd.read_excel(path,sheet_name='Results')
df
##We print the step in case there is an error
print ('Excel File was read', end='\n',flush=True)
#############################   Open Chrome Browser   ###############################################
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
## We print the step in case there is an error
print ('webdriver.Chrome was Opened', end='\n',flush=True)
############################################## Dictionaries Information Calls##############################################
#Agent Name Dicctionary
InfoAgent_dict={'Daniel':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div[3]',
                 'David':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div[4]'
                 }
#Source of the call Dicctionary
InfoSource_dict={'Incoming Call':'//*[@id="i9"]/div[3]',
                  'Outgoing Call':'//*[@id="i12"]/div[3]'
                 }
#Product Dicctionary
InfoProduct_dict={'Home repairs':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[2]/div[3]',
                   'Personal care':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[2]/div[4]'
                  }
#Call Result Dicctionary
InfoResult_dict={'Silent call':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[2]/div[3]',
                     'Call dropped':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[2]/div[4]',
                     'Unanswered call':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[2]/div[5]',
                     'Transferred to sales':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[2]/div[6]',
                     'Not interested':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[2]/div[7]'
                     }
print ('Dictionaries were Loaded', end='\n\n\n',flush=True)##We print the step in case there is an error
############################################## Python "For Loop" to iterate in the Excel file until there is an empty row  ##############################################
for row, datos in df.iterrows():
    InformationForm = 'https://forms.gle/oLq3tF4AhKnY2NE58'
    seconds = time.time()
    InfoAgent =  datos ['Agent Name']
    InfoSource =  datos ['Source']
    InfoPhone = datos ['Phone Number']
    InfoProduct = datos ['Product']
    InfoEmail = datos ['Email Adress']
    InfoResult = datos ['Call Result']
    print ('Variables were created', end='\n',flush=True)##We print the step in case there is an error

    ##open Google Form
    driver.get(InformationForm)
    print ('Information Google Form Opened', end='\n',flush=True)##We print the step in case there is an error

    ##Click on Next Button
    driver.find_element("xpath",'//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click()

    #Open Agent List
    driver.find_element("xpath",'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[1]/div[1]/span').click()

    #Delay 2 second
    time.sleep(2)

    ##Select Agent
    driver.find_element(By.XPATH,InfoAgent_dict[InfoAgent]).click()
    ##We print the step in case there is an error
    print ('Information Agente selected', InfoAgent, end='\n',flush=True)

    #Delay 2 second
    time.sleep(2)

    #Select Source
    driver.find_element(By.XPATH,InfoSource_dict[InfoSource]).click()
    ##We print the step in case there is an error
    print ('Information Source Selected', InfoSource, end='\n',flush=True)

    #Delay 2 second
    time.sleep(2)

    ##Write Phone Number
    driver.find_element("xpath",'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(int(InfoPhone))
    ##We print the step in case there is an error
    print ('Information Phone Number inserted', InfoPhone, end='\n',flush=True)
        
    #Delay 2 second
    time.sleep(2)

    ##Write Email
    driver.find_element("xpath",'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(InfoEmail)
    ##We print the step in case there is an error
    print ('Information Email inserted', InfoEmail, end='\n',flush=True)
        
    #Delay 2 second
    time.sleep(2)

    #Open product List
    driver.find_element("xpath",'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div[1]/div[1]/span').click()
        
    #Delay 2 second
    time.sleep(2)

    ##Select product
    driver.find_element(By.XPATH,InfoProduct_dict[InfoProduct]).click()
    ##We print the step in case there is an error
    print ('Information product Selected', InfoProduct, end='\n',flush=True)

    #Delay 2 second
    time.sleep(2)

    #Open Call Result List
    driver.find_element("xpath",'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div[1]/div[1]/span').click()

    #Delay 2 second 
    time.sleep(2)

    ##Select Call Result
    driver.find_element(By.XPATH,InfoResult_dict[InfoResult]).click()
    ##We print the step in case there is an error
    print ('Information Call Result Selected', InfoResult, end='\n',flush=True)

    #Delay 2 second 
    time.sleep(2)

    ##Click on Submit
    driver.find_element("xpath",'//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click()
    print ('The record was written in the Google Form, There were no errors.', end='\n\n',flush=True)

print ('All the records were written in the Google Form, There were no errors.', end='\n\n',flush=True) 
sys.stdout.close() #Stop of the log file.
# -*- coding: utf-8 -*-
"""
Created on Fri Jan 25 8:04:27 2019

@author: jlowh001
"""

import pandas as pd #v 0.23.4
import selenium #v 3.14
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import os
import time
from random import randint
#need xlread version 1.2
#need openpyxl v 2.5.12
from selenium.common.exceptions import NoSuchElementException     
from pathlib import *

home = str(Path.home())

  
def check_exists_by_class(txt):
    try:
        driver.find_element_by_class_name(txt)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_xpath(txt):
    try:
        driver.find_element_by_xpath(txt)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_id(txt):
    try:
        driver.find_element_by_id(txt)
    except NoSuchElementException:
        return False
    return True

os.chdir(home+'\\Desktop\\Mortgage Scrape2\\Mortgage Scrape\\')
currentdir = os.getcwd()

if os.path.isfile("index.txt") == True:
    f = open('index.txt', 'r')
    index = f.readlines()
    f.close()
    index = int(index[0])
    
if os.path.isfile("StartIndexController.txt") == True:
    f = open('StartIndexController.txt', 'r')
    index2 = f.readlines()
    f.close()
    if index2 != []:
        index2 = int(index2[0])
        
if index2 != []:
    index = index2       
    
if os.path.isfile("output.xlsx") == True:
    mergedata = pd.read_excel('output.xlsx',sheet_name='Sheet1')

data = pd.read_excel("scrape.xlsx",sheet_name ="ListReport")

data = data.iloc[index:,]

###needs to be changed to the path of the chrom driver
os.environ["PATH"] = currentdir

user = ''
pwd = ''

names = data.iloc[:,1]

phonesbyaddydf = pd.DataFrame()
phonesbynamedf = pd.DataFrame()
phonesbycodf = pd.DataFrame()

data['OWNER 1 LABEL NAME'] = data['OWNER 1 LABEL NAME'].str.split("\d+").str[0]
data['OWNER 2 LABEL NAME'] = data['OWNER 2 LABEL NAME'].str.split("\d+").str[0]

# Using Chrome to access web


######if you are restarting this; then you will need to select the code below and run to restart

capabilities = {
  'browserName': 'chrome',
  'chromeOptions':  {
    'useAutomationExtension': False,
    'forceDevToolsScreenshot': True,
    'args': ['--start-maximized', '--disable-infobars']
  }
}    
  
driver = webdriver.Chrome(desired_capabilities=capabilities)

driver.get('https://www.peoplesmart.com/')

driver.find_element_by_link_text('Sign In').click()

driver.find_element_by_id("UserName").send_keys(user)
driver.find_element_by_id("Password").send_keys(pwd)

driver.find_element_by_xpath('//div[@class="btnWrap btnLoginWrap"]').click()

try:
    for i in range(0,len(data)):
       
        if i == 0:
            driver.find_element_by_id('Find').send_keys(data.iloc[i,1])
            driver.find_element_by_id('Near').send_keys(data.iloc[i,4]+ ', ' + data.iloc[i,5])
            driver.find_element_by_xpath('//div[@class="M3 searchSubmitWrap"]').click()
        
        ###priority 1, address start
        
    
        driver.find_element_by_xpath('//a[@class="dropdown _dropButton"]').click()
        time.sleep(1)
        driver.find_element_by_xpath('//a[@class="address"]').click()
        driver.find_element_by_xpath('//input[@class="formEl formElClear textCptz _modalTitle"]').send_keys(data.iloc[i,3])
        driver.find_element_by_id('addressNear').send_keys(data.iloc[i,4]+', '+data.iloc[i,5])
        driver.find_element_by_xpath('//a[@class="btn btnPrime btnSearch btnLine"]').click() 
        time.sleep(1)
        ###subroutine:find names on sheet then match to data
        peoplenames = driver.find_elements_by_xpath('//a[@class="resultsTitle"]')
        varnames = []
        for elem in peoplenames:
            varnames.append(elem.text)
            
        test_name1 = data.iloc[i,1].split(' ')
        
        #remove middle initial from data set name
        if len(test_name1)==3:
            test_name1 = test_name1[0]+' '+test_name1[2]
        else:
            test_name1 = data.iloc[i,1]
        
        #remove middle initials from web scraped names
        for j in range(1,len(varnames)):
            if len(varnames[j].split(' '))==3:
                varnames[j] = varnames[j].split(' ')[0]+' '+varnames[j].split(' ')[2]
            else:
                varnames[j] = varnames[j]
        ########subroutine:end  
            
        def checkname(var):
            try:
                varnames.index(var)
            except ValueError:
                return(False)
            return(True)
            
        mainphone = 'None'
        
        checkname1 = checkname(test_name1)
        
        
        if str(data.iloc[i,2]) != 'nan':
            test_name2 = data.iloc[i,2].split(' ')
           #remove middle initial from data set name
            if len(test_name2)==3:
                test_name2 = test_name2[0]+' '+test_name2[2]
            else:
                test_name2 = data.iloc[i,1]
        else:
            test_name2 = 'None'
            
        checkname2 = checkname(test_name2)
           
        if checkname1:
            I_index = varnames.index(test_name1) + 1
            
            driver.find_element_by_xpath('(//a[@class="btn btnSndr btnResultsDetails btnLine"])['+str(I_index)+']').click()
            while  True:
                check_exists_by_class('reportCapTitle')
                if check_exists_by_class('reportCapTitle') == False:
                    break
                
            if check_exists_by_id('idContactPhone'):
                mainphone = driver.find_element_by_id('idContactPhone').text
            else:
                mainphone = 'None'
            mainphonedict = {'OWNER 1 LABEL NAME':data.iloc[i,1],'PrimaryAddyPhone':mainphone}
            mainphonedf = pd.DataFrame.from_records(mainphonedict,index=[i])
            phonesbyaddydf = phonesbyaddydf.append(mainphonedf)
        elif checkname2:
            I_index = varnames.index(test_name2) + 1
            
            driver.find_element_by_xpath('(//a[@class="btn btnSndr btnResultsDetails btnLine"])['+str(I_index)+']').click()
            while  True:
                check_exists_by_class('reportCapTitle')
                if check_exists_by_class('reportCapTitle') == False:
                    break
                
            if check_exists_by_id('idContactPhone'):
                mainphone = driver.find_element_by_id('idContactPhone').text
            else:
                mainphone = 'None'
            mainphonedict = {'OWNER 2 LABEL NAME':data.iloc[i,2],'PrimaryRelAddyPhone':mainphone}
            mainphonedf = pd.DataFrame.from_records(mainphonedict,index=[i])
            phonesbyaddydf = phonesbyaddydf.append(mainphonedf)
                  
        else: 
            driver.find_element_by_xpath('//a[@class="dropdown _dropButton"]').click()
            time.sleep(1)
            driver.find_element_by_xpath('//a[@class="person"]').click()
            
        
                ###people search
            if check_exists_by_xpath('//input[@class="formEl formElClear textCptz _modalTitle "]'):
                driver.find_element_by_xpath('//input[@class="formEl formElClear textCptz _modalTitle "]').send_keys(data.iloc[i,1])
                driver.find_element_by_xpath('//input[@class="formEl formElClear textCptz _modalCity _locationAutocomplete _locationAutocomplete_1 ui-autocomplete-input"]').send_keys(data.iloc[i,4]+ ', ' + data.iloc[i,5])
                driver.find_element_by_xpath('//div[@class="searchSubmitWrap"]').click()
            else:
                driver.find_element_by_id('Find').send_keys(data.iloc[i,1])
                driver.find_element_by_id('Near').send_keys(data.iloc[i,4]+ ', ' + data.iloc[i,5])
                driver.find_element_by_xpath('//div[@class="M3 searchSubmitWrap"]').click()
        
            driver.find_element_by_xpath('//div[@class="pseudoTd last-child"]/a').click()
            
            while  True:
                check_exists_by_class('reportCapTitle')
                if check_exists_by_class('reportCapTitle') == False:
                    break
          
            phones = []
            
            phone_numbers= driver.find_elements_by_xpath('//p[@class="barText textXLarge prime"]')
            
            for elem in phone_numbers:
                phones.append(elem.text)
                
            #add empty column names
            colnames = []
            for k,item in enumerate(phones):
                j = k + 1
                colnames.append("Phonebyname"+str(j))
            
            name = [data.iloc[i,1]]
            
            colnames = ['OWNER 1 LABEL NAME']+colnames
            phones = name+phones
            
            #clean and create the dataframe
            phonesdict = dict(zip(colnames, phones))
            phones_ = pd.DataFrame.from_records (phonesdict,index=[i])
            
            if len(phones_.columns)>=2:
            
                phonesbynamedf = phonesbynamedf.append(phones_,sort=False)
                
            elif str(data.iloc[i,2])!='nan':
                
                if check_exists_by_xpath('//input[@class="formEl formElClear textCptz _modalTitle "]'):
                    driver.find_element_by_xpath('//input[@class="formEl formElClear textCptz _modalTitle "]').send_keys(data.iloc[i,2])
                    driver.find_element_by_xpath('//input[@class="formEl formElClear textCptz _modalCity _locationAutocomplete _locationAutocomplete_1 ui-autocomplete-input"]').send_keys(data.iloc[i,4]+ ', ' + data.iloc[i,5])
                    driver.find_element_by_xpath('//div[@class="searchSubmitWrap"]').click()
                else:
                    driver.find_element_by_id('Find').send_keys(data.iloc[i,2])
                    driver.find_element_by_id('Near').send_keys(data.iloc[i,4]+ ', ' + data.iloc[i,5])
                    driver.find_element_by_xpath('//div[@class="M3 searchSubmitWrap"]').click()
            
                driver.find_element_by_xpath('//div[@class="pseudoTd last-child"]/a').click()
                
                while  True:
                    check_exists_by_class('reportCapTitle')
                    if check_exists_by_class('reportCapTitle') == False:
                        break
              
                phones = []
                
                phone_numbers= driver.find_elements_by_xpath('//p[@class="barText textXLarge prime"]')
                
                for elem in phone_numbers:
                    phones.append(elem.text)
                    
                #add empty column names
                colnames = []
                for k,item in enumerate(phones):
                    j = k + 1
                    colnames.append("PhoneByRelName"+str(j))
                
                name = [data.iloc[i,1]]
                
                colnames = ['OWNER 2 LABEL NAME']+colnames
                phones = name+phones
                
                #clean and create the dataframe
                phonesdict = dict(zip(colnames, phones))
                phones_ = pd.DataFrame.from_records (phonesdict,index=[i])
                
                if len(phones_.columns)==2:
                    phonesbycodf = phonesbycodf.append(phones_)
except:
    if len(phonesbyaddydf.columns) ==2:
        addydf1 = phonesbyaddydf
        finaldf=data.merge(addydf1,on='OWNER 1 LABEL NAME',how='left')
    else:
        addydf1 = phonesbyaddydf.iloc[:,[1,3]]
        addydf1 = addydf1.dropna()
        addydf2 = phonesbyaddydf.iloc[:,[0,2]]
        addydf2 = addydf2.dropna()
        finaldf=data.merge(addydf2, on='OWNER 1 LABEL NAME', how='left')
        finaldf=finaldf.merge(addydf1,on='OWNER 2 LABEL NAME',how='left')
        
    finaldf=finaldf.merge(phonesbynamedf,on='OWNER 1 LABEL NAME',how='left')
    
    if len(phonesbycodf) > 0:
        finaldf=finaldf.merge(phonesbycodf,on='OWNER 2 LABEL NAME',how='left')
        
    finaldf=finaldf.iloc[0:i,:]
    
    finaldf = mergedata.append(finaldf, ignore_index=True)
    
    writer = pd.ExcelWriter('output.xlsx')
    finaldf.to_excel(writer,'Sheet1')
    writer.save()
    f = open("index.txt", "w")
    f.write(repr(i+index)) 
    f.close()
    
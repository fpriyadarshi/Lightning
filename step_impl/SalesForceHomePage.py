from getgauge.python import step, before_scenario, Messages, before_suite, before_spec, after_spec, after_step
from getgauge.python import after_suite
from getgauge.python import before_suite
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
import fnmatch
import pdb
from getgauge import messages
from getgauge.python import DataStoreFactory, Screenshots
from datetime import timedelta
import xlsxwriter
import csv
from datetime import date
from datetime import datetime
from openpyxl.reader.excel import load_workbook
from openpyxl import Workbook
from simple_salesforce import Salesforce, SFType, SalesforceLogin
import yaml
from time import sleep
import logging
from step_impl import Drivers
from simple_salesforce.format import format_soql
from selenium.common.exceptions import UnexpectedAlertPresentException


class SalesForceHome():
    driverWait = WebDriverWait(Drivers.driver, 50)

    @step("Search and Login as user <userProfile>")
    def search_user(self, userProfile):

        userName = ""
        if userProfile == "US Sales Representative":
            userName = os.getenv("US_SALES_REPRESENTATIVE")
        elif userProfile == "US Sales Ops":
            userName = os.getenv("US_SALES_OPERATIONS")
        elif userProfile == "US Sales Management":
            userName = os.getenv("US_SALES_MANAGEMENT")
        elif userProfile == "CA Sales Representative":
            userName = os.getenv("CA_SALES_REPRESENTATIVE")
        elif userProfile == "CA Sales Ops":
            userName = os.getenv("CA_SALES_OPERATIONS")
        elif userProfile == "CA Sales Management":
            userName = os.getenv("CA_SALES_MANAGEMENT")
        elif userProfile == "Finance":
            userName = os.getenv("FINANACE")
        elif userProfile == "Billing/Credit":
            userName = os.getenv("BILLING_CREDIT")
        elif userProfile == "Integration":
            userName = os.getenv("INTEGRAION")
        elif userProfile == "DG Specialty Production":
            userName = os.getenv("DG_SPECIALTY_PRODUCTION")
        elif userProfile == "IS Specialty Production":
            userName = os.getenv("IS_SPECIALTY_PRODUCTION")
        elif userProfile == "MERC Specialty Production":
            userName = os.getenv("MERC_SPECIALTY_PRODUCTION")
        elif userProfile == "SSD Specialty Production":
            userName = os.getenv("SSD_SPECIALTY_PRODUCTION")
        elif userProfile == "SSMG Specialty Production":
            userName = os.getenv("SSMG_SPECIALTY_PRODUCTION")
        elif userProfile == "US_Regional Sales Approver":
            userName = os.getenv("RSA-US")
        elif userProfile == "CA_Regional Sales Approver":
            userName = os.getenv("RSA-CA")
        elif userProfile == "US_Sales Team Approver (GSM)":
            userName = os.getenv("STA-US")
        elif userProfile == "CA_Sales Team Approver (GSM)":
            userName = os.getenv("STA-CA")
        elif userProfile == "Finance_User":
            userName = os.getenv("FINANCE-USER")
        else:
            userName = os.getenv("DEFAULT_USER")

        queryData = Drivers.sf.query(format_soql(
            "SELECT Id, Name,Email FROM User WHERE Name = {} AND IsActive = True", userName))
        # pdb.set_trace()
        currentURL = Drivers.driver.current_url
        userURL = f"{currentURL.split('.com/')[0]}.com/{queryData['records'][0]['Id']}"
        Drivers.driver.get(userURL)        

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, "//div[@title='User Detail']")))
        
        btnUserDetails = Drivers.driver.find_element_by_xpath(
            "//div[@title='User Detail']")
        btnUserDetails.click()

        sleep(5)        

        title = f"User: {userName} ~ Salesforce - Unlimited Edition"
        Drivers.driverWait.until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, f"//iframe[@title='{title}']")))
            # EC.frame_to_be_available_and_switch_to_it((By.XPATH, f"//iframe[@title='{title}']")))
        # iframeElement = Drivers.driver.find_element_by_xpath(
        #     f"//iframe[@title='{title}']")
        # Drivers.driver.switch_to.frame(iframeElement)
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, "//input[@name='login']")))
        btnLogin = Drivers.driver.find_element_by_xpath(
            "//input[@name='login']")
        btnLogin.click()
        sleep(5)
        # Drivers.driver.switch_to.default_content()

        Drivers.driverWait.until(EC.visibility_of_element_located(
            (By.XPATH, f"//span[contains(text(),'Logged in as {userName}')]")))
        Messages.write_message("Logged in as: " + userName)
        # Drivers.driverWait.until(EC.visibility_of_element_located(
        #     (By.XPATH, "//img[@title='User']")))
        # btnUser = Drivers.driver.find_element_by_xpath(
        #     "//img[@title='User']")
        # sleep(5)
        # btnUser.click() 
        
        # Drivers.driverWait.until(EC.visibility_of_element_located(
        #     (By.XPATH, "//img[@title='User']")))
        # btnUser = Drivers.driver.find_element_by_xpath(
        #     "//img[@title='User']")
        # sleep(5)
        # btnUser.click()        
        
        # userXpath = f"//div[@class='profile-card-indent']/h1/a[text()='{userName}']"

        # Drivers.driverWait.until(EC.visibility_of_element_located(
        #     (By.XPATH, userXpath)))
        
        # EC.visibility_of_element_located((By.XPATH, "//img[@title='User']/span[text()='" + userName + "']")))
        

    @step("Search for <region> <objectName> record <recordName>")
    def search_record(self, region, objectName, recordName):
        if recordName == "FROM PROPERTY":
            if region == "US":
                recordName = os.getenv("US_SALES_REPRESENTATIVE_ACCOUNT")
            if region == "CA":
                recordName = os.getenv("CA_SALES_REPRESENTATIVE_ACCOUNT")

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, "//img[@title='All Tabs']")))

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'phSearchInput')))
        txtBoxGlobalSearch = Drivers.driver.find_element_by_id(
            "phSearchInput")
        txtBoxGlobalSearch.send_keys(recordName)
        txtBoxGlobalSearch.send_keys(u'\ue007')

        if objectName == "Accounts":
            Drivers.driverWait.until(
                EC.visibility_of_element_located((By.ID, 'records')))
            linkPeople = Drivers.driver.find_element_by_partial_link_text(
                "Accounts")
            linkPeople.click()
            sleep(0.5)
        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, 'phSearchButton')))
        # btnGlobalSearch = Drivers.driver.find_element_by_id(
        #     "phSearchButton")
        # btnGlobalSearch.click()

    @step("Open searched <objectName> record <recordName> of <region>")
    def open_searched_record(self, objectName, recordName, region):
        if objectName == "Accounts":
            if recordName == "FROM PROPERTY":
                if region == "US":
                    recordName = os.getenv("US_SALES_REPRESENTATIVE_ACCOUNT")
                if region == "CA":
                    recordName = os.getenv("CA_SALES_REPRESENTATIVE_ACCOUNT")

        linkSearchedRecord = Drivers.driver.find_element_by_link_text(
            recordName)
        linkSearchedRecord.click()
        Messages.write_message("Searched Record Opened: " + recordName)

    @after_step
    def after_step_hook(self, context):
        if context.step.is_failing == True:
            Messages.write_message(context.step.text)
            # Messages.write_message(context.step.message)
            Screenshots.capture_screenshot()

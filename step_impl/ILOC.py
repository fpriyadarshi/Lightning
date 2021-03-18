from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException
import sys
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from uuid import uuid1
from xlrd import open_workbook
import requests
from getgauge.python import step, Messages, after_suite, after_spec, before_suite, after_step, before_step
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import uuid
import pandas as pd
import os
from pathlib import Path
import sys
import xlsxwriter
from datetime import date
from datetime import timezone, tzinfo, timedelta
from time import sleep
from datetime import datetime
import random
from random import randint
from step_impl import Drivers
from step_impl import Common_Steps
from step_impl import Utils
from selenium.webdriver.support.select import Select
from getgauge.python import data_store, Screenshots
import re
import shutil
import yaml
import pdb
import requests
import json
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook.workbook import Workbook
from selenium.webdriver import ActionChains
import pytz
from dateutil import tz
from simple_salesforce.format import format_soql
from selenium.webdriver.support.wait import WebDriverWait

rec = 0


class ILOC:
    baseWindow = None
    childWindows = []
    lookupWindows = []

    def takeScreenShot(self, failure="E-"):
        today = date.today()
        screenShotDateTime = today.strftime("%m%d%Y-") + str(randint(1, 99999999))
        rootPath = Path(__file__).parents[1]
        
        reportDirectory = str(rootPath) + "\\reports\\" 
        failureDirectory = str(rootPath) + "\\reports\\Failures" 
        if not os.path.exists(failureDirectory):
            os.chdir(reportDirectory)
            os.mkdir("Failures")

        screenShotFileName = str(
            rootPath) + "\\reports\\Failures\\" + f"{failure}{str(screenShotDateTime)}.png"
        print("ScreenShot Path: ", screenShotFileName)
        # Drivers.driver.save_screenshot(f"{screenShotDateTime}.png")
        Drivers.driver.save_screenshot(screenShotFileName)

    def verifyATAB(self, createdByUserID, formattedISODate):
        batchATAB = ["UpdateAgreementQueueableService", "UpdateAgreementQueueableService", "DataMappingQueueableService",
                     "UpdateAgreementQueueableService", "UpdateAgreementQueueableService", "DataMappingQueueableService"]

        batchATABStatus = {}
        apexClassName = None
        apexClassStatus = None
        for batchValue in batchATAB:
            print("Searching Current Batch:", batchValue)
            isCompleted = False
            cntr = 1
            while isCompleted == False and cntr < 5:
                cntr = cntr + 1
                print("in while block")
                batchATABSOQL = f"SELECT CreatedDate, ApexClass.Name, MethodName, Status, CompletedDate FROM AsyncApexJob where ApexClass.Name = '{batchValue}' and CreatedById = '{createdByUserID}' and CreatedDate > {formattedISODate}"
                apexJobsResult = Drivers.sf.query_all(
                    query=batchATABSOQL)
                apexJobsData = apexJobsResult['records']
                print("\n", apexJobsData)
                if len(apexJobsData) > 0:
                    if apexJobsData[0]['ApexClass'] != None:
                        print(
                            f"Verifying {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
                        if apexJobsData[0]['Status'] == 'Completed':
                            print(
                                f"Verified {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
                            apexClassName = apexJobsData[0]['ApexClass']['Name']
                            apexClassStatus = apexJobsData[0]['Status']
                            batchATABStatus[apexClassName] = apexClassStatus
                            isCompleted = True
                            Messages.write_message(
                                f"Verified {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
                        elif apexJobsData[0]['Status'] == 'Failed':
                            print(
                                f"Verified {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
                            Messages.write_message(
                                f"Verified {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
                            apexClassName = apexJobsData[0]['ApexClass']['Name']
                            apexClassStatus = apexJobsData[0]['Status']
                            batchATABStatus[apexJobsData[0]
                                            ['ApexClass']['Name']] = apexClassStatus
                            isCompleted = True
                        else:
                            sleep(15)
                else:
                    sleep(15)
                    isCompleted = False
        return batchATABStatus

    def getDateTime(self):
        currentDateTime = datetime.now(
            pytz.timezone("America/Los_Angeles"))
        twoMinutes = timedelta(minutes=2)
        end = currentDateTime - twoMinutes

        current_month = end.strftime('%m')
        # current_month_text = end.strftime('%h')
        # current_month_text = end.strftime('%B')

        current_day = end.strftime('%d')
        # current_day_text = end.strftime('%a')
        # current_day_full_text = end.strftime('%A')

        # current_weekday_day_of_today = end.strftime('%w')

        current_year_full = end.strftime('%Y')
        # current_year_short = end.strftime('%y')

        current_second = end.strftime('%S')
        current_minute = end.strftime('%M')
        current_hour = end.strftime('%H')
        # current_hour_word = end.strftime('%I')

        # d = datetime.datetime.today().replace(microsecond=0)
        # print(d.isoformat())

        finalDateTime = datetime(int(current_year_full), int(current_month), int(current_day), int(
            current_hour), int(current_minute), int(current_second), tzinfo=tz.gettz("America/Los_Angeles")).isoformat()
        print(datetime(int(current_year_full), int(current_month), int(current_day), int(current_hour), int(
            current_minute), int(current_second), tzinfo=tz.gettz("America/Los_Angeles")).isoformat())

        Messages.write_message(f"ATAB RUN TIME: {finalDateTime}")
        return finalDateTime

    def autoIncrement(self):
        global rec
        pStart = 1  # adjust start value, if req'd
        pInterval = 1  # adjust interval value, if req'd
        if (rec == 0):
            rec = pStart
        else:
            rec = rec + pInterval
        return str(rec)

    def sf_api_call(self, action, access_token, instance_url, parameters={}, method='get', data={}):
        headers = {
            'Content-type': 'application/json',
            'Accept-Encoding': 'gzip',
            'Authorization': 'Bearer %s' % access_token
        }
        if method == 'get':
            r = requests.request(method, instance_url+action,
                                 headers=headers, params=parameters, timeout=50)
        elif method in ['post', 'patch']:
            r = requests.request(method, instance_url+action,
                                 headers=headers, json=data, params=parameters, timeout=50)
        else:
            # other methods not implemented in this example
            raise ValueError('Method should be get or post or patch.')
        print('Debug: API %s call: %s' % (method, r.url))
        if r.status_code < 300:
            if method == 'patch':
                return None
            else:
                return r.json()
        else:
            raise Exception('API error when calling %s : %s' %
                            (r.url, r.content))

    def get_approver(self, ilocRegion, contractId):
        params = {
            "grant_type": "password",
            "client_id": os.getenv("CLIENT_ID"),  # Consumer Key
            "client_secret": os.getenv("CLIENT_SECRET"),  # Consumer Secret
            "username": os.getenv("USER_ID"),  # The email you use to login
            # Concat your password and your security token
            "password": f"{os.getenv('USER_PASSWORD')}{os.getenv('USER_SECURITY_TOKEN')}"
            #     "security_token": 'zGfFva89UOziolmA4BGtZ0oON'
        }
        r = requests.post(
            "https://test.salesforce.com/services/oauth2/token", params=params)
        access_token = r.json().get("access_token")
        instance_url = r.json().get("instance_url")

        masterApproversSSD = []
        masterApproversALL = []
        customSettingILOCTotal = 0

        cal = r"/services/data/v48.0/query/?q=SELECT ProductLine__c, UserId__c from MasterApprover__c"
        # print(json.dumps(sf_api_call(cal), indent=2))
        print(self.sf_api_call(cal, access_token, instance_url))
        Messages.write_message(self.sf_api_call(
            cal, access_token, instance_url))
        customSettingData = self.sf_api_call(cal, access_token, instance_url)

        for customSetting in customSettingData['records']:
            print(customSetting['ProductLine__c'],
                  " ", customSetting['UserId__c'])
            soql = f"SELECT Name FROM User where Id = '{customSetting['UserId__c']}'"
            queryResult = Drivers.sf.query_all(query=soql)
            records = queryResult['records']
            queueId = records[0]['Name']
            if customSetting['ProductLine__c'] == 'SSD':
                masterApproversSSD.append(records[0]['Name'])
            elif customSetting['ProductLine__c'] == 'ALL':
                masterApproversALL.append(records[0]['Name'])
        Messages.write_message(f"SSD MASTER APPROVERS: {masterApproversSSD}")
        Messages.write_message(
            f"ALL PL MASTER APPROVERS: {masterApproversALL}")
        cal = r"/services/data/v48.0/query/?q=SELECT Key__c, Value__c from master_configuration__c"
        print(self.sf_api_call(cal, access_token, instance_url))
        customSettingData = self.sf_api_call(cal, access_token, instance_url)
        for customSetting in customSettingData['records']:
            if customSetting['Key__c'] == 'ILOC Total Cost':
                print(customSetting['Key__c'], " ", customSetting['Value__c'])
                Messages.write_message(
                    customSetting['Key__c'] + " " + customSetting['Value__c'])
                customSettingILOCTotal = int(customSetting['Value__c'])

        # Get all opportunity record types :
        # opportunityMap = {}
        ilocApprovalsMap = {}
        instoreApprovers = []
        merchandisingApprovers = []
        checkout51Approvers = []
        digitalApprovers = []
        digitalCanadaApprovers = []
        fsiApprovers = []
        smartSourceDirectApprovers = []
        smartSourceDirectCanadaApprovers = []
        ssmgApprover = []

        currencyCode = ilocRegion

        soql = f"SELECT Account.Account_Sales_Office__c,AccountId FROM Contract where Id = '{contractId}'"
        result = Drivers.sf.query_all(query=soql)
        contractData = result['records']
        userSalesOffice = contractData[0]['Account']['Account_Sales_Office__c']
        contratAccountId = contractData[0]['AccountId']

        ilocApprovalsSOQL = f"SELECT Id,Contract_Status__c,Owner.Name,Product_Line__c, Status__c FROM Contract_Approval__c WHERE Contract__c = '{contractId}'"
        ilocApprovalsResult = Drivers.sf.query_all(query=ilocApprovalsSOQL)
        ilocApprovalsData = ilocApprovalsResult['records']
        print("Total Records :", len(ilocApprovalsData))

        for ilocApproval in ilocApprovalsData:
            if ilocApproval['Product_Line__c'] == 'InStore' and ilocApproval['Status__c'] != 'Approved':
                soql = f"Select Id from Group where type='Queue' and Name='{ilocApproval['Owner']['Name']}'"
                queryResult = Drivers.sf.query_all(query=soql)
                records = queryResult['records']
                queueId = records[0]['Id']
                print("Queue ID: ", records[0]['Id'],
                      " of queue ", ilocApproval['Owner']['Name'])

                soql = f"SELECT User.Id, User.Name FROM User WHERE IsActive = True and Id IN (SELECT UserOrGroupId FROM GroupMember WHERE GroupId = '{queueId}')"
                queryResult = Drivers.sf.query_all(query=soql)
                records = queryResult['records']
                print(records)
                for idx, record in enumerate(records):
                    instoreApprovers.append(record['Name'])
                print("Instore Approver: ", instoreApprovers)
                ilocApprovalsMap['InStore'] = random.choice(instoreApprovers)

            elif ilocApproval['Product_Line__c'] == 'Merchandising' and ilocApproval['Status__c'] != 'Approved':
                soql = f"Select Id from Group where type='Queue' and Name='{ilocApproval['Owner']['Name']}'"
                queryResult = Drivers.sf.query_all(query=soql)
                records = queryResult['records']
                queueId = records[0]['Id']
                print("Queue ID: ", records[0]['Id'],
                      " of queue ", ilocApproval['Owner']['Name'])

                soql = f"SELECT User.Id, User.Name FROM User WHERE IsActive = True and Id IN (SELECT UserOrGroupId FROM GroupMember WHERE GroupId = '{queueId}')"
                queryResult = Drivers.sf.query_all(query=soql)
                records = queryResult['records']
                print(records)
                for idx, record in enumerate(records):
                    merchandisingApprovers.append(record['Name'])
                print("Merchandising Approver", merchandisingApprovers)
                ilocApprovalsMap['Merchandising'] = random.choice(
                    merchandisingApprovers)

            elif ilocApproval['Product_Line__c'] == 'Checkout 51' and ilocApproval['Status__c'] != 'Approved':
                if 'Digital Support Rep' in ilocApproval['Owner']['Name']:
                    soql = f"SELECT User__r.Id, User__r.Name FROM Account_Team__c WHERE Account__c = '{contratAccountId}' and Territory_Category__c in ('Digital','Checkout_51','Digital- Canada') and Role_In_Territory__c = 'Support'"
                    queryResult = Drivers.sf.query_all(query=soql)
                    records = queryResult['records']
                    print(records)

                    for record in records:
                        if record['User__r']['Name'] not in checkout51Approvers:
                            checkout51Approvers.append(
                                record['User__r']['Name'])
                    print("Checkout 51 Sales Rep Approver: ",
                          checkout51Approvers)
                elif 'Master Approvers' in ilocApproval['Owner']['Name']:
                    soql = f"SELECT Approver__r.Name,CurrencyIsoCode,GSMUser__r.Name,RecordType.Name,Region_Name__c FROM Approver_Matrix__c where Region_Name__c='{userSalesOffice}' and RecordType.Name='Checkout 51' and CurrencyIsoCode = '{currencyCode}'"
                    queryResult = Drivers.sf.query_all(query=soql)
                    records = queryResult['records']
                    if len(records) > 0:
                        print(records)
                        checkout51Approvers.append(
                            records[0]['Approver__r']['Name'])
                    else:
                        checkout51Approvers.append(
                            random.choice(masterApproversALL))
                    print("Checkout 51 Master Approver: ", checkout51Approvers)
                else:
                    checkout51Approvers.append(ilocApproval['Owner']['Name'])
                    print("Checkout 51 Approver: ", checkout51Approvers)
                ilocApprovalsMap['Checkout 51'] = random.choice(
                    checkout51Approvers)

            elif ilocApproval['Product_Line__c'] == 'Digital- Canada' and ilocApproval['Status__c'] != 'Approved':
                if 'Digital Support Rep' in ilocApproval['Owner']['Name']:
                    soql = f"SELECT User__r.Id, User__r.Name FROM Account_Team__c WHERE Account__c = '{contratAccountId}' and Territory_Category__c in ('Checkout_51','Digital- Canada') and Role_In_Territory__c = 'Support'"
                    queryResult = Drivers.sf.query_all(query=soql)
                    records = queryResult['records']
                    print(records)
                    i = 1
                    for record in records:
                        if record['User__r']['Name'] not in digitalCanadaApprovers:
                            digitalCanadaApprovers.append(
                                record['User__r']['Name'])
                            i = i + 1
                    print("Digital Canada Sales Rep  Approver: ",
                          digitalCanadaApprovers)
                elif 'Master Approvers' in ilocApproval['Owner']['Name']:
                    soql = f"SELECT Approver__r.Name,CurrencyIsoCode,GSMUser__r.Name,RecordType.Name,Region_Name__c FROM Approver_Matrix__c where Region_Name__c='{userSalesOffice}' and RecordType.Name='Digital- Canada' and CurrencyIsoCode = '{currencyCode}'"
                    queryResult = Drivers.sf.query_all(query=soql)
                    records = queryResult['records']
                    if len(records) > 0:
                        print(records)
                        digitalCanadaApprovers.append(
                            records[0]['Approver__r']['Name'])
                    else:
                        digitalCanadaApprovers.append(
                            random.choice(masterApproversALL))
                    print("Digital Canada Master Approver: ",
                          digitalCanadaApprovers)
                else:
                    digitalCanadaApprovers.append(
                        ilocApproval['Owner']['Name'])
                    print("Digital Canada Approver: ", digitalCanadaApprovers)
                ilocApprovalsMap['Digital- Canada'] = random.choice(
                    digitalCanadaApprovers)

            elif ilocApproval['Product_Line__c'] == 'Digital' and ilocApproval['Status__c'] != 'Approved':
                if 'Digital Support Rep' in ilocApproval['Owner']['Name']:
                    soql = f"SELECT User__r.Id, User__r.Name FROM Account_Team__c WHERE Account__c = '{contratAccountId}' and Territory_Category__c in ('Digital','Checkout_51') and Role_In_Territory__c = 'Support'"
                    queryResult = Drivers.sf.query_all(query=soql)
                    records = queryResult['records']
                    print(records)
                    i = 1
                    for record in records:
                        if record['User__r']['Name'] not in digitalApprovers:
                            digitalApprovers.append(record['User__r']['Name'])
                            i = i + 1
                    print("Digital Sales Rep Approver: ", digitalApprovers)
                elif 'Master Approvers' in ilocApproval['Owner']['Name']:
                    soql = f"SELECT Approver__r.Name,CurrencyIsoCode,GSMUser__r.Name,RecordType.Name,Region_Name__c FROM Approver_Matrix__c where Region_Name__c='{userSalesOffice}' and RecordType.Name='Digital' and CurrencyIsoCode = '{currencyCode}'"
                    queryResult = Drivers.sf.query_all(query=soql)
                    records = queryResult['records']
                    if len(records) > 0:
                        print(records)
                        digitalApprovers.append(
                            records[0]['Approver__r']['Name'])
                    else:
                        digitalApprovers.append(
                            random.choice(masterApproversALL))
                    print("Digital Master Approver: ", digitalApprovers)
                else:
                    digitalApprovers.append(ilocApproval['Owner']['Name'])
                    print("Digital Approver: ", digitalApprovers)
                ilocApprovalsMap['Digital'] = random.choice(digitalApprovers)

            elif ilocApproval['Product_Line__c'] == 'FSI' and ilocApproval['Status__c'] != 'Approved':
                if 'Master Approvers' in ilocApproval['Owner']['Name']:
                    soql = f"SELECT Approver__r.Name,CurrencyIsoCode,GSMUser__r.Name,RecordType.Name,Region_Name__c FROM Approver_Matrix__c where Region_Name__c='{userSalesOffice}' and RecordType.Name='Digital' and CurrencyIsoCode = '{currencyCode}'"
                    queryResult = Drivers.sf.query_all(query=soql)
                    records = queryResult['records']
                    if len(records) > 0:
                        print(records)
                        fsiApprovers.append(records[0]['Approver__r']['Name'])
                    else:
                        fsiApprovers.append(random.choice(masterApproversALL))
                    print("FSI Master Approver: ", fsiApprovers)
                else:
                    fsiApprovers.append(ilocApproval['Owner']['Name'])
                    print("FSI Approver: ", fsiApprovers)
                ilocApprovalsMap['FSI'] = random.choice(fsiApprovers)

            elif ilocApproval['Product_Line__c'] == 'SmartSource Direct' and ilocApproval['Status__c'] != 'Approved':
                if 'Master Approvers' in ilocApproval['Owner']['Name']:
                    soql = f"SELECT Approver__r.Name,CurrencyIsoCode,GSMUser__r.Name,RecordType.Name,Region_Name__c FROM Approver_Matrix__c where Region_Name__c='{userSalesOffice}' and RecordType.Name='Digital' and CurrencyIsoCode = '{currencyCode}'"
                    queryResult = Drivers.sf.query_all(query=soql)
                    records = queryResult['records']
                    if len(records) > 0:
                        print(records)
                        smartSourceDirectApprovers.append(
                            records[0]['Approver__r']['Name'])
                    else:
                        smartSourceDirectApprovers.append(
                            random.choice(masterApproversSSD))
                    print("SmartSource Direct Master Approver: ",
                          smartSourceDirectApprovers)
                else:
                    smartSourceDirectApprovers.append(
                        ilocApproval['Owner']['Name'])
                    print("SmartSource Direct Approver: ",
                          smartSourceDirectApprovers)
                ilocApprovalsMap['SmartSource Direct'] = random.choice(
                    smartSourceDirectApprovers)
            elif ilocApproval['Product_Line__c'] == 'SmartSource Direct- Canada' and ilocApproval['Status__c'] != 'Approved':
                if 'Master Approvers' in ilocApproval['Owner']['Name']:
                    soql = f"SELECT Approver__r.Name,CurrencyIsoCode,GSMUser__r.Name,RecordType.Name,Region_Name__c FROM Approver_Matrix__c where Region_Name__c='{userSalesOffice}' and RecordType.Name='Digital' and CurrencyIsoCode = '{currencyCode}'"
                    queryResult = Drivers.sf.query_all(query=soql)
                    records = queryResult['records']
                    if len(records) > 0:
                        print(records)
                        smartSourceDirectCanadaApprovers.append(
                            records[0]['Approver__r']['Name'])
                    else:
                        smartSourceDirectCanadaApprovers.append(
                            random.choice(masterApproversSSD))
                    print("SmartSource Direct Canada Master Approver: ",
                          smartSourceDirectApprovers)
                else:
                    smartSourceDirectCanadaApprovers.append(
                        ilocApproval['Owner']['Name'])
                    print("SmartSource Direct Canada Approver: ",
                          smartSourceDirectCanadaApprovers)
                ilocApprovalsMap['SmartSource Direct- Canada'] = random.choice(
                    smartSourceDirectCanadaApprovers)
        [print(key, value) for key, value in ilocApprovalsMap.items()]
        [Messages.write_message(f"{key} {value}")
         for key, value in ilocApprovalsMap.items()]
        return ilocApprovalsMap

    def search_user_login(self, userName):

        queryData = (Drivers.sf.query(format_soql(
            "SELECT Id, Name,Email FROM User WHERE Name = {} AND IsActive = True", userName)))['records'][0]
        currentURL = Drivers.driver.current_url
        userURL = f"{currentURL.split('.com/')[0]}.com/{queryData['Id']}"
        Drivers.driver.get(userURL)

        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, 'phSearchInput')))
        # txtBoxGlobalSearch = Drivers.driver.find_element_by_id(
        #     "phSearchInput")
        # txtBoxGlobalSearch.send_keys(userName)

        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, 'phSearchButton')))
        # buttonSearch = Drivers.driver.find_element_by_id("phSearchButton")
        # buttonSearch.click()
        # sleep(0.5)

        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, 'records')))
        # linkPeople = Drivers.driver.find_element_by_partial_link_text(
        #     "People")
        # linkPeople.click()
        # sleep(0.5)

        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, 'User_body')))
        # linkUser = Drivers.driver.find_element_by_link_text(userName)
        # linkUser.click()
        # sleep(0.5)

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'moderatorMutton')))
        btnUserMenu = Drivers.driver.find_element_by_id(
            "moderatorMutton")
        sleep(1)
        btnUserMenu.click()
        sleep(1)

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'USER_DETAIL')))
        listUserDetails = Drivers.driver.find_element_by_id(
            "USER_DETAIL")
        listUserDetails.click()
        sleep(0.5)

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, "//input[@name='login']")))
        btnLogin = Drivers.driver.find_element_by_xpath(
            "//input[@name='login']")
        btnLogin.click()
        sleep(5)

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, "//div[@id='userNavButton']/span[text()='" + userName + "']")))
        print("Logged in as: ", userName)

    def logout(self):
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, "userNavButton")))
        userMenu = Drivers.driver.find_element_by_id("userNavButton")
        userMenu.click()

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.LINK_TEXT, "Logout")))
        linkLogout = Drivers.driver.find_element_by_link_text("Logout")
        linkLogout.click()

    def get_billing_account(self, ilocRegion, billingAccountType):
        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        print(iLOCDetails)
        billingAccountQuery = ""
        url = os.getenv("DYNAMO_DB_URL")
        billingAccountName = ""
        billingSiteNumber = ""
        noAccountFound = False
        headers = {
            'x-user-name': os.getenv("x-user-name"),
            'x-user-password': os.getenv("x-user-password"),
            'x-traceId': os.getenv("x-traceId"),
            'Accept': os.getenv("Accept"),
            'Cache-Control': os.getenv("Cache-Control"),
            'Host': os.getenv("Host"),
            'Accept-Encoding': os.getenv("Accept-Encoding"),
            'Connection': os.getenv("Connection"),
            'cache-control': os.getenv("cache-control")
        }

        if ilocRegion == "US":
            billingAccountQuery = "SELECT Id,Name,Site,Currency_Code__c FROM Account WHERE Type not in ('Prospect', 'Corporate') AND Status__c = 'Active' AND Currency_Code__c = 'USD' AND Location_Type__c='Bill To'"
        elif ilocRegion == "CA":
            billingAccountQuery = "SELECT Id,Name,Site,Currency_Code__c FROM Account WHERE Type not in ('Prospect', 'Corporate') AND Status__c = 'Active' AND Currency_Code__c = 'CAD' AND Location_Type__c= 'Bill To'"

        queryResult = Drivers.sf.query_all(query=billingAccountQuery)
        recDetails = queryResult['records']
        print(f"Total Billing Accounts: {len(recDetails)}")
        Messages.write_message(f"Total Billing Accounts: {len(recDetails)}")
        cnt = 1
        for rec in recDetails:
            querystring = {"partyAccountAndSiteNumber": rec["Site"]}
            response = requests.request(
                "GET", url, headers=headers, params=querystring)
            if response.status_code == 200:
                jsonRes = response.json()
                print(jsonRes)
                Messages.write_message(f"[{cnt}] \n {jsonRes}")
                if billingAccountType == "CLIENT_ORDER_LIMIT_LT_ILOC_TOTAL" and jsonRes["CREDIT_CHECKING"] == True and jsonRes["TOLERANCE"] > 0 and jsonRes['TRX_CREDIT_LIMIT'] > 5000:
                    clientOrderLimit = int(jsonRes['TRX_CREDIT_LIMIT']) + int(
                        jsonRes['TRX_CREDIT_LIMIT']) * int(jsonRes['TOLERANCE']) / 100
                    if (int(clientOrderLimit) < int(iLOCDetails["ILOC_GRAND_TOTAL"])):
                        billingAccountName = jsonRes["ACCOUNT_NAME"]
                        billingSiteNumber = jsonRes["SITE_NUMBER"]
                        noAccountFound = False
                        print(jsonRes)
                        Messages.write_message(
                            "---------------------------------------------")
                        Messages.write_message(
                            "\CLIENT_ORDER_LIMIT_LT_ILOC_TOTAL")
                        Messages.write_message(jsonRes)
                        Messages.write_message(
                            "---------------------------------------------")
                    break
                elif billingAccountType == "CLIENT_ORDER_LIMIT_GT_ILOC_TOTAL" and jsonRes["CREDIT_CHECKING"] == True and jsonRes["TOLERANCE"] > 0 and jsonRes['TRX_CREDIT_LIMIT'] > 5000:
                    clientOrderLimit = int(jsonRes['TRX_CREDIT_LIMIT']) + int(
                        jsonRes['TRX_CREDIT_LIMIT']) * int(jsonRes['TOLERANCE']) / 100
                    print(
                        f"Client Order Limit {clientOrderLimit} > ILOC GT {int(iLOCDetails['ILOC_GRAND_TOTAL'])}")
                    Messages.write_message(
                        f"Client Order Limit {clientOrderLimit} > ILOC GT {int(iLOCDetails['ILOC_GRAND_TOTAL'])}")
                    if int(clientOrderLimit) > int(iLOCDetails["ILOC_GRAND_TOTAL"]):
                        billingAccountName = jsonRes["ACCOUNT_NAME"]
                        billingSiteNumber = jsonRes["SITE_NUMBER"]
                        noAccountFound = False
                        print(jsonRes)
                        Messages.write_message(
                            "---------------------------------------------")
                        Messages.write_message(
                            "\CLIENT_ORDER_LIMIT_GT_ILOC_TOTAL")
                        Messages.write_message(jsonRes)
                        Messages.write_message(
                            "---------------------------------------------")
                    break
                elif billingAccountType == "CREDIT_CHECKING_FALSE" and jsonRes["CREDIT_CHECKING"] == False:
                    print(jsonRes["SITE_NUMBER"], " : ",
                          jsonRes["ACCOUNT_NAME"], " : ", jsonRes["CREDIT_CHECKING"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    break
                elif billingAccountType == "PREPAY_ACCOUNT_TRUE" and jsonRes["PREPAY_ACCOUNT"] == True:
                    print(jsonRes["SITE_NUMBER"], " : ",
                          jsonRes["ACCOUNT_NAME"], " : ", jsonRes["CREDIT_CHECKING"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    break
                elif billingAccountType == "CREDIT_HOLD_TRUE" and jsonRes["CREDIT_HOLD"] == True:
                    # print(jsonRes["SITE_NUMBER"], " : ",
                    #       jsonRes["ACCOUNT_NAME"], " : ", jsonRes["CREDIT_CHECKING"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    break
                elif (billingAccountType == "TOLERANCE_NULL") and (jsonRes["TOLERANCE"] == None):
                    print(jsonRes["SITE_NUMBER"], " : ",
                          jsonRes["ACCOUNT_NAME"], " : ", jsonRes["TOLERANCE"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    Messages.write_message(
                        "---------------------------------------------")
                    Messages.write_message("\t\tTOLERANCE_NULL")
                    Messages.write_message(jsonRes)
                    Messages.write_message(
                        "---------------------------------------------")
                    print(jsonRes)
                    break
                elif billingAccountType == "PREPAY_ACCOUNT_NULL" and jsonRes["PREPAY_ACCOUNT"] == None:
                    print(jsonRes["SITE_NUMBER"], " : ",
                          jsonRes["ACCOUNT_NAME"], " : ", jsonRes["PREPAY_ACCOUNT"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    break
                elif billingAccountType == "CREDIT_HOLD_NULL" and jsonRes["CREDIT_HOLD"] == None:
                    print(jsonRes["SITE_NUMBER"], " : ",
                          jsonRes["ACCOUNT_NAME"], " : ", jsonRes["CREDIT_HOLD"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    Messages.write_message(
                        "---------------------------------------------")
                    break
                elif billingAccountType == "OVERALL_CREDIT_LIMIT_NULL" and jsonRes["OVERALL_CREDIT_LIMIT"] == None:
                    print(jsonRes["SITE_NUMBER"], " : ", jsonRes["ACCOUNT_NAME"],
                          " : ", jsonRes["OVERALL_CREDIT_LIMIT"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    Messages.write_message(
                        "---------------------------------------------")
                    break
                elif billingAccountType == "TOTAL_UNAPPLIED_CASH_NULL" and jsonRes["TOTAL_UNAPPLIED_CASH"] == None:
                    print(jsonRes["SITE_NUMBER"], " : ", jsonRes["ACCOUNT_NAME"],
                          " : ", jsonRes["TOTAL_UNAPPLIED_CASH"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    Messages.write_message(
                        "---------------------------------------------")
                    break
                elif billingAccountType == "TRX_CREDIT_LIMIT_NULL" and jsonRes["TRX_CREDIT_LIMIT"] == None:
                    print(jsonRes["SITE_NUMBER"], " : ", jsonRes["ACCOUNT_NAME"],
                          " : ", jsonRes["TRX_CREDIT_LIMIT"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    Messages.write_message(
                        "---------------------------------------------")
                    break
                elif billingAccountType == "OUTSTANDING_AMT_NULL" and jsonRes["OUTSTANDING_AMT"] == None:
                    print(jsonRes["SITE_NUMBER"], " : ",
                          jsonRes["ACCOUNT_NAME"], " : ", jsonRes["OUTSTANDING_AMT"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    Messages.write_message(
                        "---------------------------------------------")
                    break
                elif billingAccountType == "CURRENT_AMT_NULL" and jsonRes["CURRENT_AMT"] == None:
                    print(jsonRes["SITE_NUMBER"], " : ",
                          jsonRes["ACCOUNT_NAME"], " : ", jsonRes["CURRENT_AMT"])
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = False
                    print(jsonRes)
                    Messages.write_message(jsonRes)
                    Messages.write_message(
                        "---------------------------------------------")
                    break
                else:
                    print(jsonRes)
                    Messages.write_message(
                        "---------------------------------------------")
                    Messages.write_message(
                        f"Not a {billingAccountType} Billing Account")
                    Messages.write_message(jsonRes)
                    Messages.write_message(
                        "---------------------------------------------")
                    billingAccountName = jsonRes["ACCOUNT_NAME"]
                    billingSiteNumber = jsonRes["SITE_NUMBER"]
                    noAccountFound = True
                cnt = cnt + 1
        if noAccountFound:
            raise Exception(
                f"Billing Account of {billingAccountType} not found...")
        else:
            return billingAccountName, billingSiteNumber

    @step("Create <ilocRegion> ILOC with following Details <table>")
    def enter_iloc_details(self, ilocRegion, table):
        sleep(10)
        fieldsData = {}
        ILOC.baseWindow = data_store.spec.get('BASE_WINDOW')
        ILOC.childWindows.append(data_store.spec.get('BASE_WINDOW'))
        dt = datetime.today()
        print(dt.month, " ", dt.day, " ", dt.year)
        year = int(dt.year) + 2
        currentDate = f"{str(dt.month)}/{str(dt.day)}/{str(year)}"

        parentPath = Path(__file__).parents[1]
        ilocObjectRepositoryFileName = str(
            parentPath) + "\\ObjectRepository\\" + os.getenv("ILOC_OBJECT_REPOSITORY_FILE")
        ilocObjectRepositorySheet = os.getenv("ILOC_OBJECT_REPOSITORY_SHEET")
        ilocObjectRepositoryJsonFileName = str(
            parentPath) + "\\ObjectRepository\\" + f"{ilocObjectRepositorySheet}.json"

        if os.path.exists(ilocObjectRepositoryFileName):
            Utils.set_json_from_object_repository(
                ilocObjectRepositorySheet, ilocObjectRepositoryFileName)

        if os.path.exists(ilocObjectRepositoryJsonFileName):
            f = open(ilocObjectRepositoryJsonFileName)
            data_store.spec[ilocObjectRepositorySheet] = json.load(f)

        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        else:
            iLOCDetails = {}
            iLOCDetails["ILOC_GRAND_TOTAL"] = 0
        print(iLOCDetails)

        print("Last Window Handle: ", ILOC.childWindows[-1])
        print("Current Window Handle: ", Drivers.driver.window_handles[-1])
        print("Window handles:", Drivers.driver.window_handles)
        ILOC.childWindows.append(Drivers.driver.window_handles[-1])

        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])

        Utils.expected_condition_for_waiting(
            "visibility_of_element_located", locator_value="//div[@value='Generate LOC']", by_value=By.XPATH)
#         Drivers.driverWait.until(
#             EC.visibility_of_element_located((By.XPATH, "//div[@value='Generate LOC']")))

        rows = table.rows
        for row in rows:
            row0 = str(row[0]).strip()
            row1 = str(row[1]).strip()
            if row0 == "LOC Name":
                fieldsData["LOC Name"] = row1
#                 colHeaderMap, columnIndex = Utils.setILOCcolumnDetails(self, ilocRegion, ws1, columnIndex, colHeaderMap, wb, "LOC Name")
            elif row0 == "Billing Account":
                fieldsData["Billing Account"] = row1
            elif row0 == "Billing Contact Name":
                fieldsData["Billing Contact Name"] = row1
            elif row0 == "Customer Name":
                fieldsData["Customer Name"] = row1
            elif row0 == "LOC Due Date":
                fieldsData["LOC Due Date"] = row1
            elif row0 == "Target":
                fieldsData["Target"] = row1
            elif row0 == "Brand":
                fieldsData["Brand"] = row1
            elif row0 == "Custom Agreement":
                fieldsData["Custom Agreement"] = row1

        Messages.write_message(fieldsData)
        fieldType = None

        if "LOC Name" in fieldsData:
            today = date.today()
            oppDateToday = today.strftime("%Y%m%d")

            fieldType = Utils.get_element("LOC Name", sheet_name=ilocObjectRepositorySheet,
                                          file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
            fieldName = fieldsData["LOC Name"] + "-" + oppDateToday
            fieldType.send_keys(fieldName)
            Messages.write_message("ILOC Name: " + fieldName)
            sleep(0.5)
            iLOCDetails["LOC Name"] = fieldName
#             Utils.setDataInXlsx(self, ws1, fieldName, "LOC Name", ilocRegion, iLOCHeaderFilePath, rowNum)

        if "Billing Contact Name" in fieldsData:
            fieldType = Utils.get_element("Billing Contact Name", sheet_name=ilocObjectRepositorySheet,
                                          file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
            fieldType.send_keys(fieldsData["Billing Contact Name"])
            Messages.write_message(
                "Billing Contact Name: " + fieldsData["Billing Contact Name"])
            sleep(0.5)
            iLOCDetails["Billing Contact Name"] = fieldsData["Billing Contact Name"]
#             Utils.setDataInXlsx(self, ws1, fieldsData["Billing Contact Name"], "Billing Contact Name", ilocRegion, iLOCHeaderFilePath, rowNum)

        if "Target" in fieldsData:
            fieldType = Utils.get_element("Target", sheet_name=ilocObjectRepositorySheet,
                                          file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
            fieldType.send_keys(fieldsData["Target"])
            Messages.write_message("Target: " + fieldsData["Target"])
            sleep(0.5)
            iLOCDetails["Target"] = fieldsData["Target"]
#             Utils.setDataInXlsx(self, ws1, fieldsData["Target"], "Target", ilocRegion, iLOCHeaderFilePath, rowNum)

        if "Brand" in fieldsData:
            fieldType = Utils.get_element("Brand", sheet_name=ilocObjectRepositorySheet,
                                          file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
            fieldType.send_keys(fieldsData["Brand"])
            Messages.write_message("Brand: " + fieldsData["Brand"])
            sleep(0.5)
            iLOCDetails["Brand"] = fieldsData["Brand"]
#             Utils.setDataInXlsx(self, ws1, fieldsData["Brand"], "Brand", ilocRegion, iLOCHeaderFilePath, rowNum)

        if "Customer Name" in fieldsData:
            custName = ""
            fieldType = Utils.get_element("Customer Name T", sheet_name=ilocObjectRepositorySheet,
                                          file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
            custValue = fieldType.get_attribute("value")
            # pdb.set_trace()
            Messages.write_message(f"Customer Name: {custValue}")

            if len(custValue) > 0:
                print("Customer Name: ", custValue)
                custName = custValue
            else:
                custName = f'{os.getenv("ILOC_CUTOMER_NAME")} {ilocRegion}'
                fieldType = Utils.get_element("Customer Name L", sheet_name=ilocObjectRepositorySheet,
                                              file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
                fieldType.click()
                sleep(0.5)
                Drivers.driver.switch_to.window(
                    Drivers.driver.window_handles[-1])
                sleep(0.5)

                Drivers.driverWait.until(EC.visibility_of_element_located(
                    (By.XPATH, "//td[input[@value='Go']]/preceding-sibling::th/input")))
                sleep(0.5)

                txtBoxSearch = Drivers.driver.find_element_by_xpath(
                    "//td[input[@value='Go']]/preceding-sibling::th/input")
                txtBoxSearch.clear()
                txtBoxSearch.send_keys(custName)
                sleep(0.5)

                Drivers.driverWait.until(EC.visibility_of_element_located(
                    (By.XPATH, "//input[@value='Go']")))
                sleep(0.5)

                buttonSearch = Drivers.driver.find_element_by_xpath(
                    "//input[@value='Go']")
                buttonSearch.click()
                sleep(1)

                Drivers.driverWait.until(EC.visibility_of_element_located(
                    (By.LINK_TEXT, custName)))
                sleep(0.5)

                linkAccountName = Drivers.driver.find_element_by_link_text(
                    custName)
                linkAccountName.click()
                sleep(0.5)

                Drivers.driver.switch_to.window(
                    Drivers.driver.window_handles[-1])
                sleep(3)
            iLOCDetails["Customer Name"] = custName
            Messages.write_message(f"CUSTOMER NAME: {iLOCDetails['Customer Name']}")

        if "Custom Agreement" in fieldsData:
            today = date.today()
            oppDateToday = today.strftime("%Y%m%d")

            fieldType = Utils.get_element("Custom Agreement", sheet_name=ilocObjectRepositorySheet,
                                          file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
            fieldType.click()
            sleep(0.5)

            fieldType = Utils.get_element("Custom Attachment", sheet_name=ilocObjectRepositorySheet,
                                          file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))

            rootPath = Path(__file__).parents[1]
            fileList = ['.xls', '.xlsx', '.doc', '.docx', '.txt', '.pdf']
            fileToUpload = f"CUSTOM_ILOC{random.choice(fileList)}"
            screenShotFileName = str(
                rootPath) + "\\Data\\CUSTOM_ILOC_DATA\\" + fileToUpload
            fieldType.send_keys(screenShotFileName)
            sleep(0.5)
            
            Messages.write_message(f"Selected File to Upload For Custom ILOC: {fileToUpload}")

            Messages.write_message(
                "Custom Agreement: " + fieldsData["Custom Agreement"])
            iLOCDetails["Custom Agreement"] = fieldsData["Custom Agreement"]

        if "LOC Due Date" in fieldsData:
            fieldType = Utils.get_element("LOC Due Date", sheet_name=ilocObjectRepositorySheet,
                                          file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
            fieldType.click()
            sleep(2)
            Drivers.driverWait.until(EC.visibility_of_element_located(
                (By.ID, "ui-datepicker-div")))
            dt = datetime.today()
            print(dt.month, " ", dt.day, " ", dt.year)
            currentDate = f"{str(dt.month)}/{str(dt.day)}/{str(year)}"
            dtXpath = f"//div[@id='ui-datepicker-div']/table/tbody//td/a[text()='{str(dt.day)}']"
            fieldType = Drivers.driver.find_element_by_xpath(
                dtXpath)
            actions = ActionChains(Drivers.driver)
            actions.move_to_element(fieldType)
            actions.click()
            actions.perform()
            sleep(2)
            iLOCDetails["LOC Due Date"] = currentDate
#             Utils.setDataInXlsx(self, ws1, currentDate, "LOC Due Date", ilocRegion, iLOCHeaderFilePath, rowNum)

        if "Billing Account" in fieldsData:

            billingAccount, billingSiteNumber = self.get_billing_account(
                ilocRegion, fieldsData["Billing Account"])

            fieldType = Utils.get_element("Billing Account", sheet_name=ilocObjectRepositorySheet,
                                          file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
            fieldType.click()
            sleep(1)
            Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
            sleep(0.5)

            Drivers.driverWait.until(EC.visibility_of_element_located(
                (By.XPATH, "//td[input[@value='Go']]/preceding-sibling::th/input")))
            sleep(0.5)

            txtBoxSearch = Drivers.driver.find_element_by_xpath(
                "//td[input[@value='Go']]/preceding-sibling::th/input")
            txtBoxSearch.clear()
            txtBoxSearch.send_keys(billingSiteNumber)
            sleep(0.5)

            Drivers.driverWait.until(EC.visibility_of_element_located(
                (By.XPATH, "//input[@value='Go']")))
            sleep(0.5)

            buttonSearch = Drivers.driver.find_element_by_xpath(
                "//input[@value='Go']")
            buttonSearch.click()
            sleep(1)

            Drivers.driverWait.until(EC.visibility_of_element_located(
                (By.PARTIAL_LINK_TEXT, billingAccount)))
            sleep(0.5)

            linkAccountName = Drivers.driver.find_element_by_partial_link_text(
                billingAccount)
            linkAccountName.click()
            sleep(0.5)

            print("Current Window Handle: ", Drivers.driver.window_handles[-1])
            print("All Window handles:", Drivers.driver.window_handles)

            iLOCDetails["Billing Account"] = billingAccount
        print("\n----------------------------------------------------------------------------------------")
        print("\t\t\t\t ILOC DETAILS")
        print("----------------------------------------------------------------------------------------")
        print(iLOCDetails)
        print("----------------------------------------------------------------------------------------\n")
        data_store.spec[ilocRegion] = iLOCDetails

    @step("Select <oppName> Opportunity from <oppType> sheet for <ilocRegion> ILOC")
    def select_opportunities_for_iloc(self, oppName, oppType, ilocRegion):
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        parentPath = Path(__file__).parents[1]
        ilocObjectRepositoryFileName = str(
            parentPath) + "\\ObjectRepository\\" + os.getenv("ILOC_OBJECT_REPOSITORY_FILE")
        ilocObjectRepositorySheet = os.getenv("ILOC_OBJECT_REPOSITORY_SHEET")
        childOppNumber = 1

        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        else:
            iLOCDetails = {}
            iLOCDetails["ILOC_GRAND_TOTAL"] = 0
        print("\n----------------------------------------------------------------------------------------")
        print("\t\t\t\t ILOC DETAILS")
        print("----------------------------------------------------------------------------------------")
        print(iLOCDetails)
        print("----------------------------------------------------------------------------------------\n")

        Messages.write_message(
            "\n----------------------------------------------------------------------------------------")
        Messages.write_message("\t\t\t\t ILOC DETAILS")
        Messages.write_message(
            "----------------------------------------------------------------------------------------")
        Messages.write_message(iLOCDetails)
        Messages.write_message(
            "----------------------------------------------------------------------------------------\n")

        days = None
        today = date.today()
        oppToFind = None
        if "TODAY" in oppName:
            days = oppName.split("_")[0]
            oppName = oppName.split("_")[1]
            oppDateToday = today.strftime("%Y%m%d")
            oppToFind = f"{oppName}{oppDateToday}"
        elif "YESTERDAY" in oppName:
            days = oppName.split("_")[0]
            oppName = oppName.split("_")[1]
            yesterday = today - timedelta(days=1)
            oppDateYesterday = yesterday.strftime("%Y%m%d")
            oppToFind = f"{oppName}{oppDateYesterday}"
        elif "EREYESTERDAY" in oppName:
            days = oppName.split("_")[0]
            oppName = oppName.split("_")[1]
            dayBeforeYesterday = today - timedelta(days=2)
            oppDateDayBeforeYesterday = dayBeforeYesterday.strftime("%Y%m%d")
            oppToFind = f"{oppName}{oppDateDayBeforeYesterday}"
        else:
            randomDate = today - timedelta(days=randint(1, 3))
            randomDate = randomDate.strftime("%Y%m%d")
            oppToFind = f"{oppName}{randomDate}"

        opportunitiesDetailsPath = Path(__file__).parents[1]
        opportunitiesDetailsFileName = str(
            opportunitiesDetailsPath) + "\\Data\\" + os.getenv("ILOC_OPPORTUNITY_DETAILS_FILE")

        query = Drivers.sf.query_all_iter(format_soql(
            "SELECT Contract__c FROM OpportunityContractJunction__c where Opportunity__r.Name = {} AND Status__c = {}", oppToFind, "Signed by Customer"))
        for data in query:
            print(data)
            result = Drivers.sf.Contract.update(str(data['Contract__c']), {
                                                'Status': 'Pre-Sign Approved'})
            print(result)

        searchSOQL = f"SELECT EXISTS(SELECT 1 FROM Opportunity WHERE Name = '{oppToFind}')"
        fetchSOQL = f"SELECT Id,Name,Total  FROM Opportunity WHERE Name = '{oppToFind}'"
        recordDetail = Common_Steps.CommonSteps.fetch_record_details(
            self, searchSOQL, fetchSOQL)
        rec = recordDetail.fetchone()
        if rec != None:
            oppId = rec['Id']
            oppName = rec['Name']
            oppTotal = rec['Total']
            iLOCDetails["ILOC_GRAND_TOTAL"] = iLOCDetails["ILOC_GRAND_TOTAL"] + oppTotal

        Messages.write_message(
            f"\n ILOC GRAND TOTAL: {iLOCDetails['ILOC_GRAND_TOTAL']} \n")
        Utils.expected_condition_for_waiting(
            "invisibility_of_element_located", locator_value="spinner", by_value=By.ID)
        searchBox = Utils.get_element("Search T", sheet_name=ilocObjectRepositorySheet,
                                      file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
        searchButton = Utils.get_element("Search B", sheet_name=ilocObjectRepositorySheet,
                                         file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
        searchBox.clear()
        searchBox.send_keys(oppToFind)
        searchButton.click()        

        Utils.expected_condition_for_waiting(
            "invisibility_of_element_located", locator_value="spinner", by_value=By.ID)

        jsonData = data_store.spec.get(ilocObjectRepositorySheet)
        opportunityTableName = jsonData['Opportunity Table']['Value']

        rowsXPATH = f"//table[@id='{opportunityTableName}']/tbody/tr"
        trValues = Drivers.driver.find_elements_by_xpath(rowsXPATH)

        if len(trValues) > 0:

            tableXPATH = f"//table[@id='{opportunityTableName}']/thead/tr/th"
    #         pdb.set_trace()
            opportunityTableColumnList = Utils.get_column_index(
                tableXPATH)

            quantityColIndex = opportunityTableColumnList["Opportunity Name"]

            opportunityNameXPATH = f"//table[@id='{opportunityTableName}']/tbody/tr[1]/td[{quantityColIndex}]//a/span"
            oppNameField = Drivers.driver.find_element_by_xpath(
                opportunityNameXPATH)

            if str(oppToFind) == oppNameField.text:
                selectOpportunityBoxXPATH = f"//table[@id='{opportunityTableName}']/tbody/tr[1]/td[1]/input"
                selectOpportunityBox = Drivers.driver.find_element_by_xpath(
                    selectOpportunityBoxXPATH)
                selectOpportunityBox.click()

            Utils.expected_condition_for_waiting(
                "invisibility_of_element_located", locator_value="spinner", by_value=By.ID)

        data_store.spec[ilocRegion] = iLOCDetails
#         searchBox.clear()
#         searchButton.click()

    @step("Verify <ilocRegion> ILOC Status as <ilocStatus> when Billing Account is <billingAccountType>")
    def verify_iloc_status_as_when_billing_account_is(self, ilocRegion, ilocStatus, billingAccountType):

        if ilocRegion == "US":
            accountName = os.getenv(
                "OPPORUNITIES_TO_CREATE_FOR_USD_ACCOUNT_ID")
        if ilocRegion == "CA":
            accountName = os.getenv(
                "OPPORUNITIES_TO_CREATE_FOR_CAD_ACCOUNT_ID")

        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        print(iLOCDetails)
        currentURL = Drivers.driver.current_url
        contractID = currentURL.split(".com/")[1]
        iLOCDetails["CONTRACT_ID"] = contractID

        xpath = "//td[span[text()='Contract ID']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Contract ID']/following-sibling::*[position()=1][name()='td']/div"
        contractNumber = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["CONTRACT_NUMBER"] = contractNumber

        xpath = "//td[span[text()='Status']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Status']/following-sibling::*[position()=1][name()='td']/div"
        statusValue = Drivers.driver.find_element_by_xpath(xpath).text
        if statusValue == str(ilocStatus):
            iLOCDetails["CONTRACT_STATUS"] = statusValue
            data_store.spec[ilocRegion] = iLOCDetails
            Messages.write_message(
                f"ILOC Status = {ilocStatus} when Billing Account having {billingAccountType}")
            print(
                f"ILOC Status = {ilocStatus} when Billing Account having {billingAccountType}")
            Screenshots.capture_screenshot()
        else:
            iLOCDetails["CONTRACT_STATUS"] = statusValue
            data_store.spec[ilocRegion] = iLOCDetails
            Screenshots.capture_screenshot()
            raise Exception(
                f"Status of ILOC should be {ilocStatus} when Billing Account having {billingAccountType}")

        # Adding Updating ILOC details
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Id", 'id', contractID, contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "ContractNumber", 'text', contractNumber, contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "ContractName", 'text', iLOCDetails["LOC Name"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Status", 'text', statusValue, contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "BillingAccount__c", 'text', iLOCDetails["Billing Account"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "BillingContact__c", 'text', iLOCDetails["Billing Contact Name"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "CustomerName", 'text', iLOCDetails["Customer Name"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Account", 'string', accountName, contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Brand__c", 'text', iLOCDetails["Brand"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Target__c", 'text', iLOCDetails["Target"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Expiration_Date__c", 'text', iLOCDetails["LOC Due Date"], contractID)

        xpath = "//td[span[text()='Sent for Signature As']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Sent for Signature As']/following-sibling::*[position()=1][name()='td']/div"
        accoundAD = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["ACCOUNT_AD"] = accoundAD
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "AccountsAD__c", 'text', accoundAD, contractID)

        xpath = "//td[span[text()='Agreements with Custom LOC']]/following-sibling::*[position()=1][name()='td']/div/img | //td[not(span)][text()='Agreements with Custom LOC']/following-sibling::*[position()=1][name()='td']/div/img"
        isCustomTemplate = Drivers.driver.find_element_by_xpath(
            xpath).get_attribute("title")
        if isCustomTemplate == 'Not Checked':
            isCustomTemplate = 'False'
        elif isCustomTemplate == 'Checked':
            isCustomTemplate = 'True'
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "isCustomTemplateUsed__c", 'text', isCustomTemplate, contractID)

        xpath = "//td[span[text()='Credit Message']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Credit Message']/following-sibling::*[position()=1][name()='td']/div"
        ccErrorMessage = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["CC_ERROR_MESSAGE"] = ccErrorMessage
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "CreditCheckErrorMessage__c", 'textarea', ccErrorMessage, contractID)

        xpath = "//td[span[text()='Customer Comments']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Customer Comments']/following-sibling::*[position()=1][name()='td']/div"
        customerComments = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["CUSTOMER_COMMENTS"] = customerComments
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Customer_Comments__c", 'text', customerComments, contractID)

        xpath = "//td[span[text()='LOC Signed Date']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='LOC Signed Date']/following-sibling::*[position()=1][name()='td']/div"
        locSignedDate = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["CUSTOMER_SIGNED_DATE"] = locSignedDate
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "CustomerSignedDate", 'text', locSignedDate, contractID)

        xpath = "//td[span[text()='ILOC Submitter']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='ILOC Submitter']/following-sibling::*[position()=1][name()='td']/div"
        ilocSubmitter = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["ILOC_SUBMITTER"] = ilocSubmitter
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "ILOCSubmitterUser__c", 'text', ilocSubmitter, contractID)

        xpath = "//td[span[text()='Grand Total']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Grand Total']/following-sibling::*[position()=1][name()='td']/div"
        grandTotal = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["GRAND_TOTAL"] = grandTotal
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Grand_Total__c", 'text', grandTotal, contractID)

        # Utils.setILOCDataInXlsx(ilocRegion, iLOCDetails)

    @step("Click on ILOC button <buttoValue>")
    def click_on_iloc_button(self, buttoValue):
        try:
            errorMsg = None
            Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
            if buttoValue == "Generate LOC":
                xPath = f"//div[contains(@value,'{buttoValue}')]"
                Drivers.driverWait.until(
                    EC.visibility_of_element_located((By.XPATH, xPath)))
                buttonToClick = Drivers.driver.find_element_by_xpath(xPath)
                buttonToClick.click()
                sleep(5)
                xPath = "//button[text()='Confirm']"
                buttonToClick = Drivers.driver.find_element_by_xpath(xPath)
                buttonToClick.click()
                Utils.expected_condition_for_waiting(
                    "invisibility_of_element_located", locator_value="spinner", by_value=By.ID)
                sleep(5)

        except NoSuchElementException:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

    @step("Get the ILOC Approvers of <ilocRegion>")
    def get_the_iloc_approvers(self, ilocRegion):
        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        print(iLOCDetails)

        if ilocRegion == "US":
            currencyCode = "USD"
        elif ilocRegion == "CA":
            currencyCode = "CAD"
        contractId = iLOCDetails["CONTRACT_ID"]

        approvers = self.get_approver(ilocRegion, contractId)
        print(approvers)
        Messages.write_message(f"ILOC Approvers: {approvers}")

        iLOCDetails["ILOC_APPROVERS"] = str(approvers)
        data_store.spec[ilocRegion] = iLOCDetails

    @step("User <approvalAction> the <productLine> Opportunity of <ilocRegion> ILOC")
    def user_the_opportunity(self, approvalAction, productLine, ilocRegion):
        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        print(iLOCDetails)

        approvers = eval(iLOCDetails["ILOC_APPROVERS"])

        for key, value in approvers.items():
            if productLine == key:
                self.logout()
                self.search_user_login(value)

            Drivers.driverWait.until(
                EC.visibility_of_element_located((By.XPATH, "//iframe[@title='PendingILOCApprovals']")))
            iframeElement = Drivers.driver.find_element_by_xpath(
                "//iframe[@title='PendingILOCApprovals']")
            Drivers.driver.switch_to.frame(iframeElement)

            Drivers.driverWait.until(
                EC.visibility_of_element_located((By.XPATH, "//label[contains(text(),'Search:')]/input[@type='search']")))
            searchBox = Drivers.driver.find_element_by_xpath(
                "//label[contains(text(),'Search:')]/input[@type='search']")
            searchBox.send_keys(iLOCDetails["CONTRACT_NUMBER"])
            sleep(2)

            ilocApprovalTable = Drivers.driver.find_element_by_xpath(
                "//table[contains(@class,'pendingTable')]/tbody")

            ilocApprovalTableXPATH = "//table[@class='list pendingTable dataTable no-footer']/thead/tr/th"
            columnsList = Utils.get_column_index(
                ilocApprovalTableXPATH)

            contractNumberXpath = f"//table[contains(@class,'pendingTable')]/tbody/tr[1]/td[{columnsList['Contract']}]"
            contractNumberValue = Drivers.driver.find_element_by_xpath(
                contractNumberXpath).text

            if iLOCDetails["CONTRACT_NUMBER"] == contractNumberValue:
                approveRejectLinkXPATH = f"//table[contains(@class,'pendingTable')]/tbody/tr[1]/td[{columnsList['Action']}]/a"
                Drivers.driverWait.until(
                    EC.visibility_of_element_located((By.XPATH, approveRejectLinkXPATH)))
                approveRejectLink = Drivers.driver.find_element_by_xpath(
                    approveRejectLinkXPATH)
                approveRejectLink.click()
                sleep(2)
                Drivers.driver.switch_to.default_content()

                if approvalAction == "Approve":
                    approveButtonXPATH = "//input[@value='Approve']"
                    Drivers.driverWait.until(
                        EC.visibility_of_element_located((By.XPATH, approveButtonXPATH)))
                    approveButton = Drivers.driver.find_element_by_xpath(
                        approveButtonXPATH)
                    approveButton.click()
                    sleep(2)
                    alert_obj = Drivers.driver.switch_to_alert()
                    alert_obj.accept()

                elif approvalAction == "Reject":
                    commentsXPATH = "//th[label[text()='Comments']]/following-sibling::*[position()=1][name()='td']/textarea"
                    Drivers.driverWait.until(
                        EC.visibility_of_element_located((By.XPATH, commentsXPATH)))
                    commentsBox = Drivers.driver.find_element_by_xpath(
                        commentsXPATH)
                    commentsBox.send_keys("Rejected")

                    rejectButtonXPATH = "//input[@value='Reject']"
                    Drivers.driverWait.until(
                        EC.visibility_of_element_located((By.XPATH, rejectButtonXPATH)))
                    rejectButton = Drivers.driver.find_element_by_xpath(
                        rejectButtonXPATH)
                    rejectButton.click()
                    sleep(2)
                    alert_obj = Drivers.driver.switch_to_alert()
                    alert_obj.accept()

            iLOCDetails[f"{productLine}_STATUS"] = approvalAction
            data_store.spec[ilocRegion] = iLOCDetails

    @step("Verify <ilocRegion> ILOC Status as <ilocStatus>")
    def verify_iloc_status_as(self, ilocRegion, ilocStatus):
        sleep(5)
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])

        if ilocRegion == "US":
            accountName = os.getenv(
                "OPPORUNITIES_TO_CREATE_FOR_USD_ACCOUNT_ID")
        if ilocRegion == "CA":
            accountName = os.getenv(
                "OPPORUNITIES_TO_CREATE_FOR_CAD_ACCOUNT_ID")

        # Drivers.driver.refresh()
        iLOCDetails = None
        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        print(iLOCDetails)
        currentURL = Drivers.driver.current_url
        contractID = currentURL.split(".com/")[1]
        Messages.write_message(f"Record URL : {currentURL}")
        Messages.write_message(f"Contract ID : {contractID}")
        iLOCDetails["CONTRACT_ID"] = contractID
        iloctotal = iLOCDetails["ILOC_GRAND_TOTAL"]

        queryData = (Drivers.sf.query(format_soql(
            "SELECT Id, ContractNumber, Status, Grand_Total__c FROM Contract WHERE Id = {}", contractID)))['records']

        contractNumber = queryData[0]['ContractNumber']
        # xpath = "//td[span[text()='Contract ID']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Contract ID']/following-sibling::*[position()=1][name()='td']/div"
        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.XPATH, xpath)))
        # contractNumber = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["CONTRACT_NUMBER"] = contractNumber

        if "Custom Agreement" in iLOCDetails:
            if str(iLOCDetails["Custom Agreement"]) == 'Yes':
                xpath = "//td[span[text()='Agreements with Custom LOC']]/following-sibling::*[position()=1][name()='td']/div/img | //td[not(span)][text()='Agreements with Custom LOC']/following-sibling::*[position()=1][name()='td']/div/img"
                isCustomAgreement = Drivers.driver.find_element_by_xpath(
                    xpath).get_attribute("title")
                assert str(
                    isCustomAgreement) == "Checked", f"CUSTOM ILOC CHECKBOX SHOULD BE CHECKED"

                xpath = "//div[contains(@id,'RelatedNoteList_body')]/table/tbody/tr[1]/th"
                tableHeader = Drivers.driver.find_elements_by_xpath(xpath)
                labelsMap = {}
                columnIndex = 1
                sleep(5)
                for theader in tableHeader:
                    labelXpath = "//div[contains(@id,'RelatedNoteList_body')]/table/tbody/tr[1]/th[" + str(
                        columnIndex) + "]"
                    columnName = Drivers.driver.find_element_by_xpath(
                        labelXpath).text
                    # pdb.set_trace()
                    print("Attachment Column: ", columnName)
                    if columnName != None:
                        # pdb.set_trace()
                        labelsMap[columnName] = columnIndex
                        Messages.write_message(
                            columnName + " : " + str(columnIndex))
                    columnIndex = columnIndex + 1
                    print(labelsMap)

                attachmentTitle = labelsMap["Title"]
                labelXpath = f"//div[contains(@id,'RelatedNoteList_body')]/table/tbody/tr[2]/td[{int(attachmentTitle)-1}]"
                fileName = Drivers.driver.find_element_by_xpath(
                    labelXpath).text
                fileList = [f"'Contract-{contractNumber}.xlsx'", f"'Contract-{contractNumber}.xls'",
                            f"'Contract-{contractNumber}.doc'", f"'Contract-{contractNumber}.docx'", f"'Contract-{contractNumber}.txt'"]
                if any(word in fileName for word in fileList):
                    Messages.write_message(
                        f"Custom ILOC FILE {fileName} ATTACHED")
                    print("Custom ILOC FILE ATTACHED: ", fileName)

        xpath = "//td[span[text()='Status']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Status']/following-sibling::*[position()=1][name()='td']/div"
        statusValue = Drivers.driver.find_element_by_xpath(xpath).text
        
        Messages.write_message(f"ILOC STATUS ON PAGE: {statusValue}")
        # statusValue = queryData[0]['Status']

        xpath = "//img[@name='Total' and @title='Show Section - Total']"
        try:
            totalIcon = Drivers.driver.find_element_by_xpath(xpath)
            totalIcon.click()
            sleep(0.5)
        except NoSuchElementException:
            print("Total Section already expanded")

        xpath = "//td[span[text()='Grand Total']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Grand Total']/following-sibling::*[position()=1][name()='td']/div"
        grandTotal = Drivers.driver.find_element_by_xpath(xpath).text

        if grandTotal.count(",") > 0:
            grandTotal = grandTotal.split(" ")[1].replace(",", "")
        else:
            grandTotal = grandTotal.split(" ")[1]

        iloctotal = round(iloctotal, 2)
        if str(iloctotal).find(".") > 0 and len(str(iloctotal).split(".")[1]) == 1:
            iloctotal = str(iloctotal).split(
                ".")[0] + "." + str(iloctotal).split(".")[1] + "0"
            ILOC.takeScreenShot(self)
        assert str(grandTotal) == str(
            iloctotal), f"ILOC TOTAL SHOULD BE {grandTotal} AND NOT {iloctotal}"

        if str(statusValue) == str(ilocStatus):
            iLOCDetails["CONTRACT_STATUS"] = statusValue
            data_store.spec[ilocRegion] = iLOCDetails
            Utils.setILOCDataInXlsx(ilocRegion, iLOCDetails)
            Messages.write_message(
                f"Verified ILOC Status = {ilocStatus}")
            print(
                f"Verified ILOC Status = {ilocStatus}")
            ILOC.takeScreenShot(self)
        else:
            iLOCDetails["CONTRACT_STATUS"] = statusValue
            data_store.spec[ilocRegion] = iLOCDetails
            Utils.setILOCDataInXlsx(ilocRegion, iLOCDetails)
            ILOC.takeScreenShot(self)
            raise Exception(
                f"Unable to verify Status of ILOC as {ilocStatus}")
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Id", 'id', contractID, contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "ContractNumber", 'text', contractNumber, contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "ContractName", 'text', iLOCDetails["LOC Name"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Status", 'text', statusValue, contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "BillingAccount__c", 'text', iLOCDetails["Billing Account"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "BillingContact__c", 'text', iLOCDetails["Billing Contact Name"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "CustomerName", 'text', iLOCDetails["Customer Name"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "AccountId", 'string', accountName, contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Brand__c", 'text', iLOCDetails["Brand"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Target__c", 'text', iLOCDetails["Target"], contractID)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Expiration_Date__c", 'text', iLOCDetails["LOC Due Date"], contractID)
        if "Custom Agreement" in iLOCDetails:
            if str(iLOCDetails["Custom Agreement"]) == 'Yes':
                Common_Steps.CommonSteps.write_data_to_table_column(
                    self, "Contract", "Custom_Agreement", 'text', iLOCDetails["Custom Agreement"], contractID)

        xpath = "//td[span[text()='Sent for Signature As']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Sent for Signature As']/following-sibling::*[position()=1][name()='td']/div"
        accoundAD = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["ACCOUNT_AD"] = accoundAD
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "AccountsAD__c", 'text', accoundAD, contractID)

        xpath = "//td[span[text()='Agreements with Custom LOC']]/following-sibling::*[position()=1][name()='td']/div/img | //td[not(span)][text()='Agreements with Custom LOC']/following-sibling::*[position()=1][name()='td']/div/img"
        isCustomTemplate = Drivers.driver.find_element_by_xpath(
            xpath).get_attribute("title")
        if isCustomTemplate == 'Not Checked':
            isCustomTemplate = 'False'
        elif isCustomTemplate == 'Checked':
            isCustomTemplate = 'True'
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "isCustomTemplateUsed__c", 'text', isCustomTemplate, contractID)

        xpath = "//td[span[text()='Credit Message']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Credit Message']/following-sibling::*[position()=1][name()='td']/div"
        ccErrorMessage = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["CC_ERROR_MESSAGE"] = ccErrorMessage
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "CreditCheckErrorMessage__c", 'textarea', ccErrorMessage, contractID)

        xpath = "//td[span[text()='Customer Comments']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Customer Comments']/following-sibling::*[position()=1][name()='td']/div"
        customerComments = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["CUSTOMER_COMMENTS"] = customerComments
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Customer_Comments__c", 'text', customerComments, contractID)

        xpath = "//td[span[text()='LOC Signed Date']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='LOC Signed Date']/following-sibling::*[position()=1][name()='td']/div"
        locSignedDate = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["CUSTOMER_SIGNED_DATE"] = locSignedDate
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "CustomerSignedDate", 'text', locSignedDate, contractID)

        xpath = "//td[span[text()='ILOC Submitter']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='ILOC Submitter']/following-sibling::*[position()=1][name()='td']/div"
        ilocSubmitter = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["ILOC_SUBMITTER"] = ilocSubmitter
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "ILOCSubmitterUser__c", 'text', ilocSubmitter, contractID)

        xpath = "//td[span[text()='Grand Total']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Grand Total']/following-sibling::*[position()=1][name()='td']/div"
        grandTotal = Drivers.driver.find_element_by_xpath(xpath).text
        iLOCDetails["GRAND_TOTAL"] = grandTotal
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Contract", "Grand_Total__c", 'text', grandTotal, contractID)

    @step("Search <ilocName> ILOC of <ilocRegion> region")
    def search_iloc_of_region(self, ilocName, ilocRegion):

        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        print(iLOCDetails)

        contractToSearch = iLOCDetails["CONTRACT_NUMBER"]

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'phSearchInput')))

        contractID = iLOCDetails["CONTRACT_ID"]

        currentURL = Drivers.driver.current_url
        contractURL = f"{currentURL.split('.com/')[0]}.com/{contractID}"
        Drivers.driver.get(contractURL)

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'topButtonRow')))

    @step("Edit <oppName> Opportunity from <oppType> sheet for <ilocRegion> ILOC")
    def edit_opportunity_from_sheet_for_iloc(self, oppName, oppType, ilocRegion):
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        print(iLOCDetails)
        days = None
        today = date.today()
        if "TODAY" in oppName:
            days = oppName.split("_")[0]
            oppName = oppName.split("_")[1]
            oppDateToday = today.strftime("%Y%m%d")
            oppName = f"{oppName}{oppDateToday}"
        elif "YESTERDAY" in oppName:
            days = oppName.split("_")[0]
            oppName = oppName.split("_")[1]
            yesterday = today - timedelta(days=1)
            oppDateYesterday = yesterday.strftime("%Y%m%d")
            oppName = f"{oppName}{oppDateYesterday}"
        elif "EREYESTERDAY" in oppName:
            days = oppName.split("_")[0]
            oppName = oppName.split("_")[1]
            dayBeforeYesterday = today - timedelta(days=2)
            oppDateDayBeforeYesterday = dayBeforeYesterday.strftime("%Y%m%d")
            oppName = f"{oppName}{oppDateDayBeforeYesterday}"
        else:
            randomDate = today - timedelta(days=randint(1, 3))
            randomDate = randomDate.strftime("%Y%m%d")
            oppName = f"{oppName}{randomDate}"

        # oppName = iLOCDetails[f"{oppType}_{days}"]

        searchSOQL = f"SELECT EXISTS(SELECT 1 FROM Opportunity WHERE Name = '{oppName}')"
        fetchSOQL = f"SELECT Id,Name,Total  FROM Opportunity WHERE Name = '{oppName}'"
        recordDetail = Common_Steps.CommonSteps.fetch_record_details(
            self, searchSOQL, fetchSOQL)
        rec = recordDetail.fetchone()
        oppId = None
        oldAmount = 0
        if rec != None:
            oppId = rec['Id']
            oldAmount = rec['Total']
            # iLOCDetails["ILOC_GRAND_TOTAL"] = iLOCDetails["ILOC_GRAND_TOTAL"] - \
            #     oldAmount + float(newAmount)

        # soql = f"SELECT Id, Amount, Name from Opportunity where Name='{oppName}' and CurrencyIsoCode = '{ilocRegion}D'"
        # result = Drivers.sf.query_all(query=soql)
        # recDetails = result['records']
        # oldAmount = recDetails[0]['Amount']

        soql = f"SELECT Id, Quantity,Sales_Price__c,Product2Name__c FROM OpportunityLineItem WHERE OpportunityId = '{oppId}'"
        result = Drivers.sf.query_all(query=soql)
        data = result['records']
        oli = {}
        productName = None
        for record in data:
            oli[record['Id']] = {'QTY': record['Quantity']}
            productName = record['Product2Name__c']
            print(record)
        print(oli)
        lineItem = random.choice(list(oli))
        print(lineItem, " ", oli[lineItem]['QTY'])
        Messages.write_message(
            f"QTY of Line Item {lineItem} updated from {oli[lineItem]['QTY']} to {int(oli[lineItem]['QTY']) + 1}")
        Drivers.sf.OpportunityLineItem.update(str(lineItem), {
            'Quantity': int(oli[lineItem]['QTY']) + 1})

        if 'FSI' in oppType and productName != 'Remnant':
            Drivers.sf.Opportunity.update(str(oppId), {'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":6000,"Description":"OTHER CHARGE","Amount":60, "CirculationQty": 100}]',
                                                       'ILOCTotalProgramFee__c': 9000})

        soql = f"SELECT Id, Quantity,Sales_Price__c, SubTotal__c FROM OpportunityLineItem WHERE Id = '{lineItem}'"
        result = Drivers.sf.query_all(query=soql)
        data = result['records']
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "OLI", "Quantity", 'double', data[0]['Quantity'], lineItem)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "OLI", "Sales_price__c", 'double', data[0]['Sales_Price__c'], lineItem)
        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "OLI", "TotalPrice", 'double', data[0]['SubTotal__c'], lineItem)

        soql = f"SELECT Id, Commissionable_Revenue__c, Name from Opportunity where Id = '{oppId}' and CurrencyIsoCode = '{ilocRegion}D'"
        result = Drivers.sf.query_all(query=soql)
        recDetails = result['records']
        newAmount = None
        if 'FSI' in oppType and productName != 'Remnant':
            newAmount = '9000'
        else:
            newAmount = recDetails[0]['Commissionable_Revenue__c']

        iLOCDetails["ILOC_GRAND_TOTAL"] = iLOCDetails["ILOC_GRAND_TOTAL"] + \
            (float(newAmount) - oldAmount)

        Common_Steps.CommonSteps.write_data_to_table_column(
            self, "Opportunity", "Total", 'currency', newAmount, recDetails[0]['Id'])

        currentURL = Drivers.driver.current_url
        opportunityURL = f"{currentURL.split('.com/')[0]}.com/{oppId}"
        Drivers.driver.get(opportunityURL)

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'topButtonRow')))
        data_store.spec[ilocRegion] = iLOCDetails

        iframeElement = Drivers.driver.find_element_by_xpath(
            "//iframe[@title='inlineRevenueCalculation']")
        Drivers.driver.switch_to.frame(iframeElement)

        xpath = "//button[text()='Manage Products']"
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, xpath)))

        buttonManageProducts = Drivers.driver.find_element_by_xpath(xpath)
        buttonManageProducts.click()
        sleep(5)
        Drivers.driver.switch_to.default_content()

        xpath = "//input[@value='Next']"
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, xpath)))
        buttonNext = Drivers.driver.find_element_by_xpath(xpath)
        buttonNext.click()

        xpath = "//input[@value='Save']"
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, xpath)))
        buttonSave = Drivers.driver.find_element_by_xpath(xpath)
        buttonSave.click()

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'topButtonRow')))

        contractID = iLOCDetails["CONTRACT_ID"]

        currentURL = Drivers.driver.current_url
        contractURL = f"{currentURL.split('.com/')[0]}.com/{contractID}"
        Drivers.driver.get(contractURL)

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'topButtonRow')))

        data_store.spec[ilocRegion] = iLOCDetails

    @step("Deselect <oppName> Opportunity from <oppType> sheet for <ilocRegion> ILOC")
    def select_opportunities_for_iloc(self, oppName, oppType, ilocRegion):
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        parentPath = Path(__file__).parents[1]
        ilocObjectRepositoryFileName = str(
            parentPath) + "\\ObjectRepository\\" + os.getenv("ILOC_OBJECT_REPOSITORY_FILE")
        ilocObjectRepositorySheet = os.getenv("ILOC_OBJECT_REPOSITORY_SHEET")
        childOppNumber = 1

        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)

        days = None
        today = date.today()
        if "TODAY" in oppName:
            days = oppName.split("_")[0]
            oppName = oppName.split("_")[1]
            oppDateToday = today.strftime("%Y%m%d")
            oppToSearch = f"{oppName}{oppDateToday}"
        elif "YESTERDAY" in oppName:
            days = oppName.split("_")[0]
            oppName = oppName.split("_")[1]
            yesterday = today - timedelta(days=1)
            oppDateYesterday = yesterday.strftime("%Y%m%d")
            oppToSearch = f"{oppName}{oppDateYesterday}"
        elif "EREYESTERDAY" in oppName:
            days = oppName.split("_")[0]
            oppName = oppName.split("_")[1]
            dayBeforeYesterday = today - timedelta(days=2)
            oppDateDayBeforeYesterday = dayBeforeYesterday.strftime("%Y%m%d")
            oppToSearch = f"{oppName}{oppDateDayBeforeYesterday}"
        else:
            randomDate = today - timedelta(days=randint(1, 3))
            randomDate = randomDate.strftime("%Y%m%d")
            oppToSearch = f"{oppName}{randomDate}"

        # soql = f"SELECT Id, Amount, Name from Opportunity where Name='{oppToSearch}' and CurrencyIsoCode = '{ilocRegion}D'"
        # result = Drivers.sf.query_all(query=soql)
        # recDetails = result['records']
        # newAmount = recDetails[0]['Amount']

        searchSOQL = f"SELECT EXISTS(SELECT 1 FROM Opportunity WHERE Name = '{oppToSearch}')"
        fetchSOQL = f"SELECT Id,Name,Total  FROM Opportunity WHERE Name = '{oppToSearch}'"
        recordDetail = Common_Steps.CommonSteps.fetch_record_details(
            self, searchSOQL, fetchSOQL)
        rec = recordDetail.fetchone()
        if rec != None:
            oppId = rec['Id']
            oppName = rec['Name']
            oppTotal = rec['Total']
            iLOCDetails["ILOC_GRAND_TOTAL"] = iLOCDetails["ILOC_GRAND_TOTAL"] - oppTotal
        else:
            raise Exception(f"{oppToSearch} not found in system")

        Utils.expected_condition_for_waiting(
            "invisibility_of_element_located", locator_value="spinner", by_value=By.ID)
        searchBox = Utils.get_element("Search T", sheet_name=ilocObjectRepositorySheet,
                                      file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
        searchButton = Utils.get_element("Search B", sheet_name=ilocObjectRepositorySheet,
                                         file_path=ilocObjectRepositoryFileName, json_data=data_store.spec.get(ilocObjectRepositorySheet))
        searchBox.clear()
        searchBox.send_keys(oppToSearch)
        searchButton.click()

        Utils.expected_condition_for_waiting(
            "invisibility_of_element_located", locator_value="spinner", by_value=By.ID)

        jsonData = data_store.spec.get(ilocObjectRepositorySheet)
        opportunityTableName = jsonData['Opportunity Table']['Value']

        rowsXPATH = f"//table[@id='{opportunityTableName}']/tbody/tr"
        trValues = Drivers.driver.find_elements_by_xpath(rowsXPATH)

        if len(trValues) > 0:

            tableXPATH = f"//table[@id='{opportunityTableName}']/thead/tr/th"
    #         pdb.set_trace()
            opportunityTableColumnList = Utils.get_column_index(
                tableXPATH)

            quantityColIndex = opportunityTableColumnList["Opportunity Name"]

            opportunityNameXPATH = f"//table[@id='{opportunityTableName}']/tbody/tr[1]/td[{quantityColIndex}]//a/span"
            oppNameField = Drivers.driver.find_element_by_xpath(
                opportunityNameXPATH)

            if str(oppToSearch) == oppNameField.text:
                selectOpportunityBoxXPATH = f"//table[@id='{opportunityTableName}']/tbody/tr[1]/td[1]/input"
                selectOpportunityBox = Drivers.driver.find_element_by_xpath(
                    selectOpportunityBoxXPATH)
                selectOpportunityBox.click()

            Utils.expected_condition_for_waiting(
                "invisibility_of_element_located", locator_value="spinner", by_value=By.ID)

        data_store.spec[ilocRegion] = iLOCDetails

    @step("Verify <ilocRegion> ILOC Approver Details")
    def verify_iloc_approver_details(self, ilocRegion):
        # try:
        sleep(10)
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])

        if "_" in ilocRegion:
            tempRegion = str(ilocRegion).split("_")[0]
            if tempRegion in data_store.spec:
                iLOCDetails = data_store.spec.get(tempRegion)
        else:
            if ilocRegion in data_store.spec:
                iLOCDetails = data_store.spec.get(ilocRegion)

        # Drivers.driverWait.until(EC.alert_is_present())
        # alert_obj = Drivers.driver.switch_to_alert()
        # alert_obj.accept()

        customerName = iLOCDetails['Customer Name']

        searchedData = (Drivers.sf.query(format_soql(
            "SELECT Id, Name,Email FROM Contact WHERE Name = {}", customerName)))['records'][0]
        print(searchedData['Id'])

        # searchedData = Drivers.sf.quick_search(customerName)
        # print(searchedData['Id'])
        contact = Drivers.sf.Contact.get(
            searchedData['Id'])
        customerEmail = contact['Email']
        print(f"Customer Email  {contact['Email']}")
        Messages.write_message(f"Customer Email  {contact['Email']}")

        print(f"Customer Name  {customerName}")
        Messages.write_message(f"Customer Name  {customerName}")

        iLOCDetails['CUSTOMER_EMAIL'] = contact['Email']

        salesRepName = None
        if ilocRegion == 'US':
            salesRepName = (Drivers.sf.query(format_soql(
                "SELECT Id, Name,Email FROM User WHERE Name = {} AND IsActive = True", os.getenv('US_SALES_REPRESENTATIVE'))))['records'][0]
        elif ilocRegion == 'CA':
            salesRepName = (Drivers.sf.query(format_soql(
                "SELECT Id, Name,Email FROM User WHERE Name = {} AND IsActive = True", os.getenv('CA_SALES_REPRESENTATIVE'))))['records'][0]
        else:
            salesRepName = (Drivers.sf.query(format_soql(
                "SELECT Id, Name,Email FROM User WHERE Name = {} AND IsActive = True", os.getenv(ilocRegion))))['records'][0]
        print(salesRepName['Id'])
        userData = Drivers.sf.User.get(salesRepName['Id'])
        salesRepEmail = userData['Email']
        salesRepName = userData['Name']
        iLOCDetails['SALESREP_EMAIL'] = salesRepEmail
        iLOCDetails['SALESREP_NAME'] = salesRepName
        print(f"SalesRep Name  {salesRepName}")
        print(f"SalesRep Email  {salesRepEmail}")
        Messages.write_message(f"SalesRep Name  {salesRepName}")
        Messages.write_message(f"SalesRep Email  {salesRepEmail}")

        data_store.spec[ilocRegion] = iLOCDetails

        # Messages.write_message("Searching for Alert...")
        # driverWait = WebDriverWait(Drivers.driver, 30)
        # alert = driverWait.until(EC.alert_is_present())
        # # alert_obj = Drivers.driver.switch_to_alert()
        # alert.accept()
        Drivers.driver.set_window_size(
            1900, 1200, Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1]))
        Drivers.driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")
        ILOC.takeScreenShot(self)
        # driverWait = WebDriverWait(Drivers.driver, 360)
        nextButtonXPATH = "//div[@class='slds-m-top_small']/button[contains(text(),'Next')]"

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, nextButtonXPATH)))
        nextButton = Drivers.driver.find_element_by_xpath(nextButtonXPATH)
        nextButton.click()
        # except UnexpectedAlertPresentException:
        #     Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        #     print("Unexpected Alert Present For Send On Behalf Of Signature Block")
        #     Messages.write_message("Unexpected Alert Present")
        #     alert_obj = Drivers.driver.switch_to_alert()
        #     alert_obj.accept()
        #     nextButtonXPATH = "//td[@id='bodyCell']//div[@class='slds-m-top_small']/button[contains(text(),'Next')]"
        #     Drivers.driverWait.until(
        #         EC.visibility_of_element_located((By.XPATH, nextButtonXPATH)))
        #     nextButton = Drivers.driver.find_element_by_xpath(nextButtonXPATH)
        #     nextButton.click()
        # except TimeoutException as T:
        #     # Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        #     Screenshots.capture_screenshot()
        #     print("Element Not Found...")
        #     exc_type, exc_obj, exc_tb = sys.exc_info()
        #     fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        #     print(exc_type, fname, exc_tb.tb_lineno)
        #     # Messages.write_message("Alert Not Found...")
        #     # nextButtonXPATH = "//td[@id='bodyCell']//div[@class='slds-m-top_small']/button[contains(text(),'Next')]"
        #     # Drivers.driverWait.until(
        #     #     EC.visibility_of_element_located((By.XPATH, nextButtonXPATH)))
        #     # nextButton = Drivers.driver.find_element_by_xpath(nextButtonXPATH)
        #     # nextButton.click()

    @step("Add signature block on <ilocRegion> ILOC Template and send to customer")
    def add_signature_block_on_iloc_template(self, ilocRegion):
        sleep(10)
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)

        customerEmail = iLOCDetails['CUSTOMER_EMAIL']
        salesRepEmail = iLOCDetails['SALESREP_EMAIL']

        signatureBlockXPATH = "//div[@role='button']//div[@class='item-label' and contains(text(),'Signature Block')]"
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, signatureBlockXPATH)))

        ILOC.takeScreenShot(self)
        pages = f"//ul/li/img[@class='page-image']"
        totalPages = Drivers.driver.find_elements_by_xpath(pages)
        
        
        firstPageXPATH = None
        secondPageXPATH = None

        if len(totalPages) == 1:
            # firstPageXPATH = f"//ul/li/img[@class='page-image'][1]"
            firstPageXPATH = totalPages[0]
            source_element = Drivers.driver.find_element_by_xpath(
                signatureBlockXPATH)
            # dest_element = Drivers.driver.find_element_by_xpath(firstPageXPATH)
            dest_element = totalPages[0]

            navPaneXPATH = f"//ul[@class='nav']//a/p[contains(text(),'{customerEmail}')]"
            nav_element = Drivers.driver.find_element_by_xpath(navPaneXPATH)
            nav_element.click()
            sleep(2)

            customerXPATH = f"//ul[@class='role-options dropdown-menu']//li[contains(@data-role-value,'SIGNER_{customerEmail}')]"
            customer_element = Drivers.driver.find_element_by_xpath(customerXPATH)
            customer_element.click()
            sleep(2)

            Drivers.driver.execute_script(
                "window.scrollTo(0, document.body.scrollHeight);")
            ILOC.takeScreenShot(self)

            ActionChains(Drivers.driver).drag_and_drop(
                source_element, dest_element).perform()
            # ActionChains(Drivers.driver).drag_and_drop_by_offset(source_element,x_position_customer_block,y_position_customer_block).perform()
            sleep(2)

            navPaneXPATH = f"//ul[@class='nav']//a/p[contains(text(),'{customerEmail}')]"
            nav_element = Drivers.driver.find_element_by_xpath(navPaneXPATH)
            nav_element.click()
            sleep(2)

            salesRepXPATH = f"//ul[@class='role-options dropdown-menu']//li[contains(@data-role-value,'SENDER_{salesRepEmail}')]"
            salesRep_element = Drivers.driver.find_element_by_xpath(salesRepXPATH)
            salesRep_element.click()
            sleep(2)
            Drivers.driver.find_element_by_tag_name('body').send_keys(Keys.END)
            # dest_element = Drivers.driver.find_element_by_xpath(firstPageXPATH)
            dest_element = totalPages[0]

            ILOC.takeScreenShot(self)

            ActionChains(Drivers.driver).drag_and_drop(
                source_element, dest_element).perform()
            # ActionChains(Drivers.driver).drag_and_drop_by_offset(source_element,x_position_salesrep_block,y_position_salesrep_block).perform()
            sleep(2)
            
        if len(totalPages) > 1:
            # secondPageXPATH = f"//ul/li/img[contains(@alt,'Page {len(totalPages)} of {len(totalPages)}.') and @class='page-image']"
            firstPageXPATH = totalPages[0]
            secondPageXPATH = totalPages[1]

            source_element = Drivers.driver.find_element_by_xpath(
                signatureBlockXPATH)
            dest_element = secondPageXPATH

            navPaneXPATH = f"//ul[@class='nav']//a/p[contains(text(),'{customerEmail}')]"
            nav_element = Drivers.driver.find_element_by_xpath(navPaneXPATH)
            nav_element.click()
            sleep(2)

            customerXPATH = f"//ul[@class='role-options dropdown-menu']//li[contains(@data-role-value,'SIGNER_{customerEmail}')]"
            customer_element = Drivers.driver.find_element_by_xpath(customerXPATH)
            customer_element.click()
            sleep(2)

            Drivers.driver.execute_script(
                "window.scrollTo(0, document.body.scrollHeight);")
            ILOC.takeScreenShot(self)

            ActionChains(Drivers.driver).drag_and_drop(
                source_element, dest_element).perform()
            sleep(2)

            navPaneXPATH = f"//ul[@class='nav']//a/p[contains(text(),'{customerEmail}')]"
            nav_element = Drivers.driver.find_element_by_xpath(navPaneXPATH)
            nav_element.click()
            sleep(2)

            salesRepXPATH = f"//ul[@class='role-options dropdown-menu']//li[contains(@data-role-value,'SENDER_{salesRepEmail}')]"
            salesRep_element = Drivers.driver.find_element_by_xpath(salesRepXPATH)
            salesRep_element.click()
            sleep(2)
            Drivers.driver.find_element_by_tag_name('body').send_keys(Keys.END)
            ILOC.takeScreenShot(self)
            dest_element = firstPageXPATH
            ActionChains(Drivers.driver).drag_and_drop(
                source_element, dest_element).perform()
            # ActionChains(Drivers.driver).drag_and_drop_by_offset(source_element,x_position_salesrep_block,y_position_salesrep_block).perform()
            sleep(2)

        sendButtonXPATH = "//button[contains(text(),'Send')]"
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, sendButtonXPATH)))
        sendButton = Drivers.driver.find_element_by_xpath(sendButtonXPATH)
        sendButton.click()

        okButtonXPATH = "//div[@id='salesforce-chrome']/a[contains(text(),'OK')]"
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, okButtonXPATH)))
        okButton = Drivers.driver.find_element_by_xpath(okButtonXPATH)
        okButton.click()

        sleep(5)
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        adobeImageXPATH = "//img[contains(@src,'echosign_dev1__EchoSignTabLogo')]"
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, adobeImageXPATH)))

        formattedDateTime = self.getDateTime()
        createdByUserID = os.getenv("ADOBE_USER_ID")

        batchATABStatus = self.verifyATAB(
            createdByUserID, formattedDateTime)
        print(batchATABStatus)

    @step("<senderCustomer> <emailAction> the <ilocRegion> ILOC")
    def signs_the_iloc(self, senderCustomer, emailAction, ilocRegion):
        sleep(10)
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)

        username = None
        password = None
        if senderCustomer == "Customer":
            username = os.getenv("CUSTOMER_USER")
            password = os.getenv("CUSTOMER_PASSWORD")
            imap = imaplib.IMAP4_SSL(
                os.getenv("CUSTOMER_SERVER"), int(os.getenv("CUSTOMER_PORT")))
        elif senderCustomer == "Sender":
            username = os.getenv("SENDER_USER")
            password = os.getenv("SENDER_PASSWORD")
            imap = imaplib.IMAP4_SSL(
                os.getenv("SENDER_SERVER"), int(os.getenv("SENDER_PORT")))

        imap.login(username, password)
        status, messages = imap.select("INBOX")
        # number of top emails to fetch
        if senderCustomer == "Customer":
            N = 1
        if senderCustomer == "Sender":
            N = 2
        # total number of emails
        messages = int(messages[0])

        url = None
        foundMessage = False
        for i in range(messages, messages-N, -1):
            # fetch the email message by ID
            res, msg = imap.fetch(str(i), "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    # parse a bytes email into a message object
                    msg = email.message_from_bytes(response[1])
                    # decode the email subject
                    subject = decode_header(msg["Subject"])[0][0]
                    if isinstance(subject, bytes):
                        # if it's a bytes, decode to str
                        subject = subject.decode()
                    # email sender
                    from_ = msg.get("From")
                    print("Subject:", subject)
                    print("From:", from_)
                    if subject.count('Your signature is required') > 0 or subject.count('Signature requested on') > 0:
                        # if the email message is multipart
                        if msg.is_multipart():
                            # iterate over email parts
                            for part in msg.walk():
                                # extract content type of email
                                content_type = part.get_content_type()
                                content_disposition = str(
                                    part.get("Content-Disposition"))
                                try:
                                    # get the email body
                                    body = part.get_payload(
                                        decode=True).decode()
                                    urlList = re.findall(
                                        "href=[\"\'](.*?)[\"\']", body)

                                    resendList = [
                                        item for item in urlList if "resend" in item]
                                    print(*resendList, sep="\n")
                                    urlsList = [
                                        item for item in urlList if "resend" not in item]
                                    print(*urlsList, sep="\n")
                                    url = random.choice(urlsList)
                                except:
                                    pass
                                if content_type == "text/plain" and "attachment" not in content_disposition:
                                    # print text/plain emails and skip attachments
                                    print(body)
                                foundMessage = True
                    else:
                        break

        if emailAction == "Signs":
            Drivers.Initialize_Window_For_Adobe()
            Drivers.driver4Adobe.get(url)

            startbuttonXPATH = "//button[div[text()='Start']]"
            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.XPATH, startbuttonXPATH)))
            startButton = Drivers.driver4Adobe.find_element_by_xpath(
                startbuttonXPATH)
            startButton.click()

            signatureXPATH = "//div[input[@name='echosign_signature']]"
            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.XPATH, signatureXPATH)))
            signature = Drivers.driver4Adobe.find_element_by_xpath(
                signatureXPATH)
            signature.click()

            optionTypeXPATH = "//ul[@role='tablist']/li[@class='option type selected']"
            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.XPATH, optionTypeXPATH)))
            optionType = Drivers.driver4Adobe.find_element_by_xpath(
                optionTypeXPATH)
            optionType.click()

            signatureBoxXPATH = "//input[@aria-label='Type your signature here' or @placeholder='Type your signature here']"
            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.XPATH, signatureBoxXPATH)))
            signatureBox = Drivers.driver4Adobe.find_element_by_xpath(
                signatureBoxXPATH)
            signatureBox.click()
            signatureBox.send_keys(senderCustomer)

            applySignatureButtonXpath = "//button[@class='btn btn-primary apply' and text()='Apply']"
            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.XPATH, applySignatureButtonXpath)))
            applySignatureButton = Drivers.driver4Adobe.find_element_by_xpath(
                applySignatureButtonXpath)
            applySignatureButton.click()
            formattedDateTime = self.getDateTime()
            clickToSignButtonXpath = "//button[text()='Click to Sign']"
            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.XPATH, clickToSignButtonXpath)))
            clickToSignButton = Drivers.driver4Adobe.find_element_by_xpath(
                clickToSignButtonXpath)
            clickToSignButton.click()
            Drivers.driver4Adobe.close()

            createdByUserID = os.getenv("ADOBE_USER_ID")
            batchATABStatus = self.verifyATAB(
                createdByUserID, formattedDateTime)
            print(batchATABStatus)
            sleep(15)

            Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
            Drivers.driver.refresh()

        elif emailAction == "Declines":
            Drivers.Initialize_Window_For_Adobe()
            Drivers.driver4Adobe.get(url)

            xpath = "//ul[@class='nav']//a[@class='dropdown-toggle esign-options hidden-xs visible-sm visible-md visible-lg']/b"
            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.XPATH, xpath)))
            optionDropdown = Drivers.driver4Adobe.find_element_by_xpath(xpath)
            optionDropdown.click()

            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.ID, 'rejectBtn')))
            rejectButton = Drivers.driver4Adobe.find_element_by_id('rejectBtn')
            rejectButton.click()

            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.ID, 'form-control-declineComments')))
            rejectcommentTextBox = Drivers.driver4Adobe.find_element_by_id(
                'form-control-declineComments')
            rejectcommentTextBox.click()
            rejectcommentTextBox.send_keys('Rejected')
            formattedDateTime = self.getDateTime()
            xpath = "//button[text()='Decline']"
            Drivers.driver4AdobeWait.until(
                EC.visibility_of_element_located((By.XPATH, xpath)))
            declineButton = Drivers.driver4Adobe.find_element_by_xpath(xpath)
            declineButton.click()
            Drivers.driver4Adobe.close()

            createdByUserID = os.getenv("ADOBE_USER_ID")

            batchATABStatus = self.verifyATAB(
                createdByUserID, formattedDateTime)
            print(batchATABStatus)
            sleep(15)

            Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
            Drivers.driver.refresh()

    @step("Accept the SendForSignature Alert")
    def accept_the_sendforsignature_alert(self):
        try:
            Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
            driverWait = WebDriverWait(Drivers.driver, 20)
            alert_obj = driverWait.until(EC.alert_is_present())
            alert_obj.accept()
        except UnexpectedAlertPresentException:
            Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
            print("Unexpected Alert Present")
            Messages.write_message("Unexpected Alert Present")
            driverWait = WebDriverWait(Drivers.driver, 20)
            alert_obj = driverWait.until(EC.alert_is_present())
            alert_obj.accept()

    @step("Select <userProfile> from the select sender window of ILOC of <ilocRegion> region")
    def select_from_the_select_sender_window_of_iloc_of_region(self, userProfile, ilocRegion):
        sleep(5)
        if ilocRegion in data_store.spec:
            iLOCDetails = data_store.spec.get(ilocRegion)
        print(iLOCDetails)
        userName = os.getenv(userProfile)
        Drivers.driver.switch_to.window(
            Drivers.driver.window_handles[-1])
        sleep(0.5)

        Drivers.driverWait.until(EC.visibility_of_element_located(
            (By.XPATH, "//td[input[@value='Go']]/preceding-sibling::th/input")))
        sleep(0.5)

        txtBoxSearch = Drivers.driver.find_element_by_xpath(
            "//td[input[@value='Go']]/preceding-sibling::th/input")
        txtBoxSearch.clear()
        txtBoxSearch.send_keys(userName)
        sleep(0.5)

        Drivers.driverWait.until(EC.visibility_of_element_located(
            (By.XPATH, "//input[@value='Go']")))
        sleep(0.5)

        buttonSearch = Drivers.driver.find_element_by_xpath(
            "//input[@value='Go']")
        buttonSearch.click()
        sleep(1)

        Drivers.driverWait.until(EC.visibility_of_element_located(
            (By.LINK_TEXT, userName)))
        sleep(0.5)

        linkAccountName = Drivers.driver.find_element_by_link_text(
            userName)
        linkAccountName.click()
        sleep(0.5)

        Drivers.driver.switch_to.window(
            Drivers.driver.window_handles[-1])
        sleep(3)

        sleep(0.5)
        Messages.write_message("Send On Behlf Of Name: " + userName)
        # iLOCDetails["Customer Name"] = userName   


    @step("Delete OPP and OLI from SF and SQLite")
    def delete_opp_and_oli_from_sf_and_sqlite(self):
        Drivers.Initialize_SalesForce_Instance()
        identity_url = Drivers.sf.restful('')['identity']
        userDetails = Drivers.sf.User.get(identity_url[-18:])
        userId = userDetails['Id']

        oppSql = f"SELECT Id,Name FROM Opportunity WHERE CreatedDate in (TODAY) AND CreatedById = '{userId}'"
        queryResult = Drivers.sf.query_all(query=oppSql)
        recDetails = queryResult['records']

        if len(recDetails) > 0:
            for oppRec in recDetails:
                oliSQL = f"SELECT Id FROM OpportunityLineItem WHERE OpportunityId = '{oppRec['Id']}'"
                queryResult = Drivers.sf.query_all(query=oliSQL)
                oliRecDetails = queryResult['records']
                if len(oliRecDetails) > 0:
                    for oliRec in oliRecDetails:
                        sql_delete_query = f"DELETE from OLI where Id LIKe '{oliRec['Id']}%'"
                        Drivers.dbCursor.execute(sql_delete_query)
                        Drivers.dbConn.commit()

                        isDeleted = Drivers.sf.OpportunityLineItem.delete(
                            oliRec['Id'])
                        if isDeleted == 204:
                            print(f"OLI record {oliRec['Id']} deleted...")
                            Messages.write_message(
                                f"Record {oliRec['Id']} deleted...")
                        else:
                            print(f"OLI record {oliRec['Id']} Not deleted...")
                            Messages.write_message(
                                f"Record {oliRec['Id']} not deleted...")
                else:
                    print("No OpportunityLineItem records found to delete")
                    Messages.write_message(
                        "No OpportunityLineItem records found to delete")
                sql_delete_query = f"DELETE from Opportunity where Id LIKe '{oppRec['Id']}%'"
                Drivers.dbCursor.execute(sql_delete_query)
                Drivers.dbConn.commit()
                isDeleted = Drivers.sf.Opportunity.delete(oppRec['Id'])
                if isDeleted == 204:
                    print(f"OPP record {oppRec['Name']} deleted...")
                    Messages.write_message(
                        f"Record {oppRec['Name']} deleted...")
                else:
                    print(f"OPP record {oppRec['Name']} not deleted...")
                    Messages.write_message(
                        f"Record {oppRec['Name']} not deleted...")
        if len(recDetails) == 0:
            print("No Opportunity records found to delete")
            Messages.write_message("No Opportunity records found to delete")

    @step("Delete <oppName> OPP and its OLI from SF and SQLite")
    def delete_opp_and_its_oli_from_sf_and_sqlite(self, oppName):
        # Drivers.Initialize_SalesForce_Instance()
        identity_url = Drivers.sf.restful('')['identity']
        userDetails = Drivers.sf.User.get(identity_url[-18:])
        userId = userDetails['Id']
        
        today = date.today()
        oppToFind = None
        oppDateToday = today.strftime("%Y%m%d")
        oppToFind = f"{oppName}{oppDateToday}"

        oppSql = f"SELECT Id,Name FROM Opportunity WHERE Name = '{oppToFind}' AND CreatedById = '{userId}'"
        queryResult = Drivers.sf.query_all(query=oppSql)
        recDetails = queryResult['records']

        if len(recDetails) > 0:
            for oppRec in recDetails:
                oliSQL = f"SELECT Id FROM OpportunityLineItem WHERE OpportunityId = '{oppRec['Id']}'"
                queryResult = Drivers.sf.query_all(query=oliSQL)
                oliRecDetails = queryResult['records']
                if len(oliRecDetails) > 0:
                    for oliRec in oliRecDetails:
                        sql_delete_query = f"DELETE from OLI where Id LIKe '{oliRec['Id']}%'"
                        Drivers.dbCursor.execute(sql_delete_query)
                        Drivers.dbConn.commit()

                        isDeleted = Drivers.sf.OpportunityLineItem.delete(
                            oliRec['Id'])
                        if isDeleted == 204:
                            print(f"OLI record {oliRec['Id']} deleted...")
                            Messages.write_message(
                                f"Record {oliRec['Id']} deleted...")
                        else:
                            print(f"OLI record {oliRec['Id']} Not deleted...")
                            Messages.write_message(
                                f"Record {oliRec['Id']} not deleted...")
                else:
                    print("No OpportunityLineItem records found to delete")
                    Messages.write_message(
                        "No OpportunityLineItem records found to delete")
                sql_delete_query = f"DELETE from Opportunity where Id LIKe '{oppRec['Id']}%'"
                Drivers.dbCursor.execute(sql_delete_query)
                Drivers.dbConn.commit()
                isDeleted = Drivers.sf.Opportunity.delete(oppRec['Id'])
                if isDeleted == 204:
                    print(f"OPP record {oppRec['Name']} deleted...")
                    Messages.write_message(
                        f"Record {oppRec['Name']} deleted...")
                else:
                    print(f"OPP record {oppRec['Name']} not deleted...")
                    Messages.write_message(
                        f"Record {oppRec['Name']} not deleted...")
        if len(recDetails) == 0:
            print("No Opportunity records found to delete")
            Messages.write_message("No Opportunity records found to delete")

    @after_spec("<CUSTOM_ILOC_CA> and <CUSTOM_ILOC_US> and <ILOC_CC_US> and <ILOC_CC_CA> and <ILOC_STATUS_UPDATE_US> and <ILOC_STATUS_UPDATE_CA>")
    def after_spec_hook(self):
        Drivers.driver.quit()

    # @before_step
    # def before_step_hook(self):
    #     today = date.today()
    #     screenShotDateTime = today.strftime("%Y%m%d %H:%M.%S")
    #     rootPath = Path(__file__).parents[1]
    #     screenShotFileName = str(
    #         rootPath) + "\\reports\\screenshots\\" + screenShotDateTime + ".png"
    #     Drivers.driver.get_screenshot_as_file(screenShotFileName)

    @after_step
    def after_step_hook(self, context):
        if context.step.is_failing == True:
            # Messages.write_message(context.step.text)
            # Messages.write_message(context.step.message)
            now = datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
            Messages.write_message(f"Current Execution Date and Time: {dt_string}")
            ILOC.takeScreenShot(self, "F-")

    @before_step
    def before_step_hook(self):
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        Messages.write_message(f"Current Execution Date and Time: {dt_string}")

# pdb.set_trace()
        # currentURL = Drivers.driver.current_url
        # opportunityURL = f"{currentURL.split('.com/')[0]}.com/{oppId}"
        # Drivers.driver.get(opportunityURL)

        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, 'topButtonRow')))
        # data_store.spec[ilocRegion] = iLOCDetails

        # iframeElement = Drivers.driver.find_element_by_xpath(
        #     "//iframe[@title='inlineRevenueCalculation']")
        # Drivers.driver.switch_to.frame(iframeElement)

        # xpath = "//button[text()='Manage Products']"
        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.XPATH, xpath)))

        # buttonManageProducts = Drivers.driver.find_element_by_xpath(xpath)
        # buttonManageProducts.click()
        # sleep(5)
        # Drivers.driver.switch_to.default_content()

        # xpath = "//input[@value='Next']"
        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.XPATH, xpath)))
        # buttonNext = Drivers.driver.find_element_by_xpath(xpath)
        # buttonNext.click()

        # xpath = "//input[@value='Save']"
        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.XPATH, xpath)))
        # buttonSave = Drivers.driver.find_element_by_xpath(xpath)
        # buttonSave.click()

        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, 'topButtonRow')))
from decimal import Decimal
from getgauge.python import step, Messages, after_suite, after_spec, before_suite, after_step
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import uuid
import json
import os
from pathlib import Path
import sys
import xlsxwriter
from getgauge.python import data_store
from datetime import date
from time import sleep
import random
from random import randint
from step_impl import Drivers
from step_impl import Common_Steps
from step_impl import Utils
from selenium.webdriver.support.select import Select
from getgauge.python import DataStoreFactory, Screenshots
from datetime import datetime
import shutil
import yaml
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook.workbook import Workbook
import pdb
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
# from itertools import product


class Form_Header():
    
    @step("Open <pageName> Page")
    def open_tab(self, pageName):
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, "setupLink")))
        setupLink = Drivers.driver.find_element_by_id("setupLink")
        setupLink.click()

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, "setupSearch")))
        setupSearchTextBox = Drivers.driver.find_element_by_id("setupSearch")
        setupSearchTextBox.send_keys(pageName)
        
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, "AsyncApexJobs_font")))
        setupLink = Drivers.driver.find_element_by_id("AsyncApexJobs_font")
        setupLink.click()
        
        apexJobsTableXPATH = "//table[@class='list']/tbody"
        apexJobsTable = Drivers.driver.find_element_by_xpath(apexJobsTableXPATH)
        tableRows = apexJobsTable.find_elements_by_tag_name("tr")
        
        apexJobsTableColumnXPATH = "//table[@class='list']/tbody/tr[1]/th"
        columnsList = Utils.get_column_index(apexJobsTableColumnXPATH)
        print(columnsList)
        
        
        batchATAB = ["AccountTerritoryAssociationBatch", "UserHierarchyQueueDataSetupBatch", "UserHierarchyQueueDataSetupBatchUUHList", "UserHierarchyQueueDataSetupBatchO2AList", "UserHierarchyQueueAccountBatch", "UserHierarchyQueueOpportunityBatch", "UserHierarchyQueueGoalBatch", "SupportUserHierarchySharingBatch", "userSharingBatch", "UserSharingBatchForFormHeaders", "UserRevokeSharingDataSetUpBatch", "UserSharingRevokeBatch", "UserSharingRevokeForFormHeaders"]
        
        for batch in batchATAB:
#             cntr = 2
            for tableRow in tableRows:                
                pdb.set_trace()
#                 a = Drivers.driver.find_element_by_xpath("//table[@class='list']/tbody/tr[2]/td[1]")
#                 apexClassXPath = f"//table[@class='list']/tbody/tr[{cntr}]/td[{columnsList['Apex Class']-1}]"             
                apexClassElement = tableRow.find_element_by_xpath(f"//td[{columnsList['Apex Class']-1}]")
                apexClassValue = apexClassElement.text
                
#                 statusXPath = f"//table[@class='list']/tbody/tr[{cntr}]/td[{columnsList['Status']-1}]"               
                statusElement = tableRow.find_element_by_xpath(f"//td[{columnsList['Status']-1}]")
                statusValue = apexClassElement.text
                
#                 cntr = cntr + 1
#                 el = WebDriverWait(Drivers.driver).until(lambda d: d.find_element_by_tag_name("p"))
    @after_spec("<FormHeader>")
    def after_spec_hook(self):
        Drivers.driver.quit()

    @after_step
    def after_step_hook(self, context):
        if context.step.is_failing == True:
            Messages.write_message(context.step.text)
            # Messages.write_message(context.step.message)
            Screenshots.capture_screenshot()    

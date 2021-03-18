from simple_salesforce.format import format_soql
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
from selenium.webdriver.support.select import Select
from getgauge.python import data_store, Screenshots
import re
import shutil
import yaml
import pdb
import requests
import json
from pages import PricingExceptionPage
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook.workbook import Workbook
from selenium.webdriver import ActionChains
import pytz
from dateutil import tz
from simple_salesforce.format import format_soql
from selenium.webdriver.support.wait import WebDriverWait
from step_impl import Drivers
from step_impl import Common_Steps
from step_impl import Utils


class PricingException():
    pricingExceptionDetails = {}
    @step("Select Pricing Excecption Type <peValue>")
    def open_tab(self, peValue):
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        pe = PricingExceptionPage.PricingExceptionPage(
            Drivers.driver, Drivers.driverWait)
        pe.select_form_header_value(peValue)

    @step("Click <buttonName> button")
    def click_button(self, buttonName):
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        pe = PricingExceptionPage.PricingExceptionPage(
            Drivers.driver, Drivers.driverWait)
        pe.click_buton(buttonName)

    @step("Create <peType> Pricing Exception for <region> region <table>")
    def create_pricing_exception_for_region(self, peType, region, table):
        Drivers.driver.switch_to.window(Drivers.driver.window_handles[-1])
        data_store.spec.clear()
        # pricingExceptionDetails = {}
        # Drivers.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, "Name")))
        fieldsData = {}
        dt = datetime.today()
        print(dt.month, " ", dt.day, " ", dt.year)
        year = int(dt.year) + 2
        currentDate = f"{str(dt.month)}/{str(dt.day)}/{str(year)}"

        rows = table.rows
        for row in rows:
            row0 = str(row[0]).strip()
            row1 = str(row[1]).strip()
            if row0 == "Form Request Name":
                fieldsData["Form Request Name"] = row1
            elif row0 == "Regional Sales Approver":
                fieldsData["Regional Sales Approver"] = row1
            elif row0 == "Sales Team Approver (GSM)":
                fieldsData["Sales Team Approver (GSM)"] = row1
            elif row0 == "Opportunity":
                fieldsData["Opportunity"] = row1
            elif row0 == "Account":
                fieldsData["Account"] = row1
            elif row0 == "Product":
                fieldsData["Product"] = row1
            elif row0 == "Integrated proposal":
                fieldsData["Integrated proposal"] = row1
            elif row0 == "Approved based on Counter Proposal?":
                fieldsData["Approved based on Counter Proposal?"] = row1
            elif row0 == "Projected Scale":
                fieldsData["Projected Scale"] = row1
            elif row0 == "Estimated Total Revenue":
                fieldsData["Estimated Total Revenue"] = row1
            elif row0 == "Proposed Pricing Request":
                fieldsData["Proposed Pricing Request"] = row1
            elif row0 == "Reasoning":
                fieldsData["Reasoning"] = row1
            elif row0 == "Insert Date/Start Date":
                fieldsData["Insert Date_Start Date"] = row1
            elif row0 == "Cycle(s)":
                fieldsData["Cycle(s)"] = row1
            elif row0 == "Trade Class/ Store Count":
                fieldsData["Trade Class/ Store Count"] = row1
            elif row0 == "Categories":
                fieldsData["Categories"] = row1
            elif row0 == "Due by date":
                fieldsData["Due by date"] = row1

            elif row0 == "AD":
                fieldsData["AD"] = row1
            elif row0 == "Brand":
                fieldsData["Brand"] = row1
            elif row0 == "Sale ID":
                fieldsData["Sale ID"] = row1
            elif row0 == "Manufacturer":
                fieldsData["Manufacturer"] = row1
            elif row0 == "Test Panel 1":
                fieldsData["Test Panel 1"] = row1
            elif row0 == "Test Panel 2":
                fieldsData["Test Panel 2"] = row1
            elif row0 == "Test Panel 3":
                fieldsData["Test Panel 3"] = row1
            elif row0 == "First Test Cycle":
                fieldsData["First Test Cycle"] = row1
            elif row0 == "Last Test Cycle":
                fieldsData["Last Test Cycle"] = row1
            elif row0 == "Desired Metrics for Analysis":
                fieldsData["Desired Metrics for Analysis"] = row1
            elif row0 == "Desired class of trade":
                fieldsData["Desired class of trade"] = row1
            elif row0 == "Retail Banner Preference":
                fieldsData["Retail Banner Preference"] = row1
            elif row0 == "InStore activity":
                fieldsData["InStore activity"] = row1
            elif row0 == "Matched Panel Test":
                fieldsData["Matched Panel Test"] = row1
            elif row0 == "Brand Has FSI Month":
                fieldsData["Brand Has FSI Month"] = row1
            elif row0 == "Type Of Matched Panel Test":
                fieldsData["Type Of Matched Panel Test"] = row1
            elif row0 == "Client Being Billed":
                fieldsData["Client Being Billed"] = row1
            elif row0 == "Sell Check Score":
                fieldsData["Sell Check Score"] = row1
            elif row0 == "No Score Reason":
                fieldsData["No Score Reason"] = row1
            elif row0 == "Additional Comments":
                fieldsData["Additional Comments"] = row1

        print(fieldsData)
        Messages.write_message(fieldsData)

        if "Form Request Name" in fieldsData:
            days = None
            today = date.today()
            dateToday = today.strftime("%Y%m%d")
            peName = f"{fieldsData['Form Request Name']}{dateToday}"
            Messages.write_message("Opportunity Name: " + peName)
            sleep(0.5)
            self.pricingExceptionDetails["Form Request Name"] = peName

            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_form_request_name(peName)

        if "Regional Sales Approver" in fieldsData:
            regionalSalesApprover = os.getenv(
                fieldsData["Regional Sales Approver"])
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_regional_sales_approver_from_lookup(
                regionalSalesApprover)
            self.pricingExceptionDetails["Regional Sales Approver"] = regionalSalesApprover

        if "Sales Team Approver (GSM)" in fieldsData:
            salesTeamApprover = os.getenv(
                fieldsData["Sales Team Approver (GSM)"])
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_sales_team_approver_from_lookup(salesTeamApprover)
            self.pricingExceptionDetails["Sales Team Approver (GSM)"] = salesTeamApprover

        if "Opportunity" in fieldsData:
            oppName = fieldsData["Opportunity"]
            today = date.today()
            dateToday = today.strftime("%Y%m%d")
            oppName = f"{oppName}{dateToday}#Parent"
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_opportunity_from_lookup(oppName)
            self.pricingExceptionDetails["Opportunity"] = oppName

        if "Account" in fieldsData:
            accName = os.getenv(fieldsData["Account"])
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_account_from_lookup(accName)
            self.pricingExceptionDetails["Account"] = accName

        if "Product" in fieldsData:
            productName = fieldsData["Product"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            productList = productName.split(",")
            for product in productList:
                pe.select_product_from_multiselectpicklist(product)
                pe.click_add_product_buton()
            self.pricingExceptionDetails["Product"] = productName

        if "Integrated proposal" in fieldsData:
            isIntegratedProposal = fieldsData["Integrated proposal"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_inegrated_proposal_from_picklist(isIntegratedProposal)
            self.pricingExceptionDetails["Integrated proposal"] = isIntegratedProposal

        # if "Approved based on Counter Proposal?" in fieldsData:
        #     iscounterProposal = fieldsData["Approved based on Counter Proposal?"]
        #     pe = PricingExceptionPage.PricingExceptionPage(
        #         Drivers.driver, Drivers.driverWait)
        #     pe.select_counter_proposal_from_picklist(iscounterProposal)
        #     self.pricingExceptionDetails["Counter proposal"] = iscounterProposal

        if "Projected Scale" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_projected_scale(fieldsData["Projected Scale"])
            self.pricingExceptionDetails["Projected Scale"] = fieldsData["Projected Scale"]

        if "Due by date" in fieldsData:
            today = date.today()
            dateToday = today.strftime("%m/%d/%Y")
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_due_by_date(dateToday)
            self.pricingExceptionDetails["Due by date"] = dateToday

        if "Insert Date_Start Date" in fieldsData:
            today = date.today()
            dateToday = today.strftime("%m/%d/%Y")
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_insert_date_start_date(dateToday)
            self.pricingExceptionDetails["Insert Date_Start Date"] = dateToday

        if "Estimated Total Revenue" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_estimated_total_revenue(
                fieldsData["Estimated Total Revenue"])
            self.pricingExceptionDetails["Estimated Total Revenue"] = fieldsData["Estimated Total Revenue"]

        if "Cycle(s)" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_cycle(
                fieldsData["Cycle(s)"])
            self.pricingExceptionDetails["Cycle(s)"] = fieldsData["Cycle(s)"]

        if "Categories" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_categories(
                fieldsData["Categories"])
            self.pricingExceptionDetails["Categories"] = fieldsData["Categories"]

        if "Trade Class/ Store Count" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_tradeClassStoreCount(
                fieldsData["Trade Class/ Store Count"])
            self.pricingExceptionDetails["Trade Class/ Store Count"] = fieldsData["Trade Class/ Store Count"]

        if "Proposed Pricing Request" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_proposed_pricing_request(
                fieldsData["Proposed Pricing Request"])
            self.pricingExceptionDetails["Proposed Pricing Request"] = fieldsData["Proposed Pricing Request"]

        if "Reasoning" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_reasoning(fieldsData["Reasoning"])
            self.pricingExceptionDetails["Reasoning"] = fieldsData["Reasoning"]

        if "AD" in fieldsData:
            ad = os.getenv(fieldsData["AD"])
            print(ad)
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_ad_from_lookup(ad)
            self.pricingExceptionDetails["AD"] = ad

        if "Brand" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_brand(fieldsData["Brand"])
            self.pricingExceptionDetails["Brand"] = fieldsData["Brand"]

        if "Sale ID" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_sale_id(fieldsData["Sale ID"])
            self.pricingExceptionDetails["Sale ID"] = fieldsData["Sale ID"]

        if "Manufacturer" in fieldsData:
            manufacturer = os.getenv(fieldsData["Manufacturer"])
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_manufacturer_from_lookup(manufacturer)
            self.pricingExceptionDetails["Manufacturer"] = manufacturer

        if "Test Panel 1" in fieldsData:
            productName = fieldsData["Test Panel 1"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            productList = productName.split(",")
            for product in productList:
                pe.select_value_from_test_panel_1(product)
                pe.click_test_panel_1_add_buton()
            self.pricingExceptionDetails["Test Panel 1"] = productName

        if "Test Panel 2" in fieldsData:
            productName = fieldsData["Test Panel 2"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            productList = productName.split(",")
            for product in productList:
                pe.select_value_from_test_panel_2(product)
                pe.click_test_panel_2_add_buton()
            self.pricingExceptionDetails["Test Panel 2"] = productName

        if "Test Panel 3" in fieldsData:
            productName = fieldsData["Test Panel 3"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            productList = productName.split(",")
            for product in productList:
                pe.select_value_from_test_panel_3(product)
                pe.click_test_panel_3_add_buton()
            self.pricingExceptionDetails["Test Panel 3"] = productName

        if "First Test Cycle" in fieldsData:
            firstTestCycle = fieldsData["First Test Cycle"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_first_test_cycle_from_lookup(firstTestCycle)
            self.pricingExceptionDetails["First Test Cycle"] = firstTestCycle

        if "Last Test Cycle" in fieldsData:
            lastTestCycle = fieldsData["Last Test Cycle"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_last_test_cycle_from_lookup(lastTestCycle)
            self.pricingExceptionDetails["Last Test Cycle"] = lastTestCycle

        if "Desired Metrics for Analysis" in fieldsData:
            value = fieldsData["Desired Metrics for Analysis"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_desired_metrics_for_analysis_from_picklist(value)
            self.pricingExceptionDetails["Desired Metrics for Analysis"] = value

        if "Desired class of trade" in fieldsData:
            value = fieldsData["Desired class of trade"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_desired_class_of_trade_from_picklist(value)
            self.pricingExceptionDetails["Desired class of trade"] = value

        if "Retail Banner Preference" in fieldsData:
            values = fieldsData["Retail Banner Preference"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            valuesList = values.split(",")
            for value in valuesList:
                pe.select_value_from_retail_banner_preference(value)
                pe.click_retail_banner_preference_add_buton()
            self.pricingExceptionDetails["Retail Banner Preference"] = values

        if "InStore activity" in fieldsData:
            value = fieldsData["InStore activity"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_instore_activity_from_picklist(value)
            self.pricingExceptionDetails["InStore activity"] = value

        if "Matched Panel Test" in fieldsData:
            value = fieldsData["Matched Panel Test"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_matched_panel_test_from_picklist(value)
            self.pricingExceptionDetails["Matched Panel Test"] = value

        if "Brand Has FSI Month" in fieldsData:
            value = fieldsData["Brand Has FSI Month"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_fsi_month_from_picklist(value)
            self.pricingExceptionDetails["Brand Has FSI Month"] = value

        if "Type Of Matched Panel Test" in fieldsData:
            value = fieldsData["Type Of Matched Panel Test"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_type_of_matched_panel_test_from_picklist(value)
            self.pricingExceptionDetails["Type Of Matched Panel Test"] = value

        if "Client Being Billed" in fieldsData:
            value = fieldsData["Client Being Billed"]
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.select_client_being_billed_from_picklist(value)
            self.pricingExceptionDetails["Client Being Billed"] = value

        if "Sell Check Score" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_sell_check_score(fieldsData["Sell Check Score"])
            self.pricingExceptionDetails["Sell Check Score"] = fieldsData["Sell Check Score"]

        if "No Score Reason" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_no_score_reason(fieldsData["No Score Reason"])
            self.pricingExceptionDetails["No Score Reason"] = fieldsData["No Score Reason"]

        if "Additional Comments" in fieldsData:
            pe = PricingExceptionPage.PricingExceptionPage(
                Drivers.driver, Drivers.driverWait)
            pe.enter_additional_comments(fieldsData["Additional Comments"])
            self.pricingExceptionDetails["Additional Comments"] = fieldsData["Additional Comments"]
        data_store.spec[peType] = self.pricingExceptionDetails
    @step("Verify details of <Pricing Exception> Form of <region> region")
    def verify_pricing_exception_form_of_region(self,peType, region):
        pe = PricingExceptionPage.PricingExceptionPage(
            Drivers.driver, Drivers.driverWait)
        pricingExceptionDetails = None
        # if peType in data_store.spec:
        #     pricingExceptionDetails = data_store.spec.get(peType)
        # currentURL = Drivers.driver.current_url
        # peFormID = currentURL.split(".com/")[1]
        # Messages.write_message(f"Record URL : {currentURL}")
        # Messages.write_message(f"Contract ID : {peFormID}")
        # pricingExceptionDetails['Form Request Name'] = peFormID

        # queryData = (Drivers.sf.query(format_soql(
        #     "SELECT Id, Name FROM Form_Header__c WHERE Id = {}", peFormID)))['records']
        # # pdb.set_trace()
        # peName = queryData[0]['Name']
        
        # Common_Steps.CommonSteps.write_data_to_table_column(
        #     self, "PricingException", "Id", 'id', peFormID, peFormID)
        # Common_Steps.CommonSteps.write_data_to_table_column(
        #     self, "PricingException", "Form Name", 'text', peName, peFormID)
        # Common_Steps.CommonSteps.write_data_to_table_column(
        #     self, "PricingException", "Form Type", 'text', peType, peFormID)
        # pe.verify_pricing_exception_details(self.pricingExceptionDetails)

    @step("Accept the Submit For Approval Alert")
    def accept_the_sendforsignature_alert(self):
        try:
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

    @step("Logout from the SalesForce Application")
    def logout_from_the_salesforce_applocation(self):
        pe = PricingExceptionPage.PricingExceptionPage(
            Drivers.driver, Drivers.driverWait)
        pe.logout_from_salesforce()

    @step("User <Approves> the Pricing Exception Form")
    def user_the_pricing_exception_form(self, approvalAction):
        pe = PricingExceptionPage.PricingExceptionPage(
            Drivers.driver, Drivers.driverWait)
        pe.user_approve_reject_pricing_exception_form(
            self.pricingExceptionDetails, approvalAction)

    @step("Delete <CO51-DG-PE-IP(N)-> Pricing Exception")
    def delete_pricing_exception(self, peName):

        identity_url = Drivers.sf.restful('')['identity']
        userDetails = Drivers.sf.User.get(identity_url[-18:])
        userId = userDetails['Id']

        today = date.today()
        dateToday = today.strftime("%Y%m%d")
        peToFind = f"{peName}{dateToday}"
        # peSql = f"SELECT Id, Name FROM Form_Header__c WHERE Name = '{peToFind}'"
        recDetails = Drivers.sf.query(format_soql(
            "SELECT Id, Name FROM Form_Header__c WHERE Name like {}", peToFind))['records']
        # queryResult = Drivers.sf.query_all(query=peSql)
        # recDetails = queryResult['records']

        if len(recDetails) > 0:
            for peRec in recDetails:
                isDeleted = Drivers.sf.Form_Header__c.delete(peRec['Id'])
                if isDeleted == 204:
                    print(
                        f"Pricing Exception Form record {peRec['Name']} deleted...")
                    Messages.write_message(
                        f"Pricing Exception Form {peRec['Name']} deleted...")
                else:
                    print(
                        f"Pricing Exception Form {peRec['Name']} not deleted...")
                    Messages.write_message(
                        f"Pricing Exception Form {peRec['Name']} not deleted...")

    @step("Verify Pricing Exception status as <peStatus>")
    def verify_pricing_exception_status_as(self, peStatus):
        pe = PricingExceptionPage.PricingExceptionPage(
            Drivers.driver, Drivers.driverWait)
        pe.verify_pricing_exception_status(peStatus)

    @step("Finance user updated the Approved based on Counter Proposal? field to <value>")
    def finance_user_updated_the_counter_proposal_field_to(self, value):
        pe = PricingExceptionPage.PricingExceptionPage(
            Drivers.driver, Drivers.driverWait)
        pe.update_counter_proposal_value(value)

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
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook.workbook import Workbook
from selenium.webdriver import ActionChains
import pytz
from dateutil import tz
from simple_salesforce.format import format_soql
from selenium.webdriver.support.wait import WebDriverWait


class PricingExceptionPage():
    def __init__(self, driver, driverWait):
        self.driver = driver
        self.driverWait = driverWait

        self.userNav_button_id = "userNavButton"
        self.logout_link_text = "Logout"
        # -----------------------------------------------------------------------------------------------------------------------------
        # Pricing Exception New Page
        # -----------------------------------------------------------------------------------------------------------------------------
        self.recordtype_selectbox_css = "table.detailList > tbody > tr > td > select"

        self.formRequestName_textbox_id = "Name"
        self.regionalSalesApprover_lookup_xpath = "//img[@title='Regional Sales Approver Lookup (New Window)']"
        self.salesTeamApprover_lookup_xpath = "//img[@title='Sales Team Approver (GSM) Lookup (New Window)']"
        self.account_lookup_xpath = "//img[@title='Account Lookup (New Window)']"
        self.counterProposal_select_xpath = "//td[label[text()='Approved based on Counter Proposal?']]/following-sibling::*[position()=1][name()='td']//select | //td[not(label)][text()='Approved based on Counter Proposal?']/following-sibling::*[position()=1][name()='td']//select"
        self.estimatedTotalRevenue_textbox_xpath = "//td[label[text()='Estimated Total Revenue']]/following-sibling::*[position()=1][name()='td']//input | //td[not(label)][text()='Estimated Total Revenue']/following-sibling::*[position()=1][name()='td']//input"

        # CO51-DG-PE
        # -----------------------------------------------------------------------------------------------------------------------------
        self.opportunity_lookup_xpath = "//img[@title='Opportunity Lookup (New Window)']"
        self.productAvailable_select_xpath = "//select[@title='Product - Available']"
        self.productAdd_button_xpath = "//img[@title='Add']"
        self.integratedProposal_select_xpath = "//td[label[text()='Integrated proposal']]/following-sibling::*[position()=1][name()='td']//select | //td[not(label)][text()='Integrated proposal']/following-sibling::*[position()=1][name()='td']//select"
        self.projectedScale_textbox_xpath = "//td[label[text()='Projected Scale']]/following-sibling::*[position()=1][name()='td']//textarea | //td[not(label)][text()='Projected Scale']/following-sibling::*[position()=1][name()='td']//textarea"
        self.proposedPricingRequest_textbox_xpath = "//td[label[text()='Proposed Pricing Request']]/following-sibling::*[position()=1][name()='td']//textarea | //td[not(label)][text()='Proposed Pricing Request']/following-sibling::*[position()=1][name()='td']//textarea"
        self.reasoning_textbox_xpath = "//td[label[text()='Reasoning']]/following-sibling::*[position()=1][name()='td']//textarea | //td[not(label)][text()='Reasoning']/following-sibling::*[position()=1][name()='td']//textarea"

        # CORPORATE RATE AUTHORIZATION
        # -----------------------------------------------------------------------------------------------------------------------------
        self.insertDateStartDate_textbox_xpath = "//td[span[label[contains(text(),'Insert Date')]]]//following-sibling::*[position()=1][name()='td']//input | //td[not(span)[label[contains(text(),'Insert Date')]]]//following-sibling::*[position()=1][name()='td']//input"
        self.insertDateStartDate_label_xpath = "//label[contains(text(),'Insert Date')]"

        # InStore Pricing Exception
        # -----------------------------------------------------------------------------------------------------------------------------
        self.dueByDate_textbox_xpath = "//td[span[label[contains(text(),'Due by date')]]]//following-sibling::*[position()=1][name()='td']//input | //td[not(span)[label[contains(text(),'Due by date')]]]//following-sibling::*[position()=1][name()='td']//input"
        self.dueByDate_label_xpath = "//label[contains(text(),'Due by date')]"
        self.cycle_textbox_xpath = "//td[span[label[text()='Cycle(s)']]]/following-sibling::*[position()=1][name()='td']//input | //td[not(label)][text()='Cycle(s)']/following-sibling::*[position()=1][name()='td']//input"
        self.categories_textbox_xpath = "//td[label[text()='Categories']]/following-sibling::*[position()=1][name()='td']//textarea | //td[not(label)][text()='Categories']/following-sibling::*[position()=1][name()='td']//textarea"        
        self.tradeClassStoreCount_textbox_xpath = "//td[span[label[text()='Trade Class/ Store Count']]]/following-sibling::*[position()=1][name()='td']//textarea | //td[not(label)][text()='Trade Class/ Store Count']/following-sibling::*[position()=1][name()='td']//textarea"                

        # -----------------------------------------------------------------------------------------------------------------------------
        # Pricing Exception Detail Page
        # -----------------------------------------------------------------------------------------------------------------------------
        self.formRequestName_xpath = "//td[span[text()='Form Request Name']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Form Request Name']/following-sibling::*[position()=1][name()='td']/div"
        self.regionalSalesApprover_xpath = "//td[span[text()='Regional Sales Approver']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Regional Sales Approver']/following-sibling::*[position()=1][name()='td']/div"
        self.salesTeamApprover_xpath = "//td[span[text()='Sales Team Approver (GSM)']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Sales Team Approver (GSM)']/following-sibling::*[position()=1][name()='td']/div"
        self.opportunity_xpath = "//td[span[text()='Opportunity']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Opportunity']/following-sibling::*[position()=1][name()='td']/div"
        self.account_xpath = "//td[span[text()='Account']]/following-sibling::*[position()=1][name()='td']/div | //td[not(span)][text()='Account']/following-sibling::*[position()=1][name()='td']/div"
        self.formRequestStatus_xpath = "//td[text()='Status' and @class='labelCol']/following-sibling::td"

        # -----------------------------------------------------------------------------------------------------------------------------
        # Items Approval Page on Home Page
        # -----------------------------------------------------------------------------------------------------------------------------
        self.item_approval_table_id = "PendingProcessWorkitemsList_body"
        self.item_approval_table_records_css = "div#PendingProcessWorkitemsList_body > table > tbody > tr"
        self.item_approval_table_header_css = "div#PendingProcessWorkitemsList_body > table > tbody > tr:first-child > th"
        self.approve_button_name = "goNext"

        # -----------------------------------------------------------------------------------------------------------------------------
        # Matched Panel Test Request Page
        # -----------------------------------------------------------------------------------------------------------------------------
        self.ad_lookup_xpath = "//img[@title='AD Lookup (New Window)']"
        self.manufacturer_lookup_xpath = "//img[@title='Manufacturer Lookup (New Window)']"
        self.brand_textbox_xpath = "//th[span[label[text()='Brand']]]/following-sibling::*[position()=1][name()='td']//input | //td[not(label)][text()='Brand']/following-sibling::*[position()=1][name()='td']//input"        
        self.saleID_textbox_xpath = "//th[label[text()='Sale ID']]/following-sibling::*[position()=1][name()='td']//input | //td[not(label)][text()='Sale ID']/following-sibling::*[position()=1][name()='td']//input"

        self.testPanel1Available_select_xpath = "//select[@title='Test Panel 1 - Available']"
        self.testPanel1Add_button_xpath = "//th[label[text()='Test Panel 1']]/following-sibling::td/descendant::img[@title='Add']"

        self.testPanel2Available_select_xpath = "//select[@title='Test Panel 2 - Available']"
        self.testPanel2Add_button_xpath = "//th[label[text()='Test Panel 2']]/following-sibling::td/descendant::img[@title='Add']"

        self.testPanel3Available_select_xpath = "//select[@title='Test Panel 3 - Available']"
        self.testPanel3Add_button_xpath = "//th[label[text()='Test Panel 3']]/following-sibling::td/descendant::img[@title='Add']"

        self.firstTestCycle_lookup_xpath = "//img[@title='First Test Cycle Lookup (New Window)']"
        self.lastTestCycle_lookup_xpath = "//img[@title='Last Test Cycle Lookup (New Window)']"

        self.desiredMetricsForAnalysis_select_xpath = "//th[label[text()='Desired Metrics for Analysis']]/following-sibling::*[position()=1][name()='td']//select | //td[not(label)][text()='Desired Metrics for Analysis']/following-sibling::*[position()=1][name()='td']//select"
        self.desiredClassOfTrade_select_xpath = "//th[label[text()='Desired class of trade']]/following-sibling::*[position()=1][name()='td']//select | //td[not(label)][text()='Desired class of trade']/following-sibling::*[position()=1][name()='td']//select"

        self.retailBannerPreferenceAvailable_select_xpath = "//select[@title='Retail Banner Preference - Available']"
        self.retailBannerPreferenceAdd_button_xpath = "//th[span[label[text()='Retail Banner Preference']]]/following-sibling::td/descendant::img[@title='Add']"

        self.instoreActivity_select_xpath = "//th[label[contains(text(),'Will this brand have any InStore')]]/following-sibling::*[position()=1][name()='td']//select | //td[not(label)][text()='Desired class of trade']/following-sibling::*[position()=1][name()='td']//select"        
        self.matchedPanelTest_select_xpath = "//th[label[contains(text(),'Will this brand have a Matched Panel')]]/following-sibling::*[position()=1][name()='td']//select | //td[not(label)][text()='Desired class of trade']/following-sibling::*[position()=1][name()='td']//select"
        self.fsiMonth_select_xpath = "//th[label[contains(text(),'Did this brand have an FSI')]]/following-sibling::*[position()=1][name()='td']//select | //td[not(label)][text()='Desired class of trade']/following-sibling::*[position()=1][name()='td']//select"
        self.typeOfMatchedPanelTest_select_xpath = "//th[label[contains(text(),'What type of Matched panel')]]/following-sibling::*[position()=1][name()='td']//select | //td[not(label)][text()='Desired class of trade']/following-sibling::*[position()=1][name()='td']//select"
        self.clientBeingBilled_select_xpath = "//th[label[contains(text(),'What is the client being billed?')]]/following-sibling::*[position()=1][name()='td']//select | //td[not(label)][text()='Desired class of trade']/following-sibling::*[position()=1][name()='td']//select"

        self.saleCheckScore_textbox_xpath = "//th[label[text()='Sell Check Score']]/following-sibling::*[position()=1][name()='td']//input | //td[not(label)][text()='Sell Check Score']/following-sibling::*[position()=1][name()='td']//input"
        self.noScoreReason_textbox_xpath = "//th[label[text()='No Score Reason']]/following-sibling::*[position()=1][name()='td']//textarea | //td[not(label)][text()='No Score Reason']/following-sibling::*[position()=1][name()='td']//textarea"
        self.additionalComments_textbox_xpath = "//th[label[text()='Additional Comments']]/following-sibling::*[position()=1][name()='td']//textarea | //td[not(label)][text()='Additional Comments']/following-sibling::*[position()=1][name()='td']//textarea"


    def select_ad_from_lookup(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.ad_lookup_xpath)))
        self.driver.find_element_by_xpath(
            self.ad_lookup_xpath).click()
        self.select_value_from_lookup(value)
        Messages.write_message("AD: " + value)

    def select_manufacturer_from_lookup(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.manufacturer_lookup_xpath)))
        self.driver.find_element_by_xpath(
            self.manufacturer_lookup_xpath).click()
        self.select_value_from_lookup(value)
        Messages.write_message("Manufacturer: " + value)

    def enter_brand(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.brand_textbox_xpath)))
        self.driver.find_element_by_xpath(self.brand_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.brand_textbox_xpath).send_keys(value)
        Messages.write_message("Brand: " + value)

    def enter_sale_id(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.saleID_textbox_xpath)))
        self.driver.find_element_by_xpath(self.saleID_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.saleID_textbox_xpath).send_keys(value)
        Messages.write_message("Sale ID: " + value)

    def select_value_from_test_panel_1(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.testPanel1Available_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.testPanel1Available_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("Test Panel 1: " + value)

    def click_test_panel_1_add_buton(self):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.testPanel1Add_button_xpath)))
        self.driver.find_element_by_xpath(self.testPanel1Add_button_xpath).click()

    def select_value_from_test_panel_2(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.testPanel2Available_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.testPanel2Available_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("Test Panel 2: " + value)

    def click_test_panel_2_add_buton(self):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.testPanel2Add_button_xpath)))
        self.driver.find_element_by_xpath(self.testPanel2Add_button_xpath).click()

    def select_value_from_test_panel_3(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.testPanel3Available_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.testPanel3Available_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("Test Panel 3: " + value)

    def click_test_panel_3_add_buton(self):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.testPanel3Add_button_xpath)))
        self.driver.find_element_by_xpath(self.testPanel3Add_button_xpath).click()

    def select_first_test_cycle_from_lookup(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.firstTestCycle_lookup_xpath)))
        self.driver.find_element_by_xpath(
            self.firstTestCycle_lookup_xpath).click()
        self.select_value_from_lookup(value)
        Messages.write_message("First Test Cycle: " + value)

    def select_last_test_cycle_from_lookup(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.lastTestCycle_lookup_xpath)))
        self.driver.find_element_by_xpath(
            self.lastTestCycle_lookup_xpath).click()
        self.select_value_from_lookup(value)
        Messages.write_message("Last Test Cycle: " + value)

    def select_desired_metrics_for_analysis_from_picklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.desiredMetricsForAnalysis_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.desiredMetricsForAnalysis_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("Desired Metrics for Analysis: " + value)

    def select_desired_class_of_trade_from_picklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.desiredClassOfTrade_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.desiredClassOfTrade_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("Desired class of trade: " + value)

    def select_value_from_retail_banner_preference(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.retailBannerPreferenceAvailable_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.retailBannerPreferenceAvailable_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("Retail Banner Preference: " + value)

    def click_retail_banner_preference_add_buton(self):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.retailBannerPreferenceAdd_button_xpath)))
        self.driver.find_element_by_xpath(self.retailBannerPreferenceAdd_button_xpath).click()
        
    def select_instore_activity_from_picklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.instoreActivity_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.instoreActivity_select_xpath))
        dropDownProductLine.select_by_value(value)
        Messages.write_message("Instore Activity: " + value)
        
    def select_matched_panel_test_from_picklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.matchedPanelTest_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.matchedPanelTest_select_xpath))
        dropDownProductLine.select_by_value(value)
        Messages.write_message("Matched Panel Test: " + value)

    def select_type_of_matched_panel_test_from_picklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.typeOfMatchedPanelTest_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.typeOfMatchedPanelTest_select_xpath))
        dropDownProductLine.select_by_value(value)
        Messages.write_message("Matched Panel Test: " + value)
        
    def select_fsi_month_from_picklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.fsiMonth_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.fsiMonth_select_xpath))
        dropDownProductLine.select_by_value(value)
        Messages.write_message("FSI Month: " + value)
        
    def select_client_being_billed_from_picklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.clientBeingBilled_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.clientBeingBilled_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("FSI Month: " + value)

    def enter_sell_check_score(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.saleCheckScore_textbox_xpath)))
        self.driver.find_element_by_xpath(self.saleCheckScore_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.saleCheckScore_textbox_xpath).send_keys(value)
        Messages.write_message("Sale Check Score: " + value)

    def enter_no_score_reason(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.noScoreReason_textbox_xpath)))
        self.driver.find_element_by_xpath(self.noScoreReason_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.noScoreReason_textbox_xpath).send_keys(value)
        Messages.write_message("No Score Reason: " + value)

    def enter_additional_comments(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.additionalComments_textbox_xpath)))
        self.driver.find_element_by_xpath(self.additionalComments_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.additionalComments_textbox_xpath).send_keys(value)
        Messages.write_message("No Score Reason: " + value)


                
    def click_buton(self, buttonName):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, f"//input[@value='{buttonName}'] | //input[@title='{buttonName}'] | //input[@name='{buttonName}']")))
        self.driver.find_element_by_xpath(
            f"//input[@value='{buttonName}'] | //input[@title='{buttonName}'] | //input[@name='{buttonName}']").click()

    # def click_new_form_header_buton(self, buttonName):
    #     self.driverWait.until(
    #         EC.visibility_of_element_located((By.XPATH, f"//input[@value='{buttonName}']")))
    #     self.driver.find_element_by_xpath(f"//input[@value='{buttonName}']").click()

    def select_form_header_value(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, self.recordtype_selectbox_css)))
        dropDownProductLine = Select(
            self.driver.find_element_by_css_selector(self.recordtype_selectbox_css))
        dropDownProductLine.select_by_visible_text(value)

    def enter_form_request_name(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.ID, self.formRequestName_textbox_id)))
        self.driver.find_element_by_id(self.formRequestName_textbox_id).clear()
        self.driver.find_element_by_id(
            self.formRequestName_textbox_id).send_keys(value)
        Messages.write_message("Form Request Name: " + value)

    def select_regional_sales_approver_from_lookup(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.regionalSalesApprover_lookup_xpath)))
        self.driver.find_element_by_xpath(
            self.regionalSalesApprover_lookup_xpath).click()
        self.select_value_from_lookup(value)
        Messages.write_message("Regional Sales Approver: " + value)

    def select_sales_team_approver_from_lookup(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.salesTeamApprover_lookup_xpath)))
        self.driver.find_element_by_xpath(
            self.salesTeamApprover_lookup_xpath).click()
        self.select_value_from_lookup(value)
        Messages.write_message("Sales Team Approver (GSM): " + value)

    def select_opportunity_from_lookup(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.opportunity_lookup_xpath)))
        self.driver.find_element_by_xpath(
            self.opportunity_lookup_xpath).click()
        self.select_value_from_lookup(value)
        Messages.write_message("Opportunity: " + value)

    def select_account_from_lookup(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.account_lookup_xpath)))
        self.driver.find_element_by_xpath(self.account_lookup_xpath).click()
        self.select_value_from_lookup(value)
        Messages.write_message("Account: " + value)

    def select_product_from_multiselectpicklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.productAvailable_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.productAvailable_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("Products: " + value)

    def click_add_product_buton(self):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.productAdd_button_xpath)))
        self.driver.find_element_by_xpath(self.productAdd_button_xpath).click()

    def select_inegrated_proposal_from_picklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.integratedProposal_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.integratedProposal_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("Integrated Proposal: " + value)

    def select_counter_proposal_from_picklist(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.counterProposal_select_xpath)))
        dropDownProductLine = Select(self.driver.find_element_by_xpath(
            self.counterProposal_select_xpath))
        dropDownProductLine.select_by_visible_text(value)
        Messages.write_message("Counter Proposal: " + value)

    def enter_insert_date_start_date(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.insertDateStartDate_label_xpath)))
        wbElement = self.driver.find_element_by_xpath(self.insertDateStartDate_label_xpath)
        insertDateStartDate_textbox_id = wbElement.get_attribute("for")
        insertDateStartDate_textbox = self.driver.find_element_by_id(insertDateStartDate_textbox_id)
        # self.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, insertDateStartDate_textbox_id)))
        insertDateStartDate_textbox.clear()
        insertDateStartDate_textbox.send_keys(value)
        Messages.write_message("Insert Date/Start Date: " + value)

    def enter_due_by_date(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.dueByDate_label_xpath)))
        wbElement = self.driver.find_element_by_xpath(self.dueByDate_label_xpath)
        dueByDate_textbox_id = wbElement.get_attribute("for")
        dueByDate_textbox = self.driver.find_element_by_id(dueByDate_textbox_id)
        # self.driverWait.until(
        #     EC.visibility_of_element_located((By.ID, insertDateStartDate_textbox_id)))
        dueByDate_textbox.clear()
        dueByDate_textbox.send_keys(value)
        Messages.write_message("Due by date: " + value)        

    def enter_categories(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.categories_textbox_xpath)))
        self.driver.find_element_by_xpath(
            self.categories_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.categories_textbox_xpath).send_keys(value)
        Messages.write_message("Categories: " + value)

    def enter_tradeClassStoreCount(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.tradeClassStoreCount_textbox_xpath)))
        self.driver.find_element_by_xpath(
            self.tradeClassStoreCount_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.tradeClassStoreCount_textbox_xpath).send_keys(value)
        Messages.write_message("Trade Class/ Store Count: " + value)
        
    def enter_cycle(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.cycle_textbox_xpath)))
        self.driver.find_element_by_xpath(
            self.cycle_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.cycle_textbox_xpath).send_keys(value)
        Messages.write_message("Cycle(s): " + value)
        
    def enter_projected_scale(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.projectedScale_textbox_xpath)))
        self.driver.find_element_by_xpath(
            self.projectedScale_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.projectedScale_textbox_xpath).send_keys(value)
        Messages.write_message("Projected Scale: " + value)

    def enter_estimated_total_revenue(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.estimatedTotalRevenue_textbox_xpath)))
        self.driver.find_element_by_xpath(
            self.estimatedTotalRevenue_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.estimatedTotalRevenue_textbox_xpath).send_keys(value)
        Messages.write_message("Estimated Total Revenue: " + value)

    def enter_proposed_pricing_request(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.proposedPricingRequest_textbox_xpath)))
        self.driver.find_element_by_xpath(
            self.proposedPricingRequest_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.proposedPricingRequest_textbox_xpath).send_keys(value)
        Messages.write_message("Proposed Pricing Request: " + value)

    def enter_reasoning(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.reasoning_textbox_xpath)))
        self.driver.find_element_by_xpath(self.reasoning_textbox_xpath).clear()
        self.driver.find_element_by_xpath(
            self.reasoning_textbox_xpath).send_keys(value)
        Messages.write_message("Reasoning: " + value)

    def select_value_from_lookup(self, value):
        self.driver.switch_to.window(self.driver.window_handles[1])
        # sleep(0.2)

        self.driverWait.until(EC.visibility_of_element_located(
            (By.ID, "searchFrame")))
        # sleep(0.2)

        self.driver.switch_to.frame("searchFrame")
        # self.driver.switch_to_frame("searchFrame")
        # sleep(0.2)

        self.driverWait.until(EC.visibility_of_element_located(
            (By.ID, "lksrch")))
        # sleep(0.2)
        txtBoxSearch = self.driver.find_element_by_id("lksrch")
        txtBoxSearch.send_keys(value)
        # sleep(0.2)

        self.driverWait.until(EC.visibility_of_element_located(
            (By.NAME, "go")))
        # sleep(0.2)
        buttonSearch = self.driver.find_element_by_name("go")
        buttonSearch.click()
        # sleep(0.2)

        self.driver.switch_to.window(self.driver.window_handles[1])
        sleep(0.2)

        self.driverWait.until(EC.visibility_of_element_located(
            (By.ID, "resultsFrame")))
        sleep(0.2)

        self.driver.switch_to.frame("resultsFrame")
        sleep(0.2)

        self.driverWait.until(EC.visibility_of_element_located(
            (By.LINK_TEXT, value)))
        sleep(0.2)

        linkAccountName = self.driver.find_element_by_link_text(
            value)
        linkAccountName.click()
        sleep(0.2)

        self.driver.switch_to.window(self.driver.window_handles[0])
        # sleep(3)

    def verify_pricing_exception_details(self, pricingExceptionDetailMap):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.formRequestName_xpath)))

        if 'Form Request Name' in pricingExceptionDetailMap:
            peName = pricingExceptionDetailMap["Form Request Name"]

            peNameOnRecord = self.driver.find_element_by_xpath(
                self.formRequestName_xpath).text

            assert peName == peNameOnRecord, f"{peName} <> {peNameOnRecord}"
            Messages.write_message(f"{peName} == {peNameOnRecord}")

        if 'Sales Team Approver (GSM)' in pricingExceptionDetailMap:
            peSalesTeamApprover = pricingExceptionDetailMap["Sales Team Approver (GSM)"]

            peSalesTeamApproverOnRecord = self.driver.find_element_by_xpath(
                self.salesTeamApprover_xpath).text

            assert peSalesTeamApprover == peSalesTeamApproverOnRecord, f"{peSalesTeamApprover} <> {peSalesTeamApproverOnRecord}"
            Messages.write_message(
                f"{peSalesTeamApprover} == {peSalesTeamApproverOnRecord}")

        if 'Regional Sales Approver' in pricingExceptionDetailMap:
            peRegionalApprover = pricingExceptionDetailMap["Regional Sales Approver"]

            peRegionalSalesApproverOnRecord = self.driver.find_element_by_xpath(
                self.regionalSalesApprover_xpath).text

            assert peRegionalApprover == peRegionalSalesApproverOnRecord, f"{peRegionalApprover} <> {peRegionalSalesApproverOnRecord}"
            Messages.write_message(
                f"{peRegionalApprover} == {peRegionalSalesApproverOnRecord}")

        if 'Account' in pricingExceptionDetailMap:
            peAccount = pricingExceptionDetailMap["Account"]

            peAccountOnRecord = self.driver.find_element_by_xpath(
                self.account_xpath).text

            assert peAccount == peAccountOnRecord, f"{peAccount} <> {peAccountOnRecord}"
            Messages.write_message(f"{peAccount} == {peAccountOnRecord}")

        if 'Opportunity' in pricingExceptionDetailMap:
            peOpportunity = pricingExceptionDetailMap["Opportunity"]

            peOpportunityOnRecord = self.driver.find_element_by_xpath(
                self.opportunity_xpath).text

            assert peOpportunity == peOpportunityOnRecord, f"{peOpportunity} <> {peOpportunityOnRecord}"
            Messages.write_message(
                f"{peOpportunity} == {peOpportunityOnRecord}")

    def logout_from_salesforce(self):
        self.driverWait.until(
            EC.visibility_of_element_located((By.ID, self.userNav_button_id)))
        userMenu = self.driver.find_element_by_id(self.userNav_button_id)
        userMenu.click()

        self.driverWait.until(
            EC.visibility_of_element_located((By.LINK_TEXT, self.logout_link_text)))
        linkLogout = self.driver.find_element_by_link_text(
            self.logout_link_text)
        linkLogout.click()

    def user_approve_reject_pricing_exception_form(self, pricingExceptionDetailMap, approvalAction):
        self.driverWait.until(
            EC.visibility_of_element_located((By.ID, self.item_approval_table_id)))
        peName = pricingExceptionDetailMap["Form Request Name"]
        getColumnIndex = self.get_column_index("css_selector",
                                               self.item_approval_table_header_css)

        tableRecords = self.driver.find_elements_by_css_selector(
            self.item_approval_table_records_css)

        for records in tableRecords:
            relatedTo = records.find_element_by_css_selector("th").text
            if relatedTo == peName:
                records.find_element_by_css_selector(
                    "td:first-child > a:nth-child(2)").click()
                self.driver.find_element_by_name(
                    self.approve_button_name).click()
                # alert_obj = self.driverWait.until(EC.alert_is_present())
                # alert_obj.accept()
                sleep(5)
                break

    def get_column_index(self, xpathCSSSelector, tableHeadersXpath):
        labelsMap = {}
        try:
            tableHeader = None
            if xpathCSSSelector == "xpath":
                tableHeader = self.driver.find_elements_by_css_selector(
                    tableHeadersXpath)

            elif xpathCSSSelector == "css_selector":
                tableHeader = self.driver.find_elements_by_css_selector(
                    tableHeadersXpath)
            columnIndex = 1
            for theader in tableHeader:
                columnName = theader.text.title()
                labelsMap[columnName] = columnIndex
                Messages.write_message(
                    columnName + " : " + str(columnIndex))
                columnIndex = columnIndex + 1
            return labelsMap
        except NoSuchElementException:
            print("element not found")

    def verify_pricing_exception_status(self, peStatus):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.formRequestStatus_xpath)))
        peStatusOnRecord = self.driver.find_element_by_xpath(
            self.formRequestStatus_xpath).text

        assert peStatusOnRecord == peStatus, f"{peStatus} <> {peStatusOnRecord}"
        Messages.write_message(
            f"Verified Pricing Exception Form Status {peStatus} == {peStatusOnRecord}")

    def update_counter_proposal_value(self, value):
        self.driverWait.until(
            EC.visibility_of_element_located((By.XPATH, self.counterProposal_select_xpath)))
        self.select_counter_proposal_from_picklist(value)

    # def get_property_value(self, wbElement, wbElementProperty):
    #     propertyValue = wbElement.get_attribute(wbElementProperty)
    #     return propertyValue

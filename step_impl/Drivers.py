from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
import os
import sqlite3
from simple_salesforce import Salesforce, SFType, SalesforceLogin
from getgauge.python import Messages, Screenshots, after_step
from datetime import datetime
import chromedriver_autoinstaller
from pathlib import Path
import pdb
driver = None
driverWait = None
sf = None
driver4Adobe = None
driver4AdobeWait = None
dbConn = None
dbCursor = None

def Initialize():
    cwd = os.path.dirname(os.path.realpath(__file__))
    print(cwd)
    # CHORME_PATH = cwd + "\\" + "chromedriver.exe"
    chromedriver_autoinstaller.install()
    global driver
    global driverWait
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": cwd,
             'download.directory_upgrade': True}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(chrome_options=chrome_options)
    driverWait = WebDriverWait(driver, 60)
    return driver, driverWait

def Initialize_Window_For_Adobe():
    chromedriver_autoinstaller.install()
    global driver4Adobe
    global driver4AdobeWait
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    driver4Adobe = webdriver.Chrome(chrome_options=chrome_options)
    driver4AdobeWait = WebDriverWait(driver4Adobe, 720)
    return driver4Adobe, driver4AdobeWait

def Initialize_SalesForce_Instance():
    global sf
    # session_id, instance = SalesforceLogin(
    #     username=os.getenv("USER_ID"), password=os.getenv("USER_PASSWORD"), security_token=os.getenv("USER_SECURITY_TOKEN"), domain='test')

    session_id, instance = SalesforceLogin(
        username=os.getenv("USER_ID"), password=os.getenv("USER_PASSWORD"), security_token=os.getenv("USER_SECURITY_TOKEN"), domain='test')
    
    print(session_id, "\n", instance)
    Messages.write_message(str(session_id) + " : " + str(instance))
    sf = Salesforce(instance=instance, session_id=session_id)
    return sf

def Initialize_Database_Instance():
    global dbConn
    global dbCursor
    root = Path(__file__).parents[1]
    dbFilePath = str(root) + "\\Data\\" + os.getenv('DB_NAME')
    print(dbFilePath)
    dbConn = sqlite3.connect(dbFilePath)
    print("Opened database successfully", dbConn)
    Messages.write_message(f"Opened DB {dbConn} Successfully")
    dbCursor = dbConn.cursor()
    print("Cursor Object: ", dbCursor)
    Messages.write_message(f"Cursor Object {dbCursor}")
    # return SQLite3Connection.cur

    return dbConn, dbCursor


def CloseDriver():
    driver.quit()

@after_step
def after_step_hook(context):
    if context.step.is_failing == True:
        Messages.write_message(context.step.text)
        # Messages.write_message(context.step.message)
        Screenshots.capture_screenshot()
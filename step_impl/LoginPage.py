from getgauge.python import step, Messages, before_spec, before_scenario, data_store, after_step
from getgauge.python import ExecutionContext, Scenario, Screenshots, Specification
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import os
from step_impl import Drivers


class Login():
    # driverWait = WebDriverWait(Drivers.driver, 50)
    # driver, driverWait = None
    # def __init__(self):
    #     self.driver,self.driverWait = Drivers.Initialize()

    # ---------------------------
    # Gauge step implementations
    # ---------------------------
    @step("Open the SalesForce <orgType> Login Page")
    def open_salesforce_login_page(self, orgType):
        Drivers.Initialize()
        Drivers.Initialize_SalesForce_Instance()
        if orgType == "SANDBOX":
            Drivers.driver.get(os.getenv("URL"))
        elif orgType == "PRODUCTION":
            Drivers.driver.get(os.getenv("URL"))
        Drivers.driver.maximize_window()
        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'Login')))

    @step("Open sqlite connection for db <dbName>")
    def open_sqlite_connection_for_db(self, dbName):
        Drivers.Initialize_Database_Instance()

    @step("User is logging in to Org")
    def log_in_as_user(self):
        data_store.spec['US'] = {"ILOC_GRAND_TOTAL": 0}
        data_store.spec['CA'] = {"ILOC_GRAND_TOTAL": 0}
        userName = ""
        userPwd = ""
        # if orgType == "SIT":
        #     userName = os.getenv("USER_ID")
        #     userPwd = os.getenv("USER_PASSWORD")
        # elif orgType == "UAT":
        #     userName = os.getenv("USER_ID")
        #     userPwd = os.getenv("USER_PASSWORD")
        # elif orgType == "PROD":
        #     userName = os.getenv("USER_ID")
        #     userPwd = os.getenv("USER_PASSWORD")

        userName = os.getenv("USER_ID")
        userPwd = os.getenv("USER_PASSWORD")

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'username')))
        txtBoxUserName = Drivers.driver.find_element_by_id("username")
        txtBoxUserName.send_keys(userName)

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'password')))
        txtBoxPassword = Drivers.driver.find_element_by_id("password")
        txtBoxPassword.send_keys(userPwd)

        Drivers.driverWait.until(
            EC.visibility_of_element_located((By.ID, 'Login')))
        btnLogin = Drivers.driver.find_element_by_id("Login")
        btnLogin.click()

        Messages.write_message("Logged in as: " + userName)

    @after_step
    def after_step_hook(self, context):
        if context.step.is_failing == True:
            Messages.write_message(context.step.text)
            # Messages.write_message(context.step.message)
            Screenshots.capture_screenshot()        

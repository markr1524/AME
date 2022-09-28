from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from login import ChromeDriver
from datetime import datetime
import time


cty_table = {"Annapolis": 5, "Anne Arundel": 9, "Baltimore City": 10, "Catonsville": 11,
             "Essex": 3, "Towson": 2, "Carroll": 6, "Fredrick": 13, "Harford": 12, "Howard": 4,
             "Montgomery": 1, "Prince George": 7}


class Reminders:

    def __init__(self, driver, counties, batch):
        self.driver = driver
        self.counties = counties
        self.batch = batch

    def login_rbnet(self):
        # -------------------- LOGIN PAGE -----------------------------
        self.driver.get("https://www.residencybureau.com/CFS_Login.html")
        company = self.driver.find_element_by_id("txtCompanyID")
        username = self.driver.find_element_by_id("txtUserID")
        password = self.driver.find_element_by_id("txtPassword")
        company.clear()
        username.clear()
        password.clear()
        username.send_keys("user1")
        password.send_keys("dino_1524")
        company.send_keys('rbnet')
        self.driver.find_element_by_name("cmdOK").click()

        self.driver.get("https://www.residencybureau.com/selCustomOption.asp")
        self.driver.find_element_by_id("ctyreminders").click()

        # Select the batch
        select = Select(self.driver.find_element_by_name("Batch"))
        select.select_by_value(self.batch)

        # Add today's date
        day = self.driver.find_element_by_name('day1')
        day.clear()
        day.send_keys(str(datetime.today().day))

        queue = [f'{cty_table[county]}' for county in self.counties]
        while queue:
            notification = queue.pop()
            select = Select(self.driver.find_element_by_name("County"))
            select.select_by_value(notification)
            self.driver.find_element_by_name("OK").click()
            time.sleep(0.5)
            self.driver.back()


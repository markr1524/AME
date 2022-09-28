from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from paths import directory




class OpsMerchant:

    def __init__(self, driver, processing_type):
        self.driver = driver
        self.processing_type = processing_type

    def login_rbnet(self, company_id: str):
        """
        Handles the login to the RBNet portal

        Parameters
        ----------
        company_id: str
            Code required to login to a specific property.
        """
        self.driver.get("https://www.residencybureau.com/CFS_Login.html")
        # -------------------- LOGIN PAGE -----------------------------
        company = self.driver.find_element_by_id("txtCompanyID")
        username = self.driver.find_element_by_id("txtUserID")
        password = self.driver.find_element_by_id("txtPassword")
        company.clear()
        username.clear()
        password.clear()
        username.send_keys("user1")
        password.send_keys("rbnet224")
        company.send_keys(company_id)
        self.driver.find_element_by_name("cmdOK").click()

    def select_suit_or_writ(self):
        """
        Selects suits or writs from the billing page.
        """
        # -------------------- SUITS OR WRITS -----------------------------
        self.driver.get(f"https://www.residencybureau.com/InvCaseTypeMM.asp")
        if self.processing_type == 0:
            self.driver.find_element_by_name('r1').click()
        else:
            self.driver.find_elements_by_name('r1')[1].click()
        self.driver.find_element_by_xpath("//input[@value='Submit']").click()
import os
import traceback
import shutil
import time
from paths import directory
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains


folder = os.getcwd() + "\\invoices"


class VendorCafe:
    """
    This class contains the logic for extracting RBNet invoices for upload to the
    Third-party processing system, Vendor Cafe. It can also optionally save the invoice
    and print the check/invoice pages.


    Parameters
    ----------
    driver : ChromeWebdriver
        Instance of the chromewebdriver object. Shares a connection with the websocket.
    processing_type : int
        Integer representing 0 for suits 1 for writs, and 2 for notices.
    printing: bool
        Determines if the invoice and check request should be printed.
    saving: bool
        Determines if the invoice should be saved on the Z-drive.

    """

    def __init__(self, driver, processing_type: int, printing=False, saving=False, invoice_loc=2):
        self.driver = driver
        self.printing = printing
        self.processing_type = processing_type
        self.saving = saving
        self.invoice_loc = invoice_loc
        self.invoice_date = None

    @staticmethod
    def clear_folder():
        """
        Clears the invoice folder of any files that may not have been cleaned up.
        """
        for f in os.listdir(folder):
            file_path = os.path.join(folder, f)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(e)

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
        password.send_keys("dino_1524")
        company.send_keys(company_id)
        self.driver.find_element_by_name("cmdOK").click()

    def extract_property_info(self):
        """
        Extracts the property name from the property info page.
        NOTE: This could probably be skipped using a built-in look-up for performance.

        Returns
        -------
        str
            The name of the property (not the alias) will be returned.
        """
        # -------------------- PROPERTY INFO PAGE -----------------------------
        self.driver.get("https://www.residencybureau.com/selProperty.asp")
        select = Select(self.driver.find_element_by_name("Property"))
        select.select_by_index(1)
        self.driver.find_element_by_name("B1").click()
        pty_name = self.driver.find_element_by_name("ptyComplexName").get_attribute("value")
        return pty_name

    def select_suit_or_writ_or_notice(self):
        """
        Selects suits, writs or notices from the billing page.
        """
        # -------------------- SUITS, WRITS, OR NOTICES  -----------------------------
        self.driver.get(f"https://www.residencybureau.com/InvCaseTypeMM.asp")
        if self.processing_type == 0:
            self.driver.find_element_by_name('r1').click()
        elif self.processing_type == 1:
            self.driver.find_elements_by_name('r1')[1].click()
        else:
            self.driver.find_elements_by_name('r1')[2].click()
        self.driver.find_element_by_xpath("//input[@value='Submit']").click()

    def extract_invoice(self):
        """
        Extracts the invoice number and the grand total from the generated invoice page.

        Returns
        -------
        tuple[str]
            Returns a tuple of two strings, one being the invoice_number and the other being the grand total.
        """
        # -------------------- INVOICE PAGE -----------------------------
        wait = WebDriverWait(self.driver, 22)

        if self.processing_type == 0:
            ele = self.driver.find_element(By.XPATH, f"//form/center/table/tbody/tr[{self.invoice_loc}]/td[1]/a")
        elif self.processing_type == 1:
            ele = self.driver.find_element(By.XPATH, f"//form/table/tbody/tr[{self.invoice_loc}]/td[1]/a")
        else:
            ele1 = self.driver.find_element(By.XPATH, f"//form/table/tbody/tr[{self.invoice_loc}]/td[7]/a[1]")
            self.driver.execute_script("arguments[0].click();", ele1)
            ele = self.driver.find_element(By.XPATH, f"//form/table/tbody/tr[{self.invoice_loc}]/td[1]/a")
        self.driver.execute_script("arguments[0].click();", ele)

        invoice_number = self.driver.find_element_by_xpath("//a").text
        total = self.driver.find_element_by_id("grandtotal").text
        self.invoice_date = self.driver.find_element_by_id("invoice_date").text
        if self.printing:
            self.driver.execute_script("window.print();")
        self.driver.find_element_by_xpath("//a").click()
        self.driver.find_element_by_xpath("//a").click()
        return invoice_number, total

    def print_check_request(self):
        """
        Only ran if printing is True. Returns to the table page to select the check request link.
        Then some javascript is executed on the page to force a page print.
        """
        # -------------------- CHECK REQUEST PAGE -----------------------------
        self.driver.back()
        self.driver.back()
        if self.processing_type == 0:
            self.driver.find_element_by_xpath(f"//form/center/table/tbody/tr[{self.invoice_loc}]/td[3]/a").click()
        else:
            self.driver.find_element_by_xpath(f"//form/table/tbody/tr[{self.invoice_loc}]/td[3]/a").click()
        self.driver.execute_script("window.print();")

    def vc_login(self):
        """
         Sends the login details to vendor-cafe
        """
        self.driver.get("https://www.vendor-cafe.com/vendorcafe/")

        username = self.driver.find_element_by_name("userId")
        password = self.driver.find_element_by_name('password')

        username.clear()
        username.send_keys("markr@residencybureau.com")
        password.clear()
        password.send_keys("VendorCafe_1$24")
        self.driver.find_element(By.XPATH, '//button[@onclick="return fnSubmit();"]').click()
        try:
            supplier = self.driver.find_element_by_link_text("RESIDENCY BUREAU, LLC")
            supplier.click()
        except Exception as e:
            pass

    def select_management_company(self, code):
        """
        Directs the driver to select a particular management company for upload.
        At the time of writing this, only Bozzuto and Kettler use Vendor Cafe.
        """
        wait = WebDriverWait(self.driver, 22)
        element = wait.until(EC.invisibility_of_element_located((By.ID, 'preloader')))
        button = wait.until(EC.presence_of_element_located((By.ID, "btnHeaderCompanySwitch")))
        element = self.driver.find_element_by_id("preloader")
        button = self.driver.find_element_by_id("btnHeaderCompanySwitch")
        self.driver.execute_script("arguments[0].click();", element)
        self.driver.execute_script("arguments[0].click();", button)
        if code.startswith('ksi') or code == 'Kettler':
            WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//input[@value='43341']")))
            element2 = self.driver.find_element_by_xpath("//input[@value='43341']")
        elif code.startswith('bm') or code == 'Bozzuto':
            WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//input[@value='27041']")))
            element2 = self.driver.find_element_by_xpath("//input[@value='27041']")
        else:
            WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//input[@value='48556']")))
            element2 = self.driver.find_element_by_xpath("//input[@value='48556']")
        self.driver.execute_script("arguments[0].click();", element2)

        wait.until(EC.presence_of_element_located((By.XPATH, "//a[@class='menutoggle']")))
        element3 = self.driver.find_element_by_xpath("//a[@class='menutoggle']")
        ActionChains(self.driver).move_to_element(element3).click(element3).perform()

    def navigate_to_invoice(self):
        """
        Since we can't just skip to the upload a pdf invoice page, we need to navigate there...
        click by click... This locates the element and uses some javascript to execute a click
        event on a child tag.
        """
        wait = WebDriverWait(self.driver, 22)
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//a[@title='Create/Upload Invoice']")))
        element = self.driver.find_element(By.XPATH, '//a[@title="Create/Upload Invoice"]')
        self.driver.execute_script("arguments[0].click();", element)

        wait.until(EC.presence_of_element_located(
           (By.XPATH, '//a[text()="Upload a PDF Invoice"]')))
        element = self.driver.find_element(By.XPATH, '//a[text()="Upload a PDF Invoice"]')
        self.driver.execute_script("arguments[0].click();", element)

    def populate_invoice(self, grand_total: str, invoice_number: str):
        """
        Populates the invoice with the information we extracted from the generated invoice.

        """
        wait = WebDriverWait(self.driver, 20)
        wait.until(EC.presence_of_element_located(
            (By.ID, "pdfPassword")))
        self.driver.find_element_by_id("pdfPassword").clear()

        invoice_no = self.driver.find_element_by_id("invoiceNo")
        invoice_no.clear()
        invoice_no.send_keys(invoice_number)

        total = self.driver.find_element_by_name("invoiceTotalDisplay")
        total.clear()
        total.send_keys(grand_total.replace("$", ""))
        wait.until(EC.presence_of_element_located(
            (By.NAME, "invoiceDateDisplay")))
        invoice_date = self.driver.find_element_by_name("invoiceDateDisplay")
        # We send today's date as the input.
        invoice_date.send_keys(self.invoice_date)
        # Removes the annoying calendar widget pop-up
        invoice_date.send_keys(Keys.ENTER)

    def select_property(self, pty_name):
        wait = WebDriverWait(self.driver, 20)
        text = ("javascript:return pickListLinkForJSP('rcashcommon/P2P_PickList.xml','VC_Building^union all^VC_"
                "Building_multi','hdnPropVenPartyId^propVenPartyId~buildingNo^Code~buildingName^Property Name',"
                "'single','','','','');")
        wait.until(EC.invisibility_of_element_located((By.ID, 'preloader')))
        element = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.XPATH,  f'//button[@onclick="{text}"]'))
        )
        ActionChains(self.driver).move_to_element(element).click(element).perform()
        wait.until(EC.presence_of_element_located(
            (By.ID, "vistaModalIframe110")))
        self.driver.switch_to.frame(self.driver.find_element_by_id("vistaModalIframe110"))

        wait = WebDriverWait(self.driver, 20)
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@name='txtValue']")))

        property_name = self.driver.find_element_by_xpath("//input[@name='txtValue']")
        property_name.clear()
        print(pty_name)
        if pty_name.upper() == "ENCLAVE AT EMERSON APARTMENTS":
            pty_name = "Enclave at Emerson"
        elif pty_name == "The Bowen (frmly Harmony Place)":
            pty_name = "The Bowen"
        elif pty_name.upper() == "WINTHROP APARTMENTS":
            pty_name = "Winthrop"
        elif pty_name.upper() == "FLATS 170 AT ACADEMY YARD":
            pty_name = "Flats170 at Academy Yard"
        elif pty_name.upper() == "THE GRAMERCY AT TOWN CENTER":
            pty_name = "Gramercy at Town Center"
        elif pty_name.upper() == "THE VINE APARTMENTS":
            pty_name = "The Vine"
        elif pty_name.upper() == "THE ESPLANADE":
            pty_name = "Esplanade"
        elif pty_name.upper() == "THE TOWNES AT HARVEST VIEW":
            pty_name = "Townes at Harvest View"
        elif pty_name.upper() == "THE WHITNEY":
            pty_name = "Whitney Apartments"
        elif pty_name.upper() == "THE COURTS OF DEVON":
            pty_name = "Courts of Devon"
        elif pty_name.upper() == "CROSSWINDS AT ANNAPOLIS TOWNE CENTRE":
            pty_name = "Crosswinds"
        elif pty_name.upper() == "THE METROPOLITAN":
            pty_name = "Metropolitan - Office"
        elif pty_name.upper() == "LAKEHOUSE RESIDENCES":
            pty_name = "Lakehouse"
        elif pty_name.upper() == "SAINT PAUL SENIOR LIVING I":
            pty_name = "St. Paul Senior Living - Phase 1"
        elif pty_name.upper() == "SAINT PAUL SENIOR LIVING II":
            pty_name = "St. Paul Senior Living - Phase 2"
        elif pty_name.upper() == "STONE POINT APARTMENTS":
            pty_name = "Stone Point"
        elif pty_name.upper() == "THE REDWOOD":
            pty_name = "Redwood Apartments"
        elif pty_name.upper() == "THE BROADVIEW APARTMENTS":
            pty_name = "Broadview Apartments"
        elif pty_name.upper() == "FIELDS OF ROCKVILLE":
            pty_name = "The Fields of Rockville"

        property_name.send_keys(pty_name)

        self.driver.find_element_by_xpath("//button[text()='Find']").click()
        time.sleep(1)
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@type='radio' and @value='0' and @name='check']")))

        f = self.driver.find_elements_by_xpath("//input[@type='radio' and @value='0']")[0]
        f.click()

        element4 = wait.until(EC.presence_of_element_located(
            (By.NAME, "form_submit")))
        ActionChains(self.driver).move_to_element(element4).click(element4).perform()
        self.driver.switch_to.default_content()

    def upload_invoice(self, invoice_number, mode=0):
        if mode == 0:
            self.driver.find_element_by_id("invUploadFile").send_keys(folder + f"\\{invoice_number}.pdf")
        else:
            self.driver.find_element_by_id("invUploadFile").send_keys(invoice_number)

    @staticmethod
    def save_invoice(invoice_number, code):
        fp = folder + f"\\{invoice_number}.pdf"
        try:
            dest = directory[code]
            shutil.copy(fp, dest)
        except KeyError:
            print("Unable to save the file. Path was not found in the directory.")
            print(f"CODE: {code} INVOICE: {invoice_number}")
        except Exception as e:
            print(e)
            print("Invoice has already been saved.")

    def atty_run(self, invoice, total, property_name, management, pdf_path):
        try:
            self.vc_login()
            self.select_management_company(management)
            self.navigate_to_invoice()
            self.populate_invoice(total, invoice)
            self.select_property(property_name)
            self.upload_invoice(pdf_path, mode=1)
        except Exception:
            print("-" * 60)
            err = traceback.format_exc()
            print('-' * 60)

            return 404, err, "", ""
        else:
            return 200, property_name, invoice, total

    def run(self, code):
        """
        try:
            code = sys.argv[1]
        except IndexError:
            raise ValueError("You must input a property code. Restart the script.")
        """
        self.clear_folder()
        try:
            self.login_rbnet(code)
            pty_name = self.extract_property_info()
            self.select_suit_or_writ_or_notice()
            invoice_number, total = self.extract_invoice()
            if self.printing:
                self.print_check_request()
            self.vc_login()
            self.select_management_company(code)
            self.navigate_to_invoice()
            self.populate_invoice(total, invoice_number)
            self.select_property(pty_name)
            os.rename(f"{folder}\\{os.listdir(folder)[0]}", f"{folder}\\{invoice_number}.pdf")
            self.upload_invoice(invoice_number)
            if self.saving:
                self.save_invoice(invoice_number, code)
        except Exception:
            print("-" * 60)
            err = traceback.format_exc()
            print('-' * 60)

            return 404, err, "", ""
        else:
            return 200, pty_name, invoice_number, total

import os
import shutil
from shutil import SameFileError
import time
import traceback
from collections import namedtuple
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from paths import directory

customer_table = {"gsbl-OLD": (97949, 1905159), "gsbl": (3101707, 3363776), "gsbf": (97949, 2083412),
                  "gseg": (97949, 2103486), "gshss": (97949, 1550932), "gslm": (97949, 2119547),
                  "gspco": (97949, 1389077), "gsshc-OLD": (97949, 2115833), "gsshc": (3101707, 3363793),
                  "gssf-OLD": (97949, 1922015), "gssf": (3101707, 3363795), "gsgg": (97949, 1438719),
                  "gscam": (97949, 2035888), "gshkp": (97949, 2114450), "gsmeb": (97949, 2137220),
                  "gsubr": (97949, 2103498), "gsso": (97949, 2044674), "gs703": (97949, 2425645),
                  "gset": (97949, 2401281), "gsscp": (97949, 1800631), "cjvms": (1351215, 1360673),
                  "gots": (924221, 1500758), "gosc": (924221, 1756833),  "brh5074": (158453, 1106190),
                  "jpica": (97949, 2795307)
                  }

folder = os.getcwd() + "\\invoices"
county_name = ""
tally = ""


class OpsMerchant:

    def __init__(self, driver, processing_type, printing=False, saving=False, inv_location=2):
        self.driver = driver
        self.printing = printing
        self.processing_type = processing_type
        self.saving = saving
        self.location = inv_location
        self.invoice_date = None

    def login_rbnet(self, company_id):
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

        # -------------------- PROPERTY INFO PAGE -----------------------------
        self.driver.get("https://www.residencybureau.com/selProperty.asp")
        select = Select(self.driver.find_element_by_name("Property"))
        select.select_by_index(1)
        self.driver.find_element_by_name("B1").click()
        requester = self.driver.find_element_by_name("ptyContact").get_attribute("value")
        complex_name = self.driver.find_element_by_name("ptyComplexName").get_attribute("value")
        select = Select(self.driver.find_element_by_name("ptyCourt"))
        county = select.first_selected_option.text
        global county_name
        county_name = select.first_selected_option.text

        # -------------------- SELECT SUITS, WRITS, OR NOTICES  -----------------------------
        self.driver.get(f"https://www.residencybureau.com/InvCaseTypeMM.asp")
        if self.processing_type == 0:
            self.driver.find_element_by_name('r1').click()
        elif self.processing_type == 1:
            self.driver.find_elements_by_name('r1')[1].click()
        else:
            self.driver.find_elements_by_name('r1')[2].click()
        self.driver.find_element_by_xpath("//input[@value='Submit']").click()

        # ---------------------------- INVOICE -----------------------------
        if self.processing_type == 0:
            self.driver.find_element_by_xpath(f"//form/center/table/tbody/tr[{self.location}]/td[1]/a").click()
        elif self.processing_type == 1:
            self.driver.find_element_by_xpath(f"//form/table/tbody/tr[{self.location}]/td[1]/a").click()
        else:
            self.driver.find_element_by_xpath(f"//form/table/tbody/tr[{self.location}]/td[1]/a").click()

        invoice_number = self.driver.find_element_by_xpath("//a").text
        self.invoice_date = self.driver.find_element_by_id("invoice_date").text
        if self.printing:
            self.driver.execute_script("window.print();")
        if self.saving:
            self.driver.find_element_by_xpath("//a").click()
            self.driver.find_element_by_xpath("//a").click()
            self.driver.back()
        if self.processing_type == 0:
            self.driver.back()
        elif self.processing_type == 1:
            self.driver.back()

        # --------------------------- CHECK REQUEST -----------------------------
        if self.processing_type == 0:
            self.driver.find_element_by_xpath(f"//form/center/table/tbody/tr[{self.location}]/td[3]/a").click()
        elif self.processing_type == 1:
            self.driver.find_element_by_xpath(f"//form/table/tbody/tr[{self.location}]/td[3]/a").click()

        #                     GATHER TENANTS BLOCK
        if self.processing_type == 0:
            tenants_one_count = self.driver.find_element_by_id("count1").text
            tenants_one_cost = self.driver.find_element_by_id("cost1").text
            tenants_two_count = self.driver.find_element_by_id("count2").text
            tenants_two_cost = self.driver.find_element_by_id("cost2").text
            tenants_three_count = self.driver.find_element_by_id("count3").text
            tenants_three_cost = self.driver.find_element_by_id("cost3").text
            tenants_four_count = self.driver.find_element_by_id("count4").text
            tenants_four_cost = self.driver.find_element_by_id("cost4").text
            tbl_block = (tenants_one_count, tenants_one_cost,
                         tenants_two_count, tenants_two_cost,
                         tenants_three_count, tenants_three_cost,
                         tenants_four_count, tenants_four_cost)
            global tally
            tally = int(tenants_one_count) + int(tenants_two_count) + int(tenants_three_count) + int(tenants_four_count)

        elif self.processing_type == 1:
            count = int(self.driver.find_element_by_id("writcount").text)
            cost = self.driver.find_element_by_id('writcost').text
            total = self.driver.find_element_by_id('writtotal').text
            total = float(total.replace('$', '').replace(',', ''))
            if county == "MONTGOMERY COUNTY":
                count += int(self.driver.find_element_by_id('abswritcount').text)
                t = self.driver.find_element_by_id('abswrittotal').text
                total += float(t.replace("$", "").replace(',', ''))
            count = f'{count}'
            total += 18.50  # Add processing fee
            tbl_block = (count, cost, total)
        else:
            count = self.driver.find_element_by_id("noticecount").text
            count = int(count.replace('Notice(s):', ''))
            cost = self.driver.find_element_by_id('noticefee').text
            cost = cost.replace('Fee:', '')
            total = self.driver.find_element_by_id('totaldue').text
            total = float(total.replace('$', '').replace(',', ''))
            total += 18.50  # Add processing fee
            tbl_block = (count, cost, total)
        if self.printing:
            self.driver.execute_script("window.print();")

        return self._load_raw_invoice(requester, complex_name, invoice_number, company_id, tbl_block)

    def login(self):
        """Enters in the login details at the ops merchant login page."""
        self.driver.get("https://merchant.opstechnology.com/index.php")
        username = self.driver.find_element_by_name("uid")
        password = self.driver.find_element_by_name('pwd')

        username.clear()
        username.send_keys("brucemo")
        password.clear()
        password.send_keys("kathy1")
        self.driver.find_element_by_id("submit").click()

    @staticmethod
    def save_invoice(invoice_number, code):
        os.rename(folder + f"\\{os.listdir(folder)[0]}", folder + f"\\{invoice_number}.pdf")
        fp = folder + f"\\{invoice_number}.pdf"
        try:
            dest = directory[code]
            shutil.move(fp, dest)
        except KeyError:
            print("Unable to save the file. Path was not found in the directory.")
            print(f"CODE: {code} INVOICE: {invoice_number}")
            os.remove(fp)
        except Exception as e:
            print(e)
            print("Invoice has already been saved.")
            os.remove(fp)

    def _load_raw_invoice(self, requester, complex_name, inv_num, company_id, table):
        """Scrapes and loads the invoice data into a namedtuple for easy access

        """
        Invoice = namedtuple('Invoice', 'company property id ordered_by name lines')
        if self.processing_type == 0:
            lines = {
                "1": {"Count": table[0], "Cost": table[1]},
                "2": {"Count": table[2], "Cost": table[3]},
                "3": {"Count": table[4], "Cost": table[5]},
                "4": {"Count": table[6], "Cost": table[7]}
            }
        else:
            lines = {"Count": table[0], "Cost": table[1], "Total": table[2]}
        cpy, pty = customer_table[company_id]

        return Invoice(company=cpy, property=pty, id=inv_num, ordered_by=requester, name=complex_name, lines=lines)

    def add_lines(self, lines):

        if self.processing_type == 0:
            for tenants, fields in lines.items():
                if int(fields['Count']) == 0:
                    continue

                quantity = self.driver.find_element_by_name("dfield_quantity")
                description = self.driver.find_element_by_name("dfield_description")
                price = self.driver.find_element_by_name("dfield_price")
                quantity.clear()
                description.clear()
                price.clear()

                quantity.send_keys(f"{fields['Count']}")
                description.send_keys(f"Failure to pay rent. Tenants ({tenants})")
                price.send_keys(f"{fields['Cost']}")

                self.driver.find_element_by_id("btn_addline").click()
                time.sleep(0.1)
        elif self.processing_type == 1:
            quantity = self.driver.find_element_by_name("dfield_quantity")
            description = self.driver.find_element_by_name("dfield_description")
            price = self.driver.find_element_by_name("dfield_price")
            quantity.clear()
            description.clear()
            price.clear()
            quantity.send_keys(f"{lines['Count']}")
            description.send_keys(f"For Warrants of Restitution")
            price.send_keys(f"{lines['Cost'].replace('$', '')}")
            self.driver.find_element_by_id("btn_addline").click()
        else:
            quantity = self.driver.find_element_by_name("dfield_quantity")
            description = self.driver.find_element_by_name("dfield_description")
            price = self.driver.find_element_by_name("dfield_price")
            quantity.clear()
            description.clear()
            price.clear()
            quantity.send_keys(f"{lines['Count']}")
            description.send_keys(f"Notice of Intent to File")
            price.send_keys(f"{lines['Cost'].replace('$', '')}")
            self.driver.find_element_by_id("btn_addline").click()

        if county_name != "PRINCE GEORGE'S COUNTY":
            quantity = self.driver.find_element_by_name("dfield_quantity")
            description = self.driver.find_element_by_name("dfield_description")
            price = self.driver.find_element_by_name("dfield_price")
            quantity.clear()
            description.clear()
            price.clear()
            quantity.send_keys(f"{tally}")
            # description.send_keys(f"CFPB Fee.")
            # price.send_keys(f"4.00")
            # self.driver.find_element_by_id("btn_addline").click()

        quantity = self.driver.find_element_by_name("dfield_quantity")
        description = self.driver.find_element_by_name("dfield_description")
        price = self.driver.find_element_by_name("dfield_price")
        quantity.clear()
        description.clear()
        price.clear()
        quantity.send_keys(f"1")
        description.send_keys(f"Processing Fee")
        price.send_keys(f"18.50")

        self.driver.find_element_by_id("btn_addline").click()

        return

    def create_invoice(self, Invoice):
        """Creates a new invoice and inputs the company details.

        Details are generated from the invoice created in rbnet.
        """
        try:
            WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, "facebox")))
        except Exception as e:
            print("Closing Ads")
            print(e)
        else:
            self.driver.find_element_by_xpath('//div[@class="close"]').click()
            WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//li[@id="id12"]/a')))

        element = self.driver.find_element_by_xpath('//*[@id="id12"]/a')
        ActionChains(self.driver).move_to_element(element).click(element).perform()
        self.driver.find_element_by_xpath(
            "//*[contains(text(), '                             New Invoice                         ')]").click()

        select = Select(self.driver.find_element_by_id(f"ddl_company"))
        select.select_by_value(str(Invoice.company))

        # Wait for the Account dropdown to populate with the properties associated with the company
        wait = WebDriverWait(self.driver, 20)
        wait.until(EC.presence_of_element_located(
            (By.ID, "ddl_property")))
        select = Select(self.driver.find_element_by_id("ddl_property"))
        # driver.find_element_by_xpath(f"//select[@id='ddl_property']/option[text()='{Invoice.property}']").click()
        select.select_by_value(str(Invoice.property))
        inv_date = self.driver.find_element_by_name("hfield_invoicedate")
        inv_date.clear()
        inv_date.send_keys(self.invoice_date)

        inv_number = self.driver.find_element_by_name("hfield_invoicenumber")
        inv_number.clear()
        inv_number.send_keys(f"{Invoice.id}")

        order = self.driver.find_element_by_name("hfield_ordernumber")
        order.clear()
        order.send_keys(f"{Invoice.id}")

        order_by = self.driver.find_element_by_name("hfield_orderedby")
        order_by.clear()
        order_by.send_keys(f"{Invoice.ordered_by}")
        self.add_lines(Invoice.lines)
        return

    def main(self, code):
        Invoice = self.login_rbnet(code)
        self.login()
        self.create_invoice(Invoice)
        if self.saving:
            self.save_invoice(Invoice.id, code)
        return 200, Invoice

    def run(self, code):
        try:
            return self.main(code)
        except Exception as e:
            print("-" * 60)
            err = traceback.format_exc()
            print('-' * 60)
            return 404, err


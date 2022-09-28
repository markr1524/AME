import os
from selenium import webdriver
from selenium.webdriver import DesiredCapabilities
from selenium.webdriver.chrome.options import Options


download_dir = os.getcwd() + "\\invoices"

profile = {"plugins.always_open_pdf_externally": True,
           "download.default_directory": download_dir, "download.extensions_to_open": "applications/pdf",
           "printing.print_header_footer": False, "printing.enabled": True}


class ChromeDriver:

    def __init__(self, debug=False):
        self.debug = debug
        self.options = Options()
        self.capabilities = DesiredCapabilities.CHROME.copy()
        self.capabilities['acceptSslCerts'] = True
        self.capabilities['acceptInsecureCerts'] = True
        self.driver_path = "chromedriver/chromedriver.exe"
        self.options.add_experimental_option("prefs", profile)
        self.options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.options.add_argument("--kiosk-printing")
        self.options.add_argument("--use-system-default-printer")
        self.driver = None

    def open_browser(self):
        self.driver = webdriver.Chrome(executable_path=self.driver_path, chrome_options=self.options,
                                       desired_capabilities=self.capabilities)

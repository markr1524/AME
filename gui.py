import os
import shutil
import slate3k as slate
from tkinter import Tk, messagebox, Label, HORIZONTAL, Button, Entry, IntVar, BooleanVar, Radiobutton, Checkbutton, CENTER, StringVar, OptionMenu
from tkinter.ttk import Separator, Notebook, OptionMenu
from tkinter import ttk
from vc import VendorCafe
from ops import OpsMerchant, customer_table
from login import ChromeDriver
from reminders import Reminders
from selenium.common.exceptions import SessionNotCreatedException
from webdrivermanager import ChromeDriverManager


property_options = ['', '', '', '',
                    '', '', '', '', '',
                    '', '', '',
                    '', '', '']


code_lookup = {"BDS": {"pty": "1305 Dock Street", "mgmt": "Bozzuto"},
               "BLL": {"pty": "Allure Apollo", "mgmt": "Bozzuto"},
               "BAH": {"pty": "Anthem House", "mgmt": "Bozzuto"},
               "ABC": {"pty": "Arbors at Baltimore Crossroads", "mgmt": "Bozzuto"},
               "AAP": {"pty": "Arundel Preserve", "mgmt": "Bozzuto"},
               "BAA": {"pty": "Aspire Apollo", "mgmt": "Bozzuto"},
               "WOS": {"pty": "Azure Oxford Square", "mgmt": "Bozzuto"},
               "BCC": {"pty": "Cadence at Crown", "mgmt": "Bozzuto"},
               "BCA": {"pty": "Central Apartments", "mgmt": "Bozzuto"},
               "CPR": {"pty": "Concord Park at Russett", "mgmt": "Bozzuto"},
               "BCW": {"pty": "Crosswinds @ A.T.C", "mgmt": "Bozzuto"},
               "BEE": {"pty": "Enclave at Emerson Apartments", "mgmt": "Bozzuto"},
               "BEX": {"pty": "Excalibur", "mgmt": "Bozzuto"},
               "BFA": {"pty": "Fenestra", "mgmt": "Bozzuto"},
               "FEN": {"pty": "Fenwick", "mgmt": "Bozzuto"},
               "BTF": {"pty": "Fitzgerald", "mgmt": "Bozzuto"},
               "FAY": {"pty": "Flats 170 at Academy Yard", "mgmt": "Bozzuto"},
               "BF8": {"pty": "Flats 8300", "mgmt": "Bozzuto"},
               "FBA": {"pty": "Flats at Bethesda Avenue", "mgmt": "Bozzuto"},
               "GWL": {"pty": "Gables at Waters Landing", "mgmt": "Bozzuto"},
               "GTC": {"pty": "Gramercy at Town Center", "mgmt": "Bozzuto"},
               "BGP": {"pty": "Greenwich Place", "mgmt": "Bozzuto"},
               "HCA": {"pty": "Hidden Creek", "mgmt": "Bozzuto"},
               "BHG": {"pty": "Hunters Glen", "mgmt": "Bozzuto"},
               "LAR": {"pty": "Lakehouse Residences", "mgmt": "Bozzuto"},
               "BLH": {"pty": "Liberty Harbor East", "mgmt": "Bozzuto"},
               "BLOL": {"pty": "Luminary One at Light", "mgmt": "Bozzuto"},
               "MSA": {"pty": "Mallory Square", "mgmt": "Bozzuto"},
               "BMB": {"pty": "Mariner Bay", "mgmt": "Bozzuto"},
               "MHR": {"pty": "McHenry Row", "mgmt": "Bozzuto"},
               "BMS": {"pty": "Milestone Apartments", "mgmt": "Bozzuto"},
               "BMK": {"pty": "Millstone at Kingsview", "mgmt": "Bozzuto"},
               "BMA": {"pty": "Monterey", "mgmt": "Bozzuto"},
               "BMMCH": {"pty": "Monarch", "mgmt": "Bozzuto"},
               "MCP": {"pty": "Monument Village", "mgmt": "Bozzuto"},
               "NBM": {"pty": "North Bethesda Market", "mgmt": "Bozzuto"},
               "BTP": {"pty": "Pinnacle at Town Center", "mgmt": "Bozzuto"},
               "BPH": {"pty": "Pilot House at Riverdale", "mgmt": "Bozzuto"},
               "BMP": {"pty": "Metropointe", "mgmt": "Bozzuto"},
               "JRR": {"pty": "Red Run", "mgmt": "Bozzuto"},
               "JRO": {"pty": "Riverstone at Owings Mills", "mgmt": "Bozzuto"},
               "RH": {"pty": "Rolling Hills", "mgmt": "Bozzuto"},
               "BSM": {"pty": "Solaire", "mgmt": "Bozzuto"},
               "BPR": {"pty": "Spinnaker Bay & Promenade", "mgmt": "Bozzuto"},
               "SPI": {"pty": "St. Paul Senior Living", "mgmt": "Bozzuto"},
               "STP": {"pty": "Stone Point", "mgmt": "Bozzuto"},
               "BSC": {"pty": "Strathmore", "mgmt": "Bozzuto"},
               "BMC": {"pty": "The Beacon @ Waugh Chapel", "mgmt": "Bozzuto"},
               "BERK": {"pty": "The Berkleigh", "mgmt": "Bozzuto"},
               "BHP": {"pty": "The Bowen (frmly Harmony Place)", "mgmt": "Bozzuto"},
               "BCD": {"pty": "The Courts of Devon", "mgmt": "Bozzuto"},
               "TEB": {"pty": "The Equitable Building", "mgmt": "Bozzuto"},
               "ESP": {"pty": "The Esplanade", "mgmt": "Bozzuto"},
               "RMTE": {"pty": "The Estates", "mgmt": "Bozzuto"},
               "BTG": {"pty": "The Glen", "mgmt": "Bozzuto"},
               "GIL": {"pty": "The Guilford", "mgmt": "Bozzuto"},
               "BTL": {"pty": "The Lindley", "mgmt": "Bozzuto"},
               "LSO": {"pty": "The Lodge at Seven Oaks", "mgmt": "Bozzuto"},
               "MET": {"pty": "The Metropolitan", "mgmt": "Bozzuto"},
               "BTM": {"pty": "The Morgan", "mgmt": "Bozzuto"},
               "BQT": {"pty": "The Quarters at Towson Town Ctr", "mgmt": "Bozzuto"},
               "THV": {"pty": "The Townes at Harvest View", "mgmt": "Bozzuto"},
               "BVA": {"pty": "The Vine Apartments", "mgmt": "Bozzuto"},
               "BTW": {"pty": "The Whitney", "mgmt": "Bozzuto"},
               "BTC": {"pty": "Timberlawn Crescent", "mgmt": "Bozzuto"},
               "BTWD": {"pty": "Towson Woods", "mgmt": "Bozzuto"},
               "UW": {"pty": "Union Wharf", "mgmt": "Bozzuto"},
               "BWA": {"pty": "Winthrop Apartments", "mgmt": "Bozzuto"},
               "BWG": {"pty": "Woodfall Greens", "mgmt": "Bozzuto"},
               "FOF": {"pty": "1405 Point", "mgmt": "Kettler"},
               "TTF": {"pty": "225 North Calvert", "mgmt": "Kettler"},
               "AO": {"pty": "Avondale Overlook", "mgmt": "Kettler"},
               "FP": {"pty": "Fireside Park", "mgmt": "Kettler"},
               "FS": {"pty": "Forrest Street", "mgmt": "Kettler"},
               "MFL": {"pty": "M. Flats Downtown Columbia", "mgmt": "Kettler"},
               "TEN": {"pty": "Ten.M", "mgmt": "Kettler"},
               "BV": {"pty": "The Broadview", "mgmt": "Kettler"},
               "KC": {"pty": "The Fields of Germantown", "mgmt": "Kettler"},
               "FC": {"pty": "The Fields of Rockville", "mgmt": "Kettler"},
               "WC": {"pty": "The Fields of Silver Spring", "mgmt": "Kettler"},
               "TG": {"pty": "The Gunther", "mgmt": "Kettler"},
               "TL": {"pty": "The Lenore", "mgmt": "Kettler"},
               "MDC": {"pty": "The Metropolitan Downtown Columbia", "mgmt": "Kettler"},
               "BTZ": {"pty": "Zenith", "mgmt": "Bozzuto"},
               "RW": {"pty": "Redwood", "mgmt": "Kettler"},
               "BAC": {"pty": "Ascend Apollo", "mgmt": "Bozzuto"}
               }


class Interface(ChromeDriver):

    def __init__(self):
        
        super().__init__(debug=True)
        self.controller = Tk()
        self.instance = None
        self.flag = False

        self.controller.title('AME Automation Made Easy     Ver 1.5.5')
        self.controller.geometry("475x375")
        self.controller.iconbitmap(os.getcwd() + "\\icons\\icon.ico")
        self.controller.protocol("WM_DELETE_WINDOW", self._delete_window)

        self.tab_controller = Notebook(self.controller)

        self.tab1 = ttk.Frame(self.tab_controller)
        self.tab_controller.add(self.tab1, text="Invoices")
        self.tab_controller.pack(expand=1, fill='both')

        self.tab2 = ttk.Frame(self.tab_controller)
        self.tab_controller.add(self.tab2, text="Reminders")
        self.tab_controller.pack(expand=1, fill='both')

        self.tab3 = ttk.Frame(self.tab_controller)
        self.tab_controller.add(self.tab3, text="Legal Fees")
        self.tab_controller.pack(expand=1, fill='both')

        # Labels
        self.l1 = Label(self.tab1, text='Company ID:').grid(row=0, column=0, sticky='W', padx=20)
        self.l2 = Label(self.tab1, text='Status:').grid(row=4, column=0)
        self.status = Label(self.tab1, text="")
        self.status.grid(row=4, column=1)
        self.status2 = Label(self.tab2, text="")
        self.status2.grid(row=2, column=2)
        self.l3 = Label(self.tab3, text='Invoice Number:').grid(row=0, column=0, sticky='W', padx=20)

        Separator(self.tab1, orient=HORIZONTAL).grid(row=5, columnspan=5, sticky='ew')
        # Output Section
        self.account_lb = Label(self.tab1, text="")
        self.account_hdr = Label(self.tab1, text="", font='Helvetica 16 bold')
        self.account_hdr.grid(row=6, column=0)
        self.account_lb.grid(row=7, column=0)

        self.invoice_lb = Label(self.tab1, text="")
        self.invoice_hdr = Label(self.tab1, text="", font='Helvetica 15 bold')
        self.invoice_hdr.grid(row=6, column=1)
        self.invoice_lb.grid(row=7, column=1)

        self.total_lb = Label(self.tab1, text="")
        self.total_hdr = Label(self.tab1, text="", font='Helvetica 15 bold')
        self.total_hdr.grid(row=6, column=2)
        self.total_lb.grid(row=7, column=2)

        self.ordered_by_lb = Label(self.tab1, text="")
        self.ordered_by_hdr = Label(self.tab1, text="", font='Helvetica 15 bold')
        self.ordered_by_hdr.grid(row=8, column=1)
        self.ordered_by_lb.grid(row=9, column=1)

        self.tenants_1_lb = Label(self.tab1, text="")
        self.tenants_1_hdr = Label(self.tab1, text="", font='Helvetica 15 bold')
        self.tenants_1_hdr.grid(row=10, column=0)
        self.tenants_1_lb.grid(row=11, column=0)

        self.tenants_2_lb = Label(self.tab1, text="")
        self.tenants_2_hdr = Label(self.tab1, text="", font='Helvetica 15 bold')
        self.tenants_2_hdr.grid(row=10, column=1)
        self.tenants_2_lb.grid(row=11, column=1)

        self.tenants_3_lb = Label(self.tab1, text="")
        self.tenants_3_hdr = Label(self.tab1, text="", font='Helvetica 15 bold')
        self.tenants_3_hdr.grid(row=12, column=0)
        self.tenants_3_lb.grid(row=13, column=0)

        self.tenants_4_lb = Label(self.tab1, text="")
        self.tenants_4_hdr = Label(self.tab1, text="", font='Helvetica 15 bold')
        self.tenants_4_hdr.grid(row=12, column=1)
        self.tenants_4_lb.grid(row=13, column=1)

        # Buttons
        self.button = Button(self.tab1, text='Process Invoice', width=18, command=self.button_process)
        self.button2 = Button(self.tab1, text="Close Browser", width=18, command=self.close_browser)
        self.button3 = Button(self.tab1, text="Submit & Print", width=18, command=self.submit_and_print)
        self.button4 = Button(self.tab2, text="Send Notifications", width=18, command=self.submit_notifications)
        self.button5 = Button(self.tab3, text="Process Legal Fees", width=20, command=self.add_legal_fees)
        self.button6 = Button(self.tab3, text="Close Browser", width=18, command=self.close_browser)

        self.button.grid(row=3, column=1)
        self.button2.grid(row=3, column=0)

        self.button3.grid(row=3, column=2)
        self.button4.grid(row=0, column=2)
        self.button3.configure(state="disabled")
        self.button2.configure(state="disabled")
        self.button5.grid(row=1, column=0)
        self.button6.grid(row=2, column=0)

        # Radio Buttons
        self.target = IntVar()
        self.process_type = IntVar()

        self.vc_button = Radiobutton(self.tab1, text='Vendor Cafe', variable=self.target, value=2).grid(row=2, column=0, sticky="W", padx=20)
        self.ops_button = Radiobutton(self.tab1, text='Ops Merchant', variable=self.target, value=1).grid(row=2, column=1, sticky="E", padx=20)
        self.normal_button = Radiobutton(self.tab1, text="Normal", variable=self.target, value=3)
        self.suits_button = Radiobutton(self.tab1, text="Suits", variable=self.process_type, value=0)
        self.writs_button = Radiobutton(self.tab1, text="Writs", variable=self.process_type, value=1)
        self.notices_button = Radiobutton(self.tab1, text="Notices", variable=self.process_type, value=2)

        self.suits_button.grid(row=1, column=0, sticky="W", padx=20)
        self.writs_button.grid(row=1, column=1)
        self.notices_button.grid(row=1, column=2)

        self.normal_button.configure(state="disabled")
        self.normal_button.grid(row=2, column=2, sticky="E", padx=28)

        # Checkbox
        self.chk_target = BooleanVar()
        self.chk_target2 = BooleanVar()
        self.print_box = Checkbutton(self.tab1, text='Print', variable=self.chk_target).grid(row=0, column=2)
        self.save_box = Checkbutton(self.tab1, text='Save', variable=self.chk_target2).grid(row=0, column=3)

        self.aa_val = BooleanVar()
        self.ann_val = BooleanVar()
        self.hc_val = BooleanVar()
        self.tb_val = BooleanVar()
        self.eb_val = BooleanVar()
        self.cb_val = BooleanVar()
        self.bc_val = BooleanVar()
        self.mc_val = BooleanVar()
        self.har_val = BooleanVar()
        self.cc_val = BooleanVar()
        self.fc_val = BooleanVar()
        self.pg_val = BooleanVar()
        self.all_val = BooleanVar()

        self.aa = Checkbutton(self.tab2, text='Annapolis', variable=self.aa_val)
        self.aa.grid(row=0, column=0, sticky='W', padx=20)
        self.ann = Checkbutton(self.tab2, text='Anne Arundel', variable=self.ann_val)
        self.ann.grid(row=1, column=0, sticky='W', padx=20)
        self.hc = Checkbutton(self.tab2, text='Howard', variable=self.hc_val)
        self.hc.grid(row=2, column=0, sticky='W', padx=20)
        self.tb = Checkbutton(self.tab2, text='Towson', variable=self.tb_val)
        self.tb.grid(row=3, column=0, sticky='W', padx=20)
        self.eb = Checkbutton(self.tab2, text='Essex', variable=self.eb_val)
        self.eb.grid(row=4, column=0, sticky='W', padx=20)
        self.cb = Checkbutton(self.tab2, text='Catonsville', variable=self.cb_val)
        self.cb.grid(row=5, column=0, sticky='W', padx=20)
        self.bc = Checkbutton(self.tab2, text='Baltimore City', variable=self.bc_val)
        self.bc.grid(row=6, column=0, sticky='W', padx=20)
        self.mc = Checkbutton(self.tab2, text='Montgomery', variable=self.mc_val)
        self.mc.grid(row=7, column=0, sticky='W', padx=20)
        self.har = Checkbutton(self.tab2, text='Harford', variable=self.har_val)
        self.har.grid(row=8, column=0, sticky='W', padx=20)
        self.cc = Checkbutton(self.tab2, text='Carroll', variable=self.cc_val)
        self.cc.grid(row=9, column=0, sticky='W', padx=20)
        self.fc = Checkbutton(self.tab2, text='Frederick', variable=self.fc_val)
        self.fc.grid(row=10, column=0, sticky='W', padx=20)
        self.pg = Checkbutton(self.tab2, text='Prince George', variable=self.pg_val)
        self.pg.grid(row=11, column=0, sticky='W', padx=20)
        self.chk_all = Checkbutton(self.tab2, text="Select All", variable=self.all_val, command=self.select_all)
        self.chk_all.grid(row=0, column=1)

        # Text Entry
        self.batch_var = StringVar(self.tab2)
        options = ('Early', 'Late', 'NSF')
        self.batch_var.set(options[0])

        self.e1 = Entry(self.tab1)
        self.e1.grid(row=0, column=1)
        self.e2 = Entry(self.tab3)
        self.e2.grid(row=0, column=1)
        self.e3 = OptionMenu(self.tab2, self.batch_var, options[0], *options)

        self.l6 = Label(self.tab2, text='Batch:').grid(row=1, column=1, sticky='W', padx=20)
        self.e3.grid(row=1, column=2)

    def select_all(self):
        for switch in (self.aa_val, self.ann_val, self.hc_val, self.tb_val, self.eb_val, self.cb_val, self.bc_val,
                       self.mc_val, self.har_val, self.cc_val, self.fc_val):
            if self.all_val.get():
                switch.set(1)
            else:
                switch.set(0)

    def reset_screen(self):
        self.tenants_4_lb.configure(text="")
        self.tenants_3_lb.configure(text="")
        self.tenants_2_lb.configure(text="")
        self.tenants_1_lb.configure(text="")
        self.ordered_by_lb.configure(text="")
        self.invoice_lb.configure(text="")
        self.account_lb.configure(text="")
        self.total_lb.configure(text="")
        self.account_hdr.configure(text="")
        self.tenants_1_hdr.configure(text="")
        self.tenants_2_hdr.configure(text="")
        self.tenants_3_hdr.configure(text="")
        self.tenants_4_hdr.configure(text="")
        self.total_hdr.configure(text="")
        self.invoice_hdr.configure(text="")
        self.ordered_by_hdr.configure(text="")
        self.status.configure(text="")

    def display_screen_ops(self, info, process):
        self.account_hdr.configure(text="Property")
        self.total_hdr.configure(text="Total")
        self.invoice_hdr.configure(text="Invoice #")

        self.account_lb.configure(text=f'{info.name}')
        self.invoice_lb.configure(text=f'{info.id}')

        if process == 0:
            self.tenants_1_lb.configure(text=f"Count: {info.lines['1']['Count']}  Cost: {info.lines['1']['Cost']}")
            self.tenants_2_lb.configure(text=f"Count: {info.lines['2']['Count']}  Cost: {info.lines['2']['Cost']}")
            self.tenants_3_lb.configure(text=f"Count: {info.lines['3']['Count']}  Cost: {info.lines['3']['Cost']}")
            self.tenants_4_lb.configure(text=f"Count: {info.lines['4']['Count']}  Cost: {info.lines['4']['Cost']}")
            self.tenants_1_hdr.configure(text="Tenants (1)")
            self.tenants_2_hdr.configure(text="Tenants (2)")
            self.tenants_3_hdr.configure(text="Tenants (3)")
            self.tenants_4_hdr.configure(text="Tenants (4)")
            self.ordered_by_hdr.configure(text="Ordered By")
            self.ordered_by_lb.configure(text=f'{info.ordered_by}')
            total = sum(float(value['Cost']) * int(value['Count']) for key, value in info.lines.items()) + 18.50
        else:
            total = info.lines['Total']
        self.total_lb.configure(text=f'${total:,.2f}')

    @staticmethod
    def popup_alert(main_window, message):
        popup = Tk()
        popup.title("ERROR MESSAGE TRACEBACK")
        popup_label = Label(popup, compound=CENTER, text=message)
        popup_label.pack(side='top')
        popup_button = Button(popup, text="Close", command=popup.destroy)
        popup_button.pack()
        x = main_window.winfo_rootx()
        y = main_window.winfo_rooty()
        height = main_window.winfo_height()
        geometry = "+%d+%d" % (x, y + height)
        popup.wm_geometry(geometry)
        popup.mainloop()

    def display_screen_vc(self, pty, invoice, total):
        self.invoice_hdr.configure(text="Invoice #")
        self.invoice_lb.configure(text=invoice)
        self.account_hdr.configure(text="Property")
        self.account_lb.configure(text=pty)
        self.total_hdr.configure(text="Total")
        self.total_lb.configure(text=total)

    def gather_selections(self):
        reminder_selections = []
        for switch, choice in ((self.aa_val, self.aa), (self.ann_val, self.ann), (self.hc_val, self.hc),
                               (self.tb_val, self.tb), (self.eb_val, self.eb), (self.cb_val, self.cb),
                               (self.bc_val, self.bc), (self.mc_val, self.mc), (self.har_val, self.har),
                               (self.cc_val, self.cc), (self.fc_val, self.fc), (self.pg_val, self.pg)):
            if switch.get():
                reminder_selections.append(choice.cget("text"))
        return reminder_selections

    def submit_notifications(self):
        try:
            selections = self.gather_selections()
            if not selections:
                return
            if self.driver is None:
                super().open_browser()
            rem = Reminders(self.driver, selections, self.batch_var.get())
            rem.login_rbnet()
            self.status2.configure(text="REMINDERS SENT!")
        except Exception as e:
            print(e)

    def add_legal_fees(self):
        entry = self.e2.get()
        if entry == "":
            return
        self.button5.configure(state="disabled")
        if self.driver is None:
            super().open_browser()
        self.instance = 'vc'

        vendor = VendorCafe(self.driver, processing_type=0)
        _type, code, _ = entry.split('-')
        property_name, mgmt = code_lookup[code].values()

        path = f'Z:\\Properties\\{mgmt}\\{property_name}\\Invoices\\{entry}.pdf'
        total = self.extract_pdf_total(path)
        status_code, pty, invoice_num, total = vendor.atty_run(entry, total, property_name, mgmt, path)

        self.button5.configure(state="normal")
        if status_code == 200:
            self.status.configure(text='Success!', foreground='green')
            self.display_screen_vc(pty, invoice_num, total)
        else:
            self.status.configure(foreground='red', text="Failure")
            self.popup_alert(self.controller, pty)
        pass

    @staticmethod
    def extract_pdf_total(pdf_path):
        print(pdf_path)
        with open(pdf_path, 'rb') as f:
            extracted_text = slate.PDF(f)
        fn = extracted_text[0].split('\n')
        total = fn[fn.index('Total') + 2]
        return total

    def button_process(self):
        self.reset_screen()
        choice = self.target.get()
        entry = self.e1.get()
        process = self.process_type.get()
        if entry == "":
            return
        try:
            if choice == 1:
                if len(entry.split('-')) > 1:
                    entry, location = entry.split('-')
                    location = int(location) + 1
                else:
                    location = 2
                if entry not in customer_table:
                    self.status.configure(text=f"{entry} is not a valid code for Ops Merchant", foreground='red')
                    return
                if self.driver is None:
                    super().open_browser()
                self.instance = 'ops'
                op = OpsMerchant(self.driver, printing=self.chk_target.get(), processing_type=process,
                                 saving=self.chk_target2.get(), inv_location=location)

                status_code, info = op.run(entry)
                self.button2.configure(state="normal")
                if status_code == 200:
                    self.status.configure(text='Success!', foreground='green')
                    self.display_screen_ops(info, process)

                else:
                    self.status.configure(foreground='red', text="Failure!")
                    err, _ = info
                    self.popup_alert(self.controller, err)

            elif choice == 2:
                if self.driver is None:
                    super().open_browser()
                self.instance = 'vc'
                if len(entry.split('-')) > 1:
                    entry, location = entry.split('-')
                    location = int(location) + 1
                else:
                    location = 2
                vendor = VendorCafe(self.driver, processing_type=process, printing=self.chk_target.get(),
                                    saving=self.chk_target2.get(), invoice_loc=location)
                status_code, pty, invoice_num, total = vendor.run(entry)
                self.button2.configure(state="normal")
                if status_code == 200:
                    self.status.configure(text='Success!', foreground='green')
                    self.display_screen_vc(pty, invoice_num, total)
                else:
                    self.status.configure(foreground='red', text="Failure")
                    self.popup_alert(self.controller, pty)
            else:
                pass
        except SessionNotCreatedException as e:
            if "This version of ChromeDriver only supports Chrome version" in e.msg:
                chrome = ChromeDriverManager(download_root="chromedriver/chromedriver")
                chrome.download_and_install()
                os.remove(f"chromedriver/chromedriver.exe")
                for dirpath, dirs, files in os.walk('chromedriver'):
                    for filename in files:
                        if filename == 'chromedriver.exe':
                            shutil.move(f"{os.path.join(dirpath)}/{filename}", "chromedriver")
                shutil.rmtree("chromedriver/chromedriver")
                if self.flag is False:
                    self.flag = True
                    self.button_process()
                else:
                    self.popup_alert(self.controller, e.msg)
            self.flag = False

    def submit_and_print(self):
        if self.instance is None:
            return
        result = messagebox.askyesno("Invoice Submission", "Are you sure you want to submit and print the invoice?")
        if result is False:
            return

        if self.instance == 'vc':
            self.driver.find_element_by_id('btnSave0').click()
            self.driver.execute_script("window.print();")
        elif self.instance == 'ops':
            self.driver.find_element_by_id("btn_saveinvoice").click()
            self.driver.find_element_by_xpath("//*[contains(text(), 'Print')]").click()
            self.driver.find_element_by_xpath("//*[contains(text(), 'Print Now')]").click()
        else:
            pass

    def close_browser(self):
        self.driver.quit()
        self.instance = None
        self.reset_screen()
        self.driver = None
        self.button2.configure(state="disabled")

    def _delete_window(self):
        if self.driver is None:
            self.controller.destroy()
        else:
            try:
                self.driver.quit()
                self.controller.destroy()
            except Exception as e:
                print(e)


# Window setup
if __name__ == '__main__':
    ame = Interface()
    ame.controller.mainloop()

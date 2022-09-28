# AME
AME, Automation Made Easy, is an internal tool developed at Residency Bureau to help automate third-party invoicing. It utilizes Selenium to control the browser. It also includes automatically updating the chrome webdriver to match your current chrome version.

Just type in the agent code and select the appropriate third party invoicing company. If you need to retrieve older invoices use the following format:

`gots` *current invoice*

`gots-2` *previous invoice*

`gots-3` *invoice sent before the last two*

and so on.



## Compiling AME for Distribution
Compiling AME requires pyinstaller. Download this library via pip.

`pip install pyinstaller`

Open command prompt or PS and navigate to the root directory of the AME package.
Compiling AME is a two step process. First,  run the following command:

`pyinstaller gui.py --onefile --windowed --name="AME" --icon="icons/icon.ico"`

This will generate a **.spec** file. This specification file tells pyinstaller how to build your script.

Next, run this command: `pyinstaller ame.spec`. This will then compile the script. It will create a **dist** folder where you will find the executable. 

## GUI Component
The GUI component uses TCL/TK. It is very rudimentary, but it is cross platform. You will find all of the GUI components in gui.py.

## Additional Considerations
Please take note that for both `ops.py` and `vc.py` there will be hard coded passwords. They both contain the passwords to login to the legacy website. vc.py will need to be updated every so often, because vendor cafe forces us to update the passwords regularly. 

The goal was to ultimately run this headless and add some additional tests for fall backs, such that it could be integrated into the web platform. Since vendor cafe and ops merchant are sectioned into their own module, you do not need to strip the GUI component. All of the scraping logic are in those files.

Saving was being done by hard linking the paths based on agent code. Some of these may be missing or incomplete. They also do not work unless they are utilized in the office.

Legal fees are not 100% complete, and they currently only work for vendor cafe.

**NOTE FOR UPDATING** Version numbers are tracked inside the gui.py file. The version should be bumped before compiling and distributing any update.
Format is: **Major**.**Minor**.**Hotfix**


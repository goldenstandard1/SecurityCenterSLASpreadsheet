# SecurityCenterSLASpreadsheet
Python Code to Create a Spreadsheet of Tenable Security Center vulnerabilities that are past SLA
a.	A computer that's connected to the internal network.
b.	AP Credentials for Tenable Security Center.
c.	Python 3.x installed.  (Preferably Anaconda)
d.	Python 3.x added to your system environment path. 
e.	PIP packages requests and xlsxwriter installed. (These are part of the standard Anaconda library).
2.	Using for the first time:
a.	Copy the script to a folder on your local computer (Do not copy to OneDrive).


d.	Create a folder containing a CSV file for each system you’d like to be included in the spreadsheet. Each CSV file should contain a list of the IP address associated with that system. 
i.	Insert the file path for the folder between the quotes on line 15 of the script (shown below) replacing the file path there:
Note: If you enter the wrong file path, you’ll get an error similar to the following: 
(If you get this error while running the script, correct the file path in IDLE.)
e.	Create an output folder (where the script will create the output spreadsheet, Mastersheet.xlsx).
i.	 Insert the file path between the quotes on line 18 (shown) replacing the file path there: 
Note: If you enter the wrong file path, you’ll get an error similar to the following:
(If you get this error while running the script, correct the file path in IDLE.) 
3.	Regular Usage:
a.	Before running the script, make sure that the spreadsheet in the output file path is not open. If it is open, you’ll get an error similar to the following:
If you get this error, close the spreadsheet, exit PowerShell and start over (from 3a).
b.	Open Windows PowerShell. (Shortcut: Press  ⊞  + X and then i).
c.	Type py, press space, then drag the python script file into the PowerShell window (the PowerShell window should look similar to shown) and press Enter.
d.	Enter your AP account password and press Enter. For security reasons, you will not see anything appear in the command prompt as you type your password. 
i.	 Make sure to enter the correct password. If you enter an incorrect password, the following error will appear: 

If you get this error, close the spreadsheet, exit PowerShell and start over from 3a.

e.	Wait 10-20 minutes for the script to complete. After each HTTP request is made, the following warning will appear: 

This warning is the result of Tenable Security Center using a self-signed certificate, which is a standard practice for requests that only traverse internal networks. No action needs to be taken. 
f.	If any errors occur, check to see if they are one of the ones mentioned in this SOP, and edit your script accordingly. If the error does not match one of the those mentioned above, contact the administrator for assistance. 
g.	If no errors occur, you’ll see the following message: 
h.	Close PowerShell. (Shortcut: Press Alt + Space and then c)
i.	Open the Excel workbook that has been created. The workbook will be in the folder that you created in 2d above. 
j.	In Excel, select all of the results and click Sort: 
k.	Sort the results by Severity and First Seen Date:

l.	Select “Sort anything that looks like a number, as a number” and click “OK”:
m.	 Repeat Steps j through l for each system’s worksheet:
n.	This completes the process. 


Restatement Process Automation

Overview

Restatement_Process.py is a Python script designed to automate the retrieval and archiving of email attachments from a shared Outlook mailbox. It also updates an Excel workbook with processing status and timestamps. This tool is ideal for financial operations teams handling daily statement imports and audit trails.

Features

Reads configuration from an .ini file for flexible setup
Connects to a shared Outlook mailbox and folder
Filters emails by business day and cutoff time
Matches emails and attachments using an Excel mapping table
Saves attachments to a structured archive directory (by year/month/day)
Updates Excel workbook with status and timestamps
Logs all actions and errors for traceability

Requirements

Python 3.8+
Windows OS with Microsoft Outlook installed
Required Python packages:

pandas
openpyxl
pytz
pywin32
pythoncom

Configuration

Create a file named Restatement_Process_Config.ini in the same directory as the script. 

Usage

1.) Install dependencies (if not already installed):

  Shellpip install pandas openpyxl pytz pywin32Show more lines

2.)Configure your .ini file as shown above.

3.)Run the script:

  Shellpython Restatement_Process.pyShow more lines

4.)Check the log file in the daily archive directory for processing details and errors.

How It Works

The script calculates business dates (current, previous, and two days prior).
It creates a dated archive folder for saving attachments.
Loads the Excel mapping table and prepares the workbook for updates.
Connects to the specified Outlook mailbox and folder.
Filters emails received after a cutoff time on the previous business day.
Matches emails and attachments to mapping rules from Excel.
Saves matched attachments and updates the Excel file with status.
Cleans up resources and logs all steps.

Logging
A log file named script.log is created in the archive directory for each run, capturing all actions, warnings, and errors.
Troubleshooting

Ensure Outlook is installed and configured on your machine.
The Excel mapping file must contain the required columns: sender, subject, attachment, savename, status.
If you encounter permission or COM errors, try running Python as an administrator.
Check the log file for detailed error messages.

License
This project is intended for internal use. Contact the author for licensing details.

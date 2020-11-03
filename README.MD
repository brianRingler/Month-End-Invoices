***
# Expense Report Tracker 
This is a simple and small project that automatically imports CSV files then reads in the description and expense amount. The users then assign the expense amount to correct account.  When completed an invoice is automatically generated, saved to folder, then emailed to the appropriate individual. 

The project was originally created with VBA sending the emails via Outlook. To test the ability of Excel executing Python files it has been modified to send emails using Outlook but executed with Python. 

The module `run_python` runs the `main.py` file that send the emails with the invoice attached.  The file `outlookSend.py` is called to send the email.

The `main.py` finds the active workbook using `xlwings` which allows the user to freely change the file name without updating Python files. From there the `main.py` file can read the Excel file and locate the file path of invoices.   

***
# Excel File Overview  
 This is a brief overview of the files worksheets and features. It does not cover all of the error checks. 

 ## MASTER
This worksheet is the main file that contains the imported CSV data. The user selects the Import CSV Files button and will select the location of the files to be imported. Its configured to only read three files. If more/less in folder the macro exits. 

The user the assigns expense amounts to accounts with built in validation checks to detect errors.  They then assign a company number that is rolled up into the top section. 

The user can then select `Generate Invoices` and for any company with a dollar amount and invoice will be created and sent.


 ## INVOCIE_MASTER
The invoice master is a template used to generate all invoices. It is cleared and updated for each new invoice. The invoice number increases by one for each new invoice. The due date is the current date plus days due from the master worksheet `unique by company`.

 ## DATA
The data worksheet is a section for the user to set who is sending and receiving the email along with the email details. The invoice number counter is also contained on this sheet. 

 ## FILE_PATHS
This sheet contains the location of the invoices created that will be attached with the email. Along with the path to the Python files and Python executable which is required to run the Python scripts through Excel. 


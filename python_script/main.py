import pandas as pd
from outlookSend import send_email_outlook
import os
import xlwings as xw

# name master Excel file 
file_name_master = xw.books.active.name
# Path to master file which includes name used for pandas
full_path = xw.Book(file_name_master).fullname


# Path to folder with all invoices to be sent
wb = xw.Book(file_name_master)
sht = wb.sheets['FILE_PATHS']
path_invoices = sht.range('F1').value


# ExcelFile class
'''
To facilitate working with multiple sheets from the same file, the ExcelFile class can be used to wrap the file and can be passed into read_excel There will be a performance benefit for reading multiple sheets as the file is read into memory only once.
'''
xlsm = pd.ExcelFile(full_path)

# Passing the path and name of Workbook and the sheet name we want to import
# read id data from the sheet named "DATA"
df = pd.read_excel(xlsm,sheet_name='DATA')

# Loop through the directory containing invoices
# Extract the company name add to a list
# Files should all end with (.xlsx) and check for (.XLSX)
company_names = []
for inv in os.listdir(path_invoices):
    if inv.endswith(".xlsx"):
        idx = inv.find(".xlsx")
        company_names.append(inv[:idx])
    elif inv.endswith(".XLSX"):
        idx = inv.find(".XLSX")
        company_names.append(inv[:idx])

# How many emails will be sent is equal to len of company_names
invoice_count = len(company_names)
emails_sent = 0

# Access each element of company_names list 
for company in company_names:
    # iterate over the df looking for matching company name
    for index, row in df.iterrows():
        if row['company_name'] == company:
            to = row['to_email_address']
            sub = row['subject']
            body = row['body_email']
            path = row['path_file']
            send_email_outlook(to, sub, body, path)
            emails_sent += 1




            
    
            





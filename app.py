#requirements pandas, openpyxl, pypiwin32


import pandas as pd
import win32com.client as win32
import os
from datetime import date
import time

print('starting..')
start = time.time()

TODAY = date.today().strftime('%Y_%m_%d')


# Preparing a function to create emails with attached parts of the data
def Emailer(text, subject, recipient, attachment, filename):
    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    attachment1 = os.getcwd() + f"\\{attachment}"
    mail.Attachments.Add(attachment1)
    #mail.Display(True)
    mail.saveas(os.getcwd()+f'//{filename}.msg')

df = pd.read_excel('TP - test.xlsx', header=2)
dff = df[~df['Carrier'].isna()]
dff = dff[[
'Tour ID',
'Operational comment for route identifier',
'Flow Pilot',
 'Hybrid COFOR',
 'COFOR SELLER',
 'SELLER NAME',
 'COFOR SHIPPER',
 'SHIPPER NAME',
 'SHIPPER \nZIP CODE',
 'SHIPPER\nCITY',
 'SHIPPER COUNTRY',
 'Destination Number (Demand)',
 'Destination Name (Demand)',
 'Destination ZIP (Demand)',
 'Destination City (Demand)',
 'COFOR Recipient',
 'Recipient Name',
 'P (Parts)\nE (Empties)',
'DHEO days before DHRQ (PLE)',
 'DHEO / Hour (PLE)',
 'Truck arrival at supplier DHMD / Hour (PLE)',
 'End of loading DHEF / HEE Hour (PLE)',
 'HEF (end of loading at supplier)',
 'HEE (empties loaded)',
 'Loading dock for empties',
 'Start Hub',
 'Departure at Hub (HXC)',
 'Departure at Hub (HXC) based on PLE',
 'Pick Mon',
 'Pick Tue',
 'Pick Wed',
 'Pick Thu',
 'Pick Fri',
 'Pick Sat',
 'Pick Sun',
 'Unloading Dock',
 'Kind of dock (for unloading)',
 'Delivery code (PLE)',
 'Frequency / week',
'End Hub',
 'Arrival at Hub',
 'HDE (empties delivery at supplier)',
 'Arrival at Plant',
 'Unloading Time (Demand)',
 'Arrival at plant DHAS (PLE)',
 'Unloading at dock DHRQ (PLE) / HDE',
 'DEL Mon',
 'DEL Tue',
 'DEL Wed',
 'DEL Thu',
 'DEL Fri',
 'DEL Sat',
 'DEL Sun',
 'Total Transit time (max in days) / excluding blocked driving days',
 'Trailer Yard\ninformation',
 'Dangerous Goods information',
 'Carrier',
 'Transport Mode',
 'Means of Transportation',
 'km / tour']]

carriers = dff['Carrier'].str[:10].unique()

carriers_data = {}

i = 1
total = len(carriers)

print(f'{total} files to create..')



for carrier in carriers:
    carriers_data[carrier] = pd.DataFrame()

for carrier in carriers:
    carriers_data[carrier] = dff[dff['Carrier'].str.contains(f'{carrier}')]
    carriers_data[carrier].to_excel(os.getcwd()+f'//files/{carrier}_{TODAY}.xlsx')
    
    MailSubject= f"Forwarder {carrier}"
    MailInput="""
    #html code here
    """
    MailAdress=f"{carrier}@gmail.com"
    MailAttachment = f'//files/{carrier}_{TODAY}.xlsx'
    filename = f'{carrier}_{TODAY}'
    Emailer(MailInput, MailSubject, MailAdress, MailAttachment, filename)
    print(f'done {i} out of {total}')
    i += 1
    
end = time.time()
print(f'finished in {round(end-start,2)} seconds')




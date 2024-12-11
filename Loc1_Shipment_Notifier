import os
import openpyxl
import smtplib
from dotenv import load_dotenv, dotenv_values
from sys import argv

import pandas as pd

load_dotenv()

#set variables
EMAIL_ADDRESS = os.getenv('EMAIL_USER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASS')

#make data frame from csv file
data = pd.read_excel (r'ShipmentInquiryLoc1.xls')

#dropping duplicate values
data.drop_duplicates(subset = ["Order Number>>"],keep = 'first', inplace = True )

df=data

#New File
df.to_csv('CompanyEOD1.csv')

#Variables to be used through e-mail subject & body
Carrier = pd.read_csv('CompanyompanyEOD1.csv', usecols=['Ship-Via Code Description'])
Tracking = pd.read_csv('Company.csv', usecols=['Shipment Tracking Number'])
Customer = pd.read_csv('CompanyEOD1.csv', usecols=['Ship-To Name'])
PO = pd.read_csv('CompanyEOD1.csv', usecols=['Purchase Order Number'])
ORD = pd.read_csv('CompanyEOD1.csv', usecols=['Order Number>>'])


import chardet
rawdata=open("EMAIL_LIST.csv","rb").read()
result = chardet.detect(rawdata)

EmailName = pd.read_csv(r'EMAIL_LIST.csv', encoding='UTF-8', usecols=[0])
EmailAddress = pd.read_csv(r'EMAIL_LIST.csv', usecols=[1])

num_lines = df.shape[0]

num = 0

#Sending E-Mail
while num < num_lines:
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()

        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        CurrentCustomer = Customer.at[Customer.index[num],'Ship-To Name']
        CurrentPO = PO.at[PO.index[num],'Purchase Order Number']
        CurrentORD = ORD.at[ORD.index[num],'Order Number>>']
        CurrentCarr = Carrier.at[Carrier.index[num],'Ship-Via Code Description']
        CurrentTrack = Tracking.at[Tracking.index[num],'Shipment Tracking Number']
        CurrentTrack = str(CurrentTrack)

        #Find Current E-mail User Location & Address
        search_criteria = {'Customer_Name' : [CurrentCustomer]}
        #If Store not found, this will send back to Receiver - E-mail customer manually and add the best e-mail to list
        try:
            CurrentReceiverName = EmailName[EmailName.isin(search_criteria).all(axis=1)].index[0]
        except:
            CurrentReceiverName = 0

        CurrentReceiverAddress = EmailAddress.at[EmailAddress.index[CurrentReceiverName],'Customer_Email']

        if CurrentReceiverAddress == '':
            CurrentReceiverAddress = '[company]@[company].com'

        if CurrentCarr == 'BESTWAY PREPAID' or CurrentCarr == 'PPD Wrap and Tag Seperate Group Ship' or CurrentCarr == 'PPD Wrap and Tag Seperate - Group Ship' or CurrentCarr == 'INSERT CUSTOMER DETAIL' or CurrentCarr == 'PREPAID & CHARGE':
            CurrentCarr = ''
        else:
            CurrentCarr = 'with: ' + CurrentCarr

        if CurrentTrack == 'nan':
            CurrentTrack = ''
        else:
            CurrentTrack = 'tracking #' + CurrentTrack

        subject = f'Your Order PO# {CurrentPO} / {CurrentORD} has shipped'
        body = f'Good morning, \n\n{CurrentORD} / PO# {CurrentPO} has shipped out {CurrentCarr} {CurrentTrack}\n\nThank you,\n-[company]' #if tracking info - provide

        msg = f'Subject: {subject}\n\n{body}'

        try:
            smtp.sendmail(EMAIL_ADDRESS, CurrentReceiverAddress, msg)
        except:
            CurrentReceiverAddress = '[company]@[company].com'

        num += 1

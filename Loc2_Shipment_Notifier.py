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
data = pd.read_csv (r'[company]EOD.csv')

#dropping duplicate values
data.drop_duplicates(subset = ["Order_No"],keep = 'first', inplace = True )

df=data

#New File
df.to_csv('[company]EOD2.csv')

#Variables to be used through e-mail subject & body
Carrier = pd.read_csv('[company]EOD2.csv', usecols=[16])
Tracking = pd.read_csv('[company]EOD2.csv', usecols=[15])
Customer = pd.read_csv('[company]EOD2.csv', usecols=[7])
PO = pd.read_csv('[company]EOD2.csv', usecols=[3])
ORD = pd.read_csv('[company]EOD2.csv', usecols=[1])

import chardet
rawdata=open("EMAIL_LIST.csv","rb").read()
result = chardet.detect(rawdata)

EmailName = pd.read_csv(r'EMAIL_LIST.csv', encoding='UTF-8', usecols=[0])
EmailAddress = pd.read_csv(r'EMAIL_LIST.csv', encoding='UTF-8', usecols=[1])

#how long to loop for (While Statement - later in code)
num_lines = df.shape[0]

#print(Customer.at[Customer.index[1],'Customer_Name'])
#print(PO)

#import Customer_Name and associated Email
#while - rows - do, +1 (if need at end of loop)
#if Customer_Name not found - Send e-mail to sales@[company].com - alert to update

num = 0

#CurrEmailName =

while num < num_lines:
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()

        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        CurrentCustomer = Customer.at[Customer.index[num],'Customer_Name']
        CurrentPO = PO.at[PO.index[num],'CustomerPO']
        CurrentORD = ORD.at[ORD.index[num],'Order_No']
        CurrentCarr = Carrier.at[Carrier.index[num],'Carrier_Code']
        CurrentTrack = Tracking.at[Tracking.index[num],'Tracking_No']
        CurrentTrack = str(CurrentTrack)

        #Find Current E-mail User Location & Address
        search_criteria = {'Customer_Name' : [CurrentCustomer]}
        #If Store not found, this will send back to Receiver - E-mail customer manually and add the best e-mail to list
        try:
            CurrentReceiverName = EmailName[EmailName.isin(search_criteria).all(axis=1)].index[0]
        except:
            CurrentReceiverName = 0

        CurrentReceiverAddress = EmailAddress.at[EmailAddress.index[CurrentReceiverName],'Customer_Email']
        if CurrentReceiverAddress == '' or CurrentReceiverAddress == 'nan':
            CurrentReceiverAddress = 'sales@[company].com'

        #Deletable
        #if CurrentReceiverAddress == 'nan':
        #    CurrentReceiverAddress = 'sales@[company].com'


        if CurrentTrack == 'nan':
            CurrentTrack = ''
        else:
            CurrentTrack = 'tracking #' + CurrentTrack

        subject = f'Your Order PO# {CurrentPO} / {CurrentORD} has shipped'
        body = f'Good morning, \n\n{CurrentORD} / PO# {CurrentPO} has shipped out with {CurrentCarr} {CurrentTrack}\n\nThank you,\n-[company]' #if tracking info - provide

        msg = f'Subject: {subject}\n\n{body}'
        ##Issue
        smtp.sendmail(EMAIL_ADDRESS, CurrentReceiverAddress, msg)

        num += 1

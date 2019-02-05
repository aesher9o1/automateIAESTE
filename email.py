# -*- coding: utf-8 -*-
"""
Created on Tue Feb  5 11:25:29 2019

@author: aesher9o1
"""

import smtplib
import pandas as pd

database = pd.read_excel('test.xlsx')

for i, j in database.iterrows():
    #email
    TO = str(j[2])
    SUBJECT = 'Welcome to IAESTE family'
    TEXT = 'Dear '+ str(j["Name"])+',\n\nWe welcome you aboard to the IAESTE family and hope your participation proves to be beneficial. We will be having our first General Body Meeting (GBM) soon.\n\nThis email is a final confirmation of your IAESTE membership.\nYour IAESTE Number is:'+ str(j["IAESTE No."])+'\n\nPlease make a note of it as the number will be the reference ID with respect to IAESTE henceforth.\n\nHappy Interning!! \n\nRegards,\n\nSarthak Sarbahi \nHead Administration IAESTE LC MUJ \n+91 - 9619937704'

# Gmail Sign In
    gmail_sender = 'your_email_here'
    gmail_passwd = 'your_password_here'

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(gmail_sender, gmail_passwd)
    
    BODY = '\r\n'.join(['To: %s' % TO,
                    'From: %s' % gmail_sender,
                    'Subject: %s' % SUBJECT,
                    '', TEXT])

    try:
        server.sendmail(gmail_sender, [TO], BODY)
        print ('email sent')
    except:
        print ('error sending mail')

    server.quit()

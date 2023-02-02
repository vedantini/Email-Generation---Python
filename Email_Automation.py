# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import smtplib, ssl

#create dataframe
df = pd.read_excel(r'C:\Users\STSC\OneDrive - horizon.csueastbay.edu\Desktop\Book1.xlsx')
# dropping nulls values
df = df.dropna()

port = 587  # For starttls
smtp_server = "smtp.gmail.com"
sender_email = "ved@gmail.com" #add your email here
password = "kdilf"             #add your google account key here

#read the dataframe
for i in df.index:
    name = df['Name'][i]
    flag = df['Would you like to know more about the library ?'][i]
    r_email = df['email'][i]
    if flag == 'Yes':
        message = "Hello \n\n" + name + " \n We are reaching out to you because you have shown interest in CalFresh"
        + " \n \n Please view attachment for further information or click on below link to apply to CalFresh"
        receiver_email = r_email
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_server, port) as server:
            server.ehlo()  # Can be omitted
            server.starttls(context=context)
            server.ehlo()  # Can be omitted
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message)
        

# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 17:54:42 2023

@author: UDAY SANKAR
"""

import win32com.client as win32
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://www.googleapis.com/auth/spreadsheets' , "https://www.googleapis.com/auth/drive.file" , "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\pc\\Desktop\\codes\\sceret_key.json",scope)
client = gspread.authorize(creds)
sheet = client.open('Email_List').sheet1

# Load the user name and password
#user, password = "udaysankar.mukherjee2021@iem.edu.in", ""

# URL for Outlook connection
outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # 6 is the index of the Inbox folder

# Define a filter to get emails with a specific subject and that are unread
subject_filter = "[Subject]='New_Registration' AND [UnRead]=True"
items = inbox.Items.Restrict(subject_filter)



# Iterate through the messages and extract the subject and body
for i, item in enumerate(items):
    sender = item.SenderEmailAddress
    subject = item.Subject
    body = item.Body
    datastr = body.split(",")
    Name=datastr[0]
    Address=datastr[1]
    NHSNO=datastr[2]
    PHNO=datastr[3]
    DOB=datastr[4]
    sex=datastr[5]
    print(Name,Address,NHSNO,PHNO,DOB,sex)
    sheet.insert_row([Name,Address,NHSNO,PHNO,DOB,sex],3)
    
    # Mark the fetched emails as read
    item.UnRead = False

# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 17:54:42 2023

@author: UDAY SANKAR
"""

# Importing libraries
import imaplib
import email
from openpyxl import load_workbook
import pandas as pd

#Load the user name and passwd 
user, password = "udaysankar2003@gmail.com",""

#URL for IMAP connection
imap_url = 'imap.gmail.com'

# Connection with GMAIL using SSL
my_mail = imaplib.IMAP4_SSL(imap_url)

# Log in using your credentials
my_mail.login(user, password)

# Select the Inbox to fetch unread messages
my_mail.select('Inbox')

#Define Key and Value for email search
key = "SUBJECT"
value = "Registration"
_, data = my_mail.search(None, 'UNSEEN', key, value)  #Search for unread emails with specific key and value

mail_id_list = data[0].split()  #IDs of all unread emails that we want to fetch 

msgs = [] # empty list to capture all messages
#Iterate through messages and extract data into the msgs list
for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)') #RFC822 returns whole message (BODY fetches just body)
    msgs.append(data)


for msg in msgs[::-1]:
    for response_part in msg:
        if type(response_part) is tuple:
            my_msg=email.message_from_bytes((response_part[1]))

            print("_________________________________________")
            print ("subj:", my_msg['subject'])
            print ("from:", my_msg['from'])
            print ("body:")
            for part in my_msg.walk():  
                #print(part.get_content_type())
                if part.get_content_type() == 'text/plain':
                    print (part.get_payload())
                    
# Mark the fetched emails as read
#for num in mail_id_list:
#    my_mail.store(num, '+FLAGS', '\\Seen')

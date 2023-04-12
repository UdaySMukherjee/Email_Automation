# -*- coding: utf-8 -*-
"""
Created on Tue Apr 11 14:44:34 2023

@author: UDAY SANKAR
"""
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pywinauto.application import Application
import time
import pyautogui as auto
import win32com.client as win32

scope = ['https://www.googleapis.com/auth/spreadsheets' , "https://www.googleapis.com/auth/drive.file" , "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\UDAY SANKAR\\Desktop\\codes\\secret_key.json",scope)
client = gspread.authorize(creds)
sheet = client.open('Email_List').sheet1

# Get all the values in the first column
column_values = sheet.col_values(1)

# Count the number of non-empty rows
num_rows = len([value for value in column_values if value])

for i in range(2,num_rows+1):
    if sheet.cell(i,7).value == 'unsent':
        name = sheet.cell(i,1).value
        address = sheet.cell(i,2).value
        nhs_no = sheet.cell(i,3).value
        ph_no = sheet.cell(i,4).value
        dob = sheet.cell(i,5).value
        sex = sheet.cell(i,6).value
         
        app = Application(backend='uia').start("C:\\Users\\UDAY SANKAR\\Desktop\\codes\\dist\\Form.exe")
        app = Application(backend='uia').connect(title='Patient Information Form',timeout=5)
        #time.sleep(5)
        #app.PatientInformationForm.print_control_identifiers()
    
        Maximize = app.PatientInformationForm.child_window(title="System", control_type="MenuItem").wrapper_object()
        Maximize.click_input()

        x2,y2=auto.locateCenterOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\name.png",confidence=0.9)
        auto.moveTo(x2,y2,1)
        auto.click()
        time.sleep(1)
        auto.write(name)
        
        x3,y3=auto.locateCenterOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\address.png",confidence=0.9)
        auto.moveTo(x3,y3,1)
        auto.click()
        time.sleep(1)
        auto.write(address)
        
        x4,y4=auto.locateCenterOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\nhs.png",confidence=0.9)
        auto.moveTo(x4,y4,1)
        auto.click()
        time.sleep(1)
        auto.write(nhs_no)
        
        x5,y5=auto.locateCenterOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\ph.png",confidence=0.9)
        auto.moveTo(x5,y5,1)
        auto.click()
        time.sleep(1)
        auto.write(ph_no)
        
        x6,y6=auto.locateCenterOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\dob.png",confidence=0.9)
        auto.moveTo(x6,y6,1)
        auto.click()
        time.sleep(1)
        auto.write(dob)
        
        x7,y7=auto.locateCenterOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\dropdown.png",confidence=0.9)
        auto.moveTo(x7,y7,1)
        auto.click()
        time.sleep(1)
        
        if sex == ' Male':
            x71,y71, width, height=auto.locateOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\male.png",confidence=0.9)
            auto.moveTo(x71,y71,1)
            auto.click()
            time.sleep(1)
        if sex == ' Female':
            x72, y72, width, height = auto.locateOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\female.png", confidence=0.9)
            auto.moveTo(x72,y72,1)
            auto.click()
            time.sleep(1)
        if sex == ' Trans':
            x73,y73, width, height=auto.locateOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\trans.png",confidence=0.9)
            auto.moveTo(x73,y73,1)
            auto.click()
            time.sleep(1)
            
            
        x8,y8=auto.locateCenterOnScreen(r"C:\Users\UDAY SANKAR\Desktop\codes\submit.png",confidence=0.9)
        auto.moveTo(x8,y8,1)
        auto.click()
        time.sleep(1)

        Maximize = app.PatientInformationForm.child_window(title="Close", control_type="Button").wrapper_object()
        Maximize.click_input()

        olApp = win32.Dispatch('Outlook.Application')
        OINS = olApp.GetNameSpace('MAPI')

        mailItem = olApp.CreateItem(0)

        mailItem.Subject = 'Registration Successfully'
        mailItem.BodyFormat = 1
        name = sheet.cell(i,1).value
        mailItem.Body = "Dear "+name+",\n\n Congratulations!! \n Thank you so much for registering our service. \n\n\n Thanks and Regards, \n NHS Team"
        mailItem.Sender = 'udaysankar.mukherjee2021@iem.edu.in'
        mailItem.To = sheet.cell(i,8).value

        mailItem.Display()
        mailItem.Save()
        mailItem.Send()
        sheet.update_cell(i,7,"sent")
    

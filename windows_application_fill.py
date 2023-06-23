# Importing the required libraries
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pywinauto.application import Application
import time
import pyautogui as auto
import win32com.client as win32

# Define the scope and credentials for accessing the Google Sheets API
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\UDAY SANKAR\\Desktop\\codes\\secret_key.json", scope)
client = gspread.authorize(creds)

# Open the 'Email_List' spreadsheet and select the first sheet
sheet = client.open('Email_List').sheet1

# Get all the values in the first column
column_values = sheet.col_values(1)

# Count the number of non-empty rows
num_rows = len([value for value in column_values if value])

# Loop through each row starting from the second row
for i in range(2, num_rows + 1):
    # Check if the value in column 7 (status) is 'unsent'
    if sheet.cell(i, 7).value == 'unsent':
        # Retrieve data from specific columns in the current row
        name = sheet.cell(i, 1).value
        address = sheet.cell(i, 2).value
        nhs_no = sheet.cell(i, 3).value
        ph_no = sheet.cell(i, 4).value
        dob = sheet.cell(i, 5).value
        sex = sheet.cell(i, 6).value
        
        # Start the Windows Form application
        app = Application(backend='uia').start("Windows Form\\Form.exe")
        app = Application(backend='uia').connect(title='Patient Information Form', timeout=5)
        
        # Maximize the application window
        Maximize = app.PatientInformationForm.child_window(title="System", control_type="MenuItem").wrapper_object()
        Maximize.click_input()
        
        # Fill in the 'name' field
        x2, y2 = auto.locateCenterOnScreen(r"images\name.png", confidence=0.9)
        auto.moveTo(x2, y2, 1)
        auto.click()
        time.sleep(1)
        auto.write(name)
        
        # Fill in the 'address' field
        x3, y3 = auto.locateCenterOnScreen(r"images\address.png", confidence=0.9)
        auto.moveTo(x3, y3, 1)
        auto.click()
        time.sleep(1)
        auto.write(address)
        
        # Fill in the 'NHS number' field
        x4, y4 = auto.locateCenterOnScreen(r"immages\nhs.png", confidence=0.9)
        auto.moveTo(x4, y4, 1)
        auto.click()
        time.sleep(1)
        auto.write(nhs_no)
        
        # Fill in the 'phone number' field
        x5, y5 = auto.locateCenterOnScreen(r"images\ph.png", confidence=0.9)
        auto.moveTo(x5, y5, 1)
        auto.click()
        time.sleep(1)
        auto.write(ph_no)
        
        # Fill in the 'date of birth' field
        x6, y6 = auto.locateCenterOnScreen(r"images\dob.png", confidence=0.9)
        auto.moveTo(x6, y6, 1)
        auto.click()
        time.sleep(1)
        auto.write(dob)
        
        # Click on the dropdown menu for selecting sex
        x7, y7 = auto.locateCenterOnScreen(r"images\dropdown.png", confidence=0.9)
        auto.moveTo(x7, y7, 1)
        auto.click()
        time.sleep(1)
        
        # Select the appropriate option based on the 'sex' value
        if sex == ' Male':
            x71, y71, width, height = auto.locateOnScreen(r"images\male.png", confidence=0.9)
            auto.moveTo(x71, y71, 1)
            auto.click()
            time.sleep(1)
        if sex == ' Female':
            x72, y72, width, height = auto.locateOnScreen(r"images\female.png", confidence=0.9)
            auto.moveTo(x72, y72, 1)
            auto.click()
            time.sleep(1)
        if sex == ' Trans':
            x73, y73, width, height = auto.locateOnScreen(r"images\trans.png", confidence=0.9)
            auto.moveTo(x73, y73, 1)
            auto.click()
            time.sleep(1)
            
        # Click on the submit button
        x8, y8 = auto.locateCenterOnScreen(r"images\submit.png", confidence=0.9)
        auto.moveTo(x8, y8, 1)
        auto.click()
        time.sleep(1)
        
        # Close the application window
        Maximize = app.PatientInformationForm.child_window(title="Close", control_type="Button").wrapper_object()
        Maximize.click_input()
        
        # Create and send an email using Outlook
        olApp = win32.Dispatch('Outlook.Application')
        OINS = olApp.GetNameSpace('MAPI')
        
        mailItem = olApp.CreateItem(0)
        
        mailItem.Subject = 'Registration Successfully'
        mailItem.BodyFormat = 1
        name = sheet.cell(i, 1).value
        mailItem.Body = "Dear " + name + ",\n\n Congratulations!! \n Thank you so much for registering our service. \n\n\n Thanks and Regards, \n NHS Team"
        mailItem.Sender = 'udaysankar.mukherjee2021@iem.edu.in'
        mailItem.To = sheet.cell(i, 8).value
        
        # Display the email in Outlook for review
        mailItem.Display()
        
        # Save and send the email
        mailItem.Save()
        mailItem.Send()
        
        # Update the status column to 'sent'
        sheet.update_cell(i, 7, "sent")

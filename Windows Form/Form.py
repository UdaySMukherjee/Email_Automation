# -*- coding: utf-8 -*-
"""
Created on Thu Mar 30 11:38:03 2023

@author: UDAY SANKAR
"""

# Import tkinter library
import tkinter as tk
from tkinter import ttk

# Create the main window
root = tk.Tk()
root.title("Patient Information Form")

# Create the input labels and hint text
tk.Label(root, text="Name: ").grid(row=5, column=5, pady=25, padx=(200, 0))
name_entry = tk.Entry(root)
name_entry.insert(0, "Enter your name...")  # Set the hint text for name entry box
name_entry.bind("<FocusIn>", lambda args: name_entry.delete('0', 'end'))  # Clear the hint text on focus
name_entry.grid(row=5, column=8, pady=25, padx=(0, 200))

tk.Label(root, text="Address: ").grid(row=10, column=5, pady=25, padx=(200, 0))
address_entry = tk.Entry(root)
address_entry.insert(0, "Enter your address...")  # Set the hint text for address entry box
address_entry.bind("<FocusIn>", lambda args: address_entry.delete('0', 'end'))  # Clear the hint text on focus
address_entry.grid(row=10, column=8, pady=25, padx=(0, 200))

tk.Label(root, text="NHS Number: ").grid(row=15, column=5, pady=25, padx=(200, 0))
nhs_entry = tk.Entry(root)
nhs_entry.insert(0, "Enter your NHS number...")  # Set the hint text for NHS entry box
nhs_entry.bind("<FocusIn>", lambda args: nhs_entry.delete('0', 'end'))  # Clear the hint text on focus
nhs_entry.grid(row=15, column=8, pady=25, padx=(0, 200))

tk.Label(root, text="Ph Number: ").grid(row=20, column=5, pady=25, padx=(200, 0))
ph_entry = tk.Entry(root)
ph_entry.insert(0, "Enter your phone number...")  # Set the hint text for phone entry box
ph_entry.bind("<FocusIn>", lambda args: ph_entry.delete('0', 'end'))  # Clear the hint text on focus
ph_entry.grid(row=20, column=8, pady=25, padx=(0, 200))

tk.Label(root, text="DOB: ").grid(row=25, column=5, pady=25, padx=(200, 0))
dob_entry = tk.Entry(root)
dob_entry.insert(0, "Enter your date of birth...")  # Set the hint text for DOB entry box
dob_entry.bind("<FocusIn>", lambda args: dob_entry.delete('0', 'end'))  # Clear the hint text on focus
dob_entry.grid(row=25, column=8, pady=25, padx=(0, 200))

tk.Label(root, text="Sex: ").grid(row=30, column=5, pady=25, padx=(200, 0))
sex_entry = ttk.Combobox(values=[ "Male", "Female", "Trans"])  # Create the sex selection combobox
sex_entry.current(0)  # Set the default value to "Male"
sex_entry.grid(row=30, column=8, pady=25, padx=(0, 200))

# Create a button to submit the form
submit_button = tk.Button(root, text="Submit")
submit_button.grid(row=40, column=5, pady=25, padx=(350, 0))

# Start the main loop of the program
root.mainloop()

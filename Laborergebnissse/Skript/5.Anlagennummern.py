# -*- coding: utf-8 -*-
"""
Created on Tue Jun 20 10:05:33 2023

@author: zeh
"""

import openpyxl

# Load the workbook
workbook = openpyxl.load_workbook( 'C:/Users/zeh/Desktop/Laborergebnissse/AnlagenNummer.xlsx')

# Get the list of sheet names
sheet_names = workbook.sheetnames

# Counter for the numeric variable
n = 1

# Iterate through each sheet name
for sheet_name in sheet_names:
    if sheet_name.startswith("Anlage_"):
        # Generate the new sheet name
        new_sheet_name = f"Anlage_4.{n}"
        
        # Rename the sheet
        workbook[sheet_name].title = new_sheet_name
        
        # Get the renamed sheet
        renamed_sheet = workbook[new_sheet_name]
        
        # Update the name in cell K1
        name_with_space = new_sheet_name.replace("_", " ")
        renamed_sheet['K1'].value = name_with_space
        
        # Increment the numeric variable
        n += 1

# Save the modified workbook
workbook.save('C:/Users/zeh/Desktop/Laborergebnissse/AnlagenNummer.xlsx')


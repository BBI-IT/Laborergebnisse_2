# -*- coding: utf-8 -*-
"""
Created on Wed Jun 28 16:51:39 2023

@author: zeh
"""
import os

# Set the path to the directory containing the scripts
scripts_dir = r'C:\Users\zeh\Desktop\Laborergebnissse'

# Define the order and names of the scripts
script_order = [
    '1.Consolidate_Files.py',
    '2.Formating_Cells.py',
    '3.AddingFrame.py',
    '4.Transfered.py',
    '5.EraseFreeRows.py',
    '6.Anlagennummern.py',
    '7.RenameBW.py',
    '8.RenameBS.py',
    '9.SortinginGeneralNoInput.py',
    '10.TransferBackEaraaseRedundantData.py',
    '11.Formating_final.py',
]

# Execute each script in the specified order
for script_name in script_order:
    script_path = os.path.join(scripts_dir, script_name)
    with open(script_path, 'r') as script_file:
        script_code = script_file.read()
    exec(script_code)

print("All scripts executed successfully.")




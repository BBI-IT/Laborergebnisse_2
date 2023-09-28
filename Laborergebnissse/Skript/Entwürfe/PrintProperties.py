# -*- coding: utf-8 -*-
"""
Created on Tue Jul  4 11:18:03 2023

@author: zeh
"""

import win32com.client as win32

# Pfad zur Excel-Datei
source_file = r'C:\Users\zeh\Desktop\Laborergebnissse\TestTransfer.xlsx'
destination_file = r'C:\Users\zeh\Desktop\Laborergebnissse\TestPrint.xlsx'

# Excel-Anwendung öffnen
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
workbook = excel.Workbooks.Open(source_file)

# Druckbereich des ersten Arbeitsblatts ermitteln
first_sheet = workbook.Sheets(1)
print_area = first_sheet.PageSetup.PrintArea

# Iteriere über die Arbeitsblätter im Workbook
for sheet in workbook.Sheets:
    if sheet.Name.startswith('Anlage_'):
        sheet.PageSetup.PrintArea = print_area

        # Setze den Druckbereich auf eine Seite
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = 1

# Speichern und schließen der Excel-Datei
workbook.SaveAs(destination_file)
workbook.Close()
excel.Quit()

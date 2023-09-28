import openpyxl

# Pfad zur Excel-Datei
source_file = r'C:\Users\zeh\Desktop\Laborergebnissse\TestTransfer.xlsx'
destination_file = r'C:\Users\zeh\Desktop\Laborergebnissse\TestPrint.xlsx'

# Lade die Excel-Datei
source_workbook = openpyxl.load_workbook(source_file)

# Iterate over the sheets in the source workbook
for sheet_name in source_workbook.sheetnames:
    # Überprüfe, ob das Tabellenblatt mit "2Anlage_" beginnt
    if sheet_name.startswith('Anlage_'):
        # Get the sheet from the workbook
        worksheet = source_workbook[sheet_name]

        # Definiere den Druckbereich (A1:L27)
        print_area = 'A1:L27'

        # Setze den Druckbereich im Arbeitsblatt
        worksheet.print_area = print_area

        # Setze das Seitenlayout auf "Fit to Page" (eine Seite)
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = False

# Speichere die Änderungen in der Excel-Datei
source_workbook.save(destination_file)

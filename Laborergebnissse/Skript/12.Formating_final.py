import openpyxl

# Path of the consolidated file
consolidated_file_path = 'C:/Users/zeh/Desktop/Laborergebnissse/01.LaborergebnisseFinal.xlsx'

# Open the consolidated workbook
consolidated_workbook = openpyxl.load_workbook(consolidated_file_path)

# Iterate over the sheets in the consolidated workbook
for sheet in consolidated_workbook.sheetnames:
    # Get the sheet from the workbook
    consolidated_sheet = consolidated_workbook[sheet]

    # Set individual column widths
    column_widths = {'A': 0.75, 'B': 26.14, 'C': 10.71, 'D': 12.14, 'E': 12.14, 'F': 12.14, 'G': 12.14, 'H': 12.14, 'I': 12.14, 'J': 12.14, 'K': 12.14, 'L': 1.71}
    for col, width in column_widths.items():
        consolidated_sheet.column_dimensions[col].width = width

# Save the changes to the consolidated file
consolidated_workbook.save(consolidated_file_path)

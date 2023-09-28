import openpyxl

# Load the Transferred workbook
transferred_workbook = openpyxl.load_workbook('C:/Users/zeh/Desktop/Laborergebnissse/Transferred.xlsx')

# Select the "Transferred" sheet
transferred_sheet = transferred_workbook['Transferred']

# Find the last row with data in the "Transferred" sheet
last_row = transferred_sheet.max_row

# Iterate through the rows in reverse order (from the last row to the first row)
for row in range(last_row, 0, -1):
    is_empty = True

    # Check if all cells in the row are empty
    for cell in transferred_sheet[row]:
        if cell.value is not None:
            is_empty = False
            break

    # If the row is empty, delete it
    if is_empty:
        transferred_sheet.delete_rows(row)

# Save the updated Transferred workbook
transferred_workbook.save('C:/Users/zeh/Desktop/Laborergebnissse/AnlagenNummer.xlsx')

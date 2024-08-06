import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import os

# Updated file path
file_path = 'C:/Users/Jiraphan.Detchokul/Documents/SAR Parts tracking.xlsx'

# Read the Excel file into a DataFrame
df = pd.read_excel(file_path, engine='openpyxl')

# Convert columns 7 and 8 to datetime if they are not already
df.iloc[:, 6] = pd.to_datetime(df.iloc[:, 6], errors='coerce').dt.date
df.iloc[:, 7] = pd.to_datetime(df.iloc[:, 7], errors='coerce').dt.date

# Filter rows where column 6 is blank
condition1 = df.iloc[:, 5].isna()

# Additional conditions: column 12 is blank or "NO PO" or the date in column 8 exceeds the date in column 7
condition2 = (df.iloc[:, 11].isna() | df.iloc[:, 11].str.strip().eq('NO PO')) | (df.iloc[:, 7] > df.iloc[:, 6])

# Combine both conditions
filtered_df = df[condition1 & condition2]

# Select the specified columns by their indices (including column 8)
selected_columns = filtered_df.iloc[:, [0, 1, 2, 3, 4, 5, 6, 7, 10, 11, 12]]

# Define the path for the new Excel file
output_file_path = os.path.join(os.path.dirname(file_path), 'filtered_rows_selected_columns.xlsx')

# Write the filtered and selected columns to the new Excel file
selected_columns.to_excel(output_file_path, index=False, engine='openpyxl')

# Load the workbook and select the active worksheet
wb = openpyxl.load_workbook(output_file_path)
ws = wb.active

# Define the fill for highlighting
highlight_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

# Format columns 7 and 8 to only show the date
date_format = 'dd-mm-yyyy'
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=7, max_col=8):
    for cell in row:
        cell.number_format = date_format

# Loop through the rows and highlight rows where the date in column 8 is later than the date in column 7
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    col_7_date = row[6].value
    col_8_date = row[7].value
    if col_7_date and col_8_date and col_8_date > col_7_date:
        for cell in row:
            cell.fill = highlight_fill

# Save the workbook
wb.save(output_file_path)

print(f'Filtered rows with selected columns and highlighted rows have been written to {output_file_path}')

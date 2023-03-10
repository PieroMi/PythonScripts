from pandas import io
import openpyxl
from openpyxl import worksheet
from openpyxl.utils import get_column_letter

# This Script will get the main data set and clean it.
# Making deletion of certain rows with certain values to clean it

workbook = openpyxl.load_workbook('Client Transaction February 2023.xlsx')

sheet = workbook['Documento1']

sheet.unmerge_cells('A1:AH1')
sheet.unmerge_cells('A2:AH2')

sheet.delete_rows(1, 3)

skip_cols = ["H", "N", "W", "AA", "AB", "AD", "AK"]

for col_idx in range(sheet.max_column, 0, -1):
    col_letter = get_column_letter(col_idx)
    if col_letter not in skip_cols:
        sheet.delete_cols(col_idx)

names_to_drop = ['Consumidor Final']

for row in range(sheet.max_row, 1, -1):
    customer_name = sheet.cell(row=row, column=1).value
    discount_value = sheet.cell(row=row, column=6).value
    if customer_name in names_to_drop and discount_value <= 0:
        sheet.delete_rows(row)


workbook.save('Modified Client Transaction February 2023.xlsx')

# This Script will then be used with the GroupBy.py script.


from openpyxl import load_workbook
from openpyxl.styles import Border, Side

def mergeDate():
    for col in range (4, 38):
        sheet.merge_cells(start_row=9, start_column=col, end_row=11, end_column=col)

def newEmployee(n):
    firstRow = n*2 + 11
    for col in range (1, 4):
        sheet.merge_cells(start_row=firstRow, start_column=col, end_row=firstRow + 1, end_column=col)
    for col in range (36, 42):
        sheet.merge_cells(start_row=firstRow, start_column=col, end_row=firstRow + 1, end_column=col)
    sheet.merge_cells(start_row=firstRow, start_column=45, end_row=firstRow + 1, end_column=45)
    for col in range (1, 46):
        sheet.cell(row=firstRow, column=col).border = Border (top=thin, left=thin, right=thin, bottom=thin)
        sheet.cell(row=firstRow + 1, column=col).border = Border (top=thin, left=thin, right=thin, bottom=thin)
        

wb = load_workbook('c:/Temp/Табель.xlsx')
sheet = wb['tab']
thin = Side(border_style="thin", color="000000")
newEmployee(1)
newEmployee(2)
newEmployee(3)

    
wb.save('c:/Temp/Табель.xlsx')
print('The end.')



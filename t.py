from openpyxl import load_workbook, formula
from openpyxl.styles import Border, Side
import sqlite3
import calendar


"""
index
eventday
fullname
rank
startdate
stopdate
sickleave
vacation
"""

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
    sheet.cell( column = 19, row = firstRow, value = "=SUM(RC[-15]:RC[-1]")
    #sheet.cell( column = 36, row = firstRow, value = "=SUM(RC[-15]:RC[-1]")
    
    
    for col in range (1, 46):
        sheet.cell(row=firstRow, column=col).border = Border (top=thin, left=thin, right=thin, bottom=thin)
        sheet.cell(row=firstRow + 1, column=col).border = Border (top=thin, left=thin, right=thin, bottom=thin)
        
    for day in range (1, 15+1):
        col = day + 3
        dayOfWeek = calendar.weekday(year, month, day)
        sheet.cell( column = col, row = firstRow, value = 'в' if dayOfWeek == 5 or dayOfWeek == 6 else 8 )
        
    for day in range (16, calendar.monthrange(year, month)[1]+1):
        col = day + 4
        dayOfWeek = calendar.weekday(year, month, day)
        sheet.cell( column = col, row = firstRow, value = 'в' if dayOfWeek == 5 or dayOfWeek == 6 else 8 )
        
        

conn = sqlite3.connect('c:/temp/tabel.db3')
wb = load_workbook('c:/Temp/Табель.xlsx')

year = 2020
month = 7

sheet = wb['tab']
thin = Side(border_style="thin", color="000000")
newEmployee(1)
newEmployee(2)
newEmployee(3)
newEmployee(4)
newEmployee(5)
newEmployee(6)
    
wb.save('c:/Temp/Табель.xlsx')
conn.close()
print('The end.')



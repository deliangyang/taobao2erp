import openpyxl

filename = '/Users/ydl/Downloads/ExportOrderList5217472008.xlsx'
workbook = openpyxl.load_workbook(filename)
print(2)
workbook.security.workbookPassword = 'aFUAt0nH'
print(3)
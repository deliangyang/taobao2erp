from xlrd import *
import win32com.client
import csv
import sys

xlApp = win32com.client.Dispatch("Excel.Application")
filename, password = r"C:\Users\ydl\Desktop\encryptedtest.xlsx", 'encryptedtest'
xlwb = xlApp.Workbooks.Open(filename, False, True, None, Password=password)
print(xlwb.Sheets(1).Cells(1, 1))
print(xlwb.Sheets(1))
xlApp.ActiveWorkbook.SaveAs(r"C:\Users\ydl\Desktop\text233.csv", 62, "", "")
xlApp.Quit()
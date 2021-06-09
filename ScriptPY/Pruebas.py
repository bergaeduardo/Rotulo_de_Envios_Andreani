import win32com.client as win32
from selenium import webdriver
import time
import openpyxl
import math
import autoit

plantillaExcel = r'X:\Logistica\1. Logistica Interno\Historiales_de_Logistica\ImportacionAndreani.xlsx'

fname = plantillaExcel
#os.system ('taskkill / IM EXCEL.exe / F')
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(fname)
#wb.DisplayAlerts = False
wb.SaveAs(fname, FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
# excel.Application.Quit()
excel.DisplayAlerts = True